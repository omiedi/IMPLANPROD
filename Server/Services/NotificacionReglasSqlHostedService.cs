using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;

namespace IMPLANPROD.Server.Services
{
    public class NotificacionReglasSqlHostedService : BackgroundService
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly IConfiguration _configuration;
        private readonly ILogger<NotificacionReglasSqlHostedService> _logger;

        public NotificacionReglasSqlHostedService(
            IServiceProvider serviceProvider,
            IConfiguration configuration,
            ILogger<NotificacionReglasSqlHostedService> logger)
        {
            _serviceProvider = serviceProvider;
            _configuration = configuration;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("NotificacionReglasSqlHostedService iniciado");
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await ProcessOnceAsync(stoppingToken);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error en NotificacionReglasSqlHostedService");
                }

                await Task.Delay(TimeSpan.FromMinutes(1), stoppingToken);
            }
        }

        private async Task ProcessOnceAsync(CancellationToken stoppingToken)
        {
            using var scope = _serviceProvider.CreateScope();
            var context = scope.ServiceProvider.GetRequiredService<DataContext>();

            var now = DateTime.Now;

            _logger.LogDebug("Procesando reglas SQL (tick) {Now}", now);

            var reglas = await context.NotificacionReglasSql
                .Where(x => x.Activa)
                .ToListAsync(stoppingToken);

            _logger.LogDebug("Reglas activas encontradas: {Count}", reglas.Count);

            foreach (var regla in reglas)
            {
                if (stoppingToken.IsCancellationRequested)
                {
                    return;
                }

                if (regla.IntervaloMinutos <= 0)
                {
                    regla.IntervaloMinutos = 15;
                }

                var shouldRun = !regla.UltimaEjecucion.HasValue ||
                                regla.UltimaEjecucion.Value.AddMinutes(regla.IntervaloMinutos) <= now;

                if (!shouldRun)
                {
                    continue;
                }

                try
                {
                    _logger.LogInformation(
                        "Ejecutando regla SQL {ReglaId} '{Titulo}' (IntervaloMinutos={IntervaloMinutos}, MaxFilasDetalle={MaxFilasDetalle}, SqlConnectionName={SqlConnectionName})",
                        regla.Id,
                        regla.Titulo,
                        regla.IntervaloMinutos,
                        regla.MaxFilasDetalle,
                        regla.SqlConnectionName);

                    await EjecutarReglaAsync(context, regla, stoppingToken);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error al ejecutar regla SQL {ReglaId}", regla.Id);
                    regla.UltimaEjecucion = now;
                    await context.SaveChangesAsync(stoppingToken);
                }
            }
        }

        private async Task EjecutarReglaAsync(DataContext context, NotificacionReglaSql regla, CancellationToken stoppingToken)
        {
            var now = DateTime.Now;

            // Modo 0: Hash (comportamiento actual - cambios en resultados)
            if (regla.ModoSeguimiento == 0)
            {
                await EjecutarReglaModoHashAsync(context, regla, now, stoppingToken);
            }
            // Modo 1: Incremental (seguimiento de nuevos registros)
            else if (regla.ModoSeguimiento == 1)
            {
                await EjecutarReglaModoIncrementalAsync(context, regla, now, stoppingToken);
            }
        }

        private async Task EjecutarReglaModoHashAsync(DataContext context, NotificacionReglaSql regla, DateTime now, CancellationToken stoppingToken)
        {
            var maxRows = regla.MaxFilasDetalle <= 0 ? 200 : regla.MaxFilasDetalle;

            var sql = EnsureTopSelect(regla.SqlQuery, maxRows);

            var connectionString = ResolveConnectionString(regla.SqlConnectionName);

            _logger.LogDebug("Regla {ReglaId} (Modo Hash): SQL final a ejecutar: {Sql}", regla.Id, sql);

            var data = await ExecuteQueryAsync(connectionString, sql, stoppingToken);

            _logger.LogInformation("Regla {ReglaId} (Modo Hash): filas devueltas = {RowCount}", regla.Id, data.Rows.Count);

            regla.UltimaEjecucion = now;

            if (data.Rows.Count == 0)
            {
                regla.UltimoResultadoHash = null;
                await context.SaveChangesAsync(stoppingToken);
                _logger.LogInformation("Regla {ReglaId} (Modo Hash): sin resultados, no se crea notificación", regla.Id);
                return;
            }

            var html = BuildHtmlTable(data);
            var hash = ComputeSha256Hex(html);

            _logger.LogDebug("Regla {ReglaId} (Modo Hash): hash resultado = {Hash}", regla.Id, hash);

            if (!string.IsNullOrWhiteSpace(regla.UltimoResultadoHash) &&
                string.Equals(regla.UltimoResultadoHash, hash, StringComparison.OrdinalIgnoreCase))
            {
                await context.SaveChangesAsync(stoppingToken);
                _logger.LogInformation("Regla {ReglaId} (Modo Hash): resultado sin cambios (hash igual), no se crea notificación", regla.Id);
                return;
            }

            regla.UltimoResultadoHash = hash;

            var notificacion = new Notificacion
            {
                Titulo = regla.Titulo,
                Resumen = regla.Resumen,
                EsParaTodos = regla.EsParaTodos,
                CreadorUsuarioId = regla.CreadorUsuarioId,
                DetalleHtml = html,
                DetalleHash = hash,
                EsGeneradaPorRegla = true,
                ReglaSqlId = regla.Id
            };

            context.Notificaciones.Add(notificacion);
            await context.SaveChangesAsync(stoppingToken);

            // Si no es para todos, crear destinatarios
            if (!regla.EsParaTodos && !string.IsNullOrWhiteSpace(regla.DestinatariosIds))
            {
                var destinatarios = CrearDestinatarios(regla.DestinatariosIds, notificacion.Id);
                context.NotificacionDestinatarios.AddRange(destinatarios);
                await context.SaveChangesAsync(stoppingToken);
            }

            _logger.LogInformation("Regla {ReglaId} (Modo Hash): notificación creada (NotificacionId={NotificacionId})", regla.Id, notificacion.Id);
        }

        private async Task EjecutarReglaModoIncrementalAsync(DataContext context, NotificacionReglaSql regla, DateTime now, CancellationToken stoppingToken)
        {
            // Validar que CampoSeguimiento esté configurado
            if (string.IsNullOrWhiteSpace(regla.CampoSeguimiento))
            {
                _logger.LogWarning("Regla {ReglaId} (Modo Incremental): CampoSeguimiento no configurado, se omite ejecución", regla.Id);
                return;
            }

            var maxRows = regla.MaxFilasDetalle <= 0 ? 200 : regla.MaxFilasDetalle;

            // Reemplazar @UltimoValorProcesado por el valor actual (o 0 si es null)
            var ultimoValor = regla.UltimoValorProcesado ?? "0";
            var sql = regla.SqlQuery.Replace("@UltimoValorProcesado", ultimoValor, StringComparison.OrdinalIgnoreCase);

            sql = EnsureTopSelect(sql, maxRows);

            var connectionString = ResolveConnectionString(regla.SqlConnectionName);

            _logger.LogDebug("Regla {ReglaId} (Modo Incremental): SQL final a ejecutar: {Sql}", regla.Id, sql);
            _logger.LogDebug("Regla {ReglaId} (Modo Incremental): UltimoValorProcesado = {UltimoValor}", regla.Id, ultimoValor);

            var data = await ExecuteQueryAsync(connectionString, sql, stoppingToken);

            _logger.LogInformation("Regla {ReglaId} (Modo Incremental): filas devueltas = {RowCount}", regla.Id, data.Rows.Count);

            regla.UltimaEjecucion = now;

            if (data.Rows.Count == 0)
            {
                // No hay nuevos registros, no actualizamos UltimoValorProcesado
                await context.SaveChangesAsync(stoppingToken);
                _logger.LogInformation("Regla {ReglaId} (Modo Incremental): sin resultados, no se crea notificación ni se actualiza UltimoValorProcesado", regla.Id);
                return;
            }

            // Generar notificación con todos los registros nuevos
            var html = BuildHtmlTable(data);

            var notificacion = new Notificacion
            {
                Titulo = regla.Titulo,
                Resumen = regla.Resumen,
                EsParaTodos = regla.EsParaTodos,
                CreadorUsuarioId = regla.CreadorUsuarioId,
                DetalleHtml = html,
                DetalleHash = null, // No se usa hash en modo incremental
                EsGeneradaPorRegla = true,
                ReglaSqlId = regla.Id
            };

            context.Notificaciones.Add(notificacion);
            await context.SaveChangesAsync(stoppingToken);

            // Si no es para todos, crear destinatarios
            if (!regla.EsParaTodos && !string.IsNullOrWhiteSpace(regla.DestinatariosIds))
            {
                var destinatarios = CrearDestinatarios(regla.DestinatariosIds, notificacion.Id);
                context.NotificacionDestinatarios.AddRange(destinatarios);
                await context.SaveChangesAsync(stoppingToken);
            }

            _logger.LogInformation("Regla {ReglaId} (Modo Incremental): notificación creada (NotificacionId={NotificacionId})", regla.Id, notificacion.Id);

            // Obtener el mayor valor del CampoSeguimiento de los registros procesados
            var maxValor = ObtenerMaxValorCampo(data, regla.CampoSeguimiento);

            if (!string.IsNullOrWhiteSpace(maxValor))
            {
                regla.UltimoValorProcesado = maxValor;
                await context.SaveChangesAsync(stoppingToken);
                _logger.LogInformation("Regla {ReglaId} (Modo Incremental): UltimoValorProcesado actualizado a {MaxValor}", regla.Id, maxValor);
            }
            else
            {
                _logger.LogWarning("Regla {ReglaId} (Modo Incremental): no se pudo obtener MAX({CampoSeguimiento})", regla.Id, regla.CampoSeguimiento);
            }
        }

        private static string? ObtenerMaxValorCampo(DataTable data, string campoSeguimiento)
        {
            if (data.Columns.Contains(campoSeguimiento))
            {
                var maxValue = data.AsEnumerable()
                    .Select(row => row[campoSeguimiento])
                    .Where(value => value != DBNull.Value)
                    .OrderByDescending(value => value)
                    .FirstOrDefault();

                return maxValue?.ToString();
            }

            return null;
        }

        private string ResolveConnectionString(string? connectionName)
        {
            if (string.IsNullOrWhiteSpace(connectionName))
            {
                return _configuration.GetConnectionString("DefaultConnection") ?? throw new InvalidOperationException("DefaultConnection no configurada");
            }

            if (string.Equals(connectionName, "DefaultConnection", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(connectionName, "ConnectionStrings", StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogWarning(
                    "SqlConnectionName='{SqlConnectionName}' no es un nombre de ConnectionString válido. Usando DefaultConnection.",
                    connectionName);

                return _configuration.GetConnectionString("DefaultConnection") ?? throw new InvalidOperationException("DefaultConnection no configurada");
            }

            var cs = _configuration.GetConnectionString(connectionName);
            if (string.IsNullOrWhiteSpace(cs))
            {
                throw new InvalidOperationException($"ConnectionString '{connectionName}' no configurada");
            }

            return cs;
        }

        private static string EnsureTopSelect(string sqlQuery, int maxRows)
        {
            var trimmed = sqlQuery.Trim();
            if (!trimmed.StartsWith("select", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("SqlQuery debe ser SELECT");
            }

            if (trimmed.Contains("top", StringComparison.OrdinalIgnoreCase))
            {
                return trimmed;
            }

            var idx = IndexOfSelect(trimmed);
            if (idx < 0)
            {
                return trimmed;
            }

            return trimmed.Insert(idx + 6, $" TOP ({maxRows})");
        }

        private static int IndexOfSelect(string sql)
        {
            return sql.IndexOf("select", StringComparison.OrdinalIgnoreCase);
        }

        private static async Task<DataTable> ExecuteQueryAsync(string connectionString, string sql, CancellationToken stoppingToken)
        {
            using var con = new SqlConnection(connectionString);
            await con.OpenAsync(stoppingToken);

            using var cmd = con.CreateCommand();
            cmd.CommandText = sql;
            cmd.CommandType = CommandType.Text;
            cmd.CommandTimeout = 30;

            using var reader = await cmd.ExecuteReaderAsync(stoppingToken);
            var dt = new DataTable();
            dt.Load(reader);
            return dt;
        }

        private static string BuildHtmlTable(DataTable dt)
        {
            var sb = new StringBuilder();
            sb.Append("<table class=\"table table-sm table-striped\"> ");

            sb.Append("<thead><tr>");
            foreach (DataColumn col in dt.Columns)
            {
                sb.Append("<th>");
                sb.Append(HtmlEncoder.Default.Encode(col.ColumnName));
                sb.Append("</th>");
            }
            sb.Append("</tr></thead>");

            sb.Append("<tbody>");
            foreach (DataRow row in dt.Rows)
            {
                sb.Append("<tr>");
                foreach (DataColumn col in dt.Columns)
                {
                    sb.Append("<td>");
                    var value = row[col] == DBNull.Value ? string.Empty : row[col]?.ToString() ?? string.Empty;
                    sb.Append(HtmlEncoder.Default.Encode(value));
                    sb.Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</tbody>");

            sb.Append("</table>");
            return sb.ToString();
        }

        private static string ComputeSha256Hex(string input)
        {
            var bytes = Encoding.UTF8.GetBytes(input);
            var hash = SHA256.HashData(bytes);
            return Convert.ToHexString(hash);
        }

        private static List<NotificacionDestinatario> CrearDestinatarios(string destinatariosIds, int notificacionId)
        {
            var destinatarios = new List<NotificacionDestinatario>();

            if (string.IsNullOrWhiteSpace(destinatariosIds))
            {
                return destinatarios;
            }

            var ids = destinatariosIds.Split(',', StringSplitOptions.RemoveEmptyEntries);
            foreach (var idStr in ids)
            {
                if (int.TryParse(idStr.Trim(), out int usuarioId))
                {
                    destinatarios.Add(new NotificacionDestinatario
                    {
                        NotificacionId = notificacionId,
                        UsuarioId = usuarioId
                    });
                }
            }

            return destinatarios;
        }
    }
}
