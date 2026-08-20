using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs.Notificaciones;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    public class NotificacionReglasSqlService
    {
        private readonly DataContext _context;
        private readonly IUserContextService _userContext;

        public NotificacionReglasSqlService(DataContext context, IUserContextService userContext)
        {
            _context = context;
            _userContext = userContext;
        }

        private int GetCurrentUsuarioIdOrThrow()
        {
            var usuarioId = _userContext.GetCurrentUsuarioId();
            if (!usuarioId.HasValue || usuarioId.Value <= 0)
            {
                throw new UnauthorizedAccessException("Usuario no autenticado");
            }

            return usuarioId.Value;
        }

        public async Task<List<NotificacionReglaSqlItemDTO>> GetAllAsync(CancellationToken cancellationToken = default)
        {
            return await _context.NotificacionReglasSql
                .AsNoTracking()
                .OrderByDescending(x => x.Id)
                .Select(x => new NotificacionReglaSqlItemDTO
                {
                    Id = x.Id,
                    Titulo = x.Titulo,
                    Resumen = x.Resumen,
                    Activa = x.Activa,
                    MaxFilasDetalle = x.MaxFilasDetalle,
                    IntervaloMinutos = x.IntervaloMinutos,
                    UltimaEjecucion = x.UltimaEjecucion,
                    UltimoResultadoHash = x.UltimoResultadoHash,
                    SqlConnectionName = x.SqlConnectionName,
                    ModoSeguimiento = x.ModoSeguimiento,
                    CampoSeguimiento = x.CampoSeguimiento,
                    UltimoValorProcesado = x.UltimoValorProcesado,
                    EsParaTodos = x.EsParaTodos,
                    DestinatariosIds = x.DestinatariosIds
                })
                .ToListAsync(cancellationToken);
        }

        public async Task<NotificacionReglaSqlUpsertDTO?> GetByIdAsync(int id, CancellationToken cancellationToken = default)
        {
            return await _context.NotificacionReglasSql
                .AsNoTracking()
                .Where(x => x.Id == id)
                .Select(x => new NotificacionReglaSqlUpsertDTO
                {
                    Id = x.Id,
                    Titulo = x.Titulo,
                    Resumen = x.Resumen,
                    Activa = x.Activa,
                    SqlQuery = x.SqlQuery,
                    SqlConnectionName = x.SqlConnectionName,
                    MaxFilasDetalle = x.MaxFilasDetalle,
                    IntervaloMinutos = x.IntervaloMinutos,
                    ModoSeguimiento = x.ModoSeguimiento,
                    CampoSeguimiento = x.CampoSeguimiento,
                    EsParaTodos = x.EsParaTodos,
                    DestinatariosIds = x.DestinatariosIds
                })
                .FirstOrDefaultAsync(cancellationToken);
        }

        public async Task<int> CreateAsync(NotificacionReglaSqlUpsertDTO dto, CancellationToken cancellationToken = default)
        {
            var creadorUsuarioId = GetCurrentUsuarioIdOrThrow();

            ValidateSqlQuery(dto.SqlQuery);

            // Validación: si es modo incremental, debe tener CampoSeguimiento
            if (dto.ModoSeguimiento == 1 && string.IsNullOrWhiteSpace(dto.CampoSeguimiento))
            {
                throw new InvalidOperationException("En modo Incremental debe especificar el Campo de seguimiento (ej: NroPed, IdOrden)");
            }

            var sqlConnectionName = NormalizeSqlConnectionName(dto.SqlConnectionName);

            var entity = new NotificacionReglaSql
            {
                Titulo = dto.Titulo,
                Resumen = dto.Resumen,
                Activa = dto.Activa,
                SqlQuery = dto.SqlQuery,
                SqlConnectionName = sqlConnectionName,
                MaxFilasDetalle = dto.MaxFilasDetalle <= 0 ? 200 : dto.MaxFilasDetalle,
                IntervaloMinutos = dto.IntervaloMinutos <= 0 ? 15 : dto.IntervaloMinutos,
                CreadorUsuarioId = creadorUsuarioId,
                ModoSeguimiento = dto.ModoSeguimiento,
                CampoSeguimiento = dto.CampoSeguimiento,
                UltimoValorProcesado = null, // Se inicializa en null para nuevas reglas
                EsParaTodos = dto.EsParaTodos,
                DestinatariosIds = dto.DestinatariosIds
            };

            _context.NotificacionReglasSql.Add(entity);
            await _context.SaveChangesAsync(cancellationToken);
            return entity.Id;
        }

        public async Task UpdateAsync(int id, NotificacionReglaSqlUpsertDTO dto, CancellationToken cancellationToken = default)
        {
            GetCurrentUsuarioIdOrThrow();

            ValidateSqlQuery(dto.SqlQuery);

            // Validación: si es modo incremental, debe tener CampoSeguimiento
            if (dto.ModoSeguimiento == 1 && string.IsNullOrWhiteSpace(dto.CampoSeguimiento))
            {
                throw new InvalidOperationException("En modo Incremental debe especificar el Campo de seguimiento (ej: NroPed, IdOrden)");
            }

            var sqlConnectionName = NormalizeSqlConnectionName(dto.SqlConnectionName);

            var entity = await _context.NotificacionReglasSql
                .FirstOrDefaultAsync(x => x.Id == id, cancellationToken);

            if (entity == null)
            {
                throw new InvalidOperationException("Regla no encontrada");
            }

            entity.Titulo = dto.Titulo;
            entity.Resumen = dto.Resumen;
            entity.Activa = dto.Activa;
            entity.SqlQuery = dto.SqlQuery;
            entity.SqlConnectionName = sqlConnectionName;
            entity.MaxFilasDetalle = dto.MaxFilasDetalle <= 0 ? entity.MaxFilasDetalle : dto.MaxFilasDetalle;
            entity.IntervaloMinutos = dto.IntervaloMinutos <= 0 ? entity.IntervaloMinutos : dto.IntervaloMinutos;
            entity.ModoSeguimiento = dto.ModoSeguimiento;
            entity.CampoSeguimiento = dto.CampoSeguimiento;
            entity.EsParaTodos = dto.EsParaTodos;
            entity.DestinatariosIds = dto.DestinatariosIds;
            // Nota: No reseteamos UltimoValorProcesado al editar para mantener el estado de seguimiento

            await _context.SaveChangesAsync(cancellationToken);
        }

        private static string? NormalizeSqlConnectionName(string? sqlConnectionName)
        {
            if (string.IsNullOrWhiteSpace(sqlConnectionName))
            {
                return null;
            }

            if (string.Equals(sqlConnectionName, "ConnectionStrings", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(sqlConnectionName, "DefaultConnection", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return sqlConnectionName.Trim();
        }

        public async Task DeleteAsync(int id, CancellationToken cancellationToken = default)
        {
            GetCurrentUsuarioIdOrThrow();

            var entity = await _context.NotificacionReglasSql
                .FirstOrDefaultAsync(x => x.Id == id, cancellationToken);

            if (entity == null)
            {
                return;
            }

            _context.NotificacionReglasSql.Remove(entity);
            await _context.SaveChangesAsync(cancellationToken);
        }

        private static void ValidateSqlQuery(string sqlQuery)
        {
            if (string.IsNullOrWhiteSpace(sqlQuery))
            {
                throw new InvalidOperationException("SqlQuery es obligatorio");
            }

            var trimmed = sqlQuery.TrimStart();
            if (!trimmed.StartsWith("select", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("La regla SQL debe ser un SELECT");
            }
        }
    }
}
