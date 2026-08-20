using System.Data.Common;
using System.Text.RegularExpressions;
using System.Text.Json;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs.Impresiones;
using Microsoft.EntityFrameworkCore;
using Shared.Entities;

namespace IMPLANPROD.Server.Services
{
    public class RemitoPlantillaService
    {
        private const string Prefix = "PLANTILLA_REMITO::";

        private readonly DataContext _context;
        private readonly IUserContextService _userContextService;

        public RemitoPlantillaService(DataContext context, IUserContextService userContextService)
        {
            _context = context;
            _userContextService = userContextService;
        }

        private static LienzoDesignDto? DeserializeLienzo(string? lienzoJson)
        {
            if (string.IsNullOrWhiteSpace(lienzoJson))
            {
                return null;
            }

            try
            {
                return JsonSerializer.Deserialize<LienzoDesignDto>(lienzoJson, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });
            }
            catch
            {
                return null;
            }
        }

        private static string ResolveObjectDatasetKey(LienzoElementDto element, int index)
        {
            if (!string.IsNullOrWhiteSpace(element.Id))
            {
                return $"ObjectSql_{element.Id}";
            }

            return $"ObjectSql_{index}";
        }

        public async Task<List<RemitoPlantillaItemDto>> GetAllAsync(CancellationToken cancellationToken)
        {
            return await _context.ReportesPersonalizados
                .AsNoTracking()
                .Where(x => x.Nombre.StartsWith(Prefix))
                .OrderBy(x => x.Nombre)
                .Select(x => new RemitoPlantillaItemDto
                {
                    Id = x.Id,
                    Nombre = x.Nombre.Replace(Prefix, string.Empty),
                    FechaModificacion = x.FechaModificacion
                })
                .ToListAsync(cancellationToken);
        }

        public async Task<RemitoPlantillaDetailDto?> GetByIdAsync(int id, CancellationToken cancellationToken)
        {
            var entity = await _context.ReportesPersonalizados
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.Id == id && x.Nombre.StartsWith(Prefix), cancellationToken);

            if (entity == null)
            {
                return null;
            }

            var config = DeserializeConfig(entity.ConfiguracionJson);

            return new RemitoPlantillaDetailDto
            {
                Id = entity.Id,
                Nombre = entity.Nombre.Replace(Prefix, string.Empty),
                LienzoJson = config.LienzoJson,
                Datasets = config.Datasets,
                FechaCreacion = entity.FechaCreacion,
                FechaModificacion = entity.FechaModificacion
            };
        }

        public async Task<int> SaveAsync(RemitoPlantillaUpsertDto dto, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(dto.Nombre))
            {
                throw new Exception("El nombre de la plantilla es obligatorio.");
            }

            ValidateDatasets(dto.Datasets);

            var now = DateTime.Now;
            var config = new RemitoPlantillaConfigStorageDto
            {
                LienzoJson = dto.LienzoJson,
                Datasets = dto.Datasets ?? new List<RemitoPlantillaDatasetDto>()
            };

            var userId = _userContextService.GetCurrentUserId();
            var usuario = userId.HasValue && userId.Value > 0 ? userId.Value.ToString() : "sistema";

            if (dto.Id <= 0)
            {
                var entity = new ReportePersonalizado
                {
                    Nombre = Prefix + dto.Nombre.Trim(),
                    Usuario = usuario,
                    ConfiguracionJson = JsonSerializer.Serialize(config),
                    FechaCreacion = now,
                    FechaModificacion = now
                };

                _context.ReportesPersonalizados.Add(entity);
                await _context.SaveChangesAsync(cancellationToken);
                return entity.Id;
            }

            var existente = await _context.ReportesPersonalizados
                .FirstOrDefaultAsync(x => x.Id == dto.Id && x.Nombre.StartsWith(Prefix), cancellationToken);

            if (existente == null)
            {
                throw new Exception("La plantilla no existe.");
            }

            existente.Nombre = Prefix + dto.Nombre.Trim();
            existente.Usuario = usuario;
            existente.ConfiguracionJson = JsonSerializer.Serialize(config);
            existente.FechaModificacion = now;

            await _context.SaveChangesAsync(cancellationToken);
            return existente.Id;
        }

        public async Task<RemitoPlantillaPreviewResultDto> PreviewAsync(RemitoPlantillaPreviewRequestDto dto, CancellationToken cancellationToken)
        {
            if (dto == null || dto.Datasets == null || dto.Datasets.Count == 0)
            {
                throw new Exception("Debe indicar al menos un dataset para previsualizar.");
            }

            var result = new RemitoPlantillaPreviewResultDto();
            var connection = _context.Database.GetDbConnection();
            var shouldClose = connection.State != System.Data.ConnectionState.Open;
            if (shouldClose)
            {
                await connection.OpenAsync(cancellationToken);
            }

            try
            {
                foreach (var dataset in dto.Datasets)
                {
                    if (string.IsNullOrWhiteSpace(dataset.Key))
                    {
                        continue;
                    }

                    ValidateSql(dataset.Sql);
                    var rows = await ExecuteDatasetAsync(connection, dataset, dto.Parameters, cancellationToken);
                    result.Results[dataset.Key] = rows;
                }
            }
            finally
            {
                if (shouldClose && connection.State == System.Data.ConnectionState.Open)
                {
                    await connection.CloseAsync();
                }
            }

            return result;
        }

        public async Task EnrichPreviewWithObjectSqlAsync(
            string? lienzoJson,
            RemitoPlantillaPreviewResultDto preview,
            List<RemitoPlantillaParameterDto>? parameters,
            CancellationToken cancellationToken)
        {
            if (preview == null)
            {
                throw new ArgumentNullException(nameof(preview));
            }

            var design = DeserializeLienzo(lienzoJson);
            if (design?.Elements == null || design.Elements.Count == 0)
            {
                return;
            }

            var connection = _context.Database.GetDbConnection();
            var shouldClose = connection.State != System.Data.ConnectionState.Open;
            if (shouldClose)
            {
                await connection.OpenAsync(cancellationToken);
            }

            try
            {
                var index = 1;
                foreach (var element in design.Elements)
                {
                    if (string.IsNullOrWhiteSpace(element.SqlQuery))
                    {
                        continue;
                    }

                    ValidateSql(element.SqlQuery);

                    var datasetKey = ResolveObjectDatasetKey(element, index);
                    index++;

                    var dataset = new RemitoPlantillaDatasetDto
                    {
                        Key = datasetKey,
                        Sql = element.SqlQuery,
                        MaxRows = 1
                    };

                    var rows = await ExecuteDatasetAsync(connection, dataset, parameters ?? new List<RemitoPlantillaParameterDto>(), cancellationToken);
                    preview.Results[datasetKey] = rows;
                }
            }
            finally
            {
                if (shouldClose && connection.State == System.Data.ConnectionState.Open)
                {
                    await connection.CloseAsync();
                }
            }
        }

        private static RemitoPlantillaConfigStorageDto DeserializeConfig(string? json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new RemitoPlantillaConfigStorageDto();
            }

            return JsonSerializer.Deserialize<RemitoPlantillaConfigStorageDto>(json)
                   ?? new RemitoPlantillaConfigStorageDto();
        }

        private static void ValidateSql(string? sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
            {
                throw new Exception("La consulta SQL no puede estar vacía.");
            }

            var trimmed = sql.Trim();
            var startsWithSelect = trimmed.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase)
                                   || trimmed.StartsWith("WITH", StringComparison.OrdinalIgnoreCase);
            if (!startsWithSelect)
            {
                throw new Exception("Solo se permiten consultas SELECT.");
            }

            if (trimmed.Contains(";"))
            {
                throw new Exception("No se permiten múltiples sentencias SQL.");
            }

            if (trimmed.Contains("--", StringComparison.Ordinal)
                || trimmed.Contains("/*", StringComparison.Ordinal)
                || trimmed.Contains("*/", StringComparison.Ordinal))
            {
                throw new Exception("No se permiten comentarios SQL en la consulta.");
            }

            var forbiddenPatterns = new[]
            {
                @"\bINSERT\b",
                @"\bUPDATE\b",
                @"\bDELETE\b",
                @"\bDROP\b",
                @"\bALTER\b",
                @"\bTRUNCATE\b",
                @"\bMERGE\b",
                @"\bEXEC\b",
                @"\bEXECUTE\b",
                @"\bCREATE\b",
                @"\bGRANT\b",
                @"\bREVOKE\b",
                @"\bDENY\b",
                @"\bDBCC\b",
                @"\bBACKUP\b",
                @"\bRESTORE\b",
                @"\bSP_[A-Z0-9_]*\b",
                @"\bXP_[A-Z0-9_]*\b"
            };

            var upper = trimmed.ToUpperInvariant();
            foreach (var pattern in forbiddenPatterns)
            {
                if (Regex.IsMatch(upper, pattern, RegexOptions.CultureInvariant))
                {
                    throw new Exception("La consulta contiene instrucciones no permitidas por seguridad.");
                }
            }
        }

        private static void ValidateDatasets(List<RemitoPlantillaDatasetDto>? datasets)
        {
            if (datasets == null || datasets.Count == 0)
            {
                throw new Exception("Debe indicar al menos un dataset.");
            }

            foreach (var dataset in datasets)
            {
                if (string.IsNullOrWhiteSpace(dataset.Key))
                {
                    throw new Exception("Todos los datasets deben tener una clave (Key).");
                }

                ValidateSql(dataset.Sql);
            }
        }

        private static async Task<List<Dictionary<string, object?>>> ExecuteDatasetAsync(
            DbConnection connection,
            RemitoPlantillaDatasetDto dataset,
            List<RemitoPlantillaParameterDto> parameters,
            CancellationToken cancellationToken)
        {
            var rows = new List<Dictionary<string, object?>>();
            using var command = connection.CreateCommand();
            command.CommandText = dataset.Sql;
            command.CommandTimeout = 60;

            foreach (var parameter in parameters ?? new List<RemitoPlantillaParameterDto>())
            {
                if (string.IsNullOrWhiteSpace(parameter.Name))
                {
                    continue;
                }

                var dbParameter = command.CreateParameter();
                dbParameter.ParameterName = parameter.Name.StartsWith("@") ? parameter.Name : "@" + parameter.Name;
                dbParameter.Value = parameter.Value ?? (object)DBNull.Value;
                command.Parameters.Add(dbParameter);
            }

            var maxRows = dataset.MaxRows > 0 ? dataset.MaxRows : 500;
            using var reader = await command.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken) && rows.Count < maxRows)
            {
                var row = new Dictionary<string, object?>();
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                }

                rows.Add(row);
            }

            return rows;
        }

        private sealed class LienzoDesignDto
        {
            public List<LienzoElementDto> Elements { get; set; } = new();
        }

        private sealed class LienzoElementDto
        {
            public string? Id { get; set; }
            public string? Binding { get; set; }
            public string? SqlQuery { get; set; }
        }
    }
}
