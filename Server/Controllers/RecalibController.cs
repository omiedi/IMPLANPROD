using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;
using System.Data;

namespace IMPLANPROD.Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class RecalibController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<RecalibController> _logger;

        public RecalibController(DataContext context, ILogger<RecalibController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todos los registros de calibración de un instrumento específico
        /// Verifica que el instrumento exista y que el CodIntMedIdent coincida
        /// </summary>
        /// <param name="codIntMedIdent">Código interno del instrumento</param>
        /// <returns>Lista de registros de calibración</returns>
        [HttpGet("ByInstrumento/{codIntMedIdent}")]
        public async Task<ActionResult<List<Recalib>>> GetByInstrumento(int codIntMedIdent)
        {
            try
            {
                // Verificar que el instrumento existe
                var instrumentoExiste = await _context.Instrumes
                    .AnyAsync(i => i.CodIntMedIdent == codIntMedIdent);

                if (!instrumentoExiste)
                {
                    return Ok(new List<Recalib>()); // Retorna lista vacía si no existe el instrumento
                }

                // SELECT r.* FROM INSTRUME i
                // INNER JOIN RECALIB r ON i.CodIntMedIdent = r.CodIntMedIdent
                // WHERE i.CodIntMedIdent = @codIntMedIdent
                var registros = await (from i in _context.Instrumes
                                       join r in _context.Recalibs
                                       on i.CodIntMedIdent equals r.CodIntMedIdent
                                       where i.CodIntMedIdent == codIntMedIdent
                                       orderby r.FeCalibracion descending
                                       select new Recalib
                                       {
                                           Id = r.Id,
                                           RegCalibrNro = r.RegCalibrNro,
                                           FeCalibracion = r.FeCalibracion,
                                           FeProxCalibr = r.FeProxCalibr,
                                           IcNro = r.IcNro,
                                           PatronUtilizado = r.PatronUtilizado,
                                           RequirioCalib = r.RequirioCalib,
                                           RequirioAjuste = r.RequirioAjuste,
                                           Observaciones = r.Observaciones,
                                           CodIntMedIdent = r.CodIntMedIdent,
                                           Status = r.Status,
                                           AddRecord = r.AddRecord,
                                           LastUpdate = r.LastUpdate,
                                           CodiEmprNet = r.CodiEmprNet,
                                           NumUsuar = r.NumUsuar
                                       }).ToListAsync();

                return Ok(registros);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener registros de calibración para instrumento {codIntMedIdent}");
                return StatusCode(500, new { error = "Error al obtener registros de calibración" });
            }
        }

        /// <summary>
        /// Obtiene un registro de calibración por ID
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<Recalib>> GetById(int id)
        {
            try
            {
                var recalib = await _context.Recalibs
                    .Where(r => r.Id == id && r.Status == 0)
                    .Select(r => new Recalib
                    {
                        Id = r.Id,
                        RegCalibrNro = r.RegCalibrNro,
                        FeCalibracion = r.FeCalibracion,
                        FeProxCalibr = r.FeProxCalibr,
                        IcNro = r.IcNro,
                        PatronUtilizado = r.PatronUtilizado,
                        RequirioCalib = r.RequirioCalib,
                        RequirioAjuste = r.RequirioAjuste,
                        Observaciones = r.Observaciones,
                        CodIntMedIdent = r.CodIntMedIdent,
                        Status = r.Status,
                        AddRecord = r.AddRecord,
                        LastUpdate = r.LastUpdate,
                        CodiEmprNet = r.CodiEmprNet,
                        NumUsuar = r.NumUsuar
                    })
                    .FirstOrDefaultAsync();

                if (recalib == null)
                    return NotFound();

                return Ok(recalib);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener registro de calibración {id}");
                return StatusCode(500, new { error = "Error al obtener registro de calibración" });
            }
        }

        /// <summary>
        /// Obtiene todos los registros de calibración activos
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<List<Recalib>>> GetAll()
        {
            try
            {
                var registros = await _context.Recalibs
                    .Where(r => r.Status == 0)
                    .OrderByDescending(r => r.FeCalibracion)
                    .Select(r => new Recalib
                    {
                        Id = r.Id,
                        RegCalibrNro = r.RegCalibrNro,
                        FeCalibracion = r.FeCalibracion,
                        FeProxCalibr = r.FeProxCalibr,
                        IcNro = r.IcNro,
                        PatronUtilizado = r.PatronUtilizado,
                        RequirioCalib = r.RequirioCalib,
                        RequirioAjuste = r.RequirioAjuste,
                        Observaciones = r.Observaciones,
                        CodIntMedIdent = r.CodIntMedIdent,
                        Status = r.Status,
                        AddRecord = r.AddRecord,
                        LastUpdate = r.LastUpdate,
                        CodiEmprNet = r.CodiEmprNet,
                        NumUsuar = r.NumUsuar
                    })
                    .ToListAsync();

                return Ok(registros);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener registros de calibración");
                return StatusCode(500, new { error = "Error al obtener registros de calibración" });
            }
        }

        [HttpGet("NextId")]
        public async Task<ActionResult<int>> GetNextId()
        {
            try
            {
                var maxId = await _context.Recalibs
                    .Select(r => (int?)r.Id)
                    .MaxAsync() ?? 0;

                return Ok(maxId + 1);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener próximo Id de RECALIB");
                return StatusCode(500, new { error = "Error al obtener próximo Id" });
            }
        }

        [HttpPost]
        public async Task<ActionResult<Recalib>> Create([FromBody] Recalib model)
        {
            IDbContextTransaction? transaction = null;

            try
            {
                if (model.CodIntMedIdent is null || model.CodIntMedIdent <= 0)
                {
                    return BadRequest(new { error = "CodIntMedIdent es obligatorio" });
                }

                if (model.FeCalibracion is null)
                {
                    return BadRequest(new { error = "Req.Calibración es obligatorio" });
                }

                if (model.FeProxCalibr is null)
                {
                    return BadRequest(new { error = "Prox.Calibración es obligatorio" });
                }

                if (string.IsNullOrWhiteSpace(model.IcNro)
                    || string.IsNullOrWhiteSpace(model.PatronUtilizado)
                    || string.IsNullOrWhiteSpace(model.RequirioAjuste)
                    || string.IsNullOrWhiteSpace(model.Responsable))
                {
                    return BadRequest(new { error = "Todos los campos obligatorios de RECALIB deben estar completos" });
                }

                var icNro = model.IcNro.Trim();
                var patronUtilizado = model.PatronUtilizado.Trim();
                var requirioCalib = model.RequirioCalib?.Trim();
                var requirioAjuste = model.RequirioAjuste.Trim();
                var observaciones = model.Observaciones;
                var responsable = model.Responsable.Trim();

                if (icNro.Length > 16)
                {
                    return BadRequest(new { error = "IC-NRO no puede superar 16 caracteres" });
                }

                if (patronUtilizado.Length > 16)
                {
                    return BadRequest(new { error = "Patrón Utilizado no puede superar 16 caracteres" });
                }

                if (!string.IsNullOrWhiteSpace(requirioCalib) && requirioCalib.Length > 16)
                {
                    return BadRequest(new { error = "Requirió Calibración no puede superar 16 caracteres" });
                }

                if (requirioAjuste.Length > 16)
                {
                    return BadRequest(new { error = "Requirió Ajuste no puede superar 16 caracteres" });
                }

                if (responsable.Length > 32)
                {
                    return BadRequest(new { error = "Responsable no puede superar 32 caracteres" });
                }

                var tieneColumnaResponsable = await ExisteColumnaRecalibAsync("RESPONSABLE");

                transaction = await _context.Database.BeginTransactionAsync();

                var connection = _context.Database.GetDbConnection();
                if (connection.State != ConnectionState.Open)
                {
                    await connection.OpenAsync();
                }

                await using var command = connection.CreateCommand();
                command.Transaction = transaction.GetDbTransaction();

                if (tieneColumnaResponsable)
                {
                    command.CommandText = @"
INSERT INTO [dbo].[RECALIB]
    ([RegCalibr_Nro], [FeCalibracion], [FeProxCalibr], [IC_NRO], [PatronUtilizado], [RequirioCalib], [RequirioAjuste], [OBSERVACIONES], [RESPONSABLE], [CodIntMedIdent], [Status], [AddRecord], [LastUpdate], [CodiEmprNet], [NumUsuar])
VALUES
    (NULL, @FeCalibracion, @FeProxCalibr, @IC_NRO, @PatronUtilizado, @RequirioCalib, @RequirioAjuste, @OBSERVACIONES, @RESPONSABLE, @CodIntMedIdent, @Status, @AddRecord, @LastUpdate, @CodiEmprNet, @NumUsuar);
SELECT CAST(SCOPE_IDENTITY() AS INT);";
                }
                else
                {
                    command.CommandText = @"
INSERT INTO [dbo].[RECALIB]
    ([RegCalibr_Nro], [FeCalibracion], [FeProxCalibr], [IC_NRO], [PatronUtilizado], [RequirioCalib], [RequirioAjuste], [OBSERVACIONES], [CodIntMedIdent], [Status], [AddRecord], [LastUpdate], [CodiEmprNet], [NumUsuar])
VALUES
    (NULL, @FeCalibracion, @FeProxCalibr, @IC_NRO, @PatronUtilizado, @RequirioCalib, @RequirioAjuste, @OBSERVACIONES, @CodIntMedIdent, @Status, @AddRecord, @LastUpdate, @CodiEmprNet, @NumUsuar);
SELECT CAST(SCOPE_IDENTITY() AS INT);";
                }

                static void AddParameter(IDbCommand cmd, string name, object? value)
                {
                    var p = cmd.CreateParameter();
                    p.ParameterName = name;
                    p.Value = value ?? DBNull.Value;
                    cmd.Parameters.Add(p);
                }

                var ahora = DateTime.Now;
                AddParameter(command, "@FeCalibracion", model.FeCalibracion);
                AddParameter(command, "@FeProxCalibr", model.FeProxCalibr);
                AddParameter(command, "@IC_NRO", icNro);
                AddParameter(command, "@PatronUtilizado", patronUtilizado);
                AddParameter(command, "@RequirioCalib", requirioCalib);
                AddParameter(command, "@RequirioAjuste", requirioAjuste);
                AddParameter(command, "@OBSERVACIONES", observaciones);
                if (tieneColumnaResponsable)
                {
                    AddParameter(command, "@RESPONSABLE", responsable);
                }
                AddParameter(command, "@CodIntMedIdent", model.CodIntMedIdent);
                AddParameter(command, "@Status", 0);
                AddParameter(command, "@AddRecord", model.AddRecord ?? ahora);
                AddParameter(command, "@LastUpdate", ahora);
                AddParameter(command, "@CodiEmprNet", model.CodiEmprNet ?? 0);
                AddParameter(command, "@NumUsuar", model.NumUsuar);

                var newIdObj = await command.ExecuteScalarAsync();
                if (newIdObj == null || !int.TryParse(newIdObj.ToString(), out var newId) || newId <= 0)
                {
                    throw new InvalidOperationException("No se pudo obtener el Id generado de RECALIB.");
                }

                var icNroFinal = $"IC.{newId}";
                await _context.Database.ExecuteSqlInterpolatedAsync($@"
UPDATE [dbo].[RECALIB]
SET [RegCalibr_Nro] = {newId},
    [IC_NRO] = {icNroFinal},
    [LastUpdate] = {DateTime.Now}
WHERE [Id] = {newId};");

                await transaction.CommitAsync();

                var entity = await _context.Recalibs
                    .Where(r => r.Id == newId)
                    .Select(r => new Recalib
                    {
                        Id = r.Id,
                        RegCalibrNro = r.RegCalibrNro,
                        FeCalibracion = r.FeCalibracion,
                        FeProxCalibr = r.FeProxCalibr,
                        IcNro = r.IcNro,
                        PatronUtilizado = r.PatronUtilizado,
                        RequirioCalib = r.RequirioCalib,
                        RequirioAjuste = r.RequirioAjuste,
                        Observaciones = r.Observaciones,
                        CodIntMedIdent = r.CodIntMedIdent,
                        Status = r.Status,
                        AddRecord = r.AddRecord,
                        LastUpdate = r.LastUpdate,
                        CodiEmprNet = r.CodiEmprNet,
                        NumUsuar = r.NumUsuar
                    })
                    .FirstAsync();

                return Ok(entity);
            }
            catch (Exception ex)
            {
                if (transaction is not null)
                {
                    await transaction.RollbackAsync();
                }

                _logger.LogError(ex, "Error al crear registro de calibración");
                return StatusCode(500, new { error = $"Error al crear registro de calibración: {ex.Message}" });
            }
        }

        private async Task<bool> ExisteColumnaRecalibAsync(string columnName)
        {
            var connection = _context.Database.GetDbConnection();
            if (connection.State != ConnectionState.Open)
            {
                await connection.OpenAsync();
            }

            await using var command = connection.CreateCommand();
            command.CommandText = @"
SELECT COUNT(1)
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'RECALIB'
  AND COLUMN_NAME = @ColumnName;";

            var parameter = command.CreateParameter();
            parameter.ParameterName = "@ColumnName";
            parameter.Value = columnName;
            command.Parameters.Add(parameter);

            var result = await command.ExecuteScalarAsync();
            return result != null && int.TryParse(result.ToString(), out var count) && count > 0;
        }
    }
}
