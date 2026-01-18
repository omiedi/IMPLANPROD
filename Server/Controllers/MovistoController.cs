using DocumentFormat.OpenXml.InkML;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.Utilities;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/movisto")]
    public class MovistoController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<MasterController> _logger;
        private readonly IUsuarioFlexoftService _UsuarioFlexoftService;
        private readonly ICodiempresa _Codiempresa;

        public MovistoController(DataContext context, ILogger<MasterController> logger, IUsuarioFlexoftService UsuarioFlexoftService, ICodiempresa Codiempresa)
        {
            _context = context;
            _logger = logger;
            _UsuarioFlexoftService = UsuarioFlexoftService;
            _Codiempresa = Codiempresa;
        }

        [HttpPost("movistock")]
        public async Task<ActionResult> MovimientoStock([FromBody] MovistoDTO movimiento)
        {
            try
            {
                int? idFlexoft = _UsuarioFlexoftService.InterObtenerIdFlexoft(User);
                short? codigoEmpresa = _Codiempresa.IObtenerCodiempresa(User);
                if (!idFlexoft.HasValue)
                {
                    return BadRequest("(Sesion Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }
                    if (!movimiento.Cod_Inte.HasValue || !movimiento.DepoMovi.HasValue)
                {
                    return BadRequest("El código interno y el depósito son obligatorios");
                }

                var nuevoMovimiento = new Movisto
                {
                    FeReMovi = movimiento.FeReMovi,
                    CodiMovi = movimiento.CodiMovi, //"BR", // Código fijo para este tipo de movimiento
                    Cod_Inte = movimiento.Cod_Inte,
                    Cod_Exte = movimiento.Cod_exte,
                    Descrip = movimiento.Descrip,
                    DepoMovi = movimiento.DepoMovi.Value,
                    CantIngre = movimiento.CantIngre,
                    CantSalid = movimiento.CantSalid,
                    ReferCor = movimiento.ReferCor,
                    IdenLote = movimiento.IdenLote,
                    DoPriCor = movimiento.DoPriCor,
                    DoPriLar = movimiento.DoPriLar,
                    DoSecCor = movimiento.DoSecCor,
                    DoSecLar = movimiento.DoSecLar,
                    FechMov = DateTime.Now
                };

                _context.Movistos.Add(nuevoMovimiento);
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al registrar el movimiento de stock: {ex.Message}");
            }
        }

        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination)
        { //INCLUYO UN INNER JOIN DEL MASTER PARA UNIDAD DE MEDIDA INCLUIDO EN EL DTO: ProductoConSaldoDTO
            try
            {
                if (pagination.RecordsNumber <= 0)
                    return BadRequest("El número de registros por página debe ser mayor que cero.");

                var query = from ma in _context.masters.AsNoTracking()
                            join mo in _context.Movistos.AsNoTracking()
                                on ma.Codint equals mo.Cod_Inte into movimientos
                            from m in movimientos.DefaultIfEmpty()
                            group m by new
                            {
                                ma.Codint,
                                ma.Codigo,
                                ma.Descripcion,
                                ma.Umedida
                            } into g
                            select new ProductoConSaldoDTO
                            {
                                Cod_inte = g.Key.Codint,
                                Codigo = g.Key.Codigo,
                                Descripcion = g.Key.Descripcion,
                                UnidadMedida = g.Key.Umedida,
                                Saldo = g.Sum(x => (x != null ? (x.CantIngre ?? 0) - (x.CantSalid ?? 0) : 0))
                            };

                // Aplicar filtro si existe
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var filtroLower = pagination.Filter.ToLower();
                    query = query.Where(x =>
                        x.Descripcion.ToLower().Contains(filtroLower) ||
                        x.Codigo.ToLower().Contains(filtroLower));
                }

                var result = await query
                    .OrderBy(x => x.Descripcion)
                    .Paginate(pagination)
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener productos con saldo");
                return StatusCode(500, $"Error al obtener productos con saldo: {ex.Message}");
            }
        }
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination)
        {
            try
            {
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                // Usar la misma consulta que GetAsync para contar correctamente
                var query = from ma in _context.masters.AsNoTracking()
                            join mo in _context.Movistos.AsNoTracking()
                                on ma.Codint equals mo.Cod_Inte into movimientos
                            from m in movimientos.DefaultIfEmpty()
                            group m by new
                            {
                                ma.Codint,
                                ma.Codigo,
                                ma.Descripcion,
                                ma.Umedida
                            } into g
                            select new
                            {
                                Codint = g.Key.Codint,
                                Codigo = g.Key.Codigo,
                                Descripcion = g.Key.Descripcion
                            };

                // Aplicar filtro si existe
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var filtro = pagination.Filter.ToLower();
                    query = query.Where(x =>
                        x.Descripcion.ToLower().Contains(filtro) ||
                        x.Codigo.ToLower().Contains(filtro));
                }

                var count = await query.CountAsync();

                var totalPages = count == 0
                    ? 1
                    : (int)Math.Ceiling(count / (double)pagination.RecordsNumber);

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular el total de páginas");
                return StatusCode(500, $"Error al calcular el total de páginas: {ex.Message}");
            }
        }
        /// <summary>
        /// Obtiene el stock total de un producto sumando todos los depósitos
        /// </summary>
        /// <param name="repuestoId">Código interno del repuesto</param>
        /// <returns>Stock total del repuesto</returns>
        [HttpGet("stock/total/{codinte}")]
        public async Task<ActionResult<decimal>> GetStockTotal(int codinte)
        {
            var stockTotal = await _context.Movistos
                .Where(m => m.Cod_Inte == codinte)
                .SumAsync(m => (m.CantIngre ?? 0) - (m.CantSalid ?? 0));

            return stockTotal;
        }

        /// <summary>
        /// Obtiene el stock por depósito para un producto específico
        /// </summary>
        /// <param name="repuestoId">Código interno del repuesto</param>
        /// <returns>Lista de stock por depósito</returns>
        [HttpGet("stock/{codint:int}")]
        public async Task<ActionResult<List<StockDepositoDTO>>> GetStockPorDeposito(int codint)
        {
            try
            {
                var stocks = await _context.Movistos
                    .Where(m => m.Cod_Inte == codint)
                    .GroupBy(m => m.DepoMovi)
                    .Select(g => new StockDepositoDTO
                    {
                        Deposito = g.Key,
                        Stock = g.Sum(m => (m.CantIngre ?? 0) - (m.CantSalid ?? 0))
                    })
                    .ToListAsync();

                return Ok(stocks);
            }
            catch (Exception ex)
            {
                return BadRequest("Error al obtener el stock por depósito");
            }
        }
        [HttpGet("GetIdFlexoft")]
        [Authorize] // Asegurate de protegerlo
        public ActionResult<int?> GetIdFlexoft()
        {
            var user = HttpContext.User;

            if (user?.Identity?.IsAuthenticated ?? false)
            {
                var username = user.Identity.Name;
                var email = user.FindFirst(ClaimTypes.Email)?.Value;

                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
                {
                    var usuario = _context.Usuarios.FirstOrDefault(u =>
                        u.NombreUsuario == username && u.Email == email);

                    return usuario?.IdFlexoft;
                }
            }

            return Unauthorized(); // O NotFound(), según prefieras
        }
        //[HttpPost("transferencia")]
        //public async Task<ActionResult> RealizarTransferencia([FromBody] StockTransferenciaDTO transferencia)
        //{
        //    if (!ModelState.IsValid)
        //    {
        //        return BadRequest(ModelState);
        //    }

        //    if (transferencia == null)
        //    {
        //        return BadRequest("Los datos de transferencia son nulos");
        //    }

        //    using var transaction = await _context.Database.BeginTransactionAsync();
        //    try
        //    {
        //        var fecha = DateTime.Now;
        //        var secuencial = await documentNumberGenerator.GenerateNextDocumentNumberAsync("TR");
        //        var seccion = $"AJ.{fecha:yyMMdd}";

        //        // Registro de salida del depósito origen
        //        var movimientoSalida = new Movisto
        //        {
        //            FeReMovi = fecha,
        //            CodiMovi = "TR",
        //            CantSalid = transferencia.CantidadTransferir,
        //            DepoMovi = transferencia.DepositoOrigen,
        //            Cod_Inte = transferencia.CodigoRepuesto,
        //            Cod_Exte = transferencia.CodExt,
        //            Descrip = transferencia.Descripcion,
        //            DoPriCor = $"TR.{secuencial:000000}",
        //            DoPriLar = $"TR.{secuencial:000000}",
        //            DoSecCor = seccion,
        //            DoSecLar = seccion,
        //            ReferCor = transferencia.Referencia
        //        };

        //        // Registro de ingreso al depósito destino
        //        var movimientoIngreso = new Movisto
        //        {
        //            FeReMovi = fecha,
        //            CodiMovi = "TR",
        //            CantIngre = transferencia.CantidadTransferir,
        //            DepoMovi = transferencia.DepositoDestino,
        //            Cod_Inte = transferencia.CodigoRepuesto,
        //            Cod_Exte = transferencia.CodExt,
        //            Descrip = transferencia.Descripcion,
        //            DoPriCor = $"TR.{secuencial:000000}",
        //            DoPriLar = $"TR.{secuencial:000000}",
        //            DoSecCor = seccion,
        //            DoSecLar = seccion,
        //            ReferCor = transferencia.Referencia
        //        };

        //        _context.Movistos.Add(movimientoSalida);
        //        _context.Movistos.Add(movimientoIngreso);
        //        await _context.SaveChangesAsync();
        //        await transaction.CommitAsync();

        //        return Ok(new { Message = "Transferencia realizada con éxito" });
        //    }
        //    catch (Exception ex)
        //    {
        //        await transaction.RollbackAsync();
        //        return StatusCode(500, $"Error interno del servidor: {ex.Message}");
        //    }
        //}

        /// <summary>
        /// Obtiene los movimientos de un remito específico usando clave compuesta
        /// (DoPriLar, NumeCli, FechMov) para evitar ambigüedad cuando hay números duplicados
        /// </summary>
        [HttpGet("por-remito")]
        public async Task<ActionResult<List<Movisto>>> GetPorRemito(
            [FromQuery] string doPriLar,
            [FromQuery] int numeCli,
            [FromQuery] string fechMov)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(doPriLar))
                {
                    return BadRequest("El número de remito (doPriLar) es obligatorio");
                }

                if (numeCli <= 0)
                {
                    return BadRequest("El número de cliente (numeCli) es obligatorio");
                }

                if (!DateTime.TryParse(fechMov, out DateTime fecha))
                {
                    return BadRequest("La fecha (fechMov) no tiene un formato válido");
                }

                // Buscar movimientos por clave compuesta: DoPriLar + NumeCli + FechMov
                var movimientos = await _context.Movistos
                    .Where(m => m.DoPriLar == doPriLar 
                             && m.Nume_Clie == numeCli 
                             && m.FechMov.HasValue 
                             && m.FechMov.Value.Date == fecha.Date)
                    .OrderBy(m => m.Id)
                    .ToListAsync();

                return Ok(movimientos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener movimientos del remito {DoPriLar} - Cliente {NumeCli} - Fecha {FechMov}", 
                    doPriLar, numeCli, fechMov);
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }
    }
}
