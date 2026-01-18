using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Utilities;
using IMPLANPROD.Server.Services;

namespace MantenimientoImp.Server.Controllers
{
    /// <summary>
    /// Controlador para gestionar los ajustes de stock de repuestos
    /// </summary>
    [ApiController]
    [Route("api/stockajustes")]
    public class StockAjustesController : ControllerBase
    {
        private readonly DataContext context;
        private readonly DocumentNumberGenerator documentNumberGenerator;
        private readonly IUsuarioFlexoftService _UsuarioFlexoftService;
        private readonly ICodiempresa _Codiempresa;


        /// <summary>
        /// Constructor del controlador
        /// </summary>
        /// <param name="context">Contexto de base de datos</param>
        public StockAjustesController(DataContext context, IUsuarioFlexoftService UsuarioFlexoftService, ICodiempresa Codiempresa)
        {
            this.context = context;
            this.documentNumberGenerator = new DocumentNumberGenerator(context);
            _UsuarioFlexoftService = UsuarioFlexoftService;
            _Codiempresa = Codiempresa;

        }

        /// <summary>
        /// Obtiene el stock por depósito para un repuesto específico
        /// </summary>
        /// <param name="codint">Código interno del repuesto</param>
        /// <returns>Lista de stock por depósito</returns>
        [HttpGet("stock/{codint:int}")]
        public async Task<ActionResult<List<StockDepositoDTO>>> GetStockPorDeposito(int codint)
        {
            try
            {
                var stocks = await context.Movistos
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

        /// <summary>
        /// Obtiene el stock total de un repuesto sumando todos los depósitos
        /// </summary>
        /// <param name="codint">Código interno del repuesto</param>
        /// <returns>Stock total del repuesto</returns>
        [HttpGet("stock/total/{codint}")]
        public async Task<ActionResult<decimal>> GetStockTotal(int codint)
        {
            var stockTotal = await context.Movistos
                .Where(m => m.Cod_Inte == codint)
                .SumAsync(m => (m.CantIngre ?? 0) - (m.CantSalid ?? 0));

            return stockTotal;
        }

        /// <summary>
        /// Realiza un ajuste de stock (ingreso o salida)
        /// </summary>
        /// <param name="ajuste">Datos del ajuste a realizar</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPost("ajuste")]
        public async Task<ActionResult> AjustarStock([FromBody] StockAjusteDTO ajuste)
        {
            try
            {
                if (!ajuste.CodInte.HasValue || !ajuste.Deposito.HasValue)
                {
                    return BadRequest("El código interno y el depósito son obligatorios");
                }

                var codiMovi = "AJ";
                var docNumber = await documentNumberGenerator.GenerateNextDocumentNumberAsync(codiMovi);
                var secDocNumber = documentNumberGenerator.GenerateSecondaryDocumentNumber(codiMovi);
                int? idFlexoft = _UsuarioFlexoftService.InterObtenerIdFlexoft(User);
                short? codigoEmpresa = _Codiempresa.IObtenerCodiempresa(User);
                if (!idFlexoft.HasValue)
                {
                    return BadRequest("(Sesion Caducada - Debe cerrar sesión y loguearse nuevamente.");

                }

                var movimiento = new Movisto
                {
                    FeReMovi = DateTime.Now,
                    CodiMovi = codiMovi,
                    Cod_Inte = ajuste.CodInte,
                    Cod_Exte = ajuste.CodExte,
                    Descrip = ajuste.Descripcion,
                    DepoMovi = ajuste.Deposito.Value,
                    CantIngre = ajuste.Ingreso > 0 ? ajuste.Ingreso : null,
                    CantSalid = ajuste.Salida > 0 ? ajuste.Salida : null,
                    ReferCor = ajuste.Referencia,
                    FechMov = DateTime.Now,
                    NumUsuar = idFlexoft,
                    CodiEmpr = codigoEmpresa,
                    DoPriCor = docNumber,
                    DoPriLar = docNumber,
                    DoSecCor = secDocNumber,
                    DoSecLar = secDocNumber
                };

                context.Movistos.Add(movimiento);
                await context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al ajustar el stock: {ex.Message}");
            }
        }

        /// <summary>
        /// Revierte un movimiento de stock
        /// </summary>
        /// <param name="movimiento">Movimiento a revertir</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPost("revertir")]
        public async Task<ActionResult> RevertirMovimiento([FromBody] MovimientoHistoricoDTO movimiento)
        {
            try
            {
                var movimientoOriginal = await context.Movistos.FindAsync(movimiento.Id);
                if (movimientoOriginal == null)
                    return NotFound("Movimiento no encontrado");

                // Crear el movimiento de reversión
                var movimientoReversion = new Movisto
                {
                    FeReMovi = DateTime.Now,
                    CodiMovi = "RV", // RV = Reversión
                    CantIngre = movimientoOriginal.CantSalid,
                    CantSalid = movimientoOriginal.CantIngre,
                    DepoMovi = movimientoOriginal.DepoMovi,
                    Cod_Inte = movimientoOriginal.Cod_Inte,
                    Cod_Exte = movimientoOriginal.Cod_Exte,
                    Descrip = movimientoOriginal.Descrip,
                    NumUsuar = movimientoOriginal.NumUsuar,
                    CodiEmpr = movimientoOriginal.CodiEmpr,
                    DoPriCor = movimientoOriginal.DoPriCor,
                    DoPriLar = movimientoOriginal.DoPriLar,
                    DoSecCor = movimientoOriginal.DoSecCor,
                    DoSecLar = movimientoOriginal.DoSecLar,
                    IdenLote = movimientoOriginal.IdenLote,
                    ReferCor = $"Reversión del movimiento {movimientoOriginal.DoPriCor}"
                };

                await context.Movistos.AddAsync(movimientoReversion);
                await context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al revertir el movimiento: {ex.Message}");
            }
        }

      
    }
}
