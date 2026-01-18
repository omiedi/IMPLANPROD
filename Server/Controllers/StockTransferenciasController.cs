using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.Utilities;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;

// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Actualización del controlador para usar el servicio centralizado
// de contexto de usuario en lugar de inyectar múltiples servicios relacionados con el usuario

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para gestionar las transferencias de stock entre depósitos
    /// </summary>
    [ApiController]
    [Route("api/stockTransferencia")]
    public class StockTransferenciasController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly DocumentNumberGenerator _documentNumberGenerator;
        private readonly IUserContextService _userContextService;

        /// <summary>
        /// Constructor del controlador de transferencias de stock
        /// </summary>
        /// <param name="context">Contexto de datos</param>
        /// <param name="documentNumberGenerator">Generador de números de documento</param>
        /// <param name="userContextService">Servicio centralizado de contexto de usuario</param>
        public StockTransferenciasController(
            DataContext context, 
            DocumentNumberGenerator documentNumberGenerator, 
            IUserContextService userContextService)
        {
            _context = context;
            _documentNumberGenerator = documentNumberGenerator;
            _userContextService = userContextService;
        }

        [HttpPost("transferencia")]
        public async Task<ActionResult> RealizarTransferencia([FromBody] StockTransferenciaDTO transferencia)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (transferencia == null)
            {
                return BadRequest("Los datos de transferencia son nulos");
            }

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                var fecha = DateTime.Now;
                var secuencial = await _documentNumberGenerator.GenerateNextDocumentNumberAsync("TR");
                var seccion = $"AJ.{fecha:yyMMdd}";
                
                // TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Uso del servicio centralizado para obtener información del usuario
                int? idFlexoft = _userContextService.GetCurrentUserId();
                short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();
                
                // Verificar si el usuario está autenticado usando el servicio centralizado
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }
                // Registro de salida del depósito origen
                var movimientoSalida = new Movisto
                {
                    FeReMovi = fecha,
                    CodiMovi = "TR",
                    CantSalid = transferencia.CantidadTransferir,
                    DepoMovi = transferencia.DepositoOrigen,
                    Cod_Inte = transferencia.CodigoRepuesto,
                    Cod_Exte = transferencia.CodExt,
                    Descrip = transferencia.Descripcion,
                    DoPriCor = $"TR.{secuencial:000000}",
                    DoPriLar = $"TR.{secuencial:000000}",
                    DoSecCor = seccion,
                    DoSecLar = seccion,
                    NumUsuar = idFlexoft,
                    CodiEmpr = codigoEmpresa,
                    ReferCor = transferencia.Referencia
                };

                // Registro de ingreso al depósito destino
                var movimientoIngreso = new Movisto
                {
                    FeReMovi = fecha,
                    CodiMovi = "TR",
                    CantIngre = transferencia.CantidadTransferir,
                    DepoMovi = transferencia.DepositoDestino,
                    Cod_Inte = transferencia.CodigoRepuesto,
                    Cod_Exte = transferencia.CodExt,
                    Descrip = transferencia.Descripcion,
                    DoPriCor = $"TR.{secuencial:000000}",
                    DoPriLar = $"TR.{secuencial:000000}",
                    DoSecCor = seccion,
                    DoSecLar = seccion,
                    NumUsuar = idFlexoft,
                    CodiEmpr = codigoEmpresa,
                    ReferCor = transferencia.Referencia
                };

                _context.Movistos.Add(movimientoSalida);
                _context.Movistos.Add(movimientoIngreso);
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return Ok(new { Message = "Transferencia realizada con éxito" });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }
    }
}
// codiempr = dependenciaId* 100 + empresaId;
//Si:
//- dependenciaId = 1
//- empresaId = 1
//Entonces:
//codiempr = 1 * 100 + 1 = 101
//Resultado: codiempr = 101

//Este valor es reversible fácilmente:
//dependenciaId = 101 / 100 = 1
//empresaId = 101 % 100 = 1

