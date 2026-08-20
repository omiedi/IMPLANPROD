using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.Utilities;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

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
                var secuencialCompleto = await _documentNumberGenerator.GenerateNextDocumentNumberAsync("TR");
                // Extraer solo el número del formato "TR.000000"
                var numeroSecuencial = int.Parse(secuencialCompleto.Split('.').Last());
                // Formato YMMDD: año 1 dígito (últimos 2 dígitos del año % 10) + mes + día
                var seccion = $"{fecha.Year % 10}{fecha:MMdd}";
                
                // TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Uso del servicio centralizado para obtener información del usuario
                int? idFlexoft = _userContextService.GetCurrentUserId();
                short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();
                
                // Verificar si el usuario está autenticado usando el servicio centralizado
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Validar trazabilidad y lote antes de procesar
                var producto = await _context.masters
                    .Where(m => m.Codint == transferencia.CodigoRepuesto)
                    .Select(m => new { m.Codint, m.Codigo, m.Trazable })
                    .FirstOrDefaultAsync();

                if (producto == null)
                {
                    return BadRequest($"El producto con código interno {transferencia.CodigoRepuesto} no existe.");
                }

                // Si el producto es trazable (Trazable = 1), el lote es obligatorio
                if (producto.Trazable == 1)
                {
                    if (string.IsNullOrWhiteSpace(transferencia.Lote))
                    {
                        return BadRequest($"El producto '{transferencia.CodExt}' es trazable y requiere que se especifique un lote obligatoriamente.");
                    }

                    // Validar que el lote exista en MOVISTO con al menos un ingreso
                    var loteExiste = await _context.Movistos
                        .AnyAsync(m => m.Cod_Inte == transferencia.CodigoRepuesto 
                                    && m.IdenLote == transferencia.Lote 
                                    && m.CantIngre.HasValue 
                                    && m.CantIngre.Value > 0);

                    if (!loteExiste)
                    {
                        return BadRequest($"El lote '{transferencia.Lote}' no existe para el producto '{transferencia.CodExt}'. Debe existir al menos un ingreso previo con ese lote.");
                    }
                }
                else
                {
                    // Si el producto NO es trazable (Trazable = 0 o null), el lote es opcional
                    // Pero si se ingresó un lote, validar que exista
                    if (!string.IsNullOrWhiteSpace(transferencia.Lote))
                    {
                        var loteExiste = await _context.Movistos
                            .AnyAsync(m => m.Cod_Inte == transferencia.CodigoRepuesto 
                                        && m.IdenLote == transferencia.Lote 
                                        && m.CantIngre.HasValue 
                                        && m.CantIngre.Value > 0);

                        if (!loteExiste)
                        {
                            return BadRequest($"El lote '{transferencia.Lote}' no existe para el producto '{transferencia.CodExt}'. Debe existir al menos un ingreso previo con ese lote.");
                        }
                    }
                }

                // Construir referencia completa incluyendo OC si existe
                var referenciaCompleta = transferencia.Referencia;
                if (!string.IsNullOrWhiteSpace(transferencia.OrdenCompra))
                {
                    referenciaCompleta = $"{transferencia.Referencia} | OC:{transferencia.OrdenCompra}";
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
                    DoPriCor = $"TR.{numeroSecuencial}",
                    DoPriLar = $"TR.{numeroSecuencial}",
                    DoSecCor = seccion,
                    DoSecLar = seccion,
                    NumUsuar = idFlexoft,
                    CodiEmpr = codigoEmpresa,
                    ReferCor = referenciaCompleta,
                    IdenLote = transferencia.Lote
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
                    DoPriCor = $"TR.{numeroSecuencial}",
                    DoPriLar = $"TR.{numeroSecuencial}",
                    DoSecCor = seccion,
                    DoSecLar = seccion,
                    NumUsuar = idFlexoft,
                    CodiEmpr = codigoEmpresa,
                    ReferCor = referenciaCompleta,
                    IdenLote = transferencia.Lote
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

        [HttpPost("transferencia-multiple")]
        public async Task<ActionResult> RealizarTransferenciaMultiple([FromBody] StockTransferenciaMultipleDTO transferenciaMultiple)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (transferenciaMultiple == null || !transferenciaMultiple.Items.Any())
            {
                return BadRequest("Debe proporcionar al menos un producto para transferir");
            }

            if (transferenciaMultiple.DepositoOrigen == transferenciaMultiple.DepositoDestino)
            {
                return BadRequest("El depósito origen y destino no pueden ser el mismo");
            }

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                var fecha = DateTime.Now;
                var secuencialCompleto = await _documentNumberGenerator.GenerateNextDocumentNumberAsync("TR");
                // Extraer solo el número del formato "TR.000000"
                var numeroSecuencial = int.Parse(secuencialCompleto.Split('.').Last());
                var seccion = $"{fecha.Year % 10}{fecha:MMdd}";
                
                int? idFlexoft = _userContextService.GetCurrentUserId();
                short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();
                
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Validar trazabilidad y lotes antes de procesar
                foreach (var item in transferenciaMultiple.Items)
                {
                    // Obtener información del producto para verificar trazabilidad
                    var producto = await _context.masters
                        .Where(m => m.Codint == item.CodigoInterno)
                        .Select(m => new { m.Codint, m.Codigo, m.Trazable })
                        .FirstOrDefaultAsync();

                    if (producto == null)
                    {
                        return BadRequest($"El producto con código interno {item.CodigoInterno} no existe.");
                    }

                    // Si el producto es trazable (Trazable = 1), el lote es obligatorio
                    if (producto.Trazable == 1)
                    {
                        if (string.IsNullOrWhiteSpace(item.Lote))
                        {
                            return BadRequest($"El producto '{item.CodigoExterno}' es trazable y requiere que se especifique un lote obligatoriamente.");
                        }

                        // Validar que el lote exista en MOVISTO con al menos un ingreso
                        var loteExiste = await _context.Movistos
                            .AnyAsync(m => m.Cod_Inte == item.CodigoInterno 
                                        && m.IdenLote == item.Lote 
                                        && m.CantIngre.HasValue 
                                        && m.CantIngre.Value > 0);

                        if (!loteExiste)
                        {
                            return BadRequest($"El lote '{item.Lote}' no existe para el producto '{item.CodigoExterno}'. Debe existir al menos un ingreso previo con ese lote.");
                        }
                    }
                    else
                    {
                        // Si el producto NO es trazable (Trazable = 0 o null), el lote es opcional
                        // Pero si se ingresó un lote, validar que exista
                        if (!string.IsNullOrWhiteSpace(item.Lote))
                        {
                            var loteExiste = await _context.Movistos
                                .AnyAsync(m => m.Cod_Inte == item.CodigoInterno 
                                            && m.IdenLote == item.Lote 
                                            && m.CantIngre.HasValue 
                                            && m.CantIngre.Value > 0);

                            if (!loteExiste)
                            {
                                return BadRequest($"El lote '{item.Lote}' no existe para el producto '{item.CodigoExterno}'. Debe existir al menos un ingreso previo con ese lote.");
                            }
                        }
                    }
                }

                // Procesar cada item de la transferencia
                foreach (var item in transferenciaMultiple.Items)
                {
                    // Construir referencia completa incluyendo OC si existe
                    var referenciaCompleta = transferenciaMultiple.Referencia;
                    if (!string.IsNullOrWhiteSpace(item.OrdenCompra))
                    {
                        referenciaCompleta = $"{transferenciaMultiple.Referencia} | OC:{item.OrdenCompra}";
                    }

                    // Registro de salida del depósito origen
                    var movimientoSalida = new Movisto
                    {
                        FeReMovi = fecha,
                        CodiMovi = "TR",
                        CantSalid = item.Cantidad,
                        DepoMovi = transferenciaMultiple.DepositoOrigen,
                        Cod_Inte = item.CodigoInterno,
                        Cod_Exte = item.CodigoExterno,
                        Descrip = item.Descripcion,
                        DoPriCor = $"TR.{numeroSecuencial}",
                        DoPriLar = $"TR.{numeroSecuencial}",
                        DoSecCor = seccion,
                        DoSecLar = seccion,
                        NumUsuar = idFlexoft,
                        CodiEmpr = codigoEmpresa,
                        ReferCor = referenciaCompleta,
                        IdenLote = item.Lote
                    };

                    // Registro de ingreso al depósito destino
                    var movimientoIngreso = new Movisto
                    {
                        FeReMovi = fecha,
                        CodiMovi = "TR",
                        CantIngre = item.Cantidad,
                        DepoMovi = transferenciaMultiple.DepositoDestino,
                        Cod_Inte = item.CodigoInterno,
                        Cod_Exte = item.CodigoExterno,
                        Descrip = item.Descripcion,
                        DoPriCor = $"TR.{numeroSecuencial}",
                        DoPriLar = $"TR.{numeroSecuencial}",
                        DoSecCor = seccion,
                        DoSecLar = seccion,
                        NumUsuar = idFlexoft,
                        CodiEmpr = codigoEmpresa,
                        ReferCor = referenciaCompleta,
                        IdenLote = item.Lote
                    };

                    _context.Movistos.Add(movimientoSalida);
                    _context.Movistos.Add(movimientoIngreso);
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return Ok(new { Message = $"Transferencia múltiple realizada con éxito. {transferenciaMultiple.Items.Count} producto(s) transferido(s)." });
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

