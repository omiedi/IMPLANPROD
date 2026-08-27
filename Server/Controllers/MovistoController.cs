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
            [FromQuery] string fechMov,
            [FromQuery] string? codiMovi = null)
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
                var query = _context.Movistos
                    .Where(m => m.DoPriLar == doPriLar 
                             && m.Nume_Clie == numeCli 
                             && m.FechMov.HasValue 
                             && m.FechMov.Value.Date == fecha.Date);

                // Si se especifica CodiMovi, filtrar también por ese código
                if (!string.IsNullOrWhiteSpace(codiMovi))
                {
                    query = query.Where(m => m.CodiMovi == codiMovi);
                }

                var movimientos = await query
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

        /// <summary>
        /// Verifica si hay materiales transferidos para una orden de fabricación
        /// Replica la validación VB6: WHERE DOSECCOR = NORFA$ AND CODIMOVI = 'TR'
        /// Verifica si (CANTINGRE - CANTSALID) > 0.5 para algún material
        /// </summary>
        /// <param name="numeroOrdenFabricacion">Número de orden de fabricación (NORFA$)</param>
        /// <returns>
        /// true si hay materiales transferidos con saldo pendiente (> 0.5)
        /// false si no hay materiales transferidos pendientes
        /// </returns>
        [HttpGet("verificar-materiales-transferidos/{numeroOrdenFabricacion}")]
        public async Task<ActionResult<bool>> VerificarMaterialesTransferidosAsync(string numeroOrdenFabricacion)
        {
            try
            {
                Console.WriteLine($"[MovistoController] Verificando materiales transferidos para O.F. {numeroOrdenFabricacion}");

                // Buscar en MOVISTO donde DOSECCOR = NORFA$ y CODIMOVI = 'TR'
                var movimientosTransferidos = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.DoSecCor == numeroOrdenFabricacion && m.CodiMovi == "TR")
                    .ToListAsync();

                Console.WriteLine($"[MovistoController] Se encontraron {movimientosTransferidos.Count} movimientos de transferencia");

                // Verificar si algún movimiento tiene saldo pendiente > 0.5
                // CAN = CANTINGRE - CANTSALID
                bool tieneMaterialesPendientes = movimientosTransferidos.Any(m =>
                {
                    var cantIngre = m.CantIngre ?? 0m;
                    var cantSalid = m.CantSalid ?? 0m;
                    var saldo = cantIngre - cantSalid;
                    return saldo > 0.5m;
                });

                Console.WriteLine($"[MovistoController] Resultado: {(tieneMaterialesPendientes ? "TIENE" : "NO TIENE")} materiales transferidos pendientes");

                return Ok(tieneMaterialesPendientes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[MovistoController] Error al verificar materiales transferidos: {ex.Message}");
                return StatusCode(500, $"Error al verificar materiales transferidos: {ex.Message}");
            }
        }

        /// <summary>
        /// Lista las transferencias realizadas (CodiMovi = 'TR') para la pantalla de re-impresiones.
        /// </summary>
        [HttpGet("transferencias")]
        public async Task<ActionResult<List<TransferenciaResumenDTO>>> GetTransferencias(
            [FromQuery] DateTime? desde,
            [FromQuery] DateTime? hasta)
        {
            try
            {
                if (!desde.HasValue || !hasta.HasValue)
                {
                    return BadRequest("Debe indicar fecha desde y fecha hasta");
                }

                var query = _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.CodiMovi == "TR" && !string.IsNullOrEmpty(m.DoPriCor));

                query = query.Where(m => m.FeReMovi.HasValue
                                        && m.FeReMovi.Value.Date >= desde.Value.Date
                                        && m.FeReMovi.Value.Date <= hasta.Value.Date);

                var movimientos = await query
                    .OrderByDescending(m => m.FeReMovi)
                    .ToListAsync();

                var numerosUsuarios = movimientos
                    .Where(m => m.NumUsuar.HasValue)
                    .Select(m => m.NumUsuar!.Value)
                    .Distinct()
                    .ToList();

                var usuarios = await _context.Usuarios
                    .AsNoTracking()
                    .Where(u => numerosUsuarios.Contains(u.IdFlexoft))
                    .ToDictionaryAsync(u => u.IdFlexoft, u => u.NombreUsuario ?? string.Empty);

                var depositos = await _context.Tadeposi
                    .AsNoTracking()
                    .ToDictionaryAsync(d => d.CodigoDepo, d => d.Descripcion ?? string.Empty);

                var resultado = movimientos
                    .GroupBy(m => m.DoPriCor)
                    .Select(g =>
                    {
                        var primer = g.First();
                        var productosDistintos = g.Select(m => m.Cod_Inte).Distinct().Count();
                        var movOrigen = g.FirstOrDefault(m => m.CantSalid.HasValue && m.CantSalid.Value > 0);
                        var movDestino = g.FirstOrDefault(m => m.CantIngre.HasValue && m.CantIngre.Value > 0);
                        var esMultiple = productosDistintos > 1;

                        var depoOrigen = movOrigen?.DepoMovi ?? 0;
                        var depoDestino = movDestino?.DepoMovi ?? 0;

                        var idUsuario = movOrigen?.NumUsuar ?? movDestino?.NumUsuar;
                        var usuarioNombre = idUsuario.HasValue && usuarios.ContainsKey(idUsuario.Value)
                            ? usuarios[idUsuario.Value]
                            : (idUsuario?.ToString() ?? string.Empty);

                        return new TransferenciaResumenDTO
                        {
                            Numero = g.Key ?? string.Empty,
                            DoSecCor = primer.DoSecCor ?? string.Empty,
                            Fecha = primer.FeReMovi ?? DateTime.MinValue,
                            EsMultiple = esMultiple,
                            CodigoExterno = esMultiple ? null : (movOrigen?.Cod_Exte ?? movDestino?.Cod_Exte),
                            Descripcion = esMultiple ? null : (movOrigen?.Descrip ?? movDestino?.Descrip),
                            DepositoOrigen = depoOrigen,
                            DepositoOrigenNombre = depositos.ContainsKey(depoOrigen) ? depositos[depoOrigen] : depoOrigen.ToString(),
                            DepositoDestino = depoDestino,
                            DepositoDestinoNombre = depositos.ContainsKey(depoDestino) ? depositos[depoDestino] : depoDestino.ToString(),
                            UsuarioNombre = usuarioNombre
                        };
                    })
                    .OrderByDescending(t => t.Fecha)
                    .ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener transferencias: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el detalle de una transferencia para re-imprimir el vale.
        /// </summary>
        [HttpGet("transferencia")]
        public async Task<ActionResult<ValeTransferenciaDTO>> GetTransferencia([FromQuery] string numero)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(numero))
                {
                    return BadRequest("El número de transferencia es obligatorio");
                }

                var movimientos = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.CodiMovi == "TR" && m.DoPriCor == numero)
                    .ToListAsync();

                if (!movimientos.Any())
                {
                    return NotFound("Transferencia no encontrada");
                }

                var primer = movimientos.First();
                var productosDistintos = movimientos
                    .Select(m => m.Cod_Inte)
                    .Distinct()
                    .ToList();

                var umedidas = await _context.masters
                    .AsNoTracking()
                    .Where(m => productosDistintos.Contains(m.Codint))
                    .ToDictionaryAsync(m => m.Codint, m => m.Umedida ?? string.Empty);

                // Se usa FirstOrDefault (no First) porque una transferencia puede tener,
                // por datos históricos o inconsistencias, movimientos sin salida o sin ingreso
                // registrado; en ese caso no debe explotar con excepción sino degradar
                // mostrando la información disponible.
                var movimientoSalida = movimientos.FirstOrDefault(m => m.CantSalid.HasValue && m.CantSalid.Value > 0);
                var movimientoIngreso = movimientos.FirstOrDefault(m => m.CantIngre.HasValue && m.CantIngre.Value > 0);

                var depositoOrigen = movimientoSalida?.DepoMovi ?? 0;
                var depositoDestino = movimientoIngreso?.DepoMovi ?? 0;

                var origen = await _context.Tadeposi
                    .AsNoTracking()
                    .Where(d => d.CodigoDepo == depositoOrigen)
                    .Select(d => d.Descripcion)
                    .FirstOrDefaultAsync();

                var destino = await _context.Tadeposi
                    .AsNoTracking()
                    .Where(d => d.CodigoDepo == depositoDestino)
                    .Select(d => d.Descripcion)
                    .FirstOrDefaultAsync();

                // Los items del vale se construyen a partir de los movimientos de salida.
                // Si por algún motivo no hay movimientos de salida (ej. datos legacy),
                // se usan los de ingreso para no dejar el vale sin ítems.
                var movimientosParaItems = movimientos.Any(m => m.CantSalid.HasValue && m.CantSalid.Value > 0)
                    ? movimientos.Where(m => m.CantSalid.HasValue && m.CantSalid.Value > 0)
                    : movimientos.Where(m => m.CantIngre.HasValue && m.CantIngre.Value > 0);

                var vale = new ValeTransferenciaDTO
                {
                    Numero = numero,
                    DoSecCor = primer.DoSecCor ?? string.Empty,
                    Fecha = primer.FeReMovi ?? DateTime.MinValue,
                    DepositoOrigen = depositoOrigen,
                    DepositoOrigenNombre = origen ?? string.Empty,
                    DepositoDestino = depositoDestino,
                    DepositoDestinoNombre = destino ?? string.Empty,
                    Motivo = primer.ReferCor ?? string.Empty,
                    Items = movimientosParaItems
                        .Select(m => new ValeTransferenciaItemDTO
                        {
                            CodigoExterno = m.Cod_Exte ?? string.Empty,
                            Descripcion = m.Descrip ?? string.Empty,
                            UnidadMedida = umedidas.GetValueOrDefault(m.Cod_Inte ?? 0),
                            Cantidad = m.CantSalid ?? m.CantIngre ?? 0
                        })
                        .ToList()
                };

                return Ok(vale);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la transferencia {Numero}", numero);
                return StatusCode(500, $"Error al obtener la transferencia: {ex.Message}");
            }
        }

        /// <summary>
        /// Anula una transferencia eliminando sus movimientos de entrada y salida.
        /// </summary>
        [HttpDelete("transferencia")]
        public async Task<ActionResult> AnularTransferencia([FromQuery] string numero)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(numero))
                {
                    return BadRequest("El número de transferencia es obligatorio");
                }

                var movimientos = await _context.Movistos
                    .Where(m => m.CodiMovi == "TR" && m.DoPriCor == numero)
                    .ToListAsync();

                if (!movimientos.Any())
                {
                    return NotFound("No se encontró la transferencia");
                }

                _context.Movistos.RemoveRange(movimientos);
                await _context.SaveChangesAsync();

                return Ok(new { Message = "Transferencia anulada correctamente" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al anular la transferencia: {ex.Message}");
            }
        }
    }
}
