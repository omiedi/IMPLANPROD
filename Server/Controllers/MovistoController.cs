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
using System.Data;
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
        private readonly IMoviStockService _moviStockService;

        public MovistoController(DataContext context, ILogger<MasterController> logger, IUsuarioFlexoftService UsuarioFlexoftService, ICodiempresa Codiempresa, IMoviStockService moviStockService)
        {
            _context = context;
            _logger = logger;
            _UsuarioFlexoftService = UsuarioFlexoftService;
            _Codiempresa = Codiempresa;
            _moviStockService = moviStockService;
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

        // ===================================================================
        // PRODUCCIÓN PROPIA
        // ===================================================================
        // Estos endpoints alimentan la pantalla /propia-produccion.
        // En VB6 la producción propia se identificaba en MOVISTO con:
        //   - CodiMovi = 'pf' (producto fabricado / ingreso de producción)
        //   - CodiMovi = 'cf' (consumo de fabricación / salida de componentes)
        // El documento DoPriLar de las 'pf' tiene formato "OF-XXXX" (orden de fabricación).
        // El campo DoSecLar de las 'cf' vincula el consumo con la producción: "pf-XXXX".
        // ===================================================================

        /// <summary>
        /// Obtiene la lista de producciones propias (MOVISTO con CodiMovi = 'pf' y
        /// DoPriLar LIKE 'OF-%'), ordenadas por número de documento descendente.
        /// Incluye registros anulados (CantIngre = 0) para que queden visibles en la lista.
        /// Cada registro representa un producto terminado fabricado e ingresado a stock.
        /// Soporta paginación y filtros:
        ///   - Filter: filtro combinado por código (Cod_Exte) y descripción (Descrip) con '+'.
        ///   - FilterDocumento: filtro por documento (DoPriLar).
        /// </summary>
        [HttpGet("propia-produccion")]
        public async Task<ActionResult<List<PropiaProduccionDTO>>> GetPropiaProduccionAsync(
            [FromQuery] PaginationDTO pagination,
            [FromQuery] string? filterDocumento)
        {
            try
            {
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                // Condición SQL base:
                // SELECT * FROM MOVISTO WHERE codimovi = 'pf' AND doprilar LIKE 'OF-%'
                // Se quita la condición CANTINGRE > 0 para incluir las anuladas
                // (las anuladas tienen CantIngre = 0 pero siguen siendo 'pf' con 'OF-XXXX').
                var queryable = _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.CodiMovi == "pf"
                             && m.DoPriLar != null && m.DoPriLar.StartsWith("OF-"));

                // Filtro combinado por código y/o descripción (mismo patrón que /Stock-index).
                // Cada término separado por '+' actúa como un AND.
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var terms = pagination.Filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(m =>
                            EF.Functions.Like((m.Cod_Exte ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((m.Descrip ?? string.Empty).ToLower(), like));
                    }
                }

                // Filtro por documento (DoPriLar) — coincide parcial (LIKE %...%).
                if (!string.IsNullOrWhiteSpace(filterDocumento))
                {
                    var likeDoc = $"%{filterDocumento.Trim().ToLower()}%";
                    queryable = queryable.Where(m =>
                        EF.Functions.Like((m.DoPriLar ?? string.Empty).ToLower(), likeDoc));
                }

                // Traer todos los registros filtrados para ordenar por número de documento
                // (no se puede ordenar por SUBSTRING+TRY_CAST en EF Core de forma eficiente).
                var todos = await queryable
                    .Select(m => new PropiaProduccionDTO
                    {
                        FechMov = m.FechMov,
                        DoPriLar = m.DoPriLar,
                        Cod_Exte = m.Cod_Exte,
                        Descrip = m.Descrip,
                        CantIngre = m.CantIngre
                    })
                    .ToListAsync();

                // Ordenar por número de documento descendente (extraído de DoPriLar "OF-XXXX").
                // Los registros anulados tienen CantIngre = 0 pero mantienen su DoPriLar,
                // por eso quedan visibles en la lista.
                var ordenados = todos
                    .OrderByDescending(p =>
                    {
                        if (p.DoPriLar != null && p.DoPriLar.StartsWith("OF-")
                            && int.TryParse(p.DoPriLar.Substring(3), out int nro))
                        {
                            return nro;
                        }
                        return 0;
                    })
                    .ToList();

                // Aplicar paginación en memoria después de ordenar por número de documento.
                var resultado = ordenados
                    .Skip((pagination.Page - 1) * pagination.RecordsNumber)
                    .Take(pagination.RecordsNumber)
                    .ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener producción propia");
                return StatusCode(500, $"Error al obtener producción propia: {ex.Message}");
            }
        }

        /// <summary>
        /// Calcula el total de páginas de producciones propias según los mismos filtros
        /// del endpoint de listado (Filter + FilterDocumento).
        /// </summary>
        [HttpGet("propia-produccion/totalPages")]
        public async Task<ActionResult<int>> GetPropiaProduccionPagesAsync(
            [FromQuery] PaginationDTO pagination,
            [FromQuery] string? filterDocumento)
        {
            try
            {
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                // Misma condición base que el endpoint de listado: incluye anulados (sin CantIngre > 0).
                var queryable = _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.CodiMovi == "pf"
                             && m.DoPriLar != null && m.DoPriLar.StartsWith("OF-"));

                // Replicar exactamente la lógica de filtros del endpoint de listado
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var terms = pagination.Filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(m =>
                            EF.Functions.Like((m.Cod_Exte ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((m.Descrip ?? string.Empty).ToLower(), like));
                    }
                }

                if (!string.IsNullOrWhiteSpace(filterDocumento))
                {
                    var likeDoc = $"%{filterDocumento.Trim().ToLower()}%";
                    queryable = queryable.Where(m =>
                        EF.Functions.Like((m.DoPriLar ?? string.Empty).ToLower(), likeDoc));
                }

                var count = await queryable.CountAsync();
                var totalPages = count == 0 ? 1 : (int)Math.Ceiling(count / (double)pagination.RecordsNumber);

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular total de páginas de producción propia");
                return StatusCode(500, $"Error al calcular total de páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene TODAS las producciones propias sin paginar.
        /// Se usa para el botón Imprimir/Exp.Excel de la pantalla /propia-produccion,
        /// que necesita todos los registros filtrados (sin paginación) para el reporte.
        /// </summary>
        [HttpGet("propia-produccion/todos")]
        public async Task<ActionResult<List<PropiaProduccionDTO>>> GetPropiaProduccionTodosAsync(
            [FromQuery] string? filter,
            [FromQuery] string? filterDocumento)
        {
            try
            {
                // Misma condición base que el endpoint de listado: incluye anulados.
                var queryable = _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.CodiMovi == "pf"
                             && m.DoPriLar != null && m.DoPriLar.StartsWith("OF-"));

                // Replicar la misma lógica de filtros del endpoint paginado
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    var terms = filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(m =>
                            EF.Functions.Like((m.Cod_Exte ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((m.Descrip ?? string.Empty).ToLower(), like));
                    }
                }

                if (!string.IsNullOrWhiteSpace(filterDocumento))
                {
                    var likeDoc = $"%{filterDocumento.Trim().ToLower()}%";
                    queryable = queryable.Where(m =>
                        EF.Functions.Like((m.DoPriLar ?? string.Empty).ToLower(), likeDoc));
                }

                // Traer todos los registros y ordenar por número de documento descendente
                // (igual que el endpoint de listado).
                var todos = await queryable
                    .Select(m => new PropiaProduccionDTO
                    {
                        FechMov = m.FechMov,
                        DoPriLar = m.DoPriLar,
                        Cod_Exte = m.Cod_Exte,
                        Descrip = m.Descrip,
                        CantIngre = m.CantIngre
                    })
                    .ToListAsync();

                var resultado = todos
                    .OrderByDescending(p =>
                    {
                        if (p.DoPriLar != null && p.DoPriLar.StartsWith("OF-")
                            && int.TryParse(p.DoPriLar.Substring(3), out int nro))
                        {
                            return nro;
                        }
                        return 0;
                    })
                    .ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener todas las producciones propias");
                return StatusCode(500, $"Error al obtener producciones propias: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene los consumos de componentes de una producción propia específica.
        /// Busca en MOVISTO los registros con CantSalid > 0, CodiMovi = 'cf' y
        /// DoSecLar igual al identificador de la producción (ej: 'pf-5604').
        /// Cada registro representa un componente consumido para fabricar el producto terminado.
        /// </summary>
        /// <param name="doseclar">Identificador de la producción en DoSecLar (ej: 'pf-5604')</param>
        [HttpGet("consumos-produccion/{doseclar}")]
        public async Task<ActionResult<List<ConsumoProduccionDTO>>> GetConsumosProduccionAsync(string doseclar)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(doseclar))
                {
                    return BadRequest("El identificador de la producción es obligatorio");
                }

                // Condición SQL equivalente:
                // SELECT * FROM MOVISTO WHERE cantsalid > 0
                //   AND doseclar = 'pf-5604' AND codimovi = 'cf' ORDER BY FECHMOV DESC
                var query = _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.CantSalid.HasValue && m.CantSalid.Value > 0
                             && m.DoSecLar == doseclar
                             && m.CodiMovi == "cf");

                var resultado = await query
                    .OrderByDescending(m => m.FechMov)
                    .Select(m => new ConsumoProduccionDTO
                    {
                        FechMov = m.FechMov,
                        Cod_Exte = m.Cod_Exte,
                        Descrip = m.Descrip,
                        CantSalid = m.CantSalid
                    })
                    .ToListAsync();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener consumos de producción {DoSecLar}", doseclar);
                return StatusCode(500, $"Error al obtener consumos de producción: {ex.Message}");
            }
        }

        // ===================================================================
        // GENERAR PRODUCCIÓN PROPIA (ALTA + CONSUMO)
        // ===================================================================
        // A diferencia del circuito de órdenes de fabricación, este endpoint
        // genera directamente los movimientos de stock en MOVISTO:
        //   1. Un movimiento "PF" (producto fabricado) con CantIngre > 0
        //      que ingresa el producto terminado al depósito de entrada.
        //   2. N movimientos "CF" (consumo de fabricación) con CantSalid > 0
        //      que descuentan los componentes del depósito de consumo.
        //
        // La numeración del documento (OF-XXXX) se obtiene con:
        //   SELECT MAX(TRY_CAST(SUBSTRING(DOPRILAR,4,LEN(DOPRILAR)) AS INT))
        //   FROM MOVISTO WHERE DOPRILAR LIKE 'OF-%' AND DOPRILAR NOT LIKE '%-1'
        // y se incrementa en 1.
        //
        // DoPriLar = "OF-XXXX" (documento primario del ingreso)
        // DoSecLar = "PF-XXXX" (documento secundario, vincula con los consumos CF)
        // ===================================================================

        /// <summary>
        /// Genera el alta y consumo de una producción propia en una sola transacción.
        /// Equivalente al cierre de una orden de fabricación, pero sin orden previa.
        /// </summary>
        [HttpPost("generar-propia-produccion")]
        public async Task<ActionResult> GenerarPropiaProduccionAsync([FromBody] GenerarPropiaProduccionDTO dto)
        {
            // Validaciones básicas
            if (dto == null)
            {
                return BadRequest("No se recibieron datos.");
            }

            if (dto.Cod_Inte <= 0)
            {
                return BadRequest("Debe seleccionar un producto a fabricar.");
            }

            if (dto.Cantidad == null || dto.Cantidad <= 0)
            {
                return BadRequest("La cantidad a fabricar debe ser mayor a 0.");
            }

            if (dto.DepositoEntrada <= 0)
            {
                return BadRequest("Debe seleccionar un depósito de entrada.");
            }

            if (dto.DepositoConsumo <= 0)
            {
                return BadRequest("Debe seleccionar un depósito de consumo.");
            }

            if (dto.Componentes == null || !dto.Componentes.Any())
            {
                return BadRequest("No hay componentes para consumir. Debe existir una fórmula.");
            }

            // Iniciar transacción: todo o nada. Si falla el ingreso o cualquier consumo, se revierte todo.
            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);

            try
            {
                // ---------------------------------------------------------------
                // 1. Calcular NRO& (numeración del documento OF-XXXX)
                // ---------------------------------------------------------------
                // SQL equivalente:
                //   SELECT MAX(TRY_CAST(SUBSTRING(DOPRILAR,4,LEN(DOPRILAR)) AS INT))
                //   FROM MOVISTO WHERE DOPRILAR LIKE 'OF-%' AND DOPRILAR NOT LIKE '%-1'
                //
                // En EF no podemos usar TRY_CAST directamente, pero podemos
                // traer los candidatos y parsear en memoria (la cantidad de
                // registros con DOPRILAR LIKE 'OF-%' es acotada).
                // ---------------------------------------------------------------
                var candidatos = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.DoPriLar != null
                             && m.DoPriLar.StartsWith("OF-")
                             && !m.DoPriLar.EndsWith("-1"))
                    .Select(m => m.DoPriLar!)
                    .ToListAsync();

                int nro = 1;
                foreach (var dop in candidatos)
                {
                    // Quitar el prefijo "OF-" y parsear el resto como entero
                    var resto = dop.Substring(3);
                    if (int.TryParse(resto, out int valor) && valor >= nro)
                    {
                        nro = valor + 1;
                    }
                }

                string dopri = $"OF-{nro}";
                string dosec = $"PF-{nro}";

                // Referencia: si está vacía, usar "Informe de Produccion" (regla VB6)
                string refe = string.IsNullOrWhiteSpace(dto.Referencia)
                    ? "Informe de Produccion"
                    : dto.Referencia.Trim();

                // ---------------------------------------------------------------
                // 2. Registrar el ALTA del producto terminado (movimiento PF)
                // ---------------------------------------------------------------
                // Equivalente VB6:
                //   CMOV$ = "PF"
                //   DOPRI$ = "OF-" + NRO
                //   DOSEC$ = "PF-" + NRO
                //   Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$,
                //                 DOSEC$, DOSEC$, REFE$, NUORIT%=0, VUNI#=0,
                //                 MONEII%=1, NC%=0, DEPX%=DepositoEntrada, CANTI)
                // ---------------------------------------------------------------
                var movimientoAlta = new MoviStockDTO
                {
                    Fecha = dto.Fecha,
                    CodigoInterno = dto.Cod_Inte,
                    Lote = string.Empty, // LOTE$ = ""
                    CodigoMovimiento = "PF",
                    DocumentoPrimarioLargo = dopri,
                    DocumentoPrimarioCorto = dopri,
                    DocumentoSecundarioLargo = dosec,
                    DocumentoSecundarioCorto = dosec,
                    Referencia = refe,
                    NumeroOrdenItem = 0, // NUORIT% = 0
                    ValorUnitario = 0,   // VUNI# = 0
                    Moneda = 1,          // MONEII% = 1
                    NumeroCliente = 0,   // NC% = 0
                    Deposito = dto.DepositoEntrada,
                    Cantidad = dto.Cantidad.Value // Positivo → CantIngre > 0
                };

                var resultadoAlta = await _moviStockService.RegistrarMovimientoAsync(movimientoAlta, User);
                if (resultadoAlta is BadRequestObjectResult badAlta)
                {
                    await transaction.RollbackAsync();
                    var msg = badAlta.Value?.ToString() ?? "Error al registrar el alta del producto terminado.";
                    return BadRequest(msg);
                }

                // Capturar el ID del movimiento PF generado para vincular los consumos CF
                // (mismo patrón que el cierre de OF, usando IdenMovi).
                int? idMovimientoPF = null;
                if (resultadoAlta is OkObjectResult okAlta && okAlta.Value != null)
                {
                    var resultDict = okAlta.Value as dynamic;
                    idMovimientoPF = resultDict?.Id;
                }

                _logger.LogInformation($"Alta de producción propia generada: {dopri} (ID={idMovimientoPF})");

                // ---------------------------------------------------------------
                // 3. Registrar el CONSUMO de cada componente (movimientos CF)
                // ---------------------------------------------------------------
                // Equivalente VB6:
                //   For each componente:
                //     CMOV$ = "CF"
                //     REFE$ = "CONSUMO S/DESP. " + CODCON$  (CODCON$ = código externo del PT)
                //     DEPX% = DepositoConsumo
                //     NUORIT% = NUORIT% + 1
                //     Cantidad = -CantidadAUtilizar (negativo → CantSalid > 0)
                // ---------------------------------------------------------------
                // Obtener el código externo del producto terminado para la referencia
                var productoTerminado = await _context.masters
                    .AsNoTracking()
                    .Where(m => m.Codint == dto.Cod_Inte)
                    .Select(m => m.Codigo)
                    .FirstOrDefaultAsync();
                string codCon = productoTerminado ?? string.Empty;

                int nuorit = 0;
                foreach (var componente in dto.Componentes)
                {
                    if (componente == null) continue;
                    if (componente.CodIElem <= 0) continue;
                    if (componente.CantidadAUtilizar <= 0) continue;

                    nuorit++; // NUORIT% = NUORIT% + 1

                    // Referencia del consumo: "CONSUMO S/DESP. " + código externo del PT
                    string refeConsumo = $"CONSUMO S/DESP. {codCon}";

                    var movimientoConsumo = new MoviStockDTO
                    {
                        Fecha = dto.Fecha,
                        CodigoInterno = componente.CodIElem,
                        Lote = string.Empty,
                        CodigoMovimiento = "CF",
                        DocumentoPrimarioLargo = dopri,
                        DocumentoPrimarioCorto = dopri,
                        DocumentoSecundarioLargo = dosec,
                        DocumentoSecundarioCorto = dosec,
                        Referencia = refeConsumo,
                        NumeroOrdenItem = nuorit,
                        ValorUnitario = 0,
                        Moneda = 1,
                        NumeroCliente = 0,
                        Deposito = dto.DepositoConsumo,
                        Cantidad = -Math.Abs(componente.CantidadAUtilizar), // Negativo → CantSalid > 0
                        // Vinculación con el movimiento PF (mismo patrón que el cierre de OF)
                        IdentificadorMovimiento = idMovimientoPF.HasValue ? idMovimientoPF.Value.ToString() : null
                    };

                    var resultadoConsumo = await _moviStockService.RegistrarMovimientoAsync(movimientoConsumo, User);
                    if (resultadoConsumo is BadRequestObjectResult badConsumo)
                    {
                        await transaction.RollbackAsync();
                        var msg = badConsumo.Value?.ToString() ?? $"Error al consumir el componente {componente.CodXElem}.";
                        return BadRequest(msg);
                    }

                    _logger.LogInformation($"Consumo registrado: {componente.CodXElem} x {componente.CantidadAUtilizar} ({dopri})");
                }

                // Confirmar la transacción: tanto el alta como todos los consumos quedan grabados
                await transaction.CommitAsync();

                _logger.LogInformation($"Producción propia {dopri} generada exitosamente. Alta + {nuorit} consumos.");

                return Ok(new
                {
                    Documento = dopri,
                    DocumentoSecundario = dosec,
                    Cantidad = dto.Cantidad.Value,
                    Consumos = nuorit
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al generar producción propia");
                return StatusCode(500, $"Error al generar producción propia: {ex.Message}");
            }
        }

        // ===================================================================
        // ANULAR PRODUCCIÓN PROPIA
        // ===================================================================
        // Anula un movimiento de ingreso (PF) y todos sus consumos (CF)
        // asociados, poniendo en 0 los campos:
        //   - CantIngre
        //   - CantSalid
        //   - PESOKG
        //   - METROS
        //
        // La vinculación entre el PF y sus CF se hace por DoPriLar (OF-XXXX),
        // ya que ambos comparten el mismo documento primario.
        // ===================================================================

        /// <summary>
        /// Anula una producción propia identificada por su documento (OF-XXXX).
        /// Pone en 0 las cantidades del movimiento de ingreso (PF) y de todos
        /// los consumos (CF) asociados.
        /// </summary>
        /// <param name="documento">Documento primario de la producción (ej: "OF-5604")</param>
        [HttpPut("anular-propia-produccion/{documento}")]
        public async Task<ActionResult> AnularPropiaProduccionAsync(string documento)
        {
            if (string.IsNullOrWhiteSpace(documento))
            {
                return BadRequest("El documento es obligatorio.");
            }

            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);

            try
            {
                // 1. Localizar el movimiento de ingreso (PF) por DoPriLar
                var movimientoPF = await _context.Movistos
                    .AsNoTracking()
                    .FirstOrDefaultAsync(m => m.DoPriLar == documento && m.CodiMovi == "PF");

                if (movimientoPF == null)
                {
                    return NotFound($"No se encontró el movimiento de producción con documento {documento}.");
                }

                if ((movimientoPF.CantIngre ?? 0) <= 0)
                {
                    return BadRequest("El movimiento ya está anulado (cantidad en 0).");
                }

                // 2. Anular el movimiento de ingreso (PF): CantIngre=0, CantSalid=0, PESOKG=null, METROS=null
                //    y ReferCor = "Anul. " + ReferCor (mismo patrón que el anular-cierre de OF).
                // Se usa una expresión lambda en SetProperty para concatenar el prefijo
                // con el valor actual de ReferCor en SQL (funciona para cualquier texto existente).
                var filasPF = await _context.Movistos
                    .Where(m => m.DoPriLar == documento && m.CodiMovi == "PF")
                    .ExecuteUpdateAsync(s => s
                        .SetProperty(m => m.ReferCor, m => "Anul. " + (m.ReferCor ?? ""))
                        .SetProperty(m => m.CantIngre, 0m)
                        .SetProperty(m => m.CantSalid, 0m)
                        .SetProperty(m => m.PESOKG, (decimal?)null)
                        .SetProperty(m => m.METROS, (decimal?)null));

                if (filasPF == 0)
                {
                    await transaction.RollbackAsync();
                    return Conflict("No se pudo anular el movimiento de ingreso.");
                }

                _logger.LogInformation($"Movimiento PF {documento} anulado (cantidades en 0).");

                // 3. Anular todos los consumos (CF) asociados al mismo documento
                // Los consumos CF comparten el DoPriLar con el movimiento PF.
                // También se antepone "Anul. " al ReferCor de cada consumo (mismo patrón que el PF).
                var filasCF = await _context.Movistos
                    .Where(m => m.DoPriLar == documento && m.CodiMovi == "CF")
                    .ExecuteUpdateAsync(s => s
                        .SetProperty(m => m.ReferCor, m => "Anul. " + (m.ReferCor ?? ""))
                        .SetProperty(m => m.CantIngre, 0m)
                        .SetProperty(m => m.CantSalid, 0m)
                        .SetProperty(m => m.PESOKG, (decimal?)null)
                        .SetProperty(m => m.METROS, (decimal?)null));

                _logger.LogInformation($"Anulados {filasCF} movimientos CF asociados a {documento}.");

                await transaction.CommitAsync();

                return Ok(new
                {
                    Documento = documento,
                    MovimientoPFAnulado = filasPF,
                    ConsumosCFAnulados = filasCF
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al anular producción propia {Documento}", documento);
                return StatusCode(500, $"Error al anular producción propia: {ex.Message}");
            }
        }
    }
}
