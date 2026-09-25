using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace MantenimientoImp.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class BolRecepController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IDerechosUsuarioService _derechosUsuarioService;
        private readonly IUsuarioFlexoftService _usuarioFlexoftService;

        public BolRecepController(DataContext context, IDerechosUsuarioService derechosUsuarioService, IUsuarioFlexoftService usuarioFlexoftService)
        {
            _context = context;
            _derechosUsuarioService = derechosUsuarioService;
            _usuarioFlexoftService = usuarioFlexoftService;
        }

        [HttpGet("validar-remito")]
        public async Task<ActionResult> ValidarRemitoAsync([FromQuery] string numeroRemito, [FromQuery] int numeroProveedor)
        {
            if (string.IsNullOrWhiteSpace(numeroRemito) || numeroProveedor <= 0)
            {
                return Ok(new { Valid = true });
            }

            var remitoExistente = await _context.Bolreceps
                .AsNoTracking()
                .AnyAsync(b => b.NumRemProv == numeroRemito && b.NumProv == numeroProveedor);

            if (remitoExistente)
            {
                return BadRequest($"Ya existe una recepción con el remito {numeroRemito} para el proveedor {numeroProveedor}");
            }

            return Ok(new { Valid = true });
        }

        [HttpGet("pendientes-por-producto/{codint:int}")]
        public async Task<ActionResult<List<StockRecepcionPendienteDTO>>> GetPendientesPorProducto(int codint)
        {
            try
            {
                var query = from b in _context.Bolreceps.AsNoTracking()
                            join p in _context.Proveedores.AsNoTracking()
                                on (b.NumProv ?? 0) equals p.Nume_Cli into provJoin
                            from p in provJoin.DefaultIfEmpty()
                            where b.Status == 0 && b.CodInte == codint
                            orderby b.NumeBoRe
                            select new StockRecepcionPendienteDTO
                            {
                                Numebore = (int)(b.NumeBoRe ?? 0m),
                                NumProv = b.NumProv ?? 0,
                                Proveedor = p != null ? (p.Raso_Cli ?? string.Empty) : string.Empty,
                                CantPendiente = (b.CantNom ?? 0m) - (b.CanToApro ?? 0m),
                                RemitoProveedor = b.NumRemProv ?? string.Empty,
                                FechaRecepcion = b.FechRecep ?? DateTime.MinValue
                            };

                var pendientes = await query.ToListAsync();

                pendientes = pendientes
                    .Where(x => x.CantPendiente > 0m)
                    .ToList();

                return Ok(pendientes);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener recepciones pendientes: {ex.Message}");
            }
        }

        [HttpPost("recibir")]
        public async Task<ActionResult> RecibirMateriales([FromBody] List<RecepcionMaterialDTO> materiales)
        {
            if (!materiales.Any())
                return BadRequest("No hay materiales para recibir");

            try
            {
                // Validar que el número de remito no esté repetido para el mismo proveedor
                var primerMaterial = materiales.First();
                var remitoExistente = await _context.Bolreceps
                    .AnyAsync(b => b.NumRemProv == primerMaterial.NumeroRemito
                                 && b.NumProv == primerMaterial.NumeroProveedor);

                if (remitoExistente)
                {
                    return BadRequest($"Ya existe una recepción con el remito {primerMaterial.NumeroRemito} para el proveedor {primerMaterial.NumeroProveedor}");
                }

                using var transaction = await _context.Database.BeginTransactionAsync();

                try
                {
                    // Obtener el último número de boleta
                    var ultimoBoleta = await _context.Bolreceps
                        .OrderByDescending(b => b.NumeBoRe)
                        .Select(b => b.NumeBoRe ?? 0)
                        .FirstOrDefaultAsync();

                    // Convertir a entero para evitar decimales en el número de boleta
                    var numeroBoleta = Math.Floor(ultimoBoleta) + 1;
                    var fechaActual = DateTime.Now;

                    var bolReceps = new List<Bolrecep>();

                    // Para cada material, necesitamos un índice secuencial
                    var index = 1;
                    foreach (var material in materiales)
                    {
                        // 1. Actualizar la cantidad recibida en la orden de compra
                        var ordenCompraDetalle = await _context.Ocomdetas
                            .FirstOrDefaultAsync(o => o.NumeOcom == material.NumeroOrdenCompra
                                                 && o.StatOcom != -1 && o.StatOcom < 8 // para que busque en las activas
                                                 && o.NuOrdItem == material.NumeroOrdenItem);

                        if (ordenCompraDetalle != null)
                        {
                            ordenCompraDetalle.CantRec = (ordenCompraDetalle.CantRec ?? 0) + material.CantidadRecibida;

                            // Obtener la unidad de medida desde la tabla MASTER
                            var producto = await _context.masters
                                .Where(m => m.Codint == ordenCompraDetalle.CodInte)
                                .Select(m => m.Umedida)
                                .FirstOrDefaultAsync();

                            // Generar IdenLote: BR-NNNNNN-II (BR- + 6 dígitos recepción + - + 2 dígitos índice)
                            var numeroRecepcion = ((int)numeroBoleta).ToString();
                            var numeroIndice = index.ToString();
                            var idenLote = $"BR-{numeroRecepcion}-{numeroIndice}"; // Resultado: BR-019881-01

                            // 2. Crear el registro en BOLRECEP
                            var bolRecep = new Bolrecep
                            {
                                NumeBoRe = numeroBoleta,
                                NumOC = material.NumeroOrdenCompra,
                                NroItemOC = (short?)material.NumeroOrdenItem,
                                NumRemProv = material.NumeroRemito?.Length > 18 ? material.NumeroRemito.Substring(0, 18) : material.NumeroRemito,
                                CodiEmpr = ordenCompraDetalle.CodiEmpr,
                                Descrip = material.Descripcion?.Length > 64 ? material.Descripcion.Substring(0, 64) : material.Descripcion,
                                IdenLote = idenLote,
                                FechRecep = fechaActual.Date,
                                FEYHORecep = fechaActual,
                                CodInte = ordenCompraDetalle.CodInte,
                                CodExte = material.CodigoExterno?.Length > 24 ? material.CodigoExterno.Substring(0, 24) : material.CodigoExterno,
                                DescripItem = material.Descripcion?.Length > 64 ? material.Descripcion.Substring(0, 64) : material.Descripcion,
                                Unidad = string.IsNullOrEmpty(producto) ? "" : (producto.Length > 4 ? producto.Substring(0, 4) : producto),
                                CantNom = material.CantidadRecibida,
                                DepoEnt = (short?)material.Deposito,
                                NumProv = material.NumeroProveedor,
                                CanToApro = 0,
                                CanToRech = 0,
                                CanToFalta = 0,
                                Status = 0, // Pendiente
                                FechLibe = null,
                                Observa = material.Observa
                            };

                            _context.Bolreceps.Add(bolRecep);
                            bolReceps.Add(bolRecep);
                            index++; // Incrementar el índice para el siguiente material
                        }
                    }

                    await _context.SaveChangesAsync();
                    await transaction.CommitAsync();
                    
                    // Convertir a DTO para la respuesta
                    var response = bolReceps.Select(b => new BolrecepResponseDTO
                    {
                        Numebore = (int)(b.NumeBoRe ?? 0),
                        Idenlote = b.IdenLote ?? "",
                        Cod_Exte = b.CodExte ?? "",
                        CantReci = b.CantNom ?? 0
                    }).ToList();
                    
                    return Ok(response);

                }
                catch (Exception ex)
                {
                    await transaction.RollbackAsync();
                    throw;
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al recibir materiales: {ex.Message}");
            }
        }


        [HttpPost("recibir-libre")]
        public async Task<ActionResult> RecibirLibre([FromBody] List<RecepcionLibreDTO> items)
        {
            if (!items.Any())
                return BadRequest("No hay materiales para recibir");

            try
            {
                // Validar que el número de remito no esté repetido para el mismo proveedor
                var primerItem = items.First();
                var remitoExistente = await _context.Bolreceps
                    .AnyAsync(b => b.NumRemProv == primerItem.NumeroRemito
                                && b.NumProv == primerItem.NumeroProveedor);

                if (remitoExistente)
                {
                    return BadRequest($"Ya existe una recepción con el remito {primerItem.NumeroRemito} para el proveedor {primerItem.NumeroProveedor}");
                }

                using var transaction = await _context.Database.BeginTransactionAsync();

                try
                {
                    // Obtener el último número de boleta
                    var ultimoBoleta = await _context.Bolreceps
                        .OrderByDescending(b => b.NumeBoRe)
                        .Select(b => b.NumeBoRe ?? 0)
                        .FirstOrDefaultAsync();

                    // Convertir a entero para evitar decimales en el número de boleta
                    var numeroBoleta = Math.Floor(ultimoBoleta) + 1;
                    var fechaActual = DateTime.Now;

                    var bolReceps = new List<Bolrecep>();
                    var index = 1;

                    foreach (var item in items)
                    {
                        // Generar IdenLote: BR-NNNNNN-II (BR- + 6 dígitos recepción + - + 2 dígitos índice)
                        var numeroRecepcion = ((int)numeroBoleta).ToString();
                        var numeroIndice = index.ToString();
                        var idenLote = $"BR-{numeroRecepcion}-{numeroIndice}"; // Resultado: BR-019881-01

                        var bolRecep = new Bolrecep
                        {
                            NumeBoRe = numeroBoleta,
                            NumOC = 0, // Recepción libre
                            NroItemOC = (short?)index,
                            NumRemProv = item.NumeroRemito?.Length > 18 ? item.NumeroRemito.Substring(0, 18) : item.NumeroRemito,
                            CodiEmpr = 0, // Valor por defecto
                            Descrip = item.Descripcion?.Length > 64 ? item.Descripcion.Substring(0, 64) : item.Descripcion,
                            IdenLote = idenLote,
                            FechRecep = fechaActual.Date,
                            FEYHORecep = fechaActual,
                            CodInte = item.CodigoInterno,
                            CodExte = item.CodigoExterno?.Length > 24 ? item.CodigoExterno.Substring(0, 24) : item.CodigoExterno,
                            DescripItem = item.Descripcion?.Length > 64 ? item.Descripcion.Substring(0, 64) : item.Descripcion,
                            Unidad = item.Unidad?.Length > 4 ? item.Unidad.Substring(0, 4) : item.Unidad,
                            CantNom = item.CantidadNominal,
                            DepoEnt = (short?)item.Deposito,
                            NumProv = item.NumeroProveedor,
                            CanToApro = 0,
                            CanToRech = 0,
                            CanToFalta = 0,
                            Status = 0, // Pendiente
                            FechLibe = null,
                            Observa = item.Observa
                        };

                        _context.Bolreceps.Add(bolRecep);
                        bolReceps.Add(bolRecep);
                        index++;
                    }

                    await _context.SaveChangesAsync();
                    await transaction.CommitAsync();
                    
                    // Convertir a DTO para la respuesta
                    var response = bolReceps.Select(b => new BolrecepResponseDTO
                    {
                        Numebore = (int)(b.NumeBoRe ?? 0),
                        Idenlote = b.IdenLote ?? "",
                        Cod_Exte = b.CodExte ?? "",
                        CantReci = b.CantNom ?? 0
                    }).ToList();
                    
                    return Ok(response);
                }
                catch (Exception)
                {
                    await transaction.RollbackAsync();
                    throw;
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al recibir materiales: {ex.Message}");
            }
        }

        [HttpGet("pendientes-aprobacion")]
        public async Task<ActionResult<List<RecepcionPendienteAprobacionDTO>>> GetRecepcionesPendientesAprobacion()
        {
            try
            {
                // Primero obtenemos las recepciones con sus proveedores en una sola consulta
                var recepcionesPendientes = await _context.Bolreceps
                    .Where(b => b.Status == 0 && b.NumeBoRe != null)
                    .GroupBy(b => new
                    {
                        b.NumeBoRe,
                        b.NumOC,
                        b.FEYHORecep,
                        b.NumRemProv,
                        b.NumProv
                    })
                    .Select(g => new RecepcionPendienteAprobacionDTO
                    {
                        NumeroRecepcion = (int)(g.Key.NumeBoRe ?? 0),
                        NumeroOrdenCompra = g.Key.NumOC ?? 0,
                        FechaRecepcion = g.Key.FEYHORecep ?? DateTime.MinValue,
                        NumeroRemito = g.Key.NumRemProv,
                        NumeroProveedor = g.Key.NumProv ?? 0,
                        // El número de proveedor de BOLRECEP (NumProv) es Proved12.Nume_Cli,
                        // no la clave primaria Id. Buscar por Nume_Cli para obtener la Razón Social.
                        RazonSocial = _context.Proveedores
                            .Where(p => p.Nume_Cli == g.Key.NumProv)
                            .Select(p => p.Raso_Cli)
                            .FirstOrDefault() ?? string.Empty,
                        Observa = g.First().Observa
                    })
                    .OrderByDescending(r => r.FechaRecepcion)
                    .ToListAsync();

                return Ok(recepcionesPendientes);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// Devuelve el detalle por lote de las recepciones pendientes de aprobación.
        /// A diferencia de "pendientes-aprobacion" (que agrupa por recepción), aquí se
        /// devuelve una fila por cada registro de BOLRECEP: una misma recepción puede
        /// aparecer varias veces, una por cada lote (BOLRECEP.IdenLote).
        /// Se usa en el modal "Detalle por Lote" de /recepcion/selproveedor.
        /// GET /api/bolrecep/pendientes-aprobacion-detalle
        /// </summary>
        [HttpGet("pendientes-aprobacion-detalle")]
        public async Task<ActionResult<List<RecepcionPendienteAprobacionDetalleDTO>>> GetRecepcionesPendientesAprobacionDetalle()
        {
            try
            {
                var detalle = await _context.Bolreceps
                    .Where(b => b.Status == 0 && b.NumeBoRe != null)
                    .Select(b => new RecepcionPendienteAprobacionDetalleDTO
                    {
                        NumeroRecepcion = (int)(b.NumeBoRe ?? 0),
                        NumeroOrdenCompra = b.NumOC ?? 0,
                        FechaRecepcion = b.FEYHORecep ?? DateTime.MinValue,
                        NumeroRemito = b.NumRemProv,
                        NumeroProveedor = b.NumProv ?? 0,
                        // Mismo criterio que en GetRecepcionesPendientesAprobacion: el número de
                        // proveedor de BOLRECEP (NumProv) es Proved12.Nume_Cli, no la clave Id.
                        RazonSocial = _context.Proveedores
                            .Where(p => p.Nume_Cli == b.NumProv)
                            .Select(p => p.Raso_Cli)
                            .FirstOrDefault() ?? string.Empty,
                        IdenLote = b.IdenLote ?? string.Empty,
                        Codigo = b.CodExte ?? string.Empty,
                        Descripcion = b.DescripItem ?? string.Empty,
                        CantidadNominal = b.CantNom ?? 0
                    })
                    .OrderByDescending(r => r.FechaRecepcion)
                    .ThenBy(r => r.NumeroRecepcion)
                    .ThenBy(r => r.IdenLote)
                    .ToListAsync();

                return Ok(detalle);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // Devuelve los items de una recepción pendiente.
        // Importante: la grilla de pendientes agrupa por (NumeBoRe + NumOC), por lo que una misma
        // recepción puede aparecer en varias filas si tiene distintas órdenes de compra.
        // Por eso este endpoint acepta un parámetro opcional numeroOrdenCompra para traer SOLO
        // los items que correspondan a esa recepción Y a esa orden de compra seleccionada.
        // Si numeroOrdenCompra es null, se devuelven todos los items de la recepción (comportamiento anterior).
        [HttpGet("items/{numeroRecepcion}")]
        public async Task<ActionResult<List<RecepcionItemDTO>>> GetItemsRecepcion(int numeroRecepcion, [FromQuery] int? numeroOrdenCompra = null)
        {
            try
            {
                var numeroRecepcionDecimal = (decimal)numeroRecepcion;

                var items = await (
                        from b in _context.Bolreceps
                        join od in _context.Ocomdetas
                            on new { NumeOcom = b.NumOC, NuOrdItem = b.NroItemOC }
                            equals new { NumeOcom = od.NumeOcom, NuOrdItem = od.NuOrdItem }
                            into odJoin
                        from od in odJoin.DefaultIfEmpty()
                        join t in _context.Tadeposi
                            on (int?)(b.DepoEnt ?? 0) equals t.CodigoDepo
                            into tJoin
                        from t in tJoin.DefaultIfEmpty()
                        where b.NumeBoRe == numeroRecepcionDecimal && b.Status == 0
                              && (numeroOrdenCompra == null || b.NumOC == numeroOrdenCompra)
                        select new RecepcionItemDTO
                        {
                            CodigoInterno = b.CodInte ?? 0,
                            NumeBore = (int?)b.NumeBoRe,
                            NumeroOrdenCompra = b.NumOC,
                            FechaEntregaSolicitada = od != null ? od.FentreSol : null,
                            FechaRecepcion = b.FechRecep,
                            CodigoExterno = b.CodExte,
                            DescripcionItem = b.DescripItem,
                            CantidadNominal = b.CantNom ?? 0,
                            Unidad = b.Unidad,
                            CantidadRechazada = b.CanToRech,
                            CantidadFaltante = b.CanToFalta,
                            CodigoDeposito = (int?)(b.DepoEnt ?? 0) ?? 0,
                            NombreDeposito = t == null ? string.Empty : $"{t.CodigoDepo} - {t.Descripcion}",
                            Status = b.Status ?? 0,
                            Observa = b.Observa
                        })
                    .ToListAsync();

                Console.WriteLine($"[BolRecep] Items pendientes encontrados para recepción {numeroRecepcion}: {items.Count}");

                return Ok(items);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[BolRecep] Error al obtener items de recepción {numeroRecepcion}: {ex.Message}");
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// Devuelve TODOS los items de una recepción, sin filtrar por Status.
        /// A diferencia de "items/{numeroRecepcion}" (que filtra Status == 0),
        /// este endpoint devuelve los items tanto si la recepción está pendiente
        /// (Status 0) como si ya fue aprobada (Status 1).
        /// Se usa para imprimir la Nota de Recepción cuando la recepción ya fue
        /// aprobada, ya que el endpoint original no devuelve items en ese caso.
        /// GET /api/bolrecep/items-todos/{numeroRecepcion}
        /// </summary>
        [HttpGet("items-todos/{numeroRecepcion}")]
        public async Task<ActionResult<List<RecepcionItemDTO>>> GetItemsRecepcionTodos(int numeroRecepcion)
        {
            try
            {
                var numeroRecepcionDecimal = (decimal)numeroRecepcion;

                var items = await (
                        from b in _context.Bolreceps
                        join od in _context.Ocomdetas
                            on new { NumeOcom = b.NumOC, NuOrdItem = b.NroItemOC }
                            equals new { NumeOcom = od.NumeOcom, NuOrdItem = od.NuOrdItem }
                            into odJoin
                        from od in odJoin.DefaultIfEmpty()
                        join t in _context.Tadeposi
                            on (int?)(b.DepoEnt ?? 0) equals t.CodigoDepo
                            into tJoin
                        from t in tJoin.DefaultIfEmpty()
                        where b.NumeBoRe == numeroRecepcionDecimal
                        select new RecepcionItemDTO
                        {
                            CodigoInterno = b.CodInte ?? 0,
                            NumeBore = (int?)b.NumeBoRe,
                            NumeroOrdenCompra = b.NumOC,
                            FechaEntregaSolicitada = od != null ? od.FentreSol : null,
                            FechaRecepcion = b.FechRecep,
                            CodigoExterno = b.CodExte,
                            DescripcionItem = b.DescripItem,
                            CantidadNominal = b.CantNom ?? 0,
                            Unidad = b.Unidad,
                            CantidadRechazada = b.CanToRech,
                            CantidadFaltante = b.CanToFalta,
                            CodigoDeposito = (int?)(b.DepoEnt ?? 0) ?? 0,
                            NombreDeposito = t == null ? string.Empty : $"{t.CodigoDepo} - {t.Descripcion}",
                            Status = b.Status ?? 0,
                            Observa = b.Observa
                        })
                    .ToListAsync();

                Console.WriteLine($"[BolRecep] Items (todos) encontrados para recepción {numeroRecepcion}: {items.Count}");

                return Ok(items);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[BolRecep] Error al obtener items (todos) de recepción {numeroRecepcion}: {ex.Message}");
                return BadRequest(ex.Message);
            }
        }

        [HttpPut("items/{numeroRecepcion}")]//ver bien por que aca se aprueba la recepcion
        public async Task<ActionResult> UpdateItemsRecepcion(int numeroRecepcion, List<RecepcionItemDTO> items)
        {
            try
            {
                var numeroRecepcionDecimal = (decimal)numeroRecepcion;

                foreach (var item in items)
                {
                    var bolrecep = await _context.Bolreceps
                        .Where(b => b.NumeBoRe == numeroRecepcionDecimal &&
                               b.CodInte == item.CodigoInterno)
                        .FirstOrDefaultAsync();

                    if (bolrecep != null)
                    {
                        // Actualizamos la cantidad rechazada y faltante
                        bolrecep.CanToRech = item.CantidadRechazada;
                        bolrecep.CanToFalta = item.CantidadFaltante;
                        bolrecep.Status = 1;

                        // Actualizamos el depósito normal


                    }
                }

                await _context.SaveChangesAsync();
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPut("aprobar/{numeroRecepcion}")]
        public async Task<ActionResult> AprobarRecepcion(int numeroRecepcion)
        {
            try
            {
                var numeroRecepcionDecimal = (decimal)numeroRecepcion;

                var recepcion = await _context.Bolreceps
                    .FirstOrDefaultAsync(b => b.NumeBoRe == numeroRecepcionDecimal);

                if (recepcion == null)
                    return NotFound($"No se encontró la recepción {numeroRecepcion}");

                if (recepcion.Status != 0)
                    return BadRequest("La recepción ya ha sido aprobada o rechazada");

                recepcion.Status = 1; // Aprobada
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al aprobar la recepción: {ex.Message}");
            }
        }

        [HttpDelete("{numeroRecepcion}")]
        public async Task<ActionResult> BorrarRecepcion(int numeroRecepcion)
        {
            try
            {
                var numeroRecepcionDecimal = (decimal)numeroRecepcion;

                var recepcionesABorrar = await _context.Bolreceps
                    .Where(b => b.NumeBoRe == numeroRecepcionDecimal)
                    .ToListAsync();

                if (!recepcionesABorrar.Any())
                {
                    return NotFound($"No se encontró la recepción #{numeroRecepcion}");
                }

                _context.Bolreceps.RemoveRange(recepcionesABorrar);
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("depositos")]
        public async Task<ActionResult<List<object>>> GetDepositos()
        {
            try
            {
                var depositos = await _context.Tadeposi
                    .Select(d => new
                    {
                        Codigo = d.Id,
                        Nombre = d.Descripcion
                    })
                    .ToListAsync();

                return Ok(depositos);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// Obtiene recepciones por rango de fechas con filtros opcionales
        /// Equivalente al SQL: SELECT * FROM BOLRECEP INNER JOIN MASTER ON Master.Codint = BolRecep.Cod_Inte
        /// </summary>
        [HttpGet("por-rango-fechas")]
        public async Task<ActionResult<List<Bolrecep>>> GetRecepcionesPorRangoFechas(
            [FromQuery] DateTime? fechaDesde,
            [FromQuery] DateTime? fechaHasta,
            [FromQuery] string? codigoDesde,
            [FromQuery] string? codigoHasta)
        {
            try
            {
                // Verificar derecho APRSINLI en DER_ACCE
                const string COD_APLI = "APRSINLI";

                // Hacer el JOIN con MASTER
                var query = from b in _context.Bolreceps
                            join m in _context.masters on b.CodInte equals m.Codint into masterJoin
                            from m in masterJoin.DefaultIfEmpty()
                            select new { Bolrecep = b, Master = m };

                // Filtro por rango de fechas
                if (fechaDesde.HasValue)
                {
                    query = query.Where(x => x.Bolrecep.FechRecep >= fechaDesde.Value.Date);
                }

                if (fechaHasta.HasValue)
                {
                    query = query.Where(x => x.Bolrecep.FechRecep <= fechaHasta.Value.Date);
                }

                // Filtro por rango de códigos (del MASTER)
                if (!string.IsNullOrWhiteSpace(codigoDesde))
                {
                    query = query.Where(x => x.Master != null && string.Compare(x.Master.Codigo, codigoDesde) >= 0);
                }

                if (!string.IsNullOrWhiteSpace(codigoHasta))
                {
                    query = query.Where(x => x.Master != null && string.Compare(x.Master.Codigo, codigoHasta) <= 0);
                }

                var recepciones = await query
                    .OrderBy(x => x.Bolrecep.NumeBoRe)
                    .Select(x => x.Bolrecep)
                    .ToListAsync();

                return Ok(recepciones);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener recepciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene recepciones con rechazos por rango de fechas y proveedor
        /// WHERE CanToRech > 0 AND FechRecep >= fechaDesde AND FechRecep <= fechaHasta AND NumProv = numeroProveedor
        /// </summary>
        [HttpGet("con-rechazos")]
        public async Task<ActionResult<List<Bolrecep>>> GetRecepcionesConRechazos(
            [FromQuery] DateTime? fechaDesde,
            [FromQuery] DateTime? fechaHasta,
            [FromQuery] int? numeroProveedor)
        {
            try
            {
                var query = _context.Bolreceps
                    .Where(b => b.CanToRech > 0) // Solo recepciones con rechazos
                    .AsQueryable();

                // Filtro por rango de fechas
                if (fechaDesde.HasValue)
                {
                    query = query.Where(b => b.FechRecep >= fechaDesde.Value.Date);
                }

                if (fechaHasta.HasValue)
                {
                    query = query.Where(b => b.FechRecep <= fechaHasta.Value.Date);
                }

                // Filtro por proveedor (opcional)
                if (numeroProveedor.HasValue)
                {
                    query = query.Where(b => b.NumProv == numeroProveedor.Value);
                }

                var recepciones = await query
                    .OrderBy(b => b.NumeBoRe)
                    .ToListAsync();

                return Ok(recepciones);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener recepciones con rechazos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene recepciones sin orden de compra por rango de fechas y proveedor
        /// WHERE (NumOC is null or NumOC < 1) AND FechRecep >= fechaDesde AND FechRecep <= fechaHasta AND NumProv = numeroProveedor
        /// </summary>
        [HttpGet("sin-orden-compra")]
        public async Task<ActionResult<List<Bolrecep>>> GetRecepcionesSinOrdenCompra(
            [FromQuery] DateTime? fechaDesde,
            [FromQuery] DateTime? fechaHasta,
            [FromQuery] int? numeroProveedor)
        {
            try
            {
                var query = _context.Bolreceps
                    .Where(b => b.NumOC == null || b.NumOC < 1) // Sin orden de compra
                    .AsQueryable();

                // Filtro por rango de fechas
                if (fechaDesde.HasValue)
                {
                    query = query.Where(b => b.FechRecep >= fechaDesde.Value.Date);
                }

                if (fechaHasta.HasValue)
                {
                    query = query.Where(b => b.FechRecep <= fechaHasta.Value.Date);
                }

                // Filtro por proveedor (opcional)
                if (numeroProveedor.HasValue)
                {
                    query = query.Where(b => b.NumProv == numeroProveedor.Value);
                }

                var recepciones = await query
                    .OrderBy(b => b.NumeBoRe)
                    .ToListAsync();

                return Ok(recepciones);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener recepciones sin orden de compra: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene datos para gráficos de recepciones
        /// </summary>
        [HttpGet("graficos")]
        public async Task<ActionResult<GraficosRecepcionesDTO>> GetDatosGraficos(
            [FromQuery] DateTime? fechaDesde,
            [FromQuery] DateTime? fechaHasta)
        {
            try
            {
                var fechaInicio = fechaDesde ?? DateTime.Now.AddMonths(-6);
                var fechaFin = fechaHasta ?? DateTime.Now;

                var recepciones = await _context.Bolreceps
                    .Where(b => b.FechRecep >= fechaInicio && b.FechRecep <= fechaFin)
                    .ToListAsync();

                var resultado = new GraficosRecepcionesDTO();

                // 1. Recepciones por mes
                resultado.RecepcionesPorMes = recepciones
                    .GroupBy(r => new { r.FechRecep!.Value.Year, r.FechRecep.Value.Month })
                    .OrderBy(g => g.Key.Year).ThenBy(g => g.Key.Month)
                    .Select(g => new DatoMensualRecepcionDTO
                    {
                        Mes = $"{g.Key.Month:00}/{g.Key.Year}",
                        Cantidad = g.Sum(r => r.CantNom ?? 0),
                        NumeroRecepciones = g.Count()
                    })
                    .ToList();

                // 2. Calificación de proveedores
                var recepcionesPorProveedor = recepciones
                    .Where(r => r.NumProv.HasValue)
                    .GroupBy(r => r.NumProv!.Value);

                var proveedores = await _context.Proveedores
                    .Where(p => recepciones.Select(r => r.NumProv).Contains(p.Nume_Cli))
                    .ToDictionaryAsync(p => p.Nume_Cli, p => p.Raso_Cli ?? "");

                resultado.CalificacionProveedores = recepcionesPorProveedor
                    .Select(g =>
                    {
                        var totalCantidad = g.Sum(r => r.CantNom ?? 0);
                        var cantidadAprobada = g.Sum(r => r.CanToApro ?? 0);
                        var cantidadRechazada = g.Sum(r => r.CanToRech ?? 0);
                        var cantidadFaltante = g.Sum(r => r.CanToFalta ?? 0);

                        var porcAprobacion = totalCantidad > 0 ? (cantidadAprobada / totalCantidad) * 100 : 0;
                        var porcRechazo = totalCantidad > 0 ? (cantidadRechazada / totalCantidad) * 100 : 0;
                        var porcFaltante = totalCantidad > 0 ? (cantidadFaltante / totalCantidad) * 100 : 0;

                        // Calificación: 100 - (% rechazo * 2) - (% faltante)
                        var calificacion = Math.Max(0, 100 - (porcRechazo * 2) - porcFaltante);

                        string categoria = calificacion >= 90 ? "Excelente" :
                                         calificacion >= 75 ? "Bueno" :
                                         calificacion >= 60 ? "Regular" : "Deficiente";

                        return new CalificacionProveedorDTO
                        {
                            NumeroProveedor = g.Key,
                            Proveedor = proveedores.GetValueOrDefault(g.Key, $"Proveedor {g.Key}"),
                            PorcentajeAprobacion = Math.Round(porcAprobacion, 2),
                            PorcentajeRechazo = Math.Round(porcRechazo, 2),
                            PorcentajeFaltante = Math.Round(porcFaltante, 2),
                            TotalRecepciones = g.Count(),
                            CalificacionGeneral = Math.Round(calificacion, 2),
                            Categoria = categoria
                        };
                    })
                    .OrderByDescending(c => c.CalificacionGeneral)
                    .Take(10)
                    .ToList();

                // 3. Rechazos por proveedor (Top 5 con más rechazos)
                resultado.RechazosPorProveedor = recepcionesPorProveedor
                    .Where(g => g.Sum(r => r.CanToRech ?? 0) > 0)
                    .Select(g =>
                    {
                        var totalCantidad = g.Sum(r => r.CantNom ?? 0);
                        var cantidadRechazada = g.Sum(r => r.CanToRech ?? 0);
                        var porcRechazo = totalCantidad > 0 ? (cantidadRechazada / totalCantidad) * 100 : 0;

                        return new DatoRechazoDTO
                        {
                            Proveedor = proveedores.GetValueOrDefault(g.Key, $"Proveedor {g.Key}"),
                            CantidadRechazada = cantidadRechazada,
                            PorcentajeRechazo = Math.Round(porcRechazo, 2)
                        };
                    })
                    .OrderByDescending(r => r.CantidadRechazada)
                    .Take(5)
                    .ToList();

                // 4. Estados de recepciones
                var conOC = recepciones.Count(r => r.NumOC.HasValue && r.NumOC > 0);
                var sinOC = recepciones.Count(r => !r.NumOC.HasValue || r.NumOC < 1);
                var conRechazos = recepciones.Count(r => r.CanToRech.HasValue && r.CanToRech > 0);
                var sinRechazos = recepciones.Count(r => !r.CanToRech.HasValue || r.CanToRech == 0);

                resultado.EstadosRecepciones = new List<DatoEstadoRecepcionDTO>
                {
                    new DatoEstadoRecepcionDTO { Estado = "Con O/C", Cantidad = conOC },
                    new DatoEstadoRecepcionDTO { Estado = "Sin O/C", Cantidad = sinOC },
                    new DatoEstadoRecepcionDTO { Estado = "Con Rechazos", Cantidad = conRechazos },
                    new DatoEstadoRecepcionDTO { Estado = "Sin Rechazos", Cantidad = sinRechazos }
                };

                // 5. Top productos recibidos
                resultado.TopProductosRecibidos = recepciones
                    .Where(r => !string.IsNullOrEmpty(r.Descrip))
                    .GroupBy(r => r.Descrip)
                    .Select(g => new DatoProductoRecepcionDTO
                    {
                        Producto = g.Key!.Length > 50 ? g.Key.Substring(0, 50) + "..." : g.Key,
                        CantidadTotal = g.Sum(r => r.CantNom ?? 0),
                        NumeroRecepciones = g.Count()
                    })
                    .OrderByDescending(p => p.CantidadTotal)
                    .Take(10)
                    .ToList();

                // 6. Estadísticas generales
                resultado.TotalRecepciones = recepciones.Count;
                resultado.CantidadTotalRecibida = recepciones.Sum(r => r.CantNom ?? 0);
                resultado.CantidadTotalRechazada = recepciones.Sum(r => r.CanToRech ?? 0);
                resultado.PorcentajeRechazo = resultado.CantidadTotalRecibida > 0
                    ? Math.Round((resultado.CantidadTotalRechazada / resultado.CantidadTotalRecibida) * 100, 2)
                    : 0;
                resultado.TotalProveedores = recepciones.Where(r => r.NumProv.HasValue).Select(r => r.NumProv).Distinct().Count();
                
                var mesesEnRango = ((fechaFin.Year - fechaInicio.Year) * 12) + fechaFin.Month - fechaInicio.Month + 1;
                resultado.PromedioMensual = mesesEnRango > 0 ? resultado.CantidadTotalRecibida / mesesEnRango : 0;

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener datos de gráficos: {ex.Message}");
            }
        }

        /// <summary>
        /// Valida si el usuario puede aprobar una recepción según la tolerancia de días configurada
        /// </summary>
        /// <param name="numeroRecepcion">Número de recepción a validar</param>
        /// <returns>Resultado de la validación con mensaje si corresponde</returns>
        [HttpGet("validar-tolerancia-aprobacion/{numeroRecepcion}")]
        public async Task<ActionResult<ValidacionToleranciaDTO>> ValidarToleranciaAprobacionAsync(int numeroRecepcion)
        {
            try
            {
                // Obtener el IdFlexoft del usuario actual usando el servicio
                int? idFlexoft = _usuarioFlexoftService.InterObtenerIdFlexoft(User);
                if (!idFlexoft.HasValue)
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Verificar derecho APRSINLI en DER_ACCE
                const string COD_APLI = "APRSINLI";
                const string DES_APLI = "";
                var nivelAcceso = await _derechosUsuarioService.VerificarDerechoAsync((short)idFlexoft.Value, COD_APLI, DES_APLI);

                // Si tiene nivel 2 (escritura), puede aprobar sin límite de días
                if (nivelAcceso >= 2)
                {
                    return Ok(new ValidacionToleranciaDTO
                    {
                        PuedeAprobar = true,
                        Mensaje = null
                    });
                }

                // Si no tiene el derecho, validar tolerancia de días
                // Obtener tolerancia desde FLEXCRL (BOLRECEP/TOLEAPRO)
                const string PROBUS = "BOLRECEP";
                const string CLABUS = "TOLEAPRO";

                var configTolerancia = await _context.Flexcrls
                    .AsNoTracking()
                    .FirstOrDefaultAsync(f => f.PROBUS == PROBUS && f.CLABUS == CLABUS);

                int toleranciaDias = 0;
                if (configTolerancia != null && !string.IsNullOrEmpty(configTolerancia.VALCONTROL))
                {
                    if (int.TryParse(configTolerancia.VALCONTROL.Trim(), out int dias))
                    {
                        toleranciaDias = dias;
                    }
                }

                // Si no hay tolerancia configurada (o es 0), no se valida
                if (toleranciaDias <= 0)
                {
                    return Ok(new ValidacionToleranciaDTO
                    {
                        PuedeAprobar = true,
                        Mensaje = null
                    });
                }

                // Obtener fecha de recepción
                var recepcion = await _context.Bolreceps
                    .AsNoTracking()
                    .Where(b => b.NumeBoRe == numeroRecepcion)
                    .Select(b => b.FechRecep)
                    .FirstOrDefaultAsync();

                if (!recepcion.HasValue)
                {
                    return BadRequest("No se encontró la recepción");
                }

                // Obtener fecha del servidor
                var fechaServidor = await _context.Database
                    .SqlQueryRaw<DateTime>("SELECT GETDATE() AS Value")
                    .FirstOrDefaultAsync();

                // Calcular días de diferencia
                var diasDiferencia = (fechaServidor - recepcion.Value).Days;

                // Si excede la tolerancia, bloquear aprobación
                if (diasDiferencia > toleranciaDias)
                {
                    return Ok(new ValidacionToleranciaDTO
                    {
                        PuedeAprobar = false,
                        Mensaje = $"Atención plazo de tolerancia entre la recepción y Aprobación Vencido\nFecha Recepción: {recepcion.Value:dd/MM/yyyy}\nDías pasados: {diasDiferencia}\nTolerancia: {toleranciaDias}"
                    });
                }

                // Puede aprobar
                return Ok(new ValidacionToleranciaDTO
                {
                    PuedeAprobar = true,
                    Mensaje = null
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[BolRecep] Error al validar tolerancia: {ex.Message}");
                return BadRequest(ex.Message);
            }
        }

        // ===================================================================
        // RECEPCIONES TODAS POR RANGO DE FECHAS
        // ===================================================================
        // A diferencia de "pendientes-aprobacion" (que filtra Status == 0),
        // estos endpoints devuelven TODAS las recepciones (cualquier Status)
        // dentro de un rango de fechas, en el mismo formato DTO para que la
        // página RecepcionesTodasXFecha reutilice la misma grilla y filtros
        // que la solapa "Aprueba Calidad" pero sin botones de Aprobar/Borrar.
        // Se usan tanto para la grilla agrupada como para el "Detalle por Lote".
        // ===================================================================

        /// <summary>
        /// Devuelve todas las recepciones (cualquier Status) agrupadas por
        /// (NumeBoRe + NumOC) dentro de un rango de fechas.
        /// GET /api/bolrecep/todas-por-rango-fechas?fechaDesde=...&fechaHasta=...
        /// </summary>
        [HttpGet("todas-por-rango-fechas")]
        public async Task<ActionResult<List<RecepcionPendienteAprobacionDTO>>> GetTodasRecepcionesPorRangoFechas(
            [FromQuery] DateTime? fechaDesde,
            [FromQuery] DateTime? fechaHasta)
        {
            try
            {
                var query = _context.Bolreceps
                    .Where(b => b.NumeBoRe != null);

                if (fechaDesde.HasValue)
                {
                    query = query.Where(b => b.FechRecep >= fechaDesde.Value.Date);
                }
                if (fechaHasta.HasValue)
                {
                    query = query.Where(b => b.FechRecep <= fechaHasta.Value.Date);
                }

                var recepciones = await query
                    .GroupBy(b => new
                    {
                        b.NumeBoRe,
                        b.NumOC,
                        b.FEYHORecep,
                        b.NumRemProv,
                        b.NumProv
                    })
                    .Select(g => new RecepcionPendienteAprobacionDTO
                    {
                        NumeroRecepcion = (int)(g.Key.NumeBoRe ?? 0),
                        NumeroOrdenCompra = g.Key.NumOC ?? 0,
                        FechaRecepcion = g.Key.FEYHORecep ?? DateTime.MinValue,
                        NumeroRemito = g.Key.NumRemProv,
                        NumeroProveedor = g.Key.NumProv ?? 0,
                        // El número de proveedor de BOLRECEP (NumProv) es Proved12.Nume_Cli,
                        // no la clave primaria Id. Buscar por Nume_Cli para obtener la Razón Social.
                        RazonSocial = _context.Proveedores
                            .Where(p => p.Nume_Cli == g.Key.NumProv)
                            .Select(p => p.Raso_Cli)
                            .FirstOrDefault() ?? string.Empty,
                        Observa = g.First().Observa
                    })
                    .OrderByDescending(r => r.FechaRecepcion)
                    .ToListAsync();

                return Ok(recepciones);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// Devuelve el detalle por lote de TODAS las recepciones (cualquier
        /// Status) dentro de un rango de fechas. Una fila por cada registro de
        /// BOLRECEP (una misma recepción puede tener varias filas, una por lote).
        /// GET /api/bolrecep/todas-por-rango-fechas-detalle?fechaDesde=...&fechaHasta=...
        /// </summary>
        [HttpGet("todas-por-rango-fechas-detalle")]
        public async Task<ActionResult<List<RecepcionPendienteAprobacionDetalleDTO>>> GetTodasRecepcionesPorRangoFechasDetalle(
            [FromQuery] DateTime? fechaDesde,
            [FromQuery] DateTime? fechaHasta)
        {
            try
            {
                var query = _context.Bolreceps
                    .Where(b => b.NumeBoRe != null);

                if (fechaDesde.HasValue)
                {
                    query = query.Where(b => b.FechRecep >= fechaDesde.Value.Date);
                }
                if (fechaHasta.HasValue)
                {
                    query = query.Where(b => b.FechRecep <= fechaHasta.Value.Date);
                }

                var detalle = await query
                    .Select(b => new RecepcionPendienteAprobacionDetalleDTO
                    {
                        NumeroRecepcion = (int)(b.NumeBoRe ?? 0),
                        NumeroOrdenCompra = b.NumOC ?? 0,
                        FechaRecepcion = b.FEYHORecep ?? DateTime.MinValue,
                        NumeroRemito = b.NumRemProv,
                        NumeroProveedor = b.NumProv ?? 0,
                        RazonSocial = _context.Proveedores
                            .Where(p => p.Nume_Cli == b.NumProv)
                            .Select(p => p.Raso_Cli)
                            .FirstOrDefault() ?? string.Empty,
                        IdenLote = b.IdenLote ?? string.Empty,
                        Codigo = b.CodExte ?? string.Empty,
                        Descripcion = b.DescripItem ?? string.Empty,
                        CantidadNominal = b.CantNom ?? 0
                    })
                    .OrderByDescending(r => r.FechaRecepcion)
                    .ThenBy(r => r.NumeroRecepcion)
                    .ThenBy(r => r.IdenLote)
                    .ToListAsync();

                return Ok(detalle);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}
