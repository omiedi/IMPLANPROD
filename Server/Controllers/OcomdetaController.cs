using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Data;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para gestionar el detalle de órdenes de compra
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class OcomdetaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<OcomdetaController> _logger;

        public OcomdetaController(DataContext context, ILogger<OcomdetaController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene el detalle de una orden de compra específica
        /// </summary>
        [HttpGet("por-orden/{numeroOrden}")]
        public async Task<ActionResult<List<Ocomdeta>>> GetPorOrdenAsync(int numeroOrden)
        {
            var detalle = await _context.Ocomdetas
                .Where(d => d.NumeOcom == numeroOrden)
                .OrderBy(d => d.NuOrdItem)
                .ToListAsync();

            return Ok(detalle);
        }

        /// <summary>
        /// Obtiene el detalle de una orden de compra por ID de encabezado (OCOMENCA_ID)
        /// Solo muestra items con StatOcom < 8
        /// </summary>
        [HttpGet("orden/{ordenId}")]
        public async Task<ActionResult<List<Ocomdeta>>> GetPorOrdenIdAsync(int ordenId)
        {
            // Primero obtener el número de orden desde el encabezado
            var encabezado = await _context.Ocomencas
                .FirstOrDefaultAsync(o => o.OCOMENCA_ID == ordenId);

            if (encabezado == null)
            {
                return NotFound($"No se encontró la orden con ID {ordenId}");
            }

            var detalle = await _context.Ocomdetas
                .Where(d => d.NumeOcom == encabezado.NumeOcom && d.StatOcom < 8)
                .OrderBy(d => d.NuOrdItem)
                .ToListAsync();

            return Ok(detalle);
        }

        /// <summary>
        /// Obtiene todas las órdenes pendientes de un proveedor específico
        /// Pendiente = CantOcom - CantRec > 0 y StatOcom < 8
        /// </summary>
        [HttpGet("pendientes-proveedor/{numeroProveedor}")]
        public async Task<ActionResult<List<Ocomdeta>>> GetOrdenesPendientesProveedorAsync(int numeroProveedor)
        {
            var ordenesPendientes = await _context.Ocomdetas
                .Where(d => d.NproOcom == numeroProveedor && 
                           d.StatOcom < 8 && // Solo items activos (StatOcom < 8)
                           (d.CantOcom - d.CantRec) > 0) // Con cantidad pendiente
                .OrderByDescending(d => d.FemiOcom)
                .ThenBy(d => d.NumeOcom)
                .ThenBy(d => d.NuOrdItem)
                .ToListAsync();

            return Ok(ordenesPendientes);
        }

        /// <summary>
        /// Obtiene el detalle enriquecido de una orden de compra para edición
        /// Incluye cálculos y referencias según lógica VB6
        /// </summary>
        [HttpGet("edicion/{numeroOrden}")]
        public async Task<ActionResult<List<OcomdetaEditDTO>>> GetDetalleEdicionAsync(int numeroOrden)
        {
            try
            {
                Console.WriteLine($"[GetDetalleEdicionAsync] Cargando detalle para orden: {numeroOrden}");

                // Obtener el encabezado para el tipo de cambio
                var encabezado = await _context.Ocomencas
                    .FirstOrDefaultAsync(o => o.NumeOcom == numeroOrden);

                if (encabezado == null)
                {
                    return NotFound($"No se encontró la orden de compra {numeroOrden}");
                }

                // Calcular tipo de cambio según lógica VB6
                decimal tipoCambio = 1m;
                if (encabezado.MoneEmis >= 2)
                {
                    // Buscar tipo de cambio en TAMONED
                    var moneda = await _context.Tamoneds
                        .FirstOrDefaultAsync(m => m.CodigoMoneda == encabezado.MoneEmis);
                    tipoCambio = moneda?.Cotizacion ?? 1m;
                }

                // Obtener detalle activo (StatOcom < 9) ordenado por NuOrdItem
                var detalleItems = await _context.Ocomdetas
                    .Where(d => d.NumeOcom == numeroOrden && d.StatOcom < 9)
                    .OrderBy(d => d.NuOrdItem)
                    .ToListAsync();

                if (!detalleItems.Any())
                {
                    return BadRequest("Imposible cargar Detalle de O/Compra - Orden de Compra Anulada");
                }

                // Convertir a DTOs con datos básicos (sin enriquecimiento por ahora)
                var detalleEnriquecido = detalleItems.Select(item =>
                    OcomdetaEditDTO.FromEntity(item, null, null, tipoCambio)
                ).ToList();

                Console.WriteLine($"[GetDetalleEdicionAsync] ✅ Cargados {detalleEnriquecido.Count} items para orden {numeroOrden}");

                return Ok(detalleEnriquecido);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GetDetalleEdicionAsync] ❌ Error: {ex.Message}");
                return StatusCode(500, $"Error al cargar detalle de edición: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene todos los detalles con paginación y filtros
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination,
            int? numeroOrden = null,
            int? codigoInterno = null)
        {
            var queryable = _context.Ocomdetas.AsQueryable();

            // Filtro por texto general (código externo, descripción)
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x =>
                    x.CodExte.Contains(pagination.Filter) ||
                    x.Descritem.Contains(pagination.Filter));
            }

            // Filtro por número de orden
            if (numeroOrden.HasValue && numeroOrden > 0)
            {
                queryable = queryable.Where(x => x.NumeOcom == numeroOrden);
            }

            // Filtro por código interno
            if (codigoInterno.HasValue && codigoInterno > 0)
            {
                queryable = queryable.Where(x => x.CodInte == codigoInterno);
            }

            return Ok(await queryable
                .OrderByDescending(x => x.NumeOcom)
                .ThenBy(x => x.NuOrdItem)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination,
            int? numeroOrden = null,
            int? codigoInterno = null)
        {
            var queryable = _context.Ocomdetas.AsQueryable();

            // Aplicar los mismos filtros que en GetAsync
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x =>
                    x.CodExte.Contains(pagination.Filter) ||
                    x.Descritem.Contains(pagination.Filter));
            }

            if (numeroOrden.HasValue && numeroOrden > 0)
            {
                queryable = queryable.Where(x => x.NumeOcom == numeroOrden);
            }

            if (codigoInterno.HasValue && codigoInterno > 0)
            {
                queryable = queryable.Where(x => x.CodInte == codigoInterno);
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene un detalle específico por ID
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<Ocomdeta>> GetAsync(int id)
        {
            var detalle = await _context.Ocomdetas.FirstOrDefaultAsync(x => x.OCOMDETA_ID == id);
            if (detalle == null)
            {
                return NotFound();
            }
            return detalle;
        }

        /// <summary>
        /// Crea un nuevo detalle de orden de compra
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<Ocomdeta>> PostAsync(Ocomdeta detalle)
        {
            try
            {
                // Asignar número de item automáticamente si no se proporciona
                if (!detalle.NuOrdItem.HasValue || detalle.NuOrdItem == 0)
                {
                    var maxItem = await _context.Ocomdetas
                        .Where(d => d.NumeOcom == detalle.NumeOcom)
                        .MaxAsync(d => (short?)d.NuOrdItem) ?? 0;
                    detalle.NuOrdItem = (short)(maxItem + 1);
                }

                // Establecer fechas de auditoría
                detalle.AddRecord = DateTime.Now;
                detalle.LastUpdate = DateTime.Now;

                // Establecer valores por defecto
                if (!detalle.FemiOcom.HasValue)
                {
                    detalle.FemiOcom = DateTime.Now;
                }

                if (!detalle.StatOcom.HasValue)
                {
                    detalle.StatOcom = 0; // Estado activo por defecto
                }

                _context.Ocomdetas.Add(detalle);
                await _context.SaveChangesAsync();

                return CreatedAtAction(nameof(GetAsync), new { id = detalle.OCOMDETA_ID }, detalle);
            }
            catch (Exception ex)
            {
                // Log más detallado del error
                var innerMessage = ex.InnerException?.Message ?? "";
                var fullMessage = $"Error al crear el detalle de orden de compra: {ex.Message}";
                if (!string.IsNullOrEmpty(innerMessage))
                {
                    fullMessage += $" Inner: {innerMessage}";
                }
                
                Console.WriteLine($"Error en OcomdetaController.PostAsync: {fullMessage}");
                Console.WriteLine($"StackTrace: {ex.StackTrace}");
                
                return StatusCode(500, fullMessage);
            }
        }

        /// <summary>
        /// Actualiza un detalle de orden de compra existente
        /// </summary>
        [HttpPut("{id}")]
        public async Task<IActionResult> PutAsync(int id, Ocomdeta detalle)
        {
            if (id != detalle.OCOMDETA_ID)
            {
                return BadRequest("El ID no coincide con el detalle");
            }

            // Actualizar fecha de modificación
            detalle.LastUpdate = DateTime.Now;

            _context.Entry(detalle).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                return NoContent();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!await _context.Ocomdetas.AnyAsync(e => e.OCOMDETA_ID == id))
                {
                    return NotFound();
                }
                throw;
            }
        }

        /// <summary>
        /// Actualiza múltiples detalles de una orden de compra
        /// </summary>
        [HttpPut("multiples")]
        public async Task<IActionResult> PutMultiplesAsync(List<Ocomdeta> detalles)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                foreach (var detalle in detalles)
                {
                    detalle.LastUpdate = DateTime.Now;
                    _context.Entry(detalle).State = EntityState.Modified;
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();
                return NoContent();
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return StatusCode(500, $"Error al actualizar los detalles: {ex.Message}");
            }
        }

        /// <summary>
        /// Elimina un detalle de orden de compra
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var detalle = await _context.Ocomdetas.FindAsync(id);
            if (detalle == null)
            {
                return NotFound();
            }

            _context.Ocomdetas.Remove(detalle);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Elimina todos los detalles de una orden de compra
        /// </summary>
        [HttpDelete("por-orden/{numeroOrden}")]
        public async Task<IActionResult> DeletePorOrdenAsync(int numeroOrden)
        {
            var detalles = await _context.Ocomdetas
                .Where(d => d.NumeOcom == numeroOrden)
                .ToListAsync();

            if (detalles.Any())
            {
                _context.Ocomdetas.RemoveRange(detalles);
                await _context.SaveChangesAsync();
            }

            return NoContent();
        }

        /// <summary>
        /// Obtiene el siguiente número de item para una orden específica
        /// </summary>
        [HttpGet("siguiente-item/{numeroOrden}")]
        public async Task<ActionResult<short>> GetSiguienteItemAsync(int numeroOrden)
        {
            var maxItem = await _context.Ocomdetas
                .Where(d => d.NumeOcom == numeroOrden)
                .MaxAsync(d => (short?)d.NuOrdItem) ?? 0;
            return Ok((short)(maxItem + 1));
        }

        /// <summary>
        /// Crea o actualiza el detalle completo de una orden de compra
        /// </summary>
        [HttpPost("crear-con-detalle")]
        public async Task<ActionResult> CrearConDetalleAsync([FromBody] OrdenCompraConDetalleDTO ordenCompleta)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // Eliminar detalle existente si existe
                var detalleExistente = await _context.Ocomdetas
                    .Where(d => d.NumeOcom == ordenCompleta.NumeroOrden)
                    .ToListAsync();

                if (detalleExistente.Any())
                {
                    _context.Ocomdetas.RemoveRange(detalleExistente);
                }

                // Crear nuevo detalle
                short numeroItem = 1;
                foreach (var item in ordenCompleta.Detalles)
                {
                    var detalle = new Ocomdeta
                    {
                        NumeOcom = ordenCompleta.NumeroOrden,
                        NuOrdItem = numeroItem++,
                        CodInte = item.CodigoInterno,
                        CodExte = item.CodigoExterno,
                        Descritem = item.Descripcion,
                        CantOcom = item.Cantidad,
                        PreUnit = item.PrecioUnitario,
                        INetItem = item.Cantidad * item.PrecioUnitario,
                        FemiOcom = DateTime.Now,
                        StatOcom = 0,
                        AddRecord = DateTime.Now,
                        LastUpdate = DateTime.Now,
                        Rowguid = Guid.NewGuid()
                    };

                    _context.Ocomdetas.Add(detalle);
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return Ok(new { mensaje = "Detalle de orden creado exitosamente" });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return StatusCode(500, $"Error al crear el detalle de la orden: {ex.Message}");
            }
        }

        /// <summary>
        /// Cambia el estado de los items anteriores de la O.C. activa a 10 (anulados)
        /// Condición: Stat_Ocom < 9 and Nume_Ocom = numeroOrden
        /// </summary>
        [HttpPost("cambiar-estado-items")]
        public async Task<IActionResult> CambiarEstadoItems([FromBody] CambiarEstadoRequest request)
        {
            try
            {
                Console.WriteLine($"[CambiarEstadoItems] Procesando orden: {request.NumeroOrden}");
                Console.WriteLine($"[CambiarEstadoItems] Nuevo estado: {request.NuevoEstado}");
                Console.WriteLine($"[CambiarEstadoItems] Condición: {request.CondicionEstado}");

                // Construir la consulta SQL para cambiar estado
                var sql = $"UPDATE OCOMDETA SET Stat_Ocom = {request.NuevoEstado} WHERE Nume_Ocom = {request.NumeroOrden} AND Stat_Ocom < 9";
                
                Console.WriteLine($"[CambiarEstadoItems] Ejecutando SQL: {sql}");

                // Ejecutar la consulta
                var registrosAfectados = await _context.Database.ExecuteSqlRawAsync(sql);

                Console.WriteLine($"[CambiarEstadoItems] Registros actualizados: {registrosAfectados}");

                return Ok(new { 
                    Success = true, 
                    Message = $"Estado de items cambiado exitosamente. Registros afectados: {registrosAfectados}",
                    RegistrosAfectados = registrosAfectados
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[CambiarEstadoItems] Error: {ex.Message}");
                return StatusCode(500, new { 
                    Success = false, 
                    Message = "Error al cambiar estado de items",
                    Error = ex.Message 
                });
            }
        }

        /// <summary>
        /// Inserta múltiples items de detalle OCOMDETA
        /// </summary>
        [HttpPost("insertar-items")]
        public async Task<IActionResult> InsertarItems([FromBody] InsertarItemsRequest request)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                Console.WriteLine($"[InsertarItems] Insertando {request.Items.Count} items para orden: {request.NumeroOrden}");

                foreach (var itemData in request.Items)
                {
                    var detalle = new Ocomdeta
                    {
                        // ===== CAMPOS OBLIGATORIOS BÁSICOS =====
                        NumeOcom = request.NumeroOrden,
                        NuOrdItem = itemData.NuOrdItem,
                        FemiOcom = itemData.FemiOcom ?? DateTime.Now,
                        MediOcom = itemData.MediOcom,
                        NproOcom = (short?)itemData.NproOcom,
                        RaSoProv = itemData.RasoProv,

                        // ===== CAMPOS DE EMPRESA Y AUDITORÍA =====
                        CodiEmpr = 0, // DEBE SER 0 según requerimiento

                        // ===== CÓDIGOS DEL PRODUCTO =====
                        CodInte = (short?)itemData.CodInte,
                        CodExte = itemData.CodExte,
                        Descritem = itemData.Descritem,
                        Delibre = itemData.Delibre,

                        // ===== CAMPOS DE UNIDADES =====
                        UniMas = itemData.UniMas,
                        CantOcom = itemData.CantOcom,
                        UniCom = itemData.UniCom,
                        
                        // VALO_UNID - Valor unitario precedido por la moneda $ , DO
                        ValoUnid = itemData.ValoUnid?.ToString() ,
                        
                        // UNIDS_OCOM - Unidades de la orden de compra (por defecto igual a CantOcom)
                        UnidsOcom = itemData.CantOcom ?? 0,
                        
                        // Coef_Unids - Coeficiente de unidades DEBE SER 1
                        CoefUnids = 1,

                        // ===== PRECIOS Y DESCUENTOS =====
                        PreUnit = itemData.PreUnit,
                        DescItem = itemData.DescItem,
                        PrUnNeto = itemData.PrunNeto ?? itemData.PreUnit ?? 0,
                        INetItem = itemData.INetItem,

                        // ===== CAMPOS DE MONEDA =====
                        // MONE_COS - Moneda de costo (mantener el valor original o usar 2 por defecto)
                        MoneCos = itemData.MoneCos ?? 2,
                        MoneEmis = itemData.MoneEmis,
                        
                        // PRUN_NETO_PES - Precio unitario neto en pesos
                        PrUnNetoPes = itemData.PrunNeto ?? itemData.PreUnit ?? 0,

                        // ===== OTROS CAMPOS =====
                        CantRec = itemData.CantRec,
                        
                        // Fecha de entrega solicitada
                        FentreSol = itemData.FentreSol,
                        
                        // Estado
                        StatOcom = itemData.StatOcom ?? 0,

                        // Imputación contable y sector
                        IContab = itemData.IContab,
                        SectOcom = itemData.SectOcom,

                        // ===== CAMPOS DE AUDITORÍA =====
                        AddRecord = DateTime.Now,
                        LastUpdate = DateTime.Now,
                        Rowguid = Guid.NewGuid()
                    };

                    _context.Ocomdetas.Add(detalle);
                    
                    Console.WriteLine($"[InsertarItems] Item configurado - CodInte: {detalle.CodInte}, CodiEmpr: {detalle.CodiEmpr}, MoneCos: {detalle.MoneCos}, CoefUnids: {detalle.CoefUnids}, FentreSol: {detalle.FentreSol?.ToString("dd/MM/yyyy") ?? "NULL"}");
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                Console.WriteLine($"[InsertarItems] ✅ Insertados {request.Items.Count} items exitosamente");

                return Ok(new { 
                    Success = true, 
                    Message = $"Items insertados exitosamente. Total: {request.Items.Count}",
                    ItemsInsertados = request.Items.Count
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                Console.WriteLine($"[InsertarItems] ❌ Error: {ex.Message}");
                return StatusCode(500, new { 
                    Success = false, 
                    Message = "Error al insertar items de detalle",
                    Error = ex.Message 
                });
            }
        }
        /// <summary>
        /// Obtiene los items de órdenes de compra pendientes de recepción
        /// Según lógica VB6: Stat_Ocom < 3 AND Stat_Ocom <> -1 AND Cant_Rec < Cant_Ocom - 0.001
        /// </summary>
        [HttpGet("pendientes-recepcion")]
        public async Task<ActionResult<List<Ocomdeta>>> GetPendientesRecepcion(int? numeroProveedor = null)
        {
            try
            {
                var query = _context.Ocomdetas
                    .Where(d => d.StatOcom < 3 
                             && d.StatOcom != -1 
                             && d.CantRec < (d.CantOcom - 0.001m));

                // Filtrar por proveedor si se proporciona
                if (numeroProveedor.HasValue && numeroProveedor > 0)
                {
                    query = query.Where(d => d.NproOcom == numeroProveedor);
                }

                var ordenesPendientes = await query
                    .OrderBy(d => d.NproOcom)
                    .ThenBy(d => d.NumeOcom)
                    .ThenBy(d => d.CodExte)
                    .ThenBy(d => d.FentreSol)
                    .ToListAsync();

                return Ok(ordenesPendientes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener órdenes pendientes de recepción");
                return StatusCode(500, $"Error al obtener órdenes pendientes: {ex.Message}");
            }
        }

        // Verifica si un proveedor tiene órdenes de compra pendientes
        [HttpGet("pendientes/proveedor/{numeroProveedor}")]
        public async Task<ActionResult<bool>> TieneOrdenesPendientes(int numeroProveedor)
        {
            try
            {
                var tieneOrdenesPendientes = await _context.Ocomdetas
                    .AnyAsync(o => o.NproOcom == numeroProveedor
                                 && o.StatOcom < 8
                                 && o.StatOcom != -1
                                 && o.CantRec < (o.CantOcom - (decimal)0.001));

                return Ok(tieneOrdenesPendientes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al consultar órdenes pendientes. Detalles completos: {ErrorMessage}", ex.Message);
                return BadRequest($"Error al consultar órdenes pendientes: {ex.Message}");
            }
        }
        // Obtiene el detalle de órdenes pendientes por proveedor
        [HttpGet("pendientes/detalle/{numeroProveedor}")]
        public async Task<ActionResult<List<OrdenCompraPendienteDTO>>> GetOrdenesPendientesDetalle(int numeroProveedor)
        {
            try
            {
                var ordenesPendientes = await _context.Ocomdetas
                    .Where(o => o.NproOcom == numeroProveedor
                               && o.CodInte > 0
                               && o.StatOcom < 8
                               && o.StatOcom != -1
                               && o.CantRec < (o.CantOcom - (decimal)0.001))
                    .Select(o => new OrdenCompraPendienteDTO
                    {
                        CantidadOrdenada = o.CantOcom ?? 0,
                        CantidadRecibida = o.CantRec ?? 0,
                        CoeficienteUnidades = o.CoefUnids ?? 0,
                        NumeroOrdenCompra = o.NumeOcom ?? 0,
                        FechaEntregaSolicitada = o.FentreSol,
                        CodigoInterno = o.CodInte ?? 0,
                        DocumentoOrigen = o.DocOrig ?? string.Empty,
                        NumeroOrdenItem = o.NuOrdItem ?? 0,
                        NumeroOrdenServicio = o.NOrServ ?? 0,
                        Cod_Exte = o.CodExte ?? string.Empty,
                        Descritem = o.Descritem ?? string.Empty
                    })
                    .OrderBy(o => o.NumeroOrdenCompra)
                    .ToListAsync();

                return Ok(ordenesPendientes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener detalle de órdenes pendientes. Detalles completos: {ErrorMessage}", ex.Message);
                return BadRequest($"Error al obtener detalle de órdenes pendientes: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// DTO para la solicitud de cambiar estado de items
    /// </summary>
    public class CambiarEstadoRequest
    {
        public int NumeroOrden { get; set; }
        public int NuevoEstado { get; set; }
        public string CondicionEstado { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
    }

    /// <summary>
    /// DTO para la solicitud de insertar items
    /// </summary>
    public class InsertarItemsRequest
    {
        public int NumeroOrden { get; set; }
        public List<OcomdetaItemData> Items { get; set; } = new();
    }

    /// <summary>
    /// DTO para los datos de un item de OCOMDETA
    /// </summary>
    public class OcomdetaItemData
    {
        // Campos básicos de identificación
        public short? NuOrdItem { get; set; }
        public DateTime? FemiOcom { get; set; }
        public string? MediOcom { get; set; }
        public int? NproOcom { get; set; }
        public string? RasoProv { get; set; }

        // Códigos del producto
        public int? CodInte { get; set; }
        public string? CodExte { get; set; }
        public string? Descritem { get; set; }
        public string? Delibre { get; set; }

        // Unidades y cantidades
        public string? UniMas { get; set; }
        public decimal? CantOcom { get; set; }
        public string? UniCom { get; set; }

        // Precios y descuentos
        public decimal? PreUnit { get; set; }
        public decimal? DescItem { get; set; }
        public decimal? PrunNeto { get; set; }
        public decimal? INetItem { get; set; }
        public string? ValoUnid { get; set; }
        // Fechas
        public DateTime? FentreSol { get; set; }

        // Otros campos
        public short? Consimat { get; set; }
        public decimal? CantRec { get; set; }
        
        // Monedas
        public short? MoneCos { get; set; }
        public short? MoneEmis { get; set; }
        public decimal? TsumOcom { get; set; }

        // Estado
        public short? StatOcom { get; set; }

        // Imputación contable y sector
        public short? IContab { get; set; }
        public string? SectOcom { get; set; }
    }

    /// <summary>
    /// DTO para crear una orden de compra con su detalle
    /// </summary>
    public class OrdenCompraConDetalleDTO
    {
        public int NumeroOrden { get; set; }
        public List<DetalleOrdenCompraDTO> Detalles { get; set; } = new();
    }

    /// <summary>
    /// DTO para el detalle de una orden de compra
    /// </summary>
    public class DetalleOrdenCompraDTO
    {
        public short? CodigoInterno { get; set; }
        public string? CodigoExterno { get; set; }
        public string? Descripcion { get; set; }
        public decimal? Cantidad { get; set; }
        public decimal? PrecioUnitario { get; set; }
    }
}
