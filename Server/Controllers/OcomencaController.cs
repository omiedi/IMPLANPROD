using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Server.Data;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Services;
using Microsoft.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para gestionar las órdenes de compra (encabezado)
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class OcomencaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly DetaivaService _detaivaService;
        private readonly ConsignacionService _consignacionService;
        private readonly ActualizacionCostoMasterService _actualizacionCostoMasterService;

        public OcomencaController(DataContext context, DetaivaService detaivaService, ConsignacionService consignacionService, ActualizacionCostoMasterService actualizacionCostoMasterService)
        {
            _context = context;
            _detaivaService = detaivaService;
            _consignacionService = consignacionService;
            _actualizacionCostoMasterService = actualizacionCostoMasterService;
        }

        /// <summary>
        /// Obtiene todas las órdenes de compra con paginación y filtros
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, 
            int? numeroProveedor = null, 
            string? filtroEstado = null,
            DateTime? fechaDesde = null,
            DateTime? fechaHasta = null)
        {
            try
            {
                Console.WriteLine($"[OcomencaController] GetAsync - Parámetros recibidos:");
                Console.WriteLine($"  - Page: {pagination.Page}, RecordsNumber: {pagination.RecordsNumber}");
                Console.WriteLine($"  - Filter: '{pagination.Filter}'");
                Console.WriteLine($"  - NumeroProveedor: {numeroProveedor}");
                Console.WriteLine($"  - FiltroEstado: '{filtroEstado}'");
                Console.WriteLine($"  - FechaDesde: {fechaDesde}");
                Console.WriteLine($"  - FechaHasta: {fechaHasta}");

                var page = pagination.Page <= 0 ? 1 : pagination.Page;
                var recordsNumber = pagination.RecordsNumber <= 0 ? 10 : pagination.RecordsNumber;

                var resultado = await EjecutarConReintentoConexionAsync(async () =>
                {
                    var queryable = ConstruirQueryOrdenesCompraFiltrada(pagination.Filter, numeroProveedor, filtroEstado, fechaDesde, fechaHasta);

                    return await queryable
                        .OrderByDescending(x => x.FemiOcom)
                        .ThenByDescending(x => x.NumeOcom)
                        .Skip((page - 1) * recordsNumber)
                        .Take(recordsNumber)
                        .ToListAsync();
                }, "GetAsync");

                Console.WriteLine($"[OcomencaController] GetAsync - Se encontraron {resultado.Count} registros");

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[OcomencaController] GetAsync - Error: {ex.Message}");
                Console.WriteLine($"[OcomencaController] GetAsync - StackTrace: {ex.StackTrace}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination,
            int? numeroProveedor = null,
            string? filtroEstado = null,
            DateTime? fechaDesde = null,
            DateTime? fechaHasta = null)
        {
            var recordsNumber = pagination.RecordsNumber <= 0 ? 10 : pagination.RecordsNumber;

            var count = await EjecutarConReintentoConexionAsync(async () =>
            {
                var queryable = ConstruirQueryOrdenesCompraFiltrada(pagination.Filter, numeroProveedor, filtroEstado, fechaDesde, fechaHasta);
                return await queryable.CountAsync();
            }, "GetPages");

            double totalPages = Math.Ceiling((double)count / recordsNumber);
            return Ok(totalPages);
        }

        private IQueryable<Ocomenca> ConstruirQueryOrdenesCompraFiltrada(
            string? filtro,
            int? numeroProveedor,
            string? filtroEstado,
            DateTime? fechaDesde,
            DateTime? fechaHasta)
        {
            var queryable = _context.Ocomencas.AsNoTracking().AsQueryable();

            if (!string.IsNullOrWhiteSpace(filtro))
            {
                var filtroNormalizado = filtro.Trim();

                if (int.TryParse(filtroNormalizado, out int numero))
                {
                    queryable = queryable.Where(x =>
                        x.NumeOcom == numero ||
                        (x.RefeOcom != null && x.RefeOcom.Contains(filtroNormalizado)) ||
                        (x.RaSoProv != null && x.RaSoProv.Contains(filtroNormalizado)));
                }
                else
                {
                    queryable = queryable.Where(x =>
                        (x.RefeOcom != null && x.RefeOcom.Contains(filtroNormalizado)) ||
                        (x.RaSoProv != null && x.RaSoProv.Contains(filtroNormalizado)));
                }
            }

            if (numeroProveedor.HasValue && numeroProveedor > 0)
            {
                queryable = queryable.Where(x => x.NproOcom == numeroProveedor);
            }

            if (!string.IsNullOrWhiteSpace(filtroEstado) && short.TryParse(filtroEstado, out short estado))
            {
                queryable = estado >= 0
                    ? queryable.Where(x => x.StatOcom >= estado)
                    : queryable.Where(x => x.StatOcom == estado);
            }

            if (fechaDesde.HasValue)
            {
                var desde = fechaDesde.Value.Date;
                queryable = queryable.Where(x => x.FemiOcom >= desde);
            }

            if (fechaHasta.HasValue)
            {
                var hastaExclusivo = fechaHasta.Value.Date.AddDays(1);
                queryable = queryable.Where(x => x.FemiOcom < hastaExclusivo);
            }

            return queryable;
        }

        private static bool EsErrorConexionCerrada(Exception ex)
        {
            if (ex is InvalidOperationException invalidOp &&
                (invalidOp.Message?.Contains("connection is closed", StringComparison.OrdinalIgnoreCase) ?? false))
            {
                return true;
            }

            if (ex.InnerException != null)
            {
                return EsErrorConexionCerrada(ex.InnerException);
            }

            return false;
        }

        private async Task<T> EjecutarConReintentoConexionAsync<T>(Func<Task<T>> operacion, string origen)
        {
            try
            {
                return await operacion();
            }
            catch (Exception ex) when (EsErrorConexionCerrada(ex))
            {
                Console.WriteLine($"[OcomencaController] {origen} - Conexión cerrada detectada. Reintentando consulta...");

                try
                {
                    await _context.Database.CloseConnectionAsync();
                }
                catch
                {
                    // No propagar error al cerrar; el objetivo es forzar reconexión en el reintento.
                }

                return await operacion();
            }
        }

        /// <summary>
        /// Obtiene una orden de compra específica por ID
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<Ocomenca>> GetAsync(int id)
        {
            var ordenCompra = await _context.Ocomencas.FirstOrDefaultAsync(x => x.OCOMENCA_ID == id);
            if (ordenCompra == null)
            {
                return NotFound();
            }
            return ordenCompra;
        }

        /// <summary>
        /// Obtiene una orden de compra por número de orden
        /// </summary>
        [HttpGet("por-numero/{numeroOrden}")]
        public async Task<ActionResult<Ocomenca>> GetPorNumeroAsync(int numeroOrden)
        {
            var ordenCompra = await _context.Ocomencas.FirstOrDefaultAsync(x => x.NumeOcom == numeroOrden);
            if (ordenCompra == null)
            {
                return NotFound();
            }
            return ordenCompra;
        }

        /// <summary>
        /// Crea una nueva orden de compra completa con detalle y distribución de IVA
        /// Implementa transacciones para garantizar integridad de datos
        /// Valida que exista al menos un item en el detalle
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<Ocomenca>> PostAsync(OrdenCompraCompletaDTO ordenCompleta)
        {
            // Logging de entrada
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine($"[OcomencaController] 🚀 INICIANDO PostAsync");
            Console.WriteLine($"[OcomencaController] Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            Console.WriteLine($"[OcomencaController] Datos recibidos:");
            Console.WriteLine($"  - NumeOcom: {ordenCompleta?.Encabezado?.NumeOcom}");
            Console.WriteLine($"  - NproOcom: {ordenCompleta?.Encabezado?.NproOcom}");
            Console.WriteLine($"  - Items detalle: {ordenCompleta?.Detalle?.Count ?? 0}");
            Console.WriteLine($"  - Actualizar IVA: {ordenCompleta?.ActualizarIva}");
            Console.WriteLine("=".PadRight(80, '='));

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // ===== VALIDACIONES OBLIGATORIAS =====
                Console.WriteLine($"[OcomencaController] 🔍 Iniciando validaciones...");
                
                // 1. Validar estructura del DTO
                var errores = ordenCompleta.Validar();
                if (errores.Any())
                {
                    Console.WriteLine($"[OcomencaController] ❌ Errores de validación: {string.Join(", ", errores)}");
                    return BadRequest(new { errores });
                }
                Console.WriteLine($"[OcomencaController] ✅ Validación de estructura OK");

                // 2. Validar que haya al menos un item (requerimiento específico)
                if (!ordenCompleta.Detalle.Any())
                {
                    return BadRequest("Debe incluir al menos un item en la orden de compra");
                }

                var ordenCompra = ordenCompleta.Encabezado;

                // 3. Validar proveedor existe y está activo
                var proveedorExiste = await _context.Proveedores
                    .AnyAsync(p => p.Nume_Cli == ordenCompra.NproOcom && p.Stat_Cli >= 0);
                
                if (!proveedorExiste)
                {
                    return BadRequest("El proveedor especificado no existe o está inactivo");
                }

                // ===== PROCESAMIENTO SEGÚN TIPO DE OPERACIÓN =====
                
                // Detectar si es edición (si la orden ya existe en BD)
                bool esEdicion = false;
                if (ordenCompra.NumeOcom.HasValue && ordenCompra.NumeOcom > 0)
                {
                    esEdicion = await _context.Ocomencas.AnyAsync(o => o.NumeOcom == ordenCompra.NumeOcom);
                }

                if (esEdicion)
                {
                    // EDICIÓN: Marcar items anteriores como STAT_OCOM = 10
                    await MarcarItemsAnterioresComoInactivos(ordenCompra.NumeOcom!.Value);
                }
                else
                {
                    // CREACIÓN: Asignar número de orden automáticamente
                    if (!ordenCompra.NumeOcom.HasValue || ordenCompra.NumeOcom == 0)
                    {
                        var maxNumero = await _context.Ocomencas.MaxAsync(x => (int?)x.NumeOcom) ?? 0;
                        ordenCompra.NumeOcom = maxNumero + 1;
                    }
                }

                // ===== CONFIGURACIÓN DEL ENCABEZADO =====
                
                // Establecer fecha de emisión
                ordenCompra.FemiOcom ??= DateTime.Now;

                // Establecer fecha de generación (FGEN_OCOM) - Solo en creación
                if (!esEdicion)
                {
                    ordenCompra.FGenOcom = DateTime.Now;
                    Console.WriteLine($"[OcomencaController] PostAsync - Estableciendo FGenOcom: {ordenCompra.FGenOcom}");
                }

                // Establecer valores por defecto
                ordenCompra.StatOcom ??= 0;
                ordenCompra.IdComprob ??= $"OC.{ordenCompra.NumeOcom}";
                ordenCompra.MoneEmis ??= 1;
                
                // Establecer código de empresa siempre en 0
                ordenCompra.CodiEmpr = 0;
                Console.WriteLine($"[OcomencaController] PostAsync - Estableciendo CodiEmpr: {ordenCompra.CodiEmpr}");
                
                // Obtener y establecer el usuario activo en UsuaGen
                var usuarioActivo = ObtenerUsuarioActivo();
                if (usuarioActivo.HasValue)
                {
                    ordenCompra.UsuaGen = usuarioActivo.Value;
                    Console.WriteLine($"[OcomencaController] PostAsync - Estableciendo UsuaGen: {ordenCompra.UsuaGen}");
                }
                else
                {
                    Console.WriteLine($"[OcomencaController] PostAsync - ⚠️ No se pudo obtener el usuario activo para UsuaGen");
                }

                // Obtener cotización del dólar
                var monedaDolar = await _context.Tamoneds
                    .FirstOrDefaultAsync(m => m.CodigoMoneda == 2);
                ordenCompra.CotizaDolar = monedaDolar?.Cotizacion ?? 1m;

                // ===== GRABAR ENCABEZADO =====
                
                if (esEdicion)
                {
                    _context.Entry(ordenCompra).State = EntityState.Modified;
                    _context.Entry(ordenCompra).Property(x => x.FGenOcom).IsModified = false;
                }
                else
                {
                    _context.Ocomencas.Add(ordenCompra);
                }

                // ===== CALCULAR CAMPOS ADICIONALES =====
                CalcularCamposAdicionales(ordenCompra, ordenCompleta);
                
                await _context.SaveChangesAsync();

                // ===== GRABAR DETALLE =====
                await GrabarDetalleOrdenCompra(ordenCompleta.Detalle, ordenCompra.NumeOcom.Value);

                // ===== GRABAR DISTRIBUCIÓN DE IVA =====
                if (ordenCompleta.DistribucionIva != null && ordenCompleta.ActualizarIva)
                {
                    await GrabarDistribucionIvaDesdeDTO(ordenCompleta.DistribucionIva, ordenCompra.NumeOcom.Value);
                }
                
                // ===== PROCESAR CONSIGNACIONES (MACONSIG) =====
                // Según lógica VB6: por cada item que se guarda en OCOMDETA, 
                // verificar si posee estructura y agregar componentes hijos a MACONSIG
                Console.WriteLine($"[OcomencaController] 🔄 Iniciando procesamiento de consignaciones...");
                
                if (esEdicion)
                {
                    // En edición: eliminar consignaciones anteriores
                    await _consignacionService.EliminarConsignacionesOrdenAsync(ordenCompra.NumeOcom.Value);
                }

                // Procesar consignaciones para la orden
                var seProcesoConsignaciones = await _consignacionService.ProcesarConsignacionesOrdenCompraAsync(
                    ordenCompra.NumeOcom.Value,
                    ordenCompra.NproOcom ?? 0,
                    ordenCompra.RaSoProv ?? "",
                    ordenCompra.FemiOcom ?? DateTime.Now,
                    (short)(ordenCompra.MoneEmis ?? 1)
                );

                if (seProcesoConsignaciones)
                {
                    Console.WriteLine($"[OcomencaController] ✅ Consignaciones procesadas exitosamente para O.C. {ordenCompra.NumeOcom}");
                }
                else
                {
                    Console.WriteLine($"[OcomencaController] ℹ️ No se encontraron items con estructura para consignaciones en O.C. {ordenCompra.NumeOcom}");
                }

                // ===== ACTUALIZAR COSTOS EN MASTER =====
                // Según lógica VB6: actualizar COSUNI, FECOSTO, Monecos y PROVHABI
                // basado en configuración ACOSREP y proveedor habitual
                Console.WriteLine($"[OcomencaController] 🔄 Iniciando actualización de costos en MASTER...");
                
                var productosActualizados = await _actualizacionCostoMasterService.ActualizarCostosMasterAsync(
                    ordenCompra.NumeOcom.Value,
                    ordenCompra.NproOcom ?? 0
                );

                if (productosActualizados > 0)
                {
                    Console.WriteLine($"[OcomencaController] ✅ Actualizados {productosActualizados} productos en MASTER para O.C. {ordenCompra.NumeOcom}");
                }
                else
                {
                    Console.WriteLine($"[OcomencaController] ℹ️ No se actualizaron productos en MASTER para O.C. {ordenCompra.NumeOcom}");
                }

                // ===== GUARDAR TODOS LOS CAMBIOS EN LA BASE DE DATOS =====
                await _context.SaveChangesAsync();
                Console.WriteLine($"[OcomencaController] 💾 Cambios guardados en la base de datos para O.C. {ordenCompra.NumeOcom}");

                await transaction.CommitAsync();

                Console.WriteLine($"[OcomencaController] PostAsync - Orden {ordenCompra.NumeOcom} grabada exitosamente");
                
                return Ok(ordenCompra);
            }
            catch (UnauthorizedAccessException ex)
            {
                await transaction.RollbackAsync();
                
                // Logging detallado del error de autorización
                Console.WriteLine("=".PadRight(80, '='));
                Console.WriteLine($"[OcomencaController] 🔒 ERROR DE AUTORIZACIÓN EN PostAsync");
                Console.WriteLine($"[OcomencaController] Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                Console.WriteLine($"[OcomencaController] Mensaje: {ex.Message}");
                Console.WriteLine("=".PadRight(80, '='));
                
                // Devolver 401 Unauthorized con el mensaje detallado del AuditInterceptor
                return Unauthorized(ex.Message);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                
                // Logging detallado del error
                Console.WriteLine("=".PadRight(80, '='));
                Console.WriteLine($"[OcomencaController] ❌ ERROR EN PostAsync");
                Console.WriteLine($"[OcomencaController] Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                Console.WriteLine($"[OcomencaController] Tipo de Excepción: {ex.GetType().Name}");
                Console.WriteLine($"[OcomencaController] Mensaje: {ex.Message}");
                
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"[OcomencaController] Inner Exception: {ex.InnerException.Message}");
                }
                
                Console.WriteLine($"[OcomencaController] Stack Trace:");
                Console.WriteLine(ex.StackTrace);
                
                // Información adicional del contexto
                if (ordenCompleta?.Encabezado != null)
                {
                    Console.WriteLine($"[OcomencaController] Contexto - NumeOcom: {ordenCompleta.Encabezado.NumeOcom}");
                    Console.WriteLine($"[OcomencaController] Contexto - NproOcom: {ordenCompleta.Encabezado.NproOcom}");
                    Console.WriteLine($"[OcomencaController] Contexto - Items en detalle: {ordenCompleta.Detalle?.Count ?? 0}");
                }
                
                Console.WriteLine("=".PadRight(80, '='));
                
                return StatusCode(500, $"Error al procesar la orden de compra: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza una orden de compra existente con validaciones
        /// </summary>
        [HttpPut("{numeCompra}")]
        public async Task<IActionResult> PutAsync(int numeCompra, Ocomenca ordenCompra)
        {
            if (numeCompra != ordenCompra.NumeOcom)
            {
                return BadRequest("El ID no coincide con la orden de compra");
            }

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // ===== VALIDACIONES OBLIGATORIAS =====
                
                // 1. Validar condición de pago
                if (ordenCompra.CpagOcom.HasValue && ordenCompra.CpagOcom > 0)
                {
                    var condicionPagoExiste = await _context.CondPags
                        .AnyAsync(cp => cp.NumCondPag == ordenCompra.CpagOcom && cp.NumCondPag >= 0);
                    
                    if (!condicionPagoExiste)
                    {
                        return BadRequest("La condición de pago especificada no es válida o está inactiva");
                    }
                }

                // 2. Validar fecha de emisión
                if (ordenCompra.FemiOcom.HasValue)
                {
                    // Validar que la fecha no sea futura
                    if (ordenCompra.FemiOcom.Value.Date > DateTime.Now.Date)
                    {
                        return BadRequest("La fecha de emisión no puede ser mayor a la fecha actual");
                    }
                }

                // 3. Validar número de orden de compra único (solo si se cambió)
                if (ordenCompra.NumeOcom.HasValue)
                {
                    var ordenValidacion = await _context.Ocomencas.AsNoTracking()
                        .FirstOrDefaultAsync(o => o.NumeOcom == numeCompra);
                    
                    if (ordenValidacion != null && ordenValidacion.NumeOcom != ordenCompra.NumeOcom)
                    {
                        // Se cambió el número, verificar que no esté ocupado
                        var numeroExiste = await _context.Ocomencas
                            .AnyAsync(o => o.NumeOcom == ordenCompra.NumeOcom && o.NumeOcom != numeCompra);
                        
                        if (numeroExiste)
                        {
                            return BadRequest($"El número de orden de compra {ordenCompra.NumeOcom} ya está ocupado");
                        }
                    }
                }

                // ===== CARGAR ENTIDAD EXISTENTE DESDE BD =====
                var ordenExistente = await _context.Ocomencas
                    .FirstOrDefaultAsync(o => o.NumeOcom == numeCompra);
                
                if (ordenExistente == null)
                {
                    return NotFound($"No se encontró la orden de compra con número {numeCompra}");
                }

                // ===== ACTUALIZAR PROPIEDADES DE LA ENTIDAD EXISTENTE =====
                
                // Fechas
                ordenExistente.FemiOcom = ordenCompra.FemiOcom;
                
                // Proveedor
                ordenExistente.NproOcom = ordenCompra.NproOcom;
                ordenExistente.RaSoProv = ordenCompra.RaSoProv;
                
                // Monedas y totales
                ordenExistente.MoneEmis = ordenCompra.MoneEmis;
                ordenExistente.MoneCos = ordenCompra.MoneCos;
                ordenExistente.TotalOcom = ordenCompra.TotalOcom;
                
                // Observaciones y estado
                ordenExistente.ObseOcom = ordenCompra.ObseOcom;
                ordenExistente.StatOcom = ordenCompra.StatOcom;
                ordenExistente.PorcIva = ordenCompra.PorcIva;
                
                // Condición de pago
                ordenExistente.CpagOcom = ordenCompra.CpagOcom;

                // Versión / revisión (OcomRevis)
                // Regla: guardar la versión enviada (pantalla). Si ya hay una versión igual o mayor en BD,
                // entonces incrementar en 1 sobre la versión actual para evitar duplicados.
                {
                    var versionActual = (short)(ordenExistente.OcomRevis ?? 0);
                    var versionEnviada = (short)(ordenCompra.OcomRevis ?? 0);

                    if (versionEnviada <= versionActual)
                    {
                        ordenExistente.OcomRevis = (short)(versionActual + 1);
                    }
                    else
                    {
                        ordenExistente.OcomRevis = versionEnviada;
                    }
                }
                
                // Campos calculados/automáticos
                ordenExistente.LastUpdate = DateTime.Now;
                ordenExistente.CodiEmpr = 0;
                
                // Usuario que modifica
                var usuarioActivoActual = ObtenerUsuarioActivo();
                if (usuarioActivoActual.HasValue)
                {
                    ordenExistente.UsuaMod = (int?)usuarioActivoActual.Value;
                    Console.WriteLine($"[OcomencaController] PutAsync - Estableciendo UsuaMod: {ordenExistente.UsuaMod}");
                }

                // Actualizar ID de comprobante si cambió el número
                if (ordenCompra.NumeOcom.HasValue && ordenCompra.NumeOcom != numeCompra)
                {
                    ordenExistente.NumeOcom = ordenCompra.NumeOcom;
                    ordenExistente.IdComprob = $"OC.{ordenCompra.NumeOcom}";
                }

                // Obtener cotización del dólar actualizada (moneda 2)
                var monedaDolar = await _context.Tamoneds
                    .FirstOrDefaultAsync(m => m.CodigoMoneda == 2);
                ordenExistente.CotizaDolar = monedaDolar?.Cotizacion ?? 1m;

                Console.WriteLine($"[OcomencaController] PutAsync - Entidad existente actualizada, preservando FGenOcom original");
                
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();
                
                return NoContent();
            }
            catch (DbUpdateConcurrencyException)
            {
                await transaction.RollbackAsync();
                if (!await _context.Ocomencas.AnyAsync(e => e.NumeOcom == numeCompra))
                {
                    return NotFound();
                }
                throw;
            }
            catch (UnauthorizedAccessException ex)
            {
                await transaction.RollbackAsync();
                
                // Logging del error de autorización
                Console.WriteLine($"[OcomencaController] 🔒 ERROR DE AUTORIZACIÓN EN PutAsync - {ex.Message}");
                
                // Devolver 401 Unauthorized con el mensaje detallado del AuditInterceptor
                return Unauthorized(ex.Message);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return StatusCode(500, $"Error al actualizar la orden de compra: {ex.Message}");
            }
        }

        /// <summary>
        /// Elimina una orden de compra (borrado lógico)
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var ordenCompra = await _context.Ocomencas.FindAsync(id);
            if (ordenCompra == null)
            {
                return NotFound();
            }

            // Borrado lógico: cambiar estado a -1
            ordenCompra.StatOcom = -1;
            ordenCompra.LastUpdate = DateTime.Now;
            
            // Establecer código de empresa siempre en 0
            ordenCompra.CodiEmpr = 0;
            
            // Obtener y establecer el usuario activo en UsuaAnul
            var usuarioActivo = ObtenerUsuarioActivo();
            if (usuarioActivo.HasValue)
            {
                ordenCompra.UsuaAnul = (int?)usuarioActivo.Value;
                Console.WriteLine($"[OcomencaController] DeleteAsync - Estableciendo UsuaAnul: {ordenCompra.UsuaAnul}");
            }
            else
            {
                Console.WriteLine($"[OcomencaController] DeleteAsync - ⚠️ No se pudo obtener el usuario activo para UsuaAnul");
            }

            // Excluir campos de auditoría que no existen en BD
            _context.Entry(ordenCompra).Property(x => x.CodiEmprNet).IsModified = false;
            _context.Entry(ordenCompra).Property(x => x.NumUsuar).IsModified = false;
            _context.Entry(ordenCompra).Property(x => x.AddRecord).IsModified = false;
            
            // Excluir FGenOcom para que mantenga la fecha original de creación
            _context.Entry(ordenCompra).Property(x => x.FGenOcom).IsModified = false;
            Console.WriteLine($"[OcomencaController] DeleteAsync - Preservando FGenOcom original en borrado lógico");

            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Obtiene el siguiente número de orden disponible
        /// </summary>
        [HttpGet("siguiente-numero")]
        public async Task<ActionResult<int>> GetSiguienteNumeroAsync()
        {
            var maxNumero = await _context.Ocomencas.MaxAsync(x => (int?)x.NumeOcom) ?? 0;
            return Ok(maxNumero + 1);
        }

        /// <summary>
        /// Obtiene las monedas activas para el combo
        /// </summary>
        [HttpGet("monedas")]
        public async Task<ActionResult> GetMonedasAsync()
        {
            var monedas = await _context.Tamoneds
                .Where(m => m.Status >= 0)
                .OrderBy(m => m.Descripcion)
                .Select(m => new { m.CodigoMoneda, m.Descripcion, m.Cotizacion })
                .ToListAsync();

            return Ok(monedas);
        }

        /// <summary>
        /// Obtiene la configuración del modo de aprobación de órdenes de compra desde FLEXCRL
        /// </summary>
        [HttpGet("configuracion/modo-aprobacion")]
        public async Task<ActionResult<string>> GetModoAprobacionAsync()
        {
            var configuracion = await _context.Flexcrls
                .AsNoTracking()
                .FirstOrDefaultAsync(f => f.CLABUS == "OCOMPRA" && f.PROBUS == "MODOAPRO");

            if (configuracion == null)
            {
                // Valor por defecto si no existe configuración
                return Ok("DIRECTA");
            }

            return Ok(configuracion.VALCONTROL ?? "DIRECTA");
        }

        /// <summary>
        /// Verifica si existe una orden de compra de materiales asociada a una orden de fabricación
        /// Replica la validación VB6: CONDI$ = "ANEX_OCOM='" & NORFA$ & "' AND STAT_OCOM < 9"
        /// </summary>
        /// <param name="numeroOrdenFabricacion">Número de orden de fabricación (NORFA$)</param>
        /// <returns>
        /// true si existe orden de compra de materiales en curso (STAT_OCOM < 9)
        /// false si no existe orden de compra asociada
        /// </returns>
        [HttpGet("verificar-orden-compra-asociada/{numeroOrdenFabricacion}")]
        public async Task<ActionResult<bool>> VerificarOrdenCompraAsociadaAsync(string numeroOrdenFabricacion)
        {
            try
            {
                Console.WriteLine($"[OcomencaController] Verificando orden de compra asociada a O.F. {numeroOrdenFabricacion}");

                // Buscar en OCOMENCA donde ANEX_OCOM = NORFA$ y STAT_OCOM < 9
                var existeOrdenCompra = await _context.Ocomencas
                    .AsNoTracking()
                    .AnyAsync(o => o.AnexOcom == numeroOrdenFabricacion && (o.StatOcom ?? 0) < 9);

                Console.WriteLine($"[OcomencaController] Resultado: {(existeOrdenCompra ? "EXISTE" : "NO EXISTE")} orden de compra asociada");

                return Ok(existeOrdenCompra);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[OcomencaController] Error al verificar orden de compra asociada: {ex.Message}");
                return StatusCode(500, $"Error al verificar orden de compra asociada: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza una orden de compra completa siguiendo la lógica VB6 GRACOMPRA
        /// Maneja encabezado, detalle, IVA y validaciones en una transacción
        /// </summary>
        //[HttpPut("actualizar-completa/{id}")]
        //public async Task<IActionResult> ActualizarCompletaAsync(int id, [FromBody] Ocomenca ordenCompra)
        //{
        //    using var transaction = await _context.Database.BeginTransactionAsync();
        //    try
        //    {
        //        // 1. VALIDACIONES INICIALES (basadas en VB6)
        //        var ordenExistente = await _context.Ocomencas
        //            .FirstOrDefaultAsync(o => o.OCOMENCA_ID == id && o.StatOcom < 9); // Excluye anuladas

        //        if (ordenExistente == null)
        //        {
        //            return NotFound("No se ha encontrado registro de encabezado de esta O/Compra.");
        //        }

        //        // Validación de proveedor
        //        if (!ordenCompra.NproOcom.HasValue || ordenCompra.NproOcom <= 0)
        //        {
        //            return BadRequest("Número de Proveedor no Válido");
        //        }

        //        // Validación de fecha de emisión
        //        if (!ordenCompra.FemiOcom.HasValue || ordenCompra.FemiOcom > DateTime.Now)
        //        {
        //            return BadRequest("Fecha de Emisión no Válida");
        //        }

        //        // Validación de condición de pago
        //        if (ordenCompra.CpagOcom.HasValue && ordenCompra.CpagOcom > 255)
        //        {
        //            return BadRequest("Condición de Pago no Válida");
        //        }

        //        // 2. OBTENER CONFIGURACIONES (equivalente a Control$ del VB6)
        //        var modoAprobacion = await ObtenerConfiguracionAsync("MODOAPRO", "DIRECTA");
        //        var enviarSucursal = await ObtenerConfiguracionAsync("ENVARSUC", "NO");
        //        // TODO: Implementar otras configuraciones cuando se requieran
        //        // var actualizarCosto = await ObtenerConfiguracionAsync("ACOSREP", "SIEMPRE");
        //        // var considerarDescuento = await ObtenerConfiguracionAsync("NODESCUE", "NO");

        //        // 3. DETERMINAR ESTADO SEGÚN MODO DE APROBACIÓN (lógica VB6)
        //        short estadoOrden = 0; // DIRECTA por defecto
        //        if (modoAprobacion.StartsWith("POSTERIOR", StringComparison.OrdinalIgnoreCase))
        //        {
        //            estadoOrden = -1; // POSTERIOR
        //        }
        //        // TODO: Agregar validación para GRABA_PENDIENTE_APROBACION (-2) cuando se implemente

        //        // 4. ACTUALIZAR ENCABEZADO (equivalente a OCOMENCA del VB6)
        //        ordenExistente.RefeOcom = $"{ordenCompra.FemiOcom:dd/MM/yyyy} - {ordenCompra.RaSoProv}";
        //        ordenExistente.FemiOcom = ordenCompra.FemiOcom;
        //        ordenExistente.NumeCli = ordenCompra.NumeCli;
        //        ordenExistente.RaSoProv = ordenCompra.RaSoProv?.Substring(0, Math.Min(48, ordenCompra.RaSoProv.Length));
        //        ordenExistente.TSumOcom = 0; // Materiales
        //        ordenExistente.FModOcom = DateTime.Now;
        //        // UsuaMod se asignará automáticamente por el AuditInterceptor
                
        //        // Monedas y tipo de cambio
        //        ordenExistente.MoneEmis = ordenCompra.MoneEmis ?? 1;
        //        decimal tipoCambio = 1;
        //        if ((ordenCompra.MoneEmis ?? 1) <= 1)
        //        {
        //            if ((ordenCompra.MoneCos ?? 1) > 1)
        //            {
        //                // TODO: Implementar TICAMONE cuando esté disponible
        //                // tipoCambio = TICAMONE(ordenCompra.MoneCos.Value);
        //            }
        //        }
        //        ordenExistente.TcamOcom = tipoCambio;

        //        // Condiciones de pago y embalaje
        //        ordenExistente.CpagOcom = ordenCompra.CpagOcom;
        //        ordenExistente.EmbalaOcom = ordenCompra.EmbalaOcom;
        //        ordenExistente.AtSrOcom = ordenCompra.AtSrOcom;

        //        // Totales (estos vendrían del formulario con IVA calculado)
        //        ordenExistente.ToBruOcom = ordenCompra.ToBruOcom ?? 0;
        //        ordenExistente.PdesOcom = ordenCompra.PdesOcom;
        //        ordenExistente.IdesOcom = ordenCompra.IdesOcom ?? 0;
        //        ordenExistente.IivaOcom = ordenCompra.IivaOcom ?? 0;
        //        ordenExistente.Iotr1Ocom = ordenCompra.Iotr1Ocom ?? 0;
        //        ordenExistente.TotalOcom = ordenCompra.TotalOcom ?? 0;
        //        ordenExistente.Lentrega = ordenCompra.Lentrega;
        //        ordenExistente.ObseOcom = ordenCompra.ObseOcom;
        //        ordenExistente.PorcIva = ordenCompra.PorcIva; // Distribución de IVA
        //        ordenExistente.StatOcom = estadoOrden;
        //        ordenExistente.ResponOcompra = ordenCompra.ResponOcompra;
        //        ordenExistente.ViaEntrega = ordenCompra.ViaEntrega;
        //        ordenExistente.OComCliente = ordenCompra.OComCliente;
        //        ordenExistente.NumeCli = ordenCompra.NumeCli;

        //        // 5. ANULAR ITEMS EXISTENTES EN DETALLE (lógica VB6: Stat_Ocom = 10)
        //        var detallesExistentes = await _context.Ocomdetas
        //            .Where(d => d.NumeOcom == ordenExistente.NumeOcom && d.StatOcom < 9)
        //            .ToListAsync();

        //        foreach (var detalle in detallesExistentes)
        //        {
        //            detalle.StatOcom = 10; // Anulado para reemplazar
        //            // LastUpdate se asigna automáticamente por AuditInterceptor
        //        }

        //        // 6. LIMPIAR DETALLE DE IVA EXISTENTE (lógica VB6)
        //        var detallesIvaExistentes = await _context.Detaivas
        //            .Where(d => d.NumeOcom == ordenExistente.NumeOcom)
        //            .ToListAsync();

        //        _context.Detaivas.RemoveRange(detallesIvaExistentes);
        //        // 7. CREAR NUEVOS REGISTROS DE DETALLE IVA (equivalente a DETIVA del VB6)
        //        // TODO: Implementar cuando se tenga la distribución de IVA desde el frontend
        //        // Esta lógica vendría del modal de IVA del formulario
        //        /*
        //        if (!string.IsNullOrEmpty(ordenCompra.PorcIva))
        //        {
        //            // Parsear distribución de IVA y crear registros DETAIVA
        //            // Por ahora se omite hasta tener la funcionalidad completa del frontend
        //        }
        //        */

        //        // 8. GUARDAR CAMBIOS
        //        // Nota: Excluir campos de auditoría que no existen en BD
        //        _context.Entry(ordenExistente).Property(x => x.CodiEmprNet).IsModified = false;
        //        _context.Entry(ordenExistente).Property(x => x.NumUsuar).IsModified = false;
        //        _context.Entry(ordenExistente).Property(x => x.AddRecord).IsModified = false;
        //        _context.Entry(ordenExistente).Property(x => x.LastUpdate).IsModified = false;
                
        //        await _context.SaveChangesAsync();

        //        // 9. FUNCIONALIDADES PENDIENTES (comentadas según especificaciones)
        //        // TODO: Implementar cuando se requiera
        //        // - Consignaciones (MACONSIG)
        //        // - Actualización de costos en MASTER
        //        // - Transferencias a sucursales
        //        // - Validaciones de stock y pedidos
        //        // - Impresión y envío por email

        //        await transaction.CommitAsync();

        //        return Ok(new { 
        //            mensaje = "Orden de compra actualizada exitosamente",
        //            numeroOrden = ordenExistente.NumeOcom,
        //            estado = estadoOrden == 0 ? "APROBADA" : "PENDIENTE DE APROBACIÓN"
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        await transaction.RollbackAsync();
        //        return StatusCode(500, $"Error al actualizar la orden de compra: {ex.Message}");
        //    }
        //}

        ///// <summary>
        ///// Obtiene una configuración específica de FLEXCRL
        ///// Equivalente a Control$() del VB6
        ///// </summary>
        //private async Task<string> ObtenerConfiguracionAsync(string probus, string valorDefecto)
        //{
        //    try
        //    {
        //        var configuracion = await _context.Flexcrls
        //            .FirstOrDefaultAsync(f => f.CLABUS == "OCOMPRA" && f.PROBUS == probus);

        //        return configuracion?.VALCONTROL ?? valorDefecto;
        //    }
        //    catch
        //    {
        //        return valorDefecto;
        //    }
        //}

        /// <summary>
        /// Actualiza la configuración del modo de aprobación de órdenes de compra en FLEXCRL
        /// </summary>
        [HttpPut("configuracion/modo-aprobacion")]
        public async Task<IActionResult> SetModoAprobacionAsync([FromBody] string modoAprobacion)
        {
            // Validar que el modo sea válido
            if (modoAprobacion != "DIRECTA" && modoAprobacion != "POSTERIOR")
            {
                return BadRequest("El modo de aprobación debe ser 'DIRECTA' o 'POSTERIOR'");
            }

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                var configuracion = await _context.Flexcrls
                    .FirstOrDefaultAsync(f => f.CLABUS == "OCOMPRA" && f.PROBUS == "MODOAPRO");

                if (configuracion == null)
                {
                    // Crear nueva configuración
                    configuracion = new FLEXCRL
                    {
                        CLABUS = "OCOMPRA",
                        PROBUS = "MODOAPRO",
                        VALCONTROL = modoAprobacion,
                        CodiEmpr = null // Por defecto null para configuración global
                    };
                    _context.Flexcrls.Add(configuracion);
                }
                else
                {
                    // Actualizar configuración existente
                    configuracion.VALCONTROL = modoAprobacion;
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return NoContent();
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return StatusCode(500, $"Error al actualizar la configuración: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea una orden de compra completa con distribución de IVA
        /// NOTA: Este método es un alias del PostAsync principal para mantener compatibilidad
        /// </summary>
        [HttpPost("crear-completa")]
        public async Task<ActionResult<Ocomenca>> CrearOrdenCompletaAsync([FromBody] OrdenCompraCompletaDTO ordenCompleta)
        {
            // Delegar al método PostAsync principal que ya maneja todo
            return await PostAsync(ordenCompleta);
        }

        /// <summary>
        /// Actualiza una orden de compra completa con distribución de IVA
        /// Replica la lógica del código VB6 para actualizar DETAIVA
        /// </summary>
        [HttpPut("actualizar-completa-con-iva/{id}")]
        public async Task<IActionResult> ActualizarOrdenCompletaConIvaAsync(int id, [FromBody] OrdenCompraCompletaDTO ordenCompleta)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // 1. Validar que el ID coincida
                if (id != ordenCompleta.Encabezado.OCOMENCA_ID)
                {
                    return BadRequest("El ID no coincide con la orden de compra");
                }

                // 2. Validar la orden completa
                var errores = ordenCompleta.Validar();
                if (errores.Any())
                {
                    return BadRequest(new { errores });
                }

                // 3. Actualizar el encabezado usando el método existente
                var resultadoEncabezado = await PutAsync(id, ordenCompleta.Encabezado);
                if (resultadoEncabezado is not NoContentResult)
                {
                    return BadRequest("Error al actualizar el encabezado de la orden de compra");
                }

                // 4. Actualizar el detalle de IVA si está presente
                if (ordenCompleta.DistribucionIva != null && ordenCompleta.ActualizarIva)
                {
                    var numeroOrden = ordenCompleta.Encabezado.NumeOcom ?? 0;
                    var codigoEmpresa = ordenCompleta.CodigoEmpresa ?? ordenCompleta.Encabezado.CodiEmpr;

                    // Eliminar detalle IVA existente y crear nuevo (lógica VB6)
                    await _detaivaService.EliminarDetalleIvaAsync(numeroOrden);
                    
                    var resultadoIva = await _detaivaService.GrabarDetalleIvaAsync(
                        numeroOrden, 
                        codigoEmpresa, 
                        ordenCompleta.DistribucionIva);

                    if (!resultadoIva)
                    {
                        await transaction.RollbackAsync();
                        return StatusCode(500, "Error al actualizar el detalle de IVA");
                    }
                }

                await transaction.CommitAsync();
                return NoContent();
            }
            catch (UnauthorizedAccessException ex)
            {
                await transaction.RollbackAsync();
                
                // Logging del error de autorización
                Console.WriteLine($"[OcomencaController] 🔒 ERROR DE AUTORIZACIÓN EN ActualizarOrdenCompletaConIvaAsync - {ex.Message}");
                
                // Devolver 401 Unauthorized con el mensaje detallado del AuditInterceptor
                return Unauthorized(ex.Message);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return StatusCode(500, $"Error al actualizar la orden de compra completa: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el detalle de IVA de una orden de compra
        /// </summary>
        [HttpGet("{numeroOrden}/detalle-iva")]
        public async Task<ActionResult<DistribucionIvaDTO>> ObtenerDetalleIvaAsync(int numeroOrden)
        {
            try
            {
                var detalles = await _detaivaService.ObtenerDetalleIvaAsync(numeroOrden);
                var distribucion = _detaivaService.ConvertirADistribucionDto(detalles);
                
                return Ok(distribucion);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener el detalle de IVA: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza una orden de compra - VERSIÓN SIMPLIFICADA PARA TESTING
        /// Solo actualiza el encabezado en OCOMENCA con manejo completo de errores
        /// </summary>
        [HttpPut("actualizar-completa-simple/{id}")]
        public async Task<IActionResult> ActualizarCompletaSimpleAsync(int id, [FromBody] Ocomenca ordenCompra)
        {
            Console.WriteLine($"[ActualizarCompletaSimple] Iniciando actualización para ID: {id}");
            
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                Console.WriteLine($"[ActualizarCompletaSimple] Transacción iniciada");

                // 1. VALIDACIONES INICIALES
                Console.WriteLine($"[ActualizarCompletaSimple] Buscando orden existente con ID: {id}");
                var ordenExistente = await _context.Ocomencas
                    .FirstOrDefaultAsync(o => o.OCOMENCA_ID == id);

                if (ordenExistente == null)
                {
                    Console.WriteLine($"[ActualizarCompletaSimple] ERROR: No se encontró orden con ID: {id}");
                    await transaction.RollbackAsync();
                    return NotFound($"No se encontró la orden de compra con ID: {id}");
                }

                Console.WriteLine($"[ActualizarCompletaSimple] Orden encontrada. NumeOcom: {ordenExistente.NumeOcom}, Estado actual: {ordenExistente.StatOcom}");

                // Validar que no esté anulada
                if (ordenExistente.StatOcom >= 9)
                {
                    Console.WriteLine($"[ActualizarCompletaSimple] ERROR: La orden está anulada (Estado: {ordenExistente.StatOcom})");
                    await transaction.RollbackAsync();
                    return BadRequest("No se puede actualizar una orden anulada");
                }

                // 2. VALIDACIONES DE DATOS
                Console.WriteLine($"[ActualizarCompletaSimple] Validando datos de entrada");
                
                if (!ordenCompra.NproOcom.HasValue || ordenCompra.NproOcom <= 0)
                {
                    Console.WriteLine($"[ActualizarCompletaSimple] ERROR: Proveedor inválido: {ordenCompra.NproOcom}");
                    await transaction.RollbackAsync();
                    return BadRequest("Número de Proveedor no válido");
                }

                if (!ordenCompra.FemiOcom.HasValue)
                {
                    Console.WriteLine($"[ActualizarCompletaSimple] ERROR: Fecha de emisión no proporcionada");
                    await transaction.RollbackAsync();
                    return BadRequest("Fecha de emisión es obligatoria");
                }

                if (ordenCompra.FemiOcom.Value.Date > DateTime.Now.Date)
                {
                    Console.WriteLine($"[ActualizarCompletaSimple] ERROR: Fecha futura: {ordenCompra.FemiOcom}");
                    await transaction.RollbackAsync();
                    return BadRequest("La fecha de emisión no puede ser futura");
                }

                Console.WriteLine($"[ActualizarCompletaSimple] Validaciones completadas exitosamente");

                // 3. ACTUALIZAR CAMPOS DEL ENCABEZADO
                Console.WriteLine($"[ActualizarCompletaSimple] Actualizando campos del encabezado");
                
                // Campos básicos
                ordenExistente.NproOcom = ordenCompra.NproOcom;
                ordenExistente.FemiOcom = ordenCompra.FemiOcom;
                ordenExistente.RaSoProv = ordenCompra.RaSoProv?.Substring(0, Math.Min(48, ordenCompra.RaSoProv.Length));
                ordenExistente.RefeOcom = $"{ordenCompra.FemiOcom:dd/MM/yyyy} - {ordenCompra.RaSoProv}";
                
                // Fechas y tiempos
                ordenExistente.FModOcom = DateTime.Now;
                
                // Monedas
                ordenExistente.MoneEmis = ordenCompra.MoneEmis ?? 1;
                ordenExistente.MoneCos = ordenCompra.MoneCos ?? 1;
                ordenExistente.TcamOcom = ordenCompra.TcamOcom ?? 1;
                
                // Condiciones
                ordenExistente.CpagOcom = ordenCompra.CpagOcom;
                ordenExistente.EmbalaOcom = ordenCompra.EmbalaOcom;
                ordenExistente.AtSrOcom = ordenCompra.AtSrOcom;
                
                // Totales
                ordenExistente.ToBruOcom = ordenCompra.ToBruOcom ?? 0;
                ordenExistente.PdesOcom = ordenCompra.PdesOcom;
                ordenExistente.IdesOcom = ordenCompra.IdesOcom ?? 0;
                ordenExistente.IivaOcom = ordenCompra.IivaOcom ?? 0;
                ordenExistente.Iotr1Ocom = ordenCompra.Iotr1Ocom ?? 0;
                ordenExistente.TotalOcom = ordenCompra.TotalOcom ?? 0;
                
                // Otros campos
                ordenExistente.Lentrega = ordenCompra.Lentrega;
                ordenExistente.ObseOcom = ordenCompra.ObseOcom;
                ordenExistente.ResponOcompra = ordenCompra.ResponOcompra;
                ordenExistente.ViaEntrega = ordenCompra.ViaEntrega;
                ordenExistente.OComCliente = ordenCompra.OComCliente;
                ordenExistente.NumeCli = ordenCompra.NumeCli;

                Console.WriteLine($"[ActualizarCompletaSimple] Campos actualizados. Preparando para guardar...");

                // 4. EXCLUIR CAMPOS DE AUDITORÍA QUE NO EXISTEN EN BD
                Console.WriteLine($"[ActualizarCompletaSimple] Excluyendo campos de auditoría");
                _context.Entry(ordenExistente).Property(x => x.CodiEmprNet).IsModified = false;
                _context.Entry(ordenExistente).Property(x => x.NumUsuar).IsModified = false;
                _context.Entry(ordenExistente).Property(x => x.AddRecord).IsModified = false;
                _context.Entry(ordenExistente).Property(x => x.LastUpdate).IsModified = false;

                // 5. GUARDAR CAMBIOS
                Console.WriteLine($"[ActualizarCompletaSimple] Guardando cambios en la base de datos");
                await _context.SaveChangesAsync();
                
                Console.WriteLine($"[ActualizarCompletaSimple] Cambios guardados. Confirmando transacción");
                await transaction.CommitAsync();

                Console.WriteLine($"[ActualizarCompletaSimple] Actualización completada exitosamente");
                return Ok(new { 
                    mensaje = "Orden de compra actualizada exitosamente (versión simple)",
                    id = ordenExistente.OCOMENCA_ID,
                    numeroOrden = ordenExistente.NumeOcom,
                    proveedor = ordenExistente.RaSoProv,
                    fechaEmision = ordenExistente.FemiOcom,
                    total = ordenExistente.TotalOcom,
                    estado = ordenExistente.StatOcom
                });
            }
            catch (DbUpdateException dbEx)
            {
                Console.WriteLine($"[ActualizarCompletaSimple] ERROR DE BASE DE DATOS: {dbEx.Message}");
                Console.WriteLine($"[ActualizarCompletaSimple] Inner Exception: {dbEx.InnerException?.Message}");
                Console.WriteLine($"[ActualizarCompletaSimple] Stack Trace: {dbEx.StackTrace}");
                
                await transaction.RollbackAsync();
                return StatusCode(500, new { 
                    error = "Error de base de datos",
                    mensaje = dbEx.Message,
                    detalles = dbEx.InnerException?.Message,
                    tipo = "DbUpdateException"
                });
            }
            catch (SqlException sqlEx)
            {
                Console.WriteLine($"[ActualizarCompletaSimple] ERROR SQL: {sqlEx.Message}");
                Console.WriteLine($"[ActualizarCompletaSimple] SQL Error Number: {sqlEx.Number}");
                Console.WriteLine($"[ActualizarCompletaSimple] SQL State: {sqlEx.State}");
                Console.WriteLine($"[ActualizarCompletaSimple] Stack Trace: {sqlEx.StackTrace}");
                
                await transaction.RollbackAsync();
                return StatusCode(500, new { 
                    error = "Error SQL",
                    mensaje = sqlEx.Message,
                    numero = sqlEx.Number,
                    estado = sqlEx.State,
                    tipo = "SqlException"
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ActualizarCompletaSimple] ERROR GENERAL: {ex.Message}");
                Console.WriteLine($"[ActualizarCompletaSimple] Tipo: {ex.GetType().Name}");
                Console.WriteLine($"[ActualizarCompletaSimple] Stack Trace: {ex.StackTrace}");
                
                await transaction.RollbackAsync();
                return StatusCode(500, new { 
                    error = "Error interno del servidor",
                    mensaje = ex.Message,
                    tipo = ex.GetType().Name,
                    stackTrace = ex.StackTrace
                });
            }
        }

        /// <summary>
        /// Marca los items anteriores de una orden como inactivos (STAT_OCOM = 10)
        /// Se usa en ediciones para mantener historial
        /// </summary>
        private async Task MarcarItemsAnterioresComoInactivos(int numeroOrden)
        {
            Console.WriteLine($"[MarcarItemsAnterioresComoInactivos] 🔍 DEBUGGING - Buscando items anteriores para O.C. {numeroOrden}");
            
            var itemsAnteriores = await _context.Ocomdetas
                .Where(d => d.NumeOcom == numeroOrden && d.StatOcom == 0)
                .ToListAsync();

            Console.WriteLine($"[MarcarItemsAnterioresComoInactivos] 📊 DEBUGGING - Items anteriores encontrados: {itemsAnteriores.Count}");

            foreach (var item in itemsAnteriores)
            {
                Console.WriteLine($"[MarcarItemsAnterioresComoInactivos] 📋 DEBUGGING - Marcando item: NuOrdItem={item.NuOrdItem}, CodInte={item.CodInte}, StatOcom={item.StatOcom} -> 10");
                item.StatOcom = 10; // Marcar como inactivo
            }

            if (itemsAnteriores.Any())
            {
                await _context.SaveChangesAsync();
                Console.WriteLine($"[MarcarItemsAnterioresComoInactivos] ✅ Marcados {itemsAnteriores.Count} items como inactivos para orden {numeroOrden}");
            }
            else
            {
                Console.WriteLine($"[MarcarItemsAnterioresComoInactivos] ⚠️ No se encontraron items anteriores para marcar como inactivos");
            }
        }

        /// <summary>
        /// Graba el detalle de items de la orden de compra
        /// Todos los nuevos items se crean con STAT_OCOM = 0 (activos)
        /// </summary>
        private async Task GrabarDetalleOrdenCompra(List<Ocomdeta> detalle, int numeroOrden)
        {
            foreach (var item in detalle)
            {
                // ===== CAMPOS OBLIGATORIOS BÁSICOS =====
                item.NumeOcom = numeroOrden;
                item.StatOcom = 0; // Nuevo item activo
                item.FemiOcom = DateTime.Now;
                
                // ===== CAMPOS DE EMPRESA Y AUDITORÍA =====
                item.CodiEmpr = 0; // DEBE SER 0 según requerimiento
                
                // ===== CAMPOS DE UNIDADES =====
                // VALO_UNID - Valor unitario (string, convertir de PreUnit si no está establecido)
                item.ValoUnid = item.ValoUnid ?? item.PreUnit?.ToString("F4") ?? "0";
                
                // UNIDS_OCOM - Unidades de la orden de compra (por defecto igual a CantOcom)
                item.UnidsOcom = item.UnidsOcom ?? item.CantOcom ?? 0;
                
                // Coef_Unids - Coeficiente de unidades DEBE SER 1
                item.CoefUnids = 1;
                
                // ===== CAMPOS DE MONEDA =====
                // MONE_COS - Moneda de costo (mantener el valor original o usar 2 por defecto)
                // Si no está establecido, usar 2 (USD), si está establecido mantenerlo
                item.MoneCos = item.MoneCos ?? 2;
                
                // ===== CÁLCULOS DE PRECIOS =====
                if (item.CantOcom.HasValue && item.PreUnit.HasValue)
                {
                    var importeBruto = item.CantOcom.Value * item.PreUnit.Value;
                    var descuento = item.DescItem ?? 0;
                    item.INetItem = importeBruto - descuento;
                    
                    // PrUnNeto - Precio unitario neto
                    item.PrUnNeto = item.PreUnit.Value - (descuento / (item.CantOcom.Value > 0 ? item.CantOcom.Value : 1));
                    
                    // PRUN_NETO_PES - Precio unitario neto en pesos (convertir según moneda)
                    if (item.MoneCos == 1) // Pesos argentinos
                    {
                        item.PrUnNetoPes = item.PrUnNeto;
                    }
                    else // Otras monedas (USD, etc.) - convertir usando cotización
                    {
                        // Obtener cotización de la moneda
                        var moneda = await _context.Tamoneds
                            .FirstOrDefaultAsync(m => m.CodigoMoneda == item.MoneCos);
                        var cotizacion = moneda?.Cotizacion ?? 1m;
                        item.PrUnNetoPes = (item.PrUnNeto ?? 0) * cotizacion;
                    }
                }
                else
                {
                    // Si no hay cantidad o precio, establecer valores por defecto
                    item.PrUnNeto = item.PreUnit ?? 0;
                    item.PrUnNetoPes = item.PreUnit ?? 0;
                }
                
                // ===== CAMPOS DE FECHAS ADICIONALES =====
                item.AddRecord = DateTime.Now;
                item.LastUpdate = DateTime.Now;
                
                // ===== GENERAR ROWGUID =====
                if (item.Rowguid == Guid.Empty)
                {
                    item.Rowguid = Guid.NewGuid();
                }

                _context.Ocomdetas.Add(item);
                
                Console.WriteLine($"[GrabarDetalleOrdenCompra] Item configurado - CodInte: {item.CodInte}, CodiEmpr: {item.CodiEmpr}, MoneCos: {item.MoneCos}, CoefUnids: {item.CoefUnids}");
            }

            await _context.SaveChangesAsync();
            Console.WriteLine($"[GrabarDetalleOrdenCompra] Grabados {detalle.Count} items para orden {numeroOrden}");
        }

        /// <summary>
        /// Graba la distribución de IVA en la tabla DETAIVA desde DistribucionIvaDTO
        /// Crea un registro por cada tasa de IVA aplicada
        /// </summary>
        private async Task GrabarDistribucionIvaDesdeDTO(DistribucionIvaDTO distribucionIva, int numeroOrden)
        {
            // Eliminar distribución anterior si existe
            var distribucionAnterior = await _context.Detaivas
                .Where(d => d.NumeOcom == numeroOrden)
                .ToListAsync();

            if (distribucionAnterior.Any())
            {
                _context.Detaivas.RemoveRange(distribucionAnterior);
            }

            // Crear registros según las tasas del DTO
            int norit = 0;

            // 0% (Exento/No Gravado)
            if (distribucionIva.BaseImponible0 > 0)
            {
                var detaiva = new Detaiva
                {
                    CodiEmpr = 0, // Por defecto 0
                    NumeOcom = numeroOrden,
                    Norit = norit++,
                    TasaIva = 0.00m, // 0.00%
                    BaseImpo = distribucionIva.BaseImponible0,
                    ImpoIva = 0.00m
                };
                _context.Detaivas.Add(detaiva);
            }

            // 10.5%
            if (distribucionIva.BaseImponible105 > 0)
            {
                var detaiva = new Detaiva
                {
                    CodiEmpr = 0, // Por defecto 0
                    NumeOcom = numeroOrden,
                    Norit = norit++,
                    TasaIva = 10.50m, // 10.50%
                    BaseImpo = distribucionIva.BaseImponible105,
                    ImpoIva = distribucionIva.ImporteIva105
                };
                _context.Detaivas.Add(detaiva);
            }

            // 21%
            if (distribucionIva.BaseImponible21 > 0)
            {
                var detaiva = new Detaiva
                {
                    CodiEmpr = 0, // Por defecto 0
                    NumeOcom = numeroOrden,
                    Norit = norit++,
                    TasaIva = 21.00m, // 21.00%
                    BaseImpo = distribucionIva.BaseImponible21,
                    ImpoIva = distribucionIva.ImporteIva21
                };
                _context.Detaivas.Add(detaiva);
            }

            // 27%
            if (distribucionIva.BaseImponible27 > 0)
            {
                var detaiva = new Detaiva
                {
                    CodiEmpr = 0, // Por defecto 0
                    NumeOcom = numeroOrden,
                    Norit = norit++,
                    TasaIva = 27.00m, // 27.00%
                    BaseImpo = distribucionIva.BaseImponible27,
                    ImpoIva = distribucionIva.ImporteIva27
                };
                _context.Detaivas.Add(detaiva);
            }

            await _context.SaveChangesAsync();
            Console.WriteLine($"[GrabarDistribucionIvaDesdeDTO] Grabados {norit} registros de IVA para orden {numeroOrden}");
        }

        /// <summary>
        /// Graba la distribución de IVA en la tabla DETAIVA
        /// Crea un registro por cada tasa de IVA aplicada
        /// </summary>
        private async Task GrabarDistribucionIva(List<DetaivaDTO> distribucionIva, int numeroOrden)
        {
            // Eliminar distribución anterior si existe
            var distribucionAnterior = await _context.Detaivas
                .Where(d => d.NumeOcom == numeroOrden)
                .ToListAsync();

            if (distribucionAnterior.Any())
            {
                _context.Detaivas.RemoveRange(distribucionAnterior);
            }

            // Crear nuevos registros de distribución
            int norit = 0;
            foreach (var item in distribucionIva.Where(d => d.BaseImpo > 0))
            {
                var detaiva = new Detaiva
                {
                    CodiEmpr = 0, // Por defecto 0
                    NumeOcom = numeroOrden,
                    Norit = norit++,
                    TasaIva = item.TasaIva ?? 0,
                    BaseImpo = item.BaseImpo ?? 0,
                    ImpoIva = item.ImpoIva ?? 0
                };

                _context.Detaivas.Add(detaiva);
            }

            await _context.SaveChangesAsync();
            Console.WriteLine($"[GrabarDistribucionIva] Grabados {distribucionIva.Count} registros de IVA para orden {numeroOrden}");
        }

        /// <summary>
        /// Calcula y asigna los campos adicionales del encabezado de la orden
        /// ToBruOcom: Total antes del descuento y del IVA
        /// IivaOcom: Suma de los importes de IVA
        /// PorcIva: Porcentaje de IVA aplicado, o "varios" si hay múltiples tasas
        /// </summary>
        private void CalcularCamposAdicionales(Ocomenca ordenCompra, OrdenCompraCompletaDTO ordenCompleta)
        {
            // 1. ToBruOcom: Total bruto (suma de items sin descuento ni IVA)
            decimal totalBruto = 0;
            if (ordenCompleta.Detalle != null && ordenCompleta.Detalle.Any())
            {
                totalBruto = ordenCompleta.Detalle.Sum(item => 
                    (item.CantOcom ?? 0) * (item.PreUnit ?? 0));
            }
            ordenCompra.ToBruOcom = totalBruto;

            // 2. IivaOcom e PorcIva: Calcular desde distribución de IVA
            if (ordenCompleta.DistribucionIva != null && ordenCompleta.ActualizarIva)
            {
                // Suma de importes de IVA
                ordenCompra.IivaOcom = ordenCompleta.DistribucionIva.TotalIva;

                // Determinar porcentaje de IVA
                var tasasAplicadas = new List<decimal>();
                
                if (ordenCompleta.DistribucionIva.BaseImponible0 > 0)
                    tasasAplicadas.Add(0.00m);
                if (ordenCompleta.DistribucionIva.BaseImponible105 > 0)
                    tasasAplicadas.Add(10.50m);
                if (ordenCompleta.DistribucionIva.BaseImponible21 > 0)
                    tasasAplicadas.Add(21.00m);
                if (ordenCompleta.DistribucionIva.BaseImponible27 > 0)
                    tasasAplicadas.Add(27.00m);

                if (tasasAplicadas.Count == 0)
                {
                    ordenCompra.PorcIva = "0.00%";
                }
                else if (tasasAplicadas.Count == 1)
                {
                    ordenCompra.PorcIva = $"{tasasAplicadas[0]:F2}%";
                }
                else
                {
                    ordenCompra.PorcIva = "varios";
                }
            }
            else
            {
                // Sin distribución de IVA
                ordenCompra.IivaOcom = 0;
                ordenCompra.PorcIva = "0.00%";
            }

            Console.WriteLine($"[CalcularCamposAdicionales] ToBruOcom: {ordenCompra.ToBruOcom:C2}");
            Console.WriteLine($"[CalcularCamposAdicionales] IivaOcom: {ordenCompra.IivaOcom:C2}");
            Console.WriteLine($"[CalcularCamposAdicionales] PorcIva: {ordenCompra.PorcIva}");
        }

        /// <summary>
        /// Obtiene el último destino y sector utilizados para un proveedor específico
        /// Busca en las órdenes de compra más recientes del proveedor donde haya datos válidos
        /// </summary>
        /// <param name="numeroProveedor">Número del proveedor</param>
        /// <returns>DTO con el último destino y sector utilizados</returns>
        [HttpGet("ultimo-destino-sector/{numeroProveedor}")]
        public async Task<ActionResult<UltimoDestinoSectorDTO>> GetUltimoDestinoSectorAsync(int numeroProveedor)
        {
            try
            {
                Console.WriteLine($"[GetUltimoDestinoSectorAsync] Buscando último destino y sector para proveedor: {numeroProveedor}");

                // Buscar directamente en OCOMDETA la orden más reciente del proveedor 
                // que tenga datos válidos de destino o sector
                var ultimoDetalle = await _context.Ocomdetas
                    .Where(d => d.NproOcom == numeroProveedor &&
                               d.StatOcom < 9 &&
                               (d.IContab > 0 || !string.IsNullOrEmpty(d.SectOcom)))
                    .OrderByDescending(d => d.FemiOcom) // Más reciente primero por fecha
                    .ThenByDescending(d => d.NumeOcom)  // Luego por número de orden
                    .ThenByDescending(d => d.NuOrdItem) // Finalmente por número de item
                    .FirstOrDefaultAsync();

                if (ultimoDetalle == null)
                {
                    Console.WriteLine($"[GetUltimoDestinoSectorAsync] No se encontraron órdenes con datos válidos para el proveedor {numeroProveedor}");
                    return Ok(new UltimoDestinoSectorDTO()); // Retornar DTO vacío
                }

                var resultado = new UltimoDestinoSectorDTO
                {
                    UltimoDestino = ultimoDetalle.IContab > 0 ? ultimoDetalle.IContab : null, // Solo si > 0
                    UltimoSector = !string.IsNullOrEmpty(ultimoDetalle.SectOcom) ? ultimoDetalle.SectOcom : string.Empty, // Solo si no está vacío
                    FechaUltimaOrden = ultimoDetalle.FemiOcom,
                    NumeroUltimaOrden = ultimoDetalle.NumeOcom
                };

                Console.WriteLine($"[GetUltimoDestinoSectorAsync] ✅ Encontrado - Destino: {resultado.UltimoDestino}, Sector: '{resultado.UltimoSector}', Fecha: {resultado.FechaUltimaOrden}, Orden: {resultado.NumeroUltimaOrden}");

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GetUltimoDestinoSectorAsync] ❌ Error: {ex.Message}");
                return StatusCode(500, $"Error al obtener último destino y sector: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el "Atte. Señor" de la última orden de compra del proveedor
        /// Busca en las órdenes más recientes del proveedor el campo AtSr_Ocom
        /// </summary>
        /// <param name="numeroProveedor">Número del proveedor</param>
        /// <returns>Atte. Señor de la última orden o cadena vacía si no se encuentra</returns>
        [HttpGet("ultimo-atte-senor/{numeroProveedor}")]
        public async Task<IActionResult> GetUltimoAtteSenorAsync(int numeroProveedor)
        {
            try
            {
                Console.WriteLine($"[GetUltimoAtteSenorAsync] Buscando Atte. Señor para proveedor: {numeroProveedor}");

                // Buscar la orden de compra más reciente del proveedor que tenga AtSr_Ocom
                var ultimaOrden = await _context.Ocomencas
                    .Where(o => o.NproOcom == numeroProveedor && 
                               o.StatOcom >= 0 && 
                               !string.IsNullOrWhiteSpace(o.AtSrOcom))
                    .OrderByDescending(o => o.FemiOcom)
                    .ThenByDescending(o => o.NumeOcom)
                    .FirstOrDefaultAsync();

                if (ultimaOrden != null && !string.IsNullOrWhiteSpace(ultimaOrden.AtSrOcom))
                {
                    Console.WriteLine($"[GetUltimoAtteSenorAsync] ✅ Encontrado en O.C. {ultimaOrden.NumeOcom}: '{ultimaOrden.AtSrOcom}'");
                    return Content(ultimaOrden.AtSrOcom, "text/plain");
                }

                Console.WriteLine($"[GetUltimoAtteSenorAsync] ⚠️ No se encontró Atte. Señor en órdenes del proveedor {numeroProveedor}");
                return Content(string.Empty, "text/plain");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GetUltimoAtteSenorAsync] ❌ Error: {ex.Message}");
                return StatusCode(500, $"Error al obtener Atte. Señor: {ex.Message}");
            }
        }

        /// <summary>
        /// MÉTODO DE DIAGNÓSTICO: Verifica si un código tiene estructura en la tabla Formulas
        /// </summary>
        /// <param name="codigoInterno">Código interno a verificar</param>
        /// <returns>Información de diagnóstico sobre la estructura</returns>
        [HttpGet("diagnostico/estructura/{codigoInterno}")]
        public async Task<ActionResult> DiagnosticoEstructuraAsync(int codigoInterno)
        {
            try
            {
                Console.WriteLine($"[DiagnosticoEstructura] Verificando estructura para código: {codigoInterno}");

                // Verificar si existe en la tabla Formulas
                var totalFormulas = await _context.Formulas
                    .Where(f => f.CodIConj == codigoInterno)
                    .CountAsync();

                var formulasActivas = await _context.Formulas
                    .Where(f => f.CodIConj == codigoInterno && f.Marborra != 1)
                    .Select(f => new
                    {
                        f.CodIConj,
                        f.CodXConj,
                        f.CodiElem,
                        f.CodXElem,
                        f.UsoBruto,
                        f.Unis_Form,
                        f.Marborra
                    })
                    .ToListAsync();

                return Ok(new
                {
                    codigoInterno = codigoInterno,
                    totalRegistros = totalFormulas,
                    registrosActivos = formulasActivas.Count,
                    componentes = formulasActivas
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DiagnosticoEstructura] ❌ Error: {ex.Message}");
                return StatusCode(500, $"Error en diagnóstico: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene las consignaciones (MACONSIG) generadas para una orden de compra específica
        /// </summary>
        /// <param name="numeroOrden">Número de orden de compra</param>
        /// <returns>Lista de consignaciones asociadas a la orden</returns>
        [HttpGet("{numeroOrden}/consignaciones")]
        public async Task<ActionResult> GetConsignacionesOrdenAsync(int numeroOrden)
        {
            try
            {
                Console.WriteLine($"[GetConsignacionesOrdenAsync] Consultando consignaciones para O.C. {numeroOrden}");

                // Buscar consignaciones que coincidan con el patrón de la orden
                string patronOrden = $"OC-{numeroOrden.ToString().PadLeft(6, '0')}-";
                
                var consignaciones = await _context.Maconsigs
                    .Where(c => c.NroOcom != null && c.NroOcom.StartsWith(patronOrden))
                    .OrderBy(c => c.NuOrdItem)
                    .Select(c => new
                    {
                        c.Id,
                        c.NroOcom,
                        c.CodInte,
                        c.CodExte,
                        c.Descritem,
                        c.Unidad,
                        c.CantAsig,
                        c.ValoUnit,
                        c.ValoTot,
                        c.Moneda,
                        c.FemiOcom,
                        c.NuOrdItem
                    })
                    .ToListAsync();

                Console.WriteLine($"[GetConsignacionesOrdenAsync] ✅ Encontradas {consignaciones.Count} consignaciones para O.C. {numeroOrden}");

                return Ok(new
                {
                    numeroOrden = numeroOrden,
                    totalConsignaciones = consignaciones.Count,
                    consignaciones = consignaciones
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GetConsignacionesOrdenAsync] ❌ Error: {ex.Message}");
                return StatusCode(500, $"Error al obtener consignaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Genera un PDF de la orden de compra usando iTextSharp
        /// </summary>
        /// <param name="numeroOrden">Número de la orden de compra</param>
        /// <returns>PDF de la orden de compra en formato base64</returns>
        [HttpGet("{numeroOrden}/pdf")]
        public async Task<ActionResult> GenerarPdfOrdenCompraAsync(int numeroOrden)
        {
            try
            {
                Console.WriteLine($"[GenerarPdfOrdenCompraAsync] Generando PDF para O.C. {numeroOrden}");

                // Obtener la orden de compra
                var orden = await _context.Ocomencas
                    .FirstOrDefaultAsync(o => o.NumeOcom == numeroOrden);

                if (orden == null)
                {
                    return NotFound($"No se encontró la orden de compra {numeroOrden}");
                }

                // Obtener el detalle de la orden (solo items activos con StatOcom < 9)
                var detalle = await _context.Ocomdetas
                    .Where(d => d.NumeOcom == numeroOrden && d.StatOcom < 9)
                    .OrderBy(d => d.NuOrdItem)
                    .ToListAsync();

                // Obtener datos de la empresa
                var empresa = await _context.Empresas
                    .FirstOrDefaultAsync(e => e.status == 1);

                Console.WriteLine($"[GenerarPdfOrdenCompraAsync] 📋 Items encontrados: {detalle.Count}");
                Console.WriteLine($"[GenerarPdfOrdenCompraAsync] 📄 Orden: {orden.NumeOcom}, Fecha: {orden.FemiOcom?.ToString("dd/MM/yyyy")}");
                Console.WriteLine($"[GenerarPdfOrdenCompraAsync] 👤 Proveedor: {orden.RaSoProv}, N°: {orden.NproOcom}");
                Console.WriteLine($"[GenerarPdfOrdenCompraAsync] 💰 Moneda: {orden.MoneEmis}, Condición: {orden.CpagOcom}");
                Console.WriteLine($"[GenerarPdfOrdenCompraAsync] 🏢 Empresa: {empresa?.razon ?? "NO ENCONTRADA"}");

                // Generar PDF replicando el formato original (Imagen 1)
                using (MemoryStream ms = new MemoryStream())
                {
                    Document document = new Document(PageSize.A4, 40, 40, 30, 30);
                    PdfWriter writer = PdfWriter.GetInstance(document, ms);

                    document.Open();
                    Console.WriteLine($"[GenerarPdfOrdenCompraAsync] 📝 Documento PDF abierto");

                    var normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);
                    var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9);
                    var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 11);
                    var smallFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);

                    // ===== ENCABEZADO: Dos columnas =====
                    PdfPTable headerTable = new PdfPTable(2);
                    headerTable.WidthPercentage = 100;
                    headerTable.SetWidths(new float[] { 65f, 35f });

                    // COLUMNA IZQUIERDA: Datos de la empresa
                    PdfPCell leftCell = new PdfPCell();
                    leftCell.Border = Rectangle.BOX;
                    leftCell.BorderWidth = 1f;
                    leftCell.MinimumHeight = 80f;
                    leftCell.Padding = 8;
                    leftCell.VerticalAlignment = Element.ALIGN_TOP;
                    
                    if (empresa != null)
                    {
                        // Crear tabla interna para el contenido de la empresa
                        PdfPTable empresaTable = new PdfPTable(1);
                        empresaTable.WidthPercentage = 100;
                        
                        // Nombre de la empresa en naranja
                        var empresaFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.ORANGE);
                        var nombreCell = new PdfPCell(new Phrase(empresa.razon ?? "SIAM ARCON", empresaFont));
                        nombreCell.Border = Rectangle.NO_BORDER;
                        nombreCell.Padding = 0;
                        nombreCell.PaddingBottom = 3;
                        empresaTable.AddCell(nombreCell);
                        
                        // Información de la empresa
                        string infoTexto = $"Dirección: {empresa.calle} {empresa.numero} {empresa.complemento}\n" +
                                          $"({empresa.CP}) {empresa.localidad} - Argentina\n" +
                                          $"Teléfono: {empresa.TEL1} - Email: {empresa.email}\n" +
                                          $"CUIT: {empresa.CUIT}";
                        
                        var infoCell = new PdfPCell(new Phrase(infoTexto, smallFont));
                        infoCell.Border = Rectangle.NO_BORDER;
                        infoCell.Padding = 0;
                        empresaTable.AddCell(infoCell);
                        
                        leftCell.AddElement(empresaTable);
                    }
                    
                    headerTable.AddCell(leftCell);

                    // COLUMNA DERECHA: Cuadro de "PEDIDO DE SUMINISTRO"
                    PdfPCell rightCell = new PdfPCell();
                    rightCell.Border = Rectangle.BOX;
                    rightCell.BorderWidth = 2f;
                    rightCell.Padding = 8;
                    rightCell.HorizontalAlignment = Element.ALIGN_CENTER;

                    var pedidoTitle = new Paragraph("PEDIDO DE SUMINISTRO", titleFont);
                    pedidoTitle.Alignment = Element.ALIGN_CENTER;
                    rightCell.AddElement(pedidoTitle);

                    var nroBox = new PdfPCell(new Phrase($"Nro: {orden.NumeOcom}", boldFont));
                    nroBox.HorizontalAlignment = Element.ALIGN_CENTER;
                    nroBox.Border = Rectangle.BOX;
                    nroBox.Padding = 5;
                    var nroTable = new PdfPTable(1);
                    nroTable.WidthPercentage = 100;
                    nroTable.AddCell(nroBox);
                    rightCell.AddElement(nroTable);

                    var fechaPara = new Paragraph($"Fecha: {orden.FemiOcom?.ToString("dd/MM/yyyy") ?? "N/A"}", boldFont);
                    fechaPara.Alignment = Element.ALIGN_CENTER;
                    fechaPara.SpacingBefore = 5;
                    rightCell.AddElement(fechaPara);

                    headerTable.AddCell(rightCell);
                    document.Add(headerTable);
                    document.Add(new Paragraph(" ", smallFont)); // Espacio pequeño

                    // ===== SECCIÓN "SEÑORES" Y DATOS DEL PROVEEDOR =====
                    PdfPTable proveedorTable = new PdfPTable(2);
                    proveedorTable.WidthPercentage = 100;
                    proveedorTable.SetWidths(new float[] { 65f, 35f });
                    proveedorTable.SpacingBefore = 5;

                    // COLUMNA IZQUIERDA: Señores (Proveedor)
                    PdfPCell senoresCell = new PdfPCell();
                    senoresCell.Border = Rectangle.BOX;
                    senoresCell.Padding = 5;

                    var senoresHeader = new Paragraph("Señores:", boldFont);
                    senoresHeader.SpacingAfter = 3;
                    senoresCell.AddElement(senoresHeader);

                    var proveedorNombre = new Paragraph(orden.RaSoProv ?? "", boldFont);
                    senoresCell.AddElement(proveedorNombre);

                    var responsable = new Paragraph($"Responsable/Contacto: {orden.ObseOcom ?? "N/A"}", smallFont);
                    senoresCell.AddElement(responsable);

                    proveedorTable.AddCell(senoresCell);

                    // COLUMNA DERECHA: Cuenta, Condición, Moneda
                    PdfPTable datosTable = new PdfPTable(2);
                    datosTable.WidthPercentage = 100;
                    datosTable.SetWidths(new float[] { 40f, 60f });

                    var descripcionMoneda = orden.MoneEmis switch
                    {
                        1 => "Pesos",
                        2 => "Dólares",
                        3 => "Euros",
                        _ => "Pesos"
                    };

                    // Cuenta
                    var cuentaLabel = new PdfPCell(new Phrase("Cuenta:", boldFont));
                    cuentaLabel.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cuentaLabel.Padding = 3;
                    datosTable.AddCell(cuentaLabel);
                    datosTable.AddCell(new PdfPCell(new Phrase(orden.NproOcom?.ToString() ?? "", normalFont)) { Padding = 3 });

                    // Condición
                    var condicionLabel = new PdfPCell(new Phrase("Condición:", boldFont));
                    condicionLabel.BackgroundColor = BaseColor.LIGHT_GRAY;
                    condicionLabel.Padding = 3;
                    datosTable.AddCell(condicionLabel);
                    datosTable.AddCell(new PdfPCell(new Phrase(orden.CpagOcom?.ToString() ?? "N/A", normalFont)) { Padding = 3 });

                    // Moneda
                    var monedaLabel = new PdfPCell(new Phrase("Moneda:", boldFont));
                    monedaLabel.BackgroundColor = BaseColor.LIGHT_GRAY;
                    monedaLabel.Padding = 3;
                    datosTable.AddCell(monedaLabel);
                    datosTable.AddCell(new PdfPCell(new Phrase(descripcionMoneda, normalFont)) { Padding = 3 });

                    PdfPCell datosCell = new PdfPCell(datosTable);
                    datosCell.Border = Rectangle.NO_BORDER;
                    proveedorTable.AddCell(datosCell);

                    document.Add(proveedorTable);
                    document.Add(new Paragraph(" ", smallFont)); // Espacio pequeño

                    // ===== TABLA DE DETALLES CON FILAS VACÍAS =====
                    PdfPTable table = new PdfPTable(7);
                    table.WidthPercentage = 100;
                    float[] widths = new float[] { 0.5f, 4f, 0.7f, 1f, 1f, 1f, 1.2f };
                    table.SetWidths(widths);

                    // Encabezados de la tabla
                    string[] headers = new string[] { "Item", "Código / Descripción", "UM", "Cantidad", "Precio", "Importe", "F. Entrega" };
                    
                    // Item
                    var cellItem = new PdfPCell(new Phrase("Item", boldFont));
                    cellItem.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cellItem.HorizontalAlignment = Element.ALIGN_CENTER;
                    cellItem.Padding = 4;
                    table.AddCell(cellItem);
                    
                    // Código / Descripción
                    var cellCodDesc = new PdfPCell(new Phrase("Código / Descripción", boldFont));
                    cellCodDesc.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cellCodDesc.HorizontalAlignment = Element.ALIGN_CENTER;
                    cellCodDesc.Padding = 4;
                    table.AddCell(cellCodDesc);
                    
                    // UM
                    var cellUM = new PdfPCell(new Phrase("UM", boldFont));
                    cellUM.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cellUM.HorizontalAlignment = Element.ALIGN_CENTER;
                    cellUM.Padding = 4;
                    table.AddCell(cellUM);
                    
                    // Cantidad
                    var cellCantidad = new PdfPCell(new Phrase("Cantidad", boldFont));
                    cellCantidad.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cellCantidad.HorizontalAlignment = Element.ALIGN_CENTER;
                    cellCantidad.Padding = 4;
                    table.AddCell(cellCantidad);
                    
                    // Precio
                    var cellPrecio = new PdfPCell(new Phrase("Precio", boldFont));
                    cellPrecio.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cellPrecio.HorizontalAlignment = Element.ALIGN_CENTER;
                    cellPrecio.Padding = 4;
                    table.AddCell(cellPrecio);
                    
                    // Importe
                    var cellImporte = new PdfPCell(new Phrase("Importe", boldFont));
                    cellImporte.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cellImporte.HorizontalAlignment = Element.ALIGN_CENTER;
                    cellImporte.Padding = 4;
                    table.AddCell(cellImporte);
                    
                    // F. Entrega
                    var cellFecha = new PdfPCell(new Phrase("F. Entrega", boldFont));
                    cellFecha.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cellFecha.HorizontalAlignment = Element.ALIGN_CENTER;
                    cellFecha.Padding = 4;
                    table.AddCell(cellFecha);

                    // Agregar items reales
                    foreach (var item in detalle)
                    {
                        // Item
                        var rowCellItem = new PdfPCell(new Phrase(item.NuOrdItem?.ToString() ?? "", smallFont));
                        rowCellItem.HorizontalAlignment = Element.ALIGN_CENTER;
                        rowCellItem.Padding = 3;
                        rowCellItem.VerticalAlignment = Element.ALIGN_MIDDLE;
                        table.AddCell(rowCellItem);

                        // Código y Descripción (en una sola celda, dos líneas)
                        var codigoDesc = $"{item.CodExte}\n{item.Descritem}";
                        var rowCellCodDesc = new PdfPCell(new Phrase(codigoDesc, smallFont));
                        rowCellCodDesc.Padding = 3;
                        table.AddCell(rowCellCodDesc);

                        // UM
                        var rowCellUM = new PdfPCell(new Phrase(item.UniMas ?? "", smallFont));
                        rowCellUM.HorizontalAlignment = Element.ALIGN_CENTER;
                        rowCellUM.Padding = 3;
                        table.AddCell(rowCellUM);

                        // Cantidad
                        var rowCellCant = new PdfPCell(new Phrase((item.CantOcom ?? 0).ToString("N1"), smallFont));
                        rowCellCant.HorizontalAlignment = Element.ALIGN_RIGHT;
                        rowCellCant.Padding = 3;
                        table.AddCell(rowCellCant);

                        // Precio
                        var rowCellPrecio = new PdfPCell(new Phrase((item.PreUnit ?? 0).ToString("N2"), smallFont));
                        rowCellPrecio.HorizontalAlignment = Element.ALIGN_RIGHT;
                        rowCellPrecio.Padding = 3;
                        table.AddCell(rowCellPrecio);

                        // Importe
                        var rowCellImporte = new PdfPCell(new Phrase((item.INetItem ?? 0).ToString("N2"), smallFont));
                        rowCellImporte.HorizontalAlignment = Element.ALIGN_RIGHT;
                        rowCellImporte.Padding = 3;
                        table.AddCell(rowCellImporte);

                        // Fecha de Entrega
                        var rowCellFecha = new PdfPCell(new Phrase(item.FentreSol?.ToString("dd/MM/yy") ?? "", smallFont));
                        rowCellFecha.HorizontalAlignment = Element.ALIGN_CENTER;
                        rowCellFecha.Padding = 3;
                        table.AddCell(rowCellFecha);
                    }

                    // Agregar filas vacías (mínimo 15 filas totales para llenar la página)
                    int minFilas = 15;
                    int filasVacias = Math.Max(0, minFilas - detalle.Count);
                    for (int i = 0; i < filasVacias; i++)
                    {
                        for (int j = 0; j < 7; j++)
                        {
                            PdfPCell emptyCell = new PdfPCell(new Phrase(" ", smallFont));
                            emptyCell.MinimumHeight = 15f;
                            emptyCell.Padding = 3;
                            table.AddCell(emptyCell);
                        }
                    }

                    document.Add(table);

                    // ===== PIE: OBS + TOTALES + NUMERACIÓN =====
                    PdfPTable footerTable = new PdfPTable(3);
                    footerTable.WidthPercentage = 100;
                    footerTable.SetWidths(new float[] { 58f, 30f, 12f });
                    footerTable.SpacingBefore = 0;

                    // COLUMNA 1: OBS (Observaciones)
                    PdfPCell obsCell = new PdfPCell();
                    obsCell.Border = Rectangle.BOX;
                    obsCell.MinimumHeight = 80f;
                    obsCell.Padding = 5;

                    var obsHeader = new PdfPCell(new Phrase("OBS:", boldFont));
                    obsHeader.BackgroundColor = BaseColor.LIGHT_GRAY;
                    obsHeader.Border = Rectangle.NO_BORDER;
                    obsHeader.Padding = 3;
                    
                    var obsTable = new PdfPTable(1);
                    obsTable.WidthPercentage = 100;
                    obsTable.AddCell(obsHeader);
                    
                    var obsContent = new PdfPCell(new Phrase(orden.ObseOcom ?? "", smallFont));
                    obsContent.Border = Rectangle.NO_BORDER;
                    obsContent.MinimumHeight = 40f;
                    obsContent.Padding = 3;
                    obsTable.AddCell(obsContent);

                    var notaImportante = new PdfPCell(new Phrase(
                        $"Todos los importes están Expresados en {descripcionMoneda}\n" +
                        "Mencionar en facturas y remitos el número de Nota de Pedido",
                        smallFont
                    ));
                    notaImportante.BorderWidthTop = 1f;
                    notaImportante.BorderWidthLeft = 0;
                    notaImportante.BorderWidthRight = 0;
                    notaImportante.BorderWidthBottom = 0;
                    notaImportante.Padding = 3;
                    obsTable.AddCell(notaImportante);

                    obsCell.AddElement(obsTable);
                    footerTable.AddCell(obsCell);

                    // COLUMNA 2: TOTALES
                    PdfPTable totalsTable = new PdfPTable(2);
                    totalsTable.WidthPercentage = 100;
                    totalsTable.SetWidths(new float[] { 50f, 50f });

                    // Subtotal
                    var subtotalLabel = new PdfPCell(new Phrase("SUBTOTAL", boldFont));
                    subtotalLabel.BackgroundColor = BaseColor.LIGHT_GRAY;
                    subtotalLabel.Padding = 3;
                    subtotalLabel.HorizontalAlignment = Element.ALIGN_LEFT;
                    totalsTable.AddCell(subtotalLabel);
                    
                    var subtotalValue = new PdfPCell(new Phrase((orden.TotalOcom ?? 0).ToString("N2"), normalFont));
                    subtotalValue.Padding = 3;
                    subtotalValue.HorizontalAlignment = Element.ALIGN_RIGHT;
                    totalsTable.AddCell(subtotalValue);

                    // DESC
                    var descLabel = new PdfPCell(new Phrase("DESC.", boldFont));
                    descLabel.BackgroundColor = BaseColor.LIGHT_GRAY;
                    descLabel.Padding = 3;
                    totalsTable.AddCell(descLabel);
                    
                    var descValue = new PdfPCell(new Phrase("0.00", normalFont));
                    descValue.Padding = 3;
                    descValue.HorizontalAlignment = Element.ALIGN_RIGHT;
                    totalsTable.AddCell(descValue);

                    // NETO
                    var netoLabel = new PdfPCell(new Phrase("NETO", boldFont));
                    netoLabel.BackgroundColor = BaseColor.LIGHT_GRAY;
                    netoLabel.Padding = 3;
                    totalsTable.AddCell(netoLabel);
                    
                    var netoValue = new PdfPCell(new Phrase((orden.TotalOcom ?? 0).ToString("N2"), normalFont));
                    netoValue.Padding = 3;
                    netoValue.HorizontalAlignment = Element.ALIGN_RIGHT;
                    totalsTable.AddCell(netoValue);

                    // IVA
                    var ivaLabel = new PdfPCell(new Phrase("IVA", boldFont));
                    ivaLabel.BackgroundColor = BaseColor.LIGHT_GRAY;
                    ivaLabel.Padding = 3;
                    totalsTable.AddCell(ivaLabel);
                    
                    var ivaValue = new PdfPCell(new Phrase("%", normalFont));
                    ivaValue.Padding = 3;
                    ivaValue.HorizontalAlignment = Element.ALIGN_RIGHT;
                    totalsTable.AddCell(ivaValue);

                    // TOTAL
                    var totalLabel = new PdfPCell(new Phrase("TOTAL", boldFont));
                    totalLabel.BackgroundColor = BaseColor.LIGHT_GRAY;
                    totalLabel.Padding = 5;
                    totalsTable.AddCell(totalLabel);
                    
                    var totalValue = new PdfPCell(new Phrase((orden.TotalOcom ?? 0).ToString("N2"), boldFont));
                    totalValue.Padding = 5;
                    totalValue.HorizontalAlignment = Element.ALIGN_RIGHT;
                    totalsTable.AddCell(totalValue);

                    PdfPCell totalsCell = new PdfPCell(totalsTable);
                    totalsCell.Border = Rectangle.NO_BORDER;
                    footerTable.AddCell(totalsCell);

                    // COLUMNA 3: NUMERACIÓN DE HOJAS
                    PdfPTable hojasTable = new PdfPTable(1);
                    hojasTable.WidthPercentage = 100;

                    var hojaLabel = new PdfPCell(new Phrase("Hoja N°", boldFont));
                    hojaLabel.HorizontalAlignment = Element.ALIGN_CENTER;
                    hojaLabel.Padding = 3;
                    hojaLabel.MinimumHeight = 35f;
                    hojasTable.AddCell(hojaLabel);

                    var hojaNum = new PdfPCell(new Phrase("1", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD)));
                    hojaNum.HorizontalAlignment = Element.ALIGN_CENTER;
                    hojaNum.VerticalAlignment = Element.ALIGN_MIDDLE;
                    hojasTable.AddCell(hojaNum);

                    var deLabel = new PdfPCell(new Phrase("De:", boldFont));
                    deLabel.HorizontalAlignment = Element.ALIGN_CENTER;
                    deLabel.Padding = 3;
                    deLabel.MinimumHeight = 35f;
                    hojasTable.AddCell(deLabel);

                    var deNum = new PdfPCell(new Phrase("1", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD)));
                    deNum.HorizontalAlignment = Element.ALIGN_CENTER;
                    deNum.VerticalAlignment = Element.ALIGN_MIDDLE;
                    hojasTable.AddCell(deNum);

                    PdfPCell hojasCell = new PdfPCell(hojasTable);
                    hojasCell.Border = Rectangle.BOX;
                    footerTable.AddCell(hojasCell);

                    document.Add(footerTable);

                    document.Close();
                    Console.WriteLine($"[GenerarPdfOrdenCompraAsync] 📝 Documento PDF cerrado");

                    // Convertir a base64
                    byte[] pdfBytes = ms.ToArray();
                    string base64Content = Convert.ToBase64String(pdfBytes);

                    Console.WriteLine($"[GenerarPdfOrdenCompraAsync] ✅ PDF generado exitosamente - {pdfBytes.Length} bytes");

                    return Ok(new { Base64Content = base64Content });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GenerarPdfOrdenCompraAsync] ❌ Error: {ex.Message}");
                return StatusCode(500, $"Error al generar PDF: {ex.Message}");
            }
        }

        /// <summary>
        /// Genera el HTML para la orden de compra
        /// </summary>
        private string GenerarHtmlOrdenCompra(Ocomenca orden, List<Ocomdeta> detalle, Empresa? empresa)
        {
            var descripcionMoneda = orden.MoneEmis switch
            {
                1 => "Pesos",
                2 => "Dólares",
                3 => "Euros",
                _ => "Pesos"
            };

            var html = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Orden de Compra {orden.NumeOcom}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: Arial, sans-serif;
            font-size: 12px;
            padding: 20px;
            background-color: #fff;
            color: #000;
        }}
        
        .container {{
            max-width: 900px;
            margin: 0 auto;
            background: white;
        }}
        
        .header {{
            border: 2px solid #000;
            padding: 15px;
            margin-bottom: 20px;
            text-align: center;
        }}
        
        .header h2 {{
            color: #ff6600;
            font-size: 18px;
            margin-bottom: 5px;
        }}
        
        .header .empresa-info {{
            font-size: 11px;
            margin-top: 10px;
            line-height: 1.5;
        }}
        
        .orden-info {{
            border: 2px solid #000;
            padding: 10px;
            margin-bottom: 20px;
            background-color: #f9f9f9;
        }}
        
        .orden-info h3 {{
            font-size: 16px;
            margin-bottom: 10px;
            text-align: center;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 15px;
            border: 1px solid #000;
        }}
        
        th, td {{
            border: 1px solid #000;
            padding: 8px;
            text-align: left;
        }}
        
        th {{
            background-color: #f0f0f0;
            font-weight: bold;
            text-align: center;
        }}
        
        .info-table td:first-child {{
            background-color: #f0f0f0;
            font-weight: bold;
            width: 150px;
        }}
        
        .items-table td {{
            font-size: 11px;
        }}
        
        .items-table td:nth-child(1) {{
            text-align: center;
            width: 50px;
        }}
        
        .items-table td:nth-child(5),
        .items-table td:nth-child(6),
        .items-table td:nth-child(7) {{
            text-align: right;
        }}
        
        .total-row {{
            background-color: #f0f0f0;
            font-weight: bold;
        }}
        
        .total-row td {{
            text-align: right;
            font-size: 14px;
        }}
        
        .observaciones {{
            border: 1px solid #000;
            padding: 10px;
            min-height: 80px;
            margin-bottom: 20px;
        }}
        
        .observaciones h4 {{
            background-color: #f0f0f0;
            padding: 5px;
            margin: -10px -10px 10px -10px;
            border-bottom: 1px solid #000;
        }}
        
        .footer {{
            text-align: center;
            font-size: 10px;
            color: #666;
            margin-top: 30px;
            padding-top: 10px;
            border-top: 1px solid #ccc;
        }}
        
        @media print {{
            body {{
                padding: 0;
            }}
            
            .container {{
                max-width: 100%;
            }}
            
            @page {{
                margin: 1cm;
            }}
        }}
    </style>
</head>
<body>
    <div class='container'>
        <!-- Encabezado con datos de la empresa -->
        <div class='header'>
            <h2>{empresa?.razon ?? "EMPRESA"}</h2>
            <div class='empresa-info'>
                {(empresa != null ? $@"
                Dirección: {empresa.calle} {empresa.numero} {empresa.complemento}<br>
                ({empresa.CP}) {empresa.localidad} - Argentina<br>
                Teléfono: {empresa.TEL1} - Email: {empresa.email}<br>
                CUIT: {empresa.CUIT}
                " : "")}
            </div>
        </div>
        
        <!-- Información de la orden -->
        <div class='orden-info'>
            <h3>PEDIDO DE SUMINISTRO N° {orden.NumeOcom}</h3>
            <p style='text-align: center;'><strong>Fecha: {orden.FemiOcom?.ToString("dd/MM/yyyy") ?? "N/A"}</strong></p>
        </div>
        
        <!-- Datos del proveedor -->
        <table class='info-table'>
            <tr>
                <td>Proveedor:</td>
                <td colspan='3'><strong>{orden.RaSoProv}</strong></td>
            </tr>
            <tr>
                <td>N° Proveedor:</td>
                <td>{orden.NproOcom}</td>
                <td style='background-color: #f0f0f0; font-weight: bold;'>Moneda:</td>
                <td>{descripcionMoneda}</td>
            </tr>
            <tr>
                <td>Condición de Pago:</td>
                <td>{orden.CpagOcom}</td>
                <td style='background-color: #f0f0f0; font-weight: bold;'>Estado:</td>
                <td>{(orden.StatOcom == 0 ? "ACTIVA" : orden.StatOcom == -1 ? "PENDIENTE" : "ANULADA")}</td>
            </tr>
        </table>
        
        <!-- Detalle de items -->
        <table class='items-table'>
            <thead>
                <tr>
                    <th>Item</th>
                    <th>Código</th>
                    <th>Descripción</th>
                    <th>UM</th>
                    <th>Cantidad</th>
                    <th>Precio Unit.</th>
                    <th>Importe</th>
                </tr>
            </thead>
            <tbody>";

            foreach (var item in detalle)
            {
                html += $@"
                <tr>
                    <td>{item.NuOrdItem}</td>
                    <td><strong>{item.CodExte}</strong></td>
                    <td>{item.Descritem}</td>
                    <td>{item.UniMas}</td>
                    <td>{(item.CantOcom ?? 0):N2}</td>
                    <td>{(item.PreUnit ?? 0):N2}</td>
                    <td>{(item.INetItem ?? 0):N2}</td>
                </tr>";
            }

            html += $@"
            </tbody>
            <tfoot>
                <tr class='total-row'>
                    <td colspan='6'>TOTAL {descripcionMoneda.ToUpper()}:</td>
                    <td>{(orden.TotalOcom ?? 0):N2}</td>
                </tr>
            </tfoot>
        </table>
        
        <!-- Observaciones -->
        <div class='observaciones'>
            <h4>OBSERVACIONES:</h4>
            <p>{(string.IsNullOrEmpty(orden.ObseOcom) ? "Sin observaciones" : orden.ObseOcom)}</p>
        </div>
        
        <!-- Nota importante -->
        <div style='border: 1px solid #000; padding: 10px; background-color: #fffacd; margin-bottom: 20px;'>
            <p style='margin: 0; font-size: 11px;'>
                <strong>IMPORTANTE:</strong> Todos los importes están expresados en {descripcionMoneda}.<br>
                Mencionar en facturas y remitos el número de esta Orden de Compra.
            </p>
        </div>
        
        <!-- Footer -->
        <div class='footer'>
            <p>Documento generado electrónicamente - {DateTime.Now:dd/MM/yyyy HH:mm}</p>
        </div>
    </div>
</body>
</html>";

            return html;
        }

        /// <summary>
        /// Método auxiliar para agregar celdas a una tabla PDF
        /// </summary>
        private void AddCellToTable(PdfPTable table, string text, iTextSharp.text.Font font, bool isHeader, int colspan = 1)
        {
            var cell = new PdfPCell(new Phrase(text, font));
            cell.Border = Rectangle.NO_BORDER;
            cell.Padding = 3;
            if (isHeader)
            {
                cell.BackgroundColor = BaseColor.LIGHT_GRAY;
            }
            if (colspan > 1)
            {
                cell.Colspan = colspan;
            }
            table.AddCell(cell);
        }

        /// <summary>
        /// Obtiene el ID del usuario activo desde el contexto HTTP
        /// </summary>
        /// <returns>ID del usuario activo o null si no se puede obtener</returns>
        private short? ObtenerUsuarioActivo()
        {
            try
            {
                var user = HttpContext.User;

                if (user?.Identity?.IsAuthenticated ?? false)
                {
                    var username = user.Identity.Name;
                    var email = user.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value;

                    Console.WriteLine($"[ObtenerUsuarioActivo] Username: {username}, Email: {email}");

                    if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
                    {
                        var usuario = _context.Usuarios.FirstOrDefault(u =>
                            u.NombreUsuario == username && u.Email == email);

                        if (usuario != null)
                        {
                            Console.WriteLine($"[ObtenerUsuarioActivo] ✅ Usuario encontrado - IdFlexoft: {usuario.IdFlexoft}");
                            return (short?)usuario.IdFlexoft;
                        }
                        else
                        {
                            Console.WriteLine($"[ObtenerUsuarioActivo] ⚠️ Usuario no encontrado en BD");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[ObtenerUsuarioActivo] ⚠️ Username o Email vacío");
                    }
                }
                else
                {
                    Console.WriteLine($"[ObtenerUsuarioActivo] ⚠️ Usuario no autenticado");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ObtenerUsuarioActivo] ❌ Error: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Obtiene múltiples órdenes de compra por sus IDs de encabezado
        /// </summary>
        [HttpPost("por-ids")]
        public async Task<ActionResult<List<Ocomenca>>> GetPorIdsAsync([FromBody] List<int> ids)
        {
            try
            {
                if (ids == null || !ids.Any())
                {
                    return Ok(new List<Ocomenca>());
                }

                var ordenes = await _context.Ocomencas
                    .Where(o => ids.Contains(o.OCOMENCA_ID))
                    .OrderByDescending(o => o.NumeOcom)
                    .ToListAsync();

                return Ok(ordenes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener órdenes por IDs: {ex.Message}");
                return StatusCode(500, $"Error al obtener órdenes: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene encabezados de órdenes de compra pendientes de recepción por proveedor
        /// Optimizado: devuelve directamente los encabezados sin necesidad de dos consultas
        /// </summary>
        [HttpGet("pendientes-recepcion/{numeroProveedor}")]
        public async Task<ActionResult<List<Ocomenca>>> GetOrdenesPendientesRecepcion(int numeroProveedor)
        {
            try
            {
                // Obtener números de orden únicos que tienen items pendientes
                var numerosOrdenPendientes = await _context.Ocomdetas
                    .Where(d => d.NproOcom == numeroProveedor
                             && d.StatOcom < 3 
                             && d.StatOcom != -1 
                             && d.CantRec < (d.CantOcom - 0.001m))
                    .Select(d => d.NumeOcom)
                    .Distinct()
                    .ToListAsync();

                if (!numerosOrdenPendientes.Any())
                {
                    return Ok(new List<Ocomenca>());
                }

                // Obtener los encabezados de esas órdenes
                var ordenesPendientes = await _context.Ocomencas
                    .Where(o => numerosOrdenPendientes.Contains(o.NumeOcom))
                    .OrderByDescending(o => o.NumeOcom)
                    .ToListAsync();

                return Ok(ordenesPendientes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener órdenes pendientes de recepción: {ex.Message}");
                return StatusCode(500, $"Error al obtener órdenes pendientes: {ex.Message}");
            }
        }
    }
}
