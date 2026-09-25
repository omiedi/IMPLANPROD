using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.Utilities;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Data.Common;
using System.Security.Claims;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la gestión de pedidos de venta (Pedenca)
    /// </summary> 
    [ApiController]
    [Route("api/[controller]")]
    public class PedencaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ConversionMonedaService _conversionMonedaService;
        private readonly DocumentNumberGenerator _documentNumberGenerator;
        private readonly UserContextService _userContextService;
        private readonly ISyncService _syncService;
        private readonly AnotacionesPedidoService _anotacionesService;

        /// <summary>
        /// Constructor del controlador
        /// </summary>
        /// <param name="context">Contexto de base de datos</param>
        /// <param name="conversionMonedaService">Servicio de conversión de moneda</param>
        /// <param name="documentNumberGenerator">Generador de números de documento</param>
        /// <param name="userContextService">Servicio de contexto de usuario</param>
        public PedencaController(
        DataContext context,
        ConversionMonedaService conversionMonedaService,
        DocumentNumberGenerator documentNumberGenerator,
        UserContextService userContextService,
        ISyncService syncService,
        AnotacionesPedidoService anotacionesService)
        {
            _context = context;
            _conversionMonedaService = conversionMonedaService;
            _documentNumberGenerator = documentNumberGenerator;
            _userContextService = userContextService;
            _syncService = syncService;
            _anotacionesService = anotacionesService;
        }

        /// <summary>
        /// Obtiene el detalle (PEDDETA) de un pedido por NroPed
        /// </summary>
        /// <param name="nroPed">Número de pedido (dato de negocio)</param>
        /// <returns>Lista de ítems del detalle</returns>
        [HttpGet("detalle/{nroPed:int}")]
        public async Task<ActionResult> GetDetalleAsync(int nroPed)
        {
            try
            {
                var detalle = await _context.Peddetas
                    .Where(d => d.NroPed == nroPed)
                    .OrderBy(d => d.NuOrItem)
                    .ToListAsync();

                return Ok(detalle);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener el detalle del pedido: {ex.Message}");
            }
        }

        /// <summary>
        /// Convierte la moneda de un pedido y actualiza todos los precios
        /// </summary>
        /// <param name="nroPed">Número de pedido</param>
        /// <param name="nuevaMoneda">Código de la nueva moneda</param>
        /// <returns>Detalle del pedido con precios convertidos</returns>
        [HttpPost("convertir-moneda/{nroPed:int}/{nuevaMoneda:int}")]
        public async Task<ActionResult> ConvertirMonedaPedidoAsync(int nroPed, int nuevaMoneda)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // Obtener el pedido
                var pedido = await _context.Pedencas.FirstOrDefaultAsync(p => p.NroPed == nroPed);
                if (pedido == null)
                {
                    return NotFound($"Pedido {nroPed} no encontrado");
                }

                // Obtener el detalle actual
                var detalle = await _context.Peddetas
                    .Where(d => d.NroPed == nroPed)
                    .ToListAsync();

                if (!detalle.Any())
                {
                    return BadRequest("El pedido no tiene items para convertir");
                }

                // Obtener moneda actual del pedido (usar la del primer item)
                var monedaActual = detalle.First().MoneEmis ?? 1;

                // Verificar si la nueva moneda existe
                var nuevaMonedaInfo = await _conversionMonedaService.ObtenerMonedaAsync(nuevaMoneda);
                if (nuevaMonedaInfo == null)
                {
                    return BadRequest($"La moneda {nuevaMoneda} no existe o está inactiva");
                }

                // Convertir todos los precios del detalle
                var detalleConvertido = await _conversionMonedaService.ConvertirDetallesPedidoAsync(
                    detalle, monedaActual, nuevaMoneda);

                // Actualizar el encabezado del pedido
                pedido.MoneEmis = (short)nuevaMoneda;
                pedido.TipoCam = await _conversionMonedaService.ObtenerCotizacionAsync(nuevaMoneda);

                // Guardar cambios
                _context.Pedencas.Update(pedido);
                _context.Peddetas.UpdateRange(detalleConvertido);
                await _context.SaveChangesAsync();

                await transaction.CommitAsync();

                return Ok(new
                {
                    NroPed = nroPed,
                    MonedaAnterior = monedaActual,
                    MonedaNueva = nuevaMoneda,
                    NombreMoneda = nuevaMonedaInfo.Descripcion,
                    Cotizacion = nuevaMonedaInfo.Cotizacion,
                    ItemsConvertidos = detalleConvertido.Count,
                    Detalle = detalleConvertido
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return BadRequest($"Error al convertir moneda del pedido: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene todas las monedas activas del sistema
        /// </summary>
        /// <returns>Lista de monedas disponibles</returns>
        [HttpGet("monedas")]
        public async Task<ActionResult> GetMonedasAsync()
        {
            try
            {
                var monedas = await _conversionMonedaService.ObtenerMonedasActivasAsync();
                return Ok(monedas);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener monedas: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza una Orden de Venta completa: PEDENCA (encabezado) + PEDDETA (detalle)
        /// Borra el detalle existente y lo vuelve a insertar con los datos provistos.
        /// </summary>
        /// <param name="request">Datos a actualizar junto con NroPed</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut("editar-con-detalle")]
        public async Task<ActionResult> EditarConDetalleAsync([FromBody] PedidoVentaUpdateDTO request)
        {
            if (request == null || request.Items == null || request.Items.Count == 0)
            {
                return BadRequest("La solicitud es inválida: debe incluir al menos un ítem.");
            }

            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.Serializable);
            try
            {
                // Buscar encabezado por NroPed
                var encabezado = await _context.Pedencas.FirstOrDefaultAsync(p => p.NroPed == request.NroPed);
                if (encabezado == null)
                {
                    return NotFound($"Pedido Nro {request.NroPed} no encontrado.");
                }

                // Actualizar campos principales del encabezado
                encabezado.FechEmis = encabezado.FechEmis ?? DateTime.Now; // mantener si ya existe
                encabezado.NroClie = request.NroClie;
                encabezado.RasoClie = request.RasoClie;
                encabezado.NroOcom = request.NroOcom;
                encabezado.Observa = request.Observa;
                encabezado.FeSolEnt = request.FeSolEnt ?? DateTime.Today; // Si está vacío, usar fecha de hoy
                // Status: mantener la lógica actual del registro
                encabezado.ListPre = request.ListPre;
                encabezado.CondPag = request.CondPag;
                encabezado.Vended = request.Vended;
                encabezado.PorComVend = request.PorComVend;
                encabezado.VenRep = string.IsNullOrEmpty(request.Motivo) ? encabezado.VenRep : request.Motivo.Substring(0, 1);
                // Grabar texto de descuento aplicado
                encabezado.TDescApli = request.TDescApli;
                encabezado.NTranspo = request.NTranspo; 
                encabezado.TranspEnt = request.TranspEnt;
                encabezado.DomiEnt = request.DomiEnt;
                encabezado.LocaEnt = request.LocaEnt;



        _context.Pedencas.Update(encabezado);
                await _context.SaveChangesAsync();

                // Obtener detalle existente y preservar valores de CantEnt antes de borrar
                var detalleActual = await _context.Peddetas.Where(d => d.NroPed == request.NroPed).ToListAsync();
                
                // Crear diccionario para preservar CantEnt por NuOrItem
                var cantEntExistentes = detalleActual.ToDictionary(
                    d => (short)(d.NuOrItem ?? 0), 
                    d => new { 
                        CantEnt = d.CantEnt ?? 0, 
                        FeUltEnt = d.FeUltEnt,
                        NroRemi = d.NroRemi ?? string.Empty 
                    }
                );

                if (detalleActual.Count > 0)
                {
                    _context.Peddetas.RemoveRange(detalleActual);
                    await _context.SaveChangesAsync();
                }

                // Insertar nuevo detalle preservando CantEnt existente
                var nuevos = new List<Peddeta>();
                foreach (var item in request.Items.OrderBy(i => i.NuOrItem))
                {
                    // Buscar valores existentes para este item
                    var valoresExistentes = cantEntExistentes.ContainsKey(item.NuOrItem) 
                        ? cantEntExistentes[item.NuOrItem] 
                        : null;

                    var det = new Peddeta
                    {
                        NroPed = request.NroPed,
                        FechEmis = encabezado.FechEmis,
                        NroClie = encabezado.NroClie,
                        RaSoClie = encabezado.RasoClie,
                        NroOcom = encabezado.NroOcom,
                        FeSolEnt = item.FeSolEnt ?? encabezado.FeSolEnt,
                        Status = encabezado.Status,
                        ListPre = encabezado.ListPre,
                        CondPag = encabezado.CondPag,
                        Vended = encabezado.Vended,
                        PorComVend = encabezado.PorComVend,

                        NuOrItem = item.NuOrItem,
                        CodiInt = item.CodiInt,
                        CodiExt = item.CodiExt,
                        SubCod = item.SubCod ?? string.Empty,
                        Descrip = item.Descrip,
                        TColor = item.TColor ?? string.Empty,
                        TTalle = item.TTalle ?? string.Empty,
                        TLargo = item.TLargo ?? 0,
                        TAncho = item.TAncho ?? 0,
                        TEspesor = item.TEspesor ?? 0,
                        VenRep = item.VenRep ?? encabezado.VenRep,
                        MoneEmis = item.MoneEmis,
                        TipoCam = item.TipoCam,
                        LISTA_DE_ITEM = item.ListaDeItem ?? 1,
                        PreList = item.PreList,
                        CanAsig = item.CanAsig,
                        CanPed = item.CanPed,
                        PreUnid = item.PreUnid,
                        ImpoTot = item.ImpoTot,
                        // PRESERVAR CantEnt existente - NO actualizar en updates
                        CantEnt = valoresExistentes?.CantEnt ?? 0,
                        // Recalcular CanSaldo basado en la nueva cantidad pedida menos lo ya entregado
                        CanSaldo = (item.CanPed ?? 0) - (valoresExistentes?.CantEnt ?? 0),
                        // Preservar fecha de última entrega y número de remito si existían
                        FeUltEnt = valoresExistentes?.FeUltEnt,
                        NroRemi = valoresExistentes?.NroRemi ?? string.Empty,
                        TPedItem = $"{request.NroPed}-{item.NuOrItem}", // Criterio VB6: NRO_PEDIDO-NRO_ITEM
                        NumeroSerie = item.NumeroSerie ?? string.Empty,
                        NROCOTI = string.Empty,
                        SIN_CARGO = item.SinCargo ?? 0,
                        ObservaItem = item.ObservaItem ?? string.Empty,
                    };
                    nuevos.Add(det);
                }

                _context.Peddetas.AddRange(nuevos);
                await _context.SaveChangesAsync();

                await transaction.CommitAsync();

                // Guardar anotaciones estructuradas en tabla ANOTAPED
                try
                {
                    var numUsuar = _userContextService.GetCurrentUserId() ?? 0;
                    var nombreUsuario = await ObtenerNombreUsuarioAsync();
                    await _anotacionesService.GuardarEnAnotapedAsync(
                        request.NroPed, nuevos, request.Observa, nombreUsuario, numUsuar, "EDITADO");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ANOTAPED] Error al guardar (editar pedido {request.NroPed}): {ex}");
                }

                // ====================================================================
                // VALIDAR SI ES PEDIDO DE SUCURSAL Y GENERAR ARCHIVO ARCHIVO_GENERO_PEDIDO
                // ====================================================================
                try
                {
                    // Verificar si el pedido existe en OC_SUCURSAL (pedido de sucursal)
                    var ocSucursal = await _context.O_Sucursals
                        .FirstOrDefaultAsync(o => o.NpedGeneradoEnAvellaneda == request.NroPed && o.Marborra != -1);

                    if (ocSucursal != null)
                    {
                        // Es un pedido de sucursal, generar archivo ARCHIVO_GENERO_PEDIDO
                        var numUsuar = _userContextService.GetCurrentUserId() ?? 0;
                        var nroClie = encabezado.NroClie ?? 0;
                        var oCompraSuc = ocSucursal.OCompraSuc ?? 0;

                        await _syncService.GenerarArchivoGeneroPedidoAsync(
                            request.NroPed,
                            nroClie,
                            oCompraSuc,
                            "MODIPEDI",
                            numUsuar
                        );
                    }
                }
                catch (Exception ex)
                {
                    // No fallar la edición si hay error al generar el archivo
                    Console.WriteLine($"Error al generar ARCHIVO_GENERO_PEDIDO al editar pedido: {ex.Message}");
                }
                return Ok(new { request.NroPed, Items = nuevos.Count });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return BadRequest($"Error al actualizar el pedido con detalle: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea una Orden de Venta completa: PEDENCA (encabezado) + PEDDETA (detalle)
        /// en una única transacción atómica y segura para multiusuario.
        /// 
        /// Estrategia de concurrencia:
        /// - Se utiliza una transacción con aislamiento SERIALIZABLE
        /// - Se obtiene un Application Lock (sp_getapplock) para serializar la obtención del NroPed
        /// - Se asigna NroPed = 1 + MAX(PEDENCA.NroPed)
        /// </summary>
        /// <param name="request">Datos del encabezado y sus ítems</param>
        /// <returns>Resultado con el número de pedido generado y datos básicos</returns>
        [HttpPost("crear-con-detalle")]
        public async Task<ActionResult> CrearConDetalleAsync([FromBody] PedidoVentaCreateDTO request)
        {
            // Validación básica de entrada
            if (request == null || request.Items == null || request.Items.Count == 0)
            {
                return BadRequest("La solicitud es inválida: debe incluir al menos un ítem.");
            }

            // Iniciar transacción SERIALIZABLE para máxima seguridad en concurrencia
            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.Serializable);

            try
            {
                // Obtención segura de NroPed dentro de la transacción SERIALIZABLE usando hints de bloqueo
                // WITH (UPDLOCK, HOLDLOCK) asegura un bloqueo de rango que evita condiciones de carrera.
                var connection = _context.Database.GetDbConnection();
                if (connection.State != ConnectionState.Open)
                {
                    await connection.OpenAsync();
                }
                using var command = connection.CreateCommand();
                // Asociar el comando a la transacción actual de EF Core
                var currentTx = _context.Database.CurrentTransaction;
                if (currentTx != null)
                {
                    command.Transaction = currentTx.GetDbTransaction();
                }
                command.CommandText = "SELECT ISNULL(MAX(NroPed), 0) FROM PEDENCA WITH (UPDLOCK, HOLDLOCK)";
                var scalar = await command.ExecuteScalarAsync();
                var maxNroPed = scalar == null || scalar == DBNull.Value ? 0 : Convert.ToInt32(scalar);
                var nuevoNroPed = maxNroPed + 1;

                // Crear encabezado PEDENCA
                var ahora = DateTime.Now;
                var encabezado = new Pedenca
                {
                    NroPed = nuevoNroPed,
                    FechEmis = ahora,
                    NroClie = request.NroClie,
                    RasoClie = request.RasoClie,
                    NroOcom = request.NroOcom,
                    Observa = request.Observa,
                    FeSolEnt = request.FeSolEnt ?? DateTime.Today, // Si está vacío, usar fecha de hoy
                    Status = 0,        // 0: pendiente de aprobación (según requerimiento)
                    ListPre = request.ListPre,
                    CondPag = request.CondPag,
                    Vended = request.Vended,
                    PorComVend = request.PorComVend,
                    VenRep = string.IsNullOrEmpty(request.Motivo) ? null : request.Motivo!.Substring(0, 1),
                    // Texto de descuento aplicado
                    TDescApli = request.TDescApli,
                    TranspEnt = request.TranspEnt,
                    NTranspo = request.NTranspo,
                    DomiEnt = request.DomiEnt,
                    LocaEnt = request.LocaEnt,
                };

                _context.Pedencas.Add(encabezado);
                await _context.SaveChangesAsync();

                // Mapear los ítems a PEDDETA reutilizando datos del encabezado
                var detalles = new List<Peddeta>();
                foreach (var item in request.Items)
                {
                    var det = new Peddeta
                    {
                        // Datos de cabecera replicados en el detalle
                        NroPed = nuevoNroPed,
                        FechEmis = encabezado.FechEmis,
                        NroClie = encabezado.NroClie,
                        RaSoClie = encabezado.RasoClie,
                        NroOcom = encabezado.NroOcom,
                        FeSolEnt = encabezado.FeSolEnt,
                        Status = encabezado.Status,
                        ListPre = encabezado.ListPre,
                        CondPag = encabezado.CondPag,
                        Vended = encabezado.Vended,
                        PorComVend = encabezado.PorComVend,

                        // Datos específicos del item (según requerimiento)
                        NuOrItem = item.NuOrItem,
                        CodiInt = item.CodiInt,
                        CodiExt = item.CodiExt,
                        SubCod = item.SubCod ?? string.Empty,
                        Descrip = item.Descrip,
                        TColor = item.TColor ?? string.Empty,
                        TTalle = item.TTalle ?? string.Empty,
                        TLargo = item.TLargo ?? 0,
                        TAncho = item.TAncho ?? 0,
                        TEspesor = item.TEspesor ?? 0,
                        VenRep = item.VenRep ?? encabezado.VenRep,
                        MoneEmis = item.MoneEmis,
                        TipoCam = item.TipoCam,

                        LISTA_DE_ITEM = item.ListaDeItem ?? 1,
                        PreList = item.PreList,
                        CanAsig = item.CanAsig,
                        CanPed = item.CanPed,
                        PreUnid = item.PreUnid,
                        ImpoTot = item.ImpoTot,
                        CantEnt = 0, // inicial
                        CanSaldo = item.CanPed, // inicial igual a lo pedido
                        FeUltEnt = null,
                        NroRemi = string.Empty,
                        TPedItem = $"{nuevoNroPed}-{item.NuOrItem}", // Criterio VB6: NRO_PEDIDO-NRO_ITEM

                        NumeroSerie = item.NumeroSerie ?? string.Empty,
                        NROCOTI = string.Empty,
                        SIN_CARGO = item.SinCargo ?? 0,
                        ObservaItem = item.ObservaItem ?? string.Empty,
                    };

                    detalles.Add(det);
                }

                _context.Peddetas.AddRange(detalles);
                await _context.SaveChangesAsync();

                // Confirmar transacción
                await transaction.CommitAsync();

                // Guardar anotaciones estructuradas en tabla ANOTAPED
                // (fuera de la transacción principal para no afectar el alta del pedido)
                try
                {
                    var numUsuar = _userContextService.GetCurrentUserId() ?? 0;
                    var nombreUsuario = await ObtenerNombreUsuarioAsync();
                    await _anotacionesService.GuardarEnAnotapedAsync(
                        nuevoNroPed, detalles, request.Observa, nombreUsuario, numUsuar, "CREADO");
                }
                catch (Exception ex)
                {
                    // Log detallado para diagnóstico: el pedido se creó correctamente,
                    // pero falló el guardado en ANOTAPED (tabla opcional/estructurada).
                    Console.WriteLine($"[ANOTAPED] Error al guardar (crear pedido {nuevoNroPed}): {ex}");
                }

                // Responder con el número de pedido generado y un resumen
                return Ok(new
                {
                    NroPed = nuevoNroPed,
                    EncabezadoId = encabezado.Id,
                    ItemsCreados = detalles.Count
                });
            }
            catch (Exception ex)
            {
                // Rollback explícito en caso de error
                await transaction.RollbackAsync();
                return BadRequest($"Error al crear la orden de venta y su detalle: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene todos los pedidos con paginación y filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="filter">Filtro de texto para búsqueda</param>
        /// <param name="estado">Estado del pedido</param>
        /// <param name="clienteId">ID del cliente</param>
        /// <param name="fechaDesde">Fecha desde</param>
        /// <param name="fechaHasta">Fecha hasta</param>
        /// <returns>Lista paginada de pedidos</returns>
        [HttpGet]
        public async Task<ActionResult> GetAsync([FromQuery] PaginationDTO pagination, 
            [FromQuery] string? filter = null,
            [FromQuery] string? estado = null,
            [FromQuery] int? clienteId = null,
            [FromQuery] int? nroPed = null,
            [FromQuery] DateTime? fechaDesde = null,
            [FromQuery] DateTime? fechaHasta = null)
        {
            try
            {
                var queryable = _context.Pedencas.AsQueryable();

                // Filtrar por estado activo si es necesario
                queryable = queryable.Where(x => x.Status >= 0);

                // Aplicar filtro por número de pedido (exacto)
                if (nroPed.HasValue)
                {
                    queryable = queryable.Where(x => x.NroPed == nroPed.Value);
                }

                // Aplicar filtro de texto
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    queryable = queryable.Where(x => 
                        (x.RasoClie != null && x.RasoClie.Contains(filter)) ||
                        (x.NroOcom != null && x.NroOcom.Contains(filter)) ||
                        (x.Referencia != null && x.Referencia.Contains(filter)) ||
                        (x.NroPed != null && x.NroPed.ToString().Contains(filter)) ||
                        (x.NroClie != null && x.NroClie.ToString().Contains(filter)));
                }

                // Aplicar filtro de estado
                if (!string.IsNullOrWhiteSpace(estado) && int.TryParse(estado, out int estadoInt))
                {
                    if (estadoInt == 0)
                    {
                        // Si es 0, filtrar por estados 0 y 1 (ambos activos)
                        queryable = queryable.Where(x => x.Status == 0 || x.Status == 1);
                    }
                    else
                    {
                        queryable = queryable.Where(x => x.Status == estadoInt);
                    }
                }

                // Aplicar filtro de cliente
                if (clienteId.HasValue)
                {
                    queryable = queryable.Where(x => x.NroClie == clienteId.Value);
                }

                // Aplicar filtro de fechas
                if (fechaDesde.HasValue)
                {
                    queryable = queryable.Where(x => x.FechEmis >= fechaDesde.Value);
                }

                if (fechaHasta.HasValue)
                {
                    queryable = queryable.Where(x => x.FechEmis <= fechaHasta.Value.AddDays(1));
                }

                return Ok(await queryable
                    .OrderByDescending(x => x.FechEmis)
                    .Paginate(pagination)
                    .ToListAsync());
            }
            catch (Exception ex)
            {
                // Manejo de errores para detectar si se genera alguna excepción en el listado
                return BadRequest($"Error al obtener pedidos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación con filtros
        /// </summary>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <param name="filter">Filtro de texto para búsqueda</param>
        /// <param name="estado">Estado del pedido</param>
        /// <param name="clienteId">ID del cliente</param>
        /// <param name="fechaDesde">Fecha desde</param>
        /// <param name="fechaHasta">Fecha hasta</param>
        /// <returns>Número total de páginas</returns>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination,
            [FromQuery] string? filter = null,
            [FromQuery] string? estado = null,
            [FromQuery] int? clienteId = null,
            [FromQuery] int? nroPed = null,
            [FromQuery] DateTime? fechaDesde = null,
            [FromQuery] DateTime? fechaHasta = null)
        {
            try
            {
                var queryable = _context.Pedencas.AsQueryable();
                queryable = queryable.Where(x => x.Status >= 0);

                // Aplicar filtro por número de pedido (exacto)
                if (nroPed.HasValue)
                {
                    queryable = queryable.Where(x => x.NroPed == nroPed.Value);
                }

                // Aplicar filtro de texto
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    queryable = queryable.Where(x => 
                        (x.RasoClie != null && x.RasoClie.Contains(filter)) ||
                        (x.NroOcom != null && x.NroOcom.Contains(filter)) ||
                        (x.Referencia != null && x.Referencia.Contains(filter)) ||
                        (x.NroPed != null && x.NroPed.ToString().Contains(filter)) ||
                        (x.NroClie != null && x.NroClie.ToString().Contains(filter)));
                }

                // Aplicar filtro de estado
                if (!string.IsNullOrWhiteSpace(estado) && int.TryParse(estado, out int estadoInt))
                {
                    if (estadoInt == 0)
                    {
                        // Si es 0, filtrar por estados 0 y 1 (ambos activos)
                        queryable = queryable.Where(x => x.Status == 0 || x.Status == 1);
                    }
                    else
                    {
                        queryable = queryable.Where(x => x.Status == estadoInt);
                    }
                }

                // Aplicar filtro de cliente
                if (clienteId.HasValue)
                {
                    queryable = queryable.Where(x => x.NroClie == clienteId.Value);
                }

                // Aplicar filtro de fechas
                if (fechaDesde.HasValue)
                {
                    queryable = queryable.Where(x => x.FechEmis >= fechaDesde.Value);
                }

                if (fechaHasta.HasValue)
                {
                    queryable = queryable.Where(x => x.FechEmis <= fechaHasta.Value.AddDays(1));
                }

                double count = await queryable.CountAsync();
                double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                // Manejo de errores para detectar problemas en el cálculo de páginas
                return BadRequest($"Error al calcular total de páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Lista de pedidos que tienen al menos un ítem pendiente de entrega.
        /// Equivale al SQL del VB6 (recordset Pedianuittemp del anulador de ítems):
        ///   SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom
        ///   FROM PEDENCA INNER JOIN PEDDETA ON (PEDENCA.NroPed = PEDDETA.NroPed)
        ///   WHERE PEDDETA.Status &lt; 8                     -- no anulada ni cumplida
        ///     AND PEDDETA.CantEnt &lt; (PEDDETA.CanPed - 0.005)  -- con saldo pendiente
        ///     AND PEDENCA.Status &lt; 9                     -- pedido no anulado
        ///   ORDER BY PEDENCA.NroPed ASC
        /// Se usa en la utilidad "Anular Items de Pedido" de /pedidos-venta.
        /// </summary>
        /// <param name="nroPed">Filtro exacto por número de pedido</param>
        /// <param name="texto">Filtro de texto libre sobre la referencia compuesta
        /// (N° pedido, N° OC del cliente y razón social)</param>
        [HttpGet("pedidos-con-items-pendientes")]
        public async Task<ActionResult<List<PedidoConItemsPendientesDTO>>> GetPedidosConItemsPendientes(
            [FromQuery] int? nroPed = null,
            [FromQuery] string? texto = null)
        {
            try
            {
                var query = _context.Pedencas
                    .AsNoTracking()
                    .Where(pe => pe.Status < 9)
                    .Where(pe => _context.Peddetas.Any(pd =>
                        pd.NroPed == pe.NroPed
                        && pd.Status < 8
                        && (pd.CantEnt ?? 0m) < (pd.CanPed ?? 0m) - 0.005m));

                if (nroPed.HasValue)
                {
                    query = query.Where(pe => pe.NroPed == nroPed.Value);
                }

                if (!string.IsNullOrWhiteSpace(texto))
                {
                    var t = texto.Trim();
                    query = query.Where(pe =>
                        (pe.NroOcom != null && pe.NroOcom.Contains(t)) ||
                        (pe.RasoClie != null && pe.RasoClie.Contains(t)) ||
                        (pe.NroPed != null && pe.NroPed.ToString().Contains(t)));
                }

                var pedidos = await query
                    .OrderBy(pe => pe.NroPed)
                    .Select(pe => new PedidoConItemsPendientesDTO
                    {
                        NroPed = pe.NroPed ?? 0,
                        NroClie = pe.NroClie ?? 0,
                        RasoClie = pe.RasoClie ?? string.Empty,
                        FechEmis = pe.FechEmis,
                        NroOcom = pe.NroOcom ?? string.Empty
                    })
                    .ToListAsync();

                return Ok(pedidos);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener pedidos con ítems pendientes: {ex.Message}");
            }
        }

        /// <summary>
        /// Marca como cumplidos (PEDDETA.Status = 8) los ítems seleccionados de un
        /// pedido de venta. Equivale al botón "Marcar Cumplidos" del anulador de
        /// ítems del VB6: la línea queda cumplida y deja de figurar como pendiente
        /// de entrega en todos los listados que filtran Status &lt; 8.
        /// Se registran los campos de auditoría LastUpdate y NumUsuar (usuario activo).
        /// </summary>
        /// <param name="request">Número de pedido y lista de NuOrItem a cumplir</param>
        [HttpPost("marcar-items-cumplidos")]
        public async Task<ActionResult> MarcarItemsCumplidos([FromBody] MarcarItemsCumplidosDTO request)
        {
            if (request == null || request.NuOrItems == null || request.NuOrItems.Count == 0)
            {
                return BadRequest("Debe indicar al menos un ítem a marcar como cumplido.");
            }

            await using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // NuOrItem es short? en PEDDETA: se convierte la lista para que EF
                // traduzca el Contains a un IN (...) sin problemas de tipo.
                var itemsSeleccionados = request.NuOrItems.Select(x => (short?)x).ToList();

                var items = await _context.Peddetas
                    .Where(d => d.NroPed == request.NroPed && itemsSeleccionados.Contains(d.NuOrItem))
                    .ToListAsync();

                if (items.Count == 0)
                {
                    return NotFound($"No se encontraron los ítems indicados del pedido {request.NroPed}.");
                }

                var numUsuar = _userContextService.GetCurrentUserId() ?? 0;
                var ahora = DateTime.Now;

                foreach (var item in items)
                {
                    // 8 = ítem cumplido (regla VB6). No se tocan cantidades ni saldos:
                    // la línea simplemente deja de considerarse pendiente.
                    item.Status = 8;
                    item.LastUpdate = ahora;
                    item.NumUsuar = numUsuar;
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                // Registrar anotación automática "Items Marcados Como Cumplidos" con
                // el detalle de los ítems (visible en /anotaciones-pedido).
                // Se hace después del commit: si la anotación falla no revierte el
                // marcado de los ítems (el servicio loguea el error internamente).
                try
                {
                    var nombreUsuario = await ObtenerNombreUsuarioAsync();
                    await _anotacionesService.RegistrarItemsCumplidosAsync(
                        request.NroPed, items, nombreUsuario, numUsuar);
                }
                catch (Exception exAnota)
                {
                    Console.WriteLine($"[ANOTAPED] Error al registrar ítems cumplidos del pedido {request.NroPed}: {exAnota}");
                }

                return Ok(new { request.NroPed, ItemsCumplidos = items.Count });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return BadRequest($"Error al marcar ítems como cumplidos: {ex.Message}");
            }
        }

        /// <summary>
        /// Registra la anotación automática "Items Remitidos" en el anotador de cada
        /// pedido asociado a un remito de cliente.
        /// Los ítems remitidos por pedido se derivan de MOVISTO: CodiMovi = "RT" y
        /// DoPriLar = remito, filtrando NuPedCli &gt; 0, por lo que los ítems del
        /// remito que NO están asociados a un pedido quedan excluidos.
        /// Se invoca desde /remitos/cliente una vez grabados todos los movimientos.
        /// </summary>
        /// <param name="request">Documento largo del remito (MOVISTO.DoPriLar)</param>
        [HttpPost("registrar-items-remitidos")]
        public async Task<ActionResult> RegistrarItemsRemitidos([FromBody] RegistrarItemsRemitidosDTO request)
        {
            try
            {
                if (request == null || string.IsNullOrWhiteSpace(request.RemitoLargo))
                {
                    return BadRequest("Debe indicar el número de remito.");
                }

                // Solo movimientos del remito asociados a pedido (NuPedCli > 0)
                var movimientos = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.DoPriLar == request.RemitoLargo
                             && m.CodiMovi == "RT"
                             && m.NuPedCli > 0)
                    .ToListAsync();

                if (!movimientos.Any())
                {
                    return Ok(new { PedidosAnotados = 0 });
                }

                var numUsuar = _userContextService.GetCurrentUserId() ?? 0;
                var nombreUsuario = await ObtenerNombreUsuarioAsync();

                // Una anotación por cada pedido referenciado por el remito
                var pedidosAnotados = 0;
                foreach (var grupo in movimientos.GroupBy(m => m.NuPedCli!.Value))
                {
                    await _anotacionesService.RegistrarItemsRemitidosAsync(
                        grupo.Key, grupo.ToList(), request.RemitoLargo, nombreUsuario, numUsuar);
                    pedidosAnotados++;
                }

                return Ok(new { request.RemitoLargo, PedidosAnotados = pedidosAnotados });
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al registrar ítems remitidos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene un pedido específico por ID
        /// </summary>
        /// <param name="id">ID del pedido</param>
        /// <returns>Pedido encontrado</returns>
        [HttpGet("{id:int}")]
        public async Task<ActionResult> GetAsync(int id)
        {
            try
            {
                var pedido = await _context.Pedencas.FirstOrDefaultAsync(x => x.NroPed == id);
                if (pedido == null)
                {
                    return NotFound();
                }
                return Ok(pedido);
            }
            catch (Exception ex)
            {
                // Manejo de errores para la obtención individual del pedido
                return BadRequest($"Error al obtener el pedido: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea un nuevo pedido
        /// </summary>
        /// <param name="pedido">Datos del pedido a crear</param>
        /// <returns>Pedido creado</returns>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Pedenca pedido)
        {
            try
            {
                // Generar nuevo número de pedido
                var maxNroPed = await _context.Pedencas.MaxAsync(x => (int?)x.NroPed) ?? 0;
                pedido.NroPed = maxNroPed + 1;
                
                // Establecer fecha de emisión si no está definida
                if (!pedido.FechEmis.HasValue)
                {
                    pedido.FechEmis = DateTime.Now;
                }

                // Establecer estado inicial
                if (!pedido.Status.HasValue)
                {
                    pedido.Status = 1; // Estado activo
                }

                _context.Add(pedido);
                await _context.SaveChangesAsync();
                return Ok(pedido);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// Actualiza un pedido existente
        /// </summary>
        /// <param name="pedido">Datos del pedido a actualizar</param>
        /// <returns>Pedido actualizado</returns>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Pedenca pedido)
        {
            try
            {
                _context.Update(pedido);
                await _context.SaveChangesAsync();
                return Ok(pedido);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// Elimina un pedido (cambio de estado)
        /// </summary>
        /// <param name="id">ID del pedido a eliminar</param>
        /// <returns>Resultado de la operación</returns>
        [HttpDelete("{id:int}")]
        public async Task<ActionResult> DeleteAsync(int id)
        {
            var pedido = await _context.Pedencas.FirstOrDefaultAsync(x => x.NroPed == id);
            if (pedido == null)
            {
                return NotFound();
            }

            // Cambiar estado en lugar de eliminar físicamente
            pedido.Status = 9;
            pedido.FechAnul = DateTime.Now;
            
            _context.Update(pedido);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Reporte "PEDIDOS C/O.Compra x Fechas" de /pedidos-venta.
        /// Lee directamente el encabezado PEDENCA (sin ir al detalle) filtrando por rango
        /// de fecha de emisión. Excluye siempre el cliente "998" (venta mostrador / genérico),
        /// igual que la consulta original:
        /// SELECT FechEmis, NroPed, NroClie, RasoClie, NroOcom FROM PEDENCA
        /// WHERE FechEmis >= @desde AND FechEmis &lt;= @hasta AND NroClie &lt;&gt; '998'
        /// </summary>
        /// <param name="fechaDesde">Fecha de emisión desde (inclusive)</param>
        /// <param name="fechaHasta">Fecha de emisión hasta (inclusive)</param>
        /// <param name="nroPed">Filtro opcional por número de pedido</param>
        /// <param name="clienteId">Filtro opcional por número de cliente</param>
        /// <param name="rasoClie">Filtro opcional por razón social (contiene)</param>
        /// <param name="nroOcom">Filtro opcional por número de O.Compra (contiene)</param>
        /// <param name="orderBy">Campo de ordenamiento (nombre de propiedad de PedidoVentaConOcomDTO)</param>
        /// <param name="asc">Orden ascendente si es true, descendente si es false</param>
        [HttpGet("con-ocom-por-fechas")]
        public async Task<ActionResult<List<PedidoVentaConOcomDTO>>> GetPedidosConOcomPorFechasAsync(
            [FromQuery] DateTime? fechaDesde = null,
            [FromQuery] DateTime? fechaHasta = null,
            [FromQuery] int? nroPed = null,
            [FromQuery] int? clienteId = null,
            [FromQuery] string? rasoClie = null,
            [FromQuery] string? nroOcom = null,
            [FromQuery] string? orderBy = null,
            [FromQuery] bool asc = true)
        {
            try
            {
                // NroClie "998" es el cliente genérico/venta mostrador: se excluye siempre,
                // igual que en la consulta original (NroClie <> '998').
                var query = _context.Pedencas.AsNoTracking()
                    .Where(p => p.NroClie != 998);

                if (fechaDesde.HasValue)
                {
                    var desde = fechaDesde.Value.Date;
                    query = query.Where(p => p.FechEmis.HasValue && p.FechEmis.Value.Date >= desde);
                }

                if (fechaHasta.HasValue)
                {
                    var hasta = fechaHasta.Value.Date;
                    query = query.Where(p => p.FechEmis.HasValue && p.FechEmis.Value.Date <= hasta);
                }

                if (nroPed.HasValue)
                {
                    query = query.Where(p => p.NroPed == nroPed.Value);
                }

                if (clienteId.HasValue)
                {
                    query = query.Where(p => p.NroClie == clienteId.Value);
                }

                if (!string.IsNullOrWhiteSpace(rasoClie))
                {
                    var r = rasoClie.Trim().ToLower();
                    query = query.Where(p => p.RasoClie != null && p.RasoClie.ToLower().Contains(r));
                }

                if (!string.IsNullOrWhiteSpace(nroOcom))
                {
                    var o = nroOcom.Trim().ToLower();
                    query = query.Where(p => p.NroOcom != null && p.NroOcom.ToLower().Contains(o));
                }

                orderBy = (orderBy ?? string.Empty).Trim();
                query = orderBy switch
                {
                    nameof(PedidoVentaConOcomDTO.NroPed) => asc ? query.OrderBy(p => p.NroPed) : query.OrderByDescending(p => p.NroPed),
                    nameof(PedidoVentaConOcomDTO.NroClie) => asc ? query.OrderBy(p => p.NroClie) : query.OrderByDescending(p => p.NroClie),
                    nameof(PedidoVentaConOcomDTO.RasoClie) => asc ? query.OrderBy(p => p.RasoClie) : query.OrderByDescending(p => p.RasoClie),
                    nameof(PedidoVentaConOcomDTO.NroOcom) => asc ? query.OrderBy(p => p.NroOcom) : query.OrderByDescending(p => p.NroOcom),
                    _ => asc ? query.OrderBy(p => p.FechEmis) : query.OrderByDescending(p => p.FechEmis)
                };

                var rows = await query
                    .Select(p => new PedidoVentaConOcomDTO
                    {
                        FechEmis = p.FechEmis,
                        NroPed = p.NroPed,
                        NroClie = p.NroClie,
                        RasoClie = p.RasoClie ?? string.Empty,
                        NroOcom = p.NroOcom ?? string.Empty
                    })
                    .ToListAsync();

                return Ok(rows);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener el reporte de pedidos con O.Compra por fechas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene combo de pedidos para selección
        /// </summary>
        /// <returns>Lista de pedidos para combo</returns>
        [HttpGet("combo")]
        public async Task<ActionResult> GetComboAsync()
        {
            return Ok(await _context.Pedencas
                .Where(x => x.Status >= 0)
                .OrderByDescending(x => x.FechEmis)
                .Select(x => new { x.NroPed, Descripcion = $"Pedido {x.NroPed} - {x.RasoClie}" })
                .ToListAsync());
        }

        /// <summary>
        /// Busca pedidos por filtros
        /// </summary>
        /// <param name="filtro">Texto de búsqueda</param>
        /// <param name="pagination">Parámetros de paginación</param>
        /// <returns>Lista filtrada de pedidos</returns>
        [HttpGet("buscar")]
        public async Task<ActionResult> BuscarAsync([FromQuery] string filtro, [FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Pedencas.AsQueryable();

            if (!string.IsNullOrWhiteSpace(filtro))
            {
                queryable = queryable.Where(x => 
                    x.RasoClie!.Contains(filtro) ||
                    x.NroOcom!.Contains(filtro) ||
                    x.Referencia!.Contains(filtro));
            }

            queryable = queryable.Where(x => x.Status >= 0);

            return Ok(await queryable
                .OrderByDescending(x => x.FechEmis)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene el seguimiento completo de un pedido (órdenes de fabricación y remitos)
        /// </summary>
        /// <param name="nroPed">Número de pedido</param>
        /// <returns>Información de seguimiento del pedido</returns>
        [HttpGet("seguimiento/{nroPed:int}")]
        public async Task<ActionResult<SeguimientoPedidoDTO>> GetSeguimientoPedidoAsync(int nroPed)
        {
            try
            {
                var seguimiento = new SeguimientoPedidoDTO
                {
                    NroPed = nroPed
                };

                // Obtener el número de cliente del pedido
                var pedido = await _context.Pedencas
                    .Where(p => p.NroPed == nroPed)
                    .Select(p => new { p.NroClie })
                    .FirstOrDefaultAsync();

                if (pedido == null)
                {
                    return NotFound($"No se encontró el pedido {nroPed}");
                }

                // Obtener órdenes de fabricación relacionadas
                var ordenesFabricacion = await _context.Ordefabrs
                    .Where(o => o.NroPed == nroPed)
                    .Select(o => new OrdenFabricacionSeguimientoDTO
                    {
                        CodExte = o.Cod_Exte,
                        Descrip = o.Descrip,
                        CanProy = o.CanProy,
                        CanApro = o.CanApro,
                        IdenLote = o.IdenLote,
                        StatuOrf = o.StatuOrf
                    })
                    .ToListAsync();

                seguimiento.OrdenesFabricacion = ordenesFabricacion;

                // Obtener remitos relacionados (de la tabla MOVISTO)
                // Filtra por: pedido, tipo 'RT' (Remito) y número de cliente
                var remitos = await _context.Movistos
                    .Where(m => m.NuPedCli == nroPed 
                             && m.CodiMovi == "RT" 
                             && m.Nume_Clie == pedido.NroClie)
                    .Select(m => new RemitoSeguimientoDTO
                    {
                        NumeroRemito = m.DoPriLar,
                        CodExte = m.Cod_Exte,
                        Descrip = m.Descrip,
                        CantSalid = m.CantSalid,
                        FechMov = m.FechMov
                    })
                    .OrderBy(m => m.FechMov)
                    .ToListAsync();

                seguimiento.Remitos = remitos;

                return Ok(seguimiento);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener seguimiento del pedido {nroPed}: {ex.Message}");
                return BadRequest($"Error al obtener seguimiento del pedido: {ex.Message}");
            }
        }

     

      
        /// <summary>
        /// Transfiere productos de un pedido entre depósitos
        /// </summary>
        [HttpPost("transferir-por-pedido")]
        public async Task<ActionResult<ValeTransferenciaDTO>> TransferirPorPedido([FromBody] TransferenciaPorPedidoDTO transferencia)
        {
            if (transferencia == null || !transferencia.Items.Any())
            {
                return BadRequest("Debe proporcionar al menos un producto para transferir");
            }

            if (transferencia.DepositoOrigen == transferencia.DepositoDestino)
            {
                return BadRequest("El depósito origen y destino no pueden ser el mismo");
            }

            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                var fecha = DateTime.Now;
                var secuencialCompleto = await _documentNumberGenerator.GenerateNextDocumentNumberAsync("TR");
                var numeroSecuencial = int.Parse(secuencialCompleto.Split('.').Last());
                var anioCorto = fecha.Year % 100;
                var seccion = $"{anioCorto:D2}{fecha:MMdd}";
                
                int? idFlexoft = _userContextService.GetCurrentUserId();
                short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();
                
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Validar trazabilidad y lotes antes de procesar
                foreach (var item in transferencia.Items)
                {
                    // Si el producto es trazable (Trazable = 1), el lote es obligatorio
                    if (item.Trazable == 1)
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

                    // Validar que el producto no sea trazable nivel 2
                    if (item.Trazable == 2)
                    {
                        return BadRequest($"El producto '{item.CodigoExterno}' tiene trazabilidad nivel 2 y no puede ser transferido.");
                    }
                }

                // Cargar unidades de medida de los productos transferidos
                var codigosInternos = transferencia.Items
                    .Select(i => i.CodigoInterno)
                    .Distinct()
                    .ToList();

                var umedidas = await _context.masters
                    .AsNoTracking()
                    .Where(m => codigosInternos.Contains(m.Codint))
                    .ToDictionaryAsync(
                        m => m.Codint,
                        m => m.Umedida ?? string.Empty);

                // Construir referencia con número de pedido
                var referenciaCompleta = $"TR. Asig. Pedido {transferencia.NumeroPedido}";
                if (!string.IsNullOrWhiteSpace(transferencia.Referencia))
                {
                    referenciaCompleta = $"{referenciaCompleta} | {transferencia.Referencia}";
                }

                // Procesar cada item de la transferencia
                foreach (var item in transferencia.Items)
                {
                    // Registro de salida del depósito origen
                    var movimientoSalida = new Movisto
                    {
                        FeReMovi = fecha,
                        CodiMovi = "TR",
                        CantSalid = item.Cantidad,
                        DepoMovi = transferencia.DepositoOrigen,
                        Cod_Inte = item.CodigoInterno,
                        Cod_Exte = item.CodigoExterno,
                        Descrip = item.Descripcion,
                        DoPriCor = $"TR.{numeroSecuencial}",
                        DoPriLar = $"TR.{numeroSecuencial}",
                        DoSecCor = $"TR.{seccion}",
                        DoSecLar = $"TR.{seccion}",
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
                        DepoMovi = transferencia.DepositoDestino,
                        Cod_Inte = item.CodigoInterno,
                        Cod_Exte = item.CodigoExterno,
                        Descrip = item.Descripcion,
                        DoPriCor = $"TR.{numeroSecuencial}",
                        DoPriLar = $"TR.{numeroSecuencial}",
                        DoSecCor = $"TR.{seccion}",
                        DoSecLar = $"TR.{seccion}",
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

                var origen = await _context.Tadeposi
                    .AsNoTracking()
                    .Where(d => d.CodigoDepo == transferencia.DepositoOrigen)
                    .Select(d => d.Descripcion)
                    .FirstOrDefaultAsync();

                var destino = await _context.Tadeposi
                    .AsNoTracking()
                    .Where(d => d.CodigoDepo == transferencia.DepositoDestino)
                    .Select(d => d.Descripcion)
                    .FirstOrDefaultAsync();

                var vale = new ValeTransferenciaDTO
                {
                    Numero = $"TR.{numeroSecuencial}",
                    DoSecCor = $"TR.{seccion}",
                    Fecha = fecha,
                    DepositoOrigen = transferencia.DepositoOrigen,
                    DepositoOrigenNombre = origen ?? transferencia.DepositoOrigen.ToString(),
                    DepositoDestino = transferencia.DepositoDestino,
                    DepositoDestinoNombre = destino ?? transferencia.DepositoDestino.ToString(),
                    Motivo = referenciaCompleta,
                    Items = transferencia.Items.Select(i => new ValeTransferenciaItemDTO
                    {
                        CodigoExterno = i.CodigoExterno,
                        Descripcion = i.Descripcion,
                        UnidadMedida = umedidas.GetValueOrDefault(i.CodigoInterno),
                        NroPed = transferencia.NumeroPedido,
                        Cantidad = i.Cantidad,
                        Lote = i.Lote
                    }).ToList()
                };

                return Ok(vale);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el nombre del usuario activo desde el contexto HTTP.
        /// Usa el claim Name (NombreUsuario) del JWT; si no está, usa el Identity.Name.
        /// </summary>
        private async Task<string?> ObtenerNombreUsuarioAsync()
        {
            var user = _userContextService.GetCurrentUser();
            if (user?.Identity?.IsAuthenticated ?? false)
            {
                // Intentar obtener el nombre desde el claim Name
                var nombre = user.FindFirst(ClaimTypes.Name)?.Value
                             ?? user.Identity?.Name;
                if (!string.IsNullOrEmpty(nombre))
                    return nombre;

                // Fallback: buscar el usuario por ID en la base
                var usuarioId = _userContextService.GetCurrentUsuarioId();
                if (usuarioId.HasValue)
                {
                    var usuario = await _context.Usuarios
                        .FirstOrDefaultAsync(u => u.Id == usuarioId.Value);
                    return usuario?.NombreUsuario;
                }
            }
            return null;
        }
    }
}
