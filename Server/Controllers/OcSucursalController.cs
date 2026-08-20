using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using System.Data.Common;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class OcSucursalController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<OcSucursalController> _logger;
        private readonly AnotacionesPedidoService _anotacionesService;
        private readonly ISyncService _syncService;
        private readonly UserContextService _userContextService;

        public OcSucursalController(
            DataContext context,
            ILogger<OcSucursalController> logger,
            AnotacionesPedidoService anotacionesService,
            ISyncService syncService,
            UserContextService userContextService)
        {
            _context = context;
            _logger = logger;
            _anotacionesService = anotacionesService;
            _syncService = syncService;
            _userContextService = userContextService;
        }

        /// <summary>
        /// Obtiene órdenes de compra de sucursal con paginación y filtros opcionales
        /// Optimizado con JOIN para evitar problema N+1 y paginación en servidor SQL
        /// </summary>
        /// <param name="pagination">Parámetros de paginación (página, registros por página, filtro)</param>
        /// <param name="filtro">1=Todos, 2=Sin pedido, 3=Con pedido</param>
        [HttpGet]
        public async Task<IActionResult> GetOcSucursales([FromQuery] PaginationDTO pagination, int filtro = 1)
        {
            try
            {
                // Consulta base agrupada según el código VB6 con JOIN para obtener razón social en una sola consulta
                // Se convierte Sucursal (short?) a int para poder hacer JOIN con Nume_Cli (int)
                var query = from oc in _context.O_Sucursals.AsNoTracking()
                            join cliente in _context.Clientes.AsNoTracking()
                                on (int?)oc.Sucursal equals (int?)cliente.Nume_Cli into clienteJoin
                            from cliente in clienteJoin.DefaultIfEmpty()
                            where oc.Marborra != -1 // Excluir borrados
                            group new { oc, cliente } by new
                            {
                                oc.Sucursal,
                                oc.OCompraSuc,
                                oc.NclienPedido,
                                oc.VersionCompraSuc,
                                oc.NpedGeneradoEnAvellaneda,
                                oc.Fechemisi,
                                oc.OComCliente,
                                oc.RazCteSuc
                            }
                            into g
                            select new OcSucursalConsultaDTO
                            {
                                Sucursal = g.Key.Sucursal,
                                OCompraSuc = g.Key.OCompraSuc,
                                Fechemisi = g.Key.Fechemisi,
                                NclienPedido = g.Key.NclienPedido,
                                RazCteSuc = g.Key.RazCteSuc,
                                NpedGeneradoEnAvellaneda = g.Key.NpedGeneradoEnAvellaneda,
                                OComCliente = g.Key.OComCliente,
                                VersionCompraSuc = g.Key.VersionCompraSuc,
                                RazonSocialSucursal = g.Max(x => x.cliente != null ? x.cliente.Raso_Cli : null)
                            };

                // Filtrar según el parámetro OPX% del código VB6
                if (filtro == 2) // Solo sin pedido
                {
                    query = query.Where(x => (x.NpedGeneradoEnAvellaneda ?? 0) <= 0);
                }
                else if (filtro == 3) // Solo con pedido
                {
                    query = query.Where(x => (x.NpedGeneradoEnAvellaneda ?? 0) > 0);
                }
                // filtro == 1 muestra todos (sin filtro adicional)

                // Ordenar por P.Repos. (OCompraSuc) de modo descendente
                // Aplicar paginación usando el método de extensión Paginate
                return Ok(await query
                    .OrderByDescending(x => x.OCompraSuc)
                    .Paginate(pagination)
                    .ToListAsync());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener órdenes de compra de sucursal");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación de órdenes de compra de sucursal
        /// </summary>
        /// <param name="pagination">Parámetros de paginación (registros por página, filtro)</param>
        /// <param name="filtro">1=Todos, 2=Sin pedido, 3=Con pedido</param>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetTotalPages([FromQuery] PaginationDTO pagination, int filtro = 1)
        {
            try
            {
                // Validar que RecordsNumber sea mayor que cero
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                // Consulta base agrupada según el código VB6
                var query = from oc in _context.O_Sucursals.AsNoTracking()
                            where oc.Marborra != -1 // Excluir borrados
                            group oc by new
                            {
                                oc.Sucursal,
                                oc.OCompraSuc,
                                oc.NclienPedido,
                                oc.VersionCompraSuc,
                                oc.NpedGeneradoEnAvellaneda,
                                oc.Fechemisi,
                                oc.OComCliente,
                                oc.RazCteSuc
                            }
                            into g
                            select new OcSucursalConsultaDTO
                            {
                                Sucursal = g.Key.Sucursal,
                                OCompraSuc = g.Key.OCompraSuc,
                                Fechemisi = g.Key.Fechemisi,
                                NclienPedido = g.Key.NclienPedido,
                                RazCteSuc = g.Key.RazCteSuc,
                                NpedGeneradoEnAvellaneda = g.Key.NpedGeneradoEnAvellaneda,
                                OComCliente = g.Key.OComCliente,
                                VersionCompraSuc = g.Key.VersionCompraSuc
                            };

                // Filtrar según el parámetro OPX% del código VB6
                if (filtro == 2) // Solo sin pedido
                {
                    query = query.Where(x => (x.NpedGeneradoEnAvellaneda ?? 0) <= 0);
                }
                else if (filtro == 3) // Solo con pedido
                {
                    query = query.Where(x => (x.NpedGeneradoEnAvellaneda ?? 0) > 0);
                }

                // Obtener el conteo total de registros
                double count = await query.CountAsync();

                // Si no hay registros, devolver 1 página
                if (count == 0)
                {
                    return Ok(1);
                }

                // Calcular el número total de páginas
                double totalPages = Math.Ceiling(count / pagination.RecordsNumber);

                // Asegurar que siempre haya al menos una página
                totalPages = Math.Max(1, totalPages);

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular el total de páginas de órdenes de compra de sucursal");
                return StatusCode(500, $"Error al calcular el total de páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el detalle de productos de una orden de compra de sucursal
        /// </summary>
        /// <param name="oCompraSuc">Orden de compra de sucursal</param>
        /// <param name="sucursal">Código de sucursal</param>
        [HttpGet("detalle")]
        public async Task<ActionResult<List<OcSucursalDetalleDTO>>> GetDetalle(int oCompraSuc, int sucursal)
        {
            try
            {
                // Consulta base para obtener productos de OC_SUCURSAL
                // El código externo se obtiene de master.Codigo ya que OC_SUCURSAL.CodExte está en null
                var query = from oc in _context.O_Sucursals.AsNoTracking()
                            join master in _context.masters.AsNoTracking()
                                on oc.Codint equals master.Codint into masterJoin
                            from master in masterJoin.DefaultIfEmpty()
                            where oc.OCompraSuc == oCompraSuc
                            && oc.Sucursal == sucursal
                            && oc.Marborra != -1
                            select new OcSucursalDetalleDTO
                            {
                                CodExt = master != null ? master.Codigo : "",
                                Descripcion = master != null ? master.Descripcion : null,
                                CantidadCompra = oc.Cantidad,
                                FechaSolicitudCompra = oc.FechaSoliciCompra,
                                NpedGeneradoEnAvellaneda = oc.NpedGeneradoEnAvellaneda
                            };

                var detalles = await query.ToListAsync();

                // Para cada detalle, si tiene pedido generado, buscar cantidades en PEDDETA
                foreach (var detalle in detalles)
                {
                    if (detalle.NpedGeneradoEnAvellaneda.HasValue && detalle.NpedGeneradoEnAvellaneda > 0)
                    {
                        // Buscar en PEDDETA usando Nroped y Codint
                        // Nota: Necesitamos el Codint de OC_SUCURSAL para hacer el JOIN
                        var ocRecord = await _context.O_Sucursals.AsNoTracking()
                            .FirstOrDefaultAsync(x => x.OCompraSuc == oCompraSuc 
                                && x.Sucursal == sucursal 
                                && x.CodExte == detalle.CodExt
                                && x.Marborra != -1);

                        if (ocRecord != null && ocRecord.Codint.HasValue)
                        {
                            var peddeta = await _context.Peddetas.AsNoTracking()
                                .FirstOrDefaultAsync(p => p.NroPed == detalle.NpedGeneradoEnAvellaneda.Value
                                    && p.CodiInt == ocRecord.Codint.Value);

                            if (peddeta != null)
                            {
                                detalle.CantidadPedida = peddeta.CanPed;
                                detalle.CantidadEntregada = peddeta.CantEnt;
                                // Calcular saldo: CANPEDIDA - CANENTRE, si es menor a 0.05 poner 0
                                var saldo = (peddeta.CanPed ?? 0) - (peddeta.CantEnt ?? 0);
                                detalle.Saldo = saldo < 0.05m ? 0 : saldo;
                            }
                        }
                    }
                }

                return Ok(detalles);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener detalle de orden de compra de sucursal");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Obtiene los mensajes de una orden de compra de sucursal
        /// </summary>
        /// <param name="oCompraSuc">Orden de compra de sucursal</param>
        /// <param name="sucursal">Código de sucursal</param>
        [HttpGet("mensajes")]
        public async Task<ActionResult<List<MensaSucursalDTO>>> GetMensajes(int oCompraSuc, int sucursal)
        {
            try
            {
                var mensajes = await _context.MensaSucursals.AsNoTracking()
                    .Where(m => m.OCompraSuc == oCompraSuc && m.Sucursal == sucursal)
                    .Select(m => new MensaSucursalDTO
                    {
                        OCompraSuc = m.OCompraSuc,
                        Mensaje = m.Mensaje,
                        Nroped = null // Se podría obtener si está relacionado con un pedido
                    })
                    .ToListAsync();

                return Ok(mensajes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener mensajes de orden de compra de sucursal");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// Genera un pedido a partir de una orden de compra de sucursal
        /// </summary>
        /// <param name="request">Datos para generar el pedido</param>
        [HttpPost("generar-pedido")]
        public async Task<ActionResult<int>> GenerarPedido([FromBody] GenerarPedidoDTO request)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // Obtener el siguiente número de pedido con bloqueo para evitar duplicados
                var connection = _context.Database.GetDbConnection();
                if (connection.State != System.Data.ConnectionState.Open)
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

                // Obtener datos de la orden de compra de sucursal
                var ocSucursal = await _context.O_Sucursals
                    .FirstOrDefaultAsync(x => x.Sucursal == request.Sucursal && x.OCompraSuc == request.OCompraSuc);

                if (ocSucursal == null)
                {
                    return NotFound("No se encontró la orden de compra de sucursal");
                }

                // El cliente del pedido es el número de sucursal (columna SUCURSAL)
                var nroClie = request.Sucursal;

                // Obtener razón social del cliente desde DEUDOR12 (DbSet Clientes) donde NUME_CLI = SUCURSAL
                var cliente = await _context.Clientes
                    .FirstOrDefaultAsync(c => c.Nume_Cli == nroClie);
                var rasoClie = cliente?.Raso_Cli ?? "";

                // Generar referencia: "PR-{OCompraSuc} - {Fecha}"
                var fechaHoy = DateTime.Now;
                var referencia = $"PR-{request.OCompraSuc} - {fechaHoy:dd/MM/yyyy}";

                // Crear encabezado PEDENCA
                var pendenca = new Pedenca
                {
                    NroPed = nuevoNroPed,
                    FechEmis = fechaHoy,
                    NroClie = nroClie,
                    RasoClie = rasoClie,
                    NroOcom = $"PR-{request.OCompraSuc}",
                    Referencia = referencia,
                    Status = 0,
                    ListPre = 0,
                    CondPag = 0,
                    Vended = 0,
                    CodiEmpr = 0
                };

                _context.Pedencas.Add(pendenca);
                await _context.SaveChangesAsync();

                // Obtener items de OC_SUCURSAL para generar PEDDETA
                var ocItems = await _context.O_Sucursals
                    .Where(x => x.Sucursal == request.Sucursal && x.OCompraSuc == request.OCompraSuc && x.Marborra != -1)
                    .ToListAsync();

                var itemNumero = 0;
                var detallesPedido = new List<Peddeta>();
                
                foreach (var item in ocItems)
                {
                    itemNumero++;
                    var codInt = item.Codint ?? 0;
                    var cantidad = item.Cantidad ?? 0;
                    var fechaSol = item.FechaSoliciCompra ?? DateTime.Now;

                    // Obtener código externo y descripción
                    var codExt = item.CodExte ?? "";
                    var descripcion = "";
                    if (codInt > 0)
                    {
                        var master = await _context.masters
                            .FirstOrDefaultAsync(m => m.Codint == codInt);
                        descripcion = master?.Descripcion ?? "";
                    }

                    // Crear detalle PEDDETA
                    var peddeta = new Peddeta
                    {
                        NroPed = nuevoNroPed,
                        FechEmis = fechaHoy,
                        NroClie = nroClie,
                        RaSoClie = rasoClie,
                        NroOcom = $"PR-{request.OCompraSuc}",
                        FeSolEnt = fechaSol,
                        Status = 0,
                        ListPre = 0,
                        CondPag = 0,
                        Vended = 0,
                        NuOrItem = (short)itemNumero,
                        CodiInt = codInt,
                        CodiExt = codExt,
                        Descrip = descripcion,
                        VenRep = "V",
                        MoneEmis = 1,
                        PreList = 0,
                        CanAsig = 0,
                        CanPed = cantidad,
                        PreUnid = 0,
                        ImpoTot = 0,
                        CantEnt = 0,
                        CanSaldo = cantidad,
                        FeUltEnt = null,
                        CodiEmpr = 0
                    };

                    _context.Peddetas.Add(peddeta);
                    detallesPedido.Add(peddeta);
                }

                await _context.SaveChangesAsync();

                // Generar anotación usando el servicio AnotacionesPedidoService con formato VB6
                // Esto genera el encabezado con fecha, usuario y texto "Generación de Orden de Venta:"
                // y luego cada item con código, descripción y cantidad en formato posicional
                await _anotacionesService.RegistrarCambiosDetalleAsync(nuevoNroPed, detallesPedido, "GENERADO DESDE OC SUCURSAL");

                // Actualizar OC_SUCURSAL con el número de pedido generado
                foreach (var item in ocItems)
                {
                    item.NpedGeneradoEnAvellaneda = nuevoNroPed;
                }
                await _context.SaveChangesAsync();

                await transaction.CommitAsync();
                // ====================================================================
                // GENERAR ARCHIVO ARCHIVO_GENERO_PEDIDO CON ACCION GENEPEDI
                // ====================================================================
                try
                {
                    var numUsuar = _userContextService.GetCurrentUserId() ?? 0;
                    var nroclie = request.Sucursal;
                    var oCompraSuc = request.OCompraSuc;

                    await _syncService.GenerarArchivoGeneroPedidoAsync(
                        nuevoNroPed,
                        nroClie,
                        oCompraSuc,
                        "GENEPEDI",
                        numUsuar
                    );
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error al generar ARCHIVO_GENERO_PEDIDO al generar pedido desde OC Sucursal");
                    // No fallar la generación del pedido si hay error al generar el archivo
                }

                return Ok(nuevoNroPed);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al generar pedido desde orden de compra de sucursal");
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
