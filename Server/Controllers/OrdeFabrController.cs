using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Claims;

namespace IMPLANPROD.Server.Controllers
{
    // TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Actualización del controlador para usar el servicio centralizado
    // de contexto de usuario en lugar de inyectar múltiples servicios relacionados con el usuariO

    /// <summary>
    /// Controlador para la gestión de órdenes de fabricación
    /// </summary>
    [ApiController]
    [Route("/api/ordefabr")]
    public class OrdeFabrController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<OrdeFabrController> _logger;
        private readonly IUserContextService _userContextService;
        private readonly IMoviStockService _moviStockService;

        /// <summary>
        /// Constructor del controlador de transferencias de stock
        /// </summary>
        /// <param name="context">Contexto de datos</param>
        /// <param name="logger">Logger para registro de eventos</param>
        /// <param name="userContextService">Servicio centralizado de contexto de usuario</param>
        /// <param name="moviStockService">Servicio para movimientos de stock</param>
        public OrdeFabrController(DataContext context, ILogger<OrdeFabrController> logger, IUserContextService userContextService, IMoviStockService moviStockService)
        {
            _context = context;
            _logger = logger;
            _userContextService = userContextService;
            _moviStockService = moviStockService;
        }

        /// <summary>
        /// Obtiene una lista paginada de órdenes de fabricación con filtro de estado y fechas
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, short statuOrf = 0, DateTime? fechaDesde = null, DateTime? fechaHasta = null, string? idSubOrPrefix = null)
        {
            try
            {
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                var queryable = _context.Ordefabrs.AsNoTracking().AsQueryable();

                // Filtro por estado
                if (statuOrf == 0)
                {
                    // Mostrar todas
                }
                else if (statuOrf < 3)
                {
                    // Mostrar abiertas (StatuOrf < 3)
                    queryable = queryable.Where(x => x.StatuOrf < 3);
                }
                else
                {
                    // Mostrar cerradas (StatuOrf >= 3)
                    queryable = queryable.Where(x => x.StatuOrf >= 3);
                }

                // Filtro por rango de fechas
                if (fechaDesde.HasValue)
                {
                    queryable = queryable.Where(x => x.FechGen >= fechaDesde);
                }

                if (fechaHasta.HasValue)
                {
                    var fechaHastaFinal = fechaHasta.Value.Date.AddDays(1).AddTicks(-1);//para que considere la fecha hassta con las horas 
                    queryable = queryable.Where(x => x.FechGen <= fechaHastaFinal);
                }

                // Filtro por texto
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    queryable = queryable.Where(x =>
                        (x.IdSubOr != null && x.IdSubOr.ToLower().Contains(pagination.Filter.ToLower())) ||
                        (x.Id_Orfa != null && x.Id_Orfa.ToLower().Contains(pagination.Filter.ToLower())) ||
                        (x.Descrip != null && x.Descrip.ToLower().Contains(pagination.Filter.ToLower()))
                    );
                }

                // Filtro por prefijo de suborden (ej: "DL-" para limpieza)
                if (!string.IsNullOrWhiteSpace(idSubOrPrefix))
                {
                    queryable = queryable.Where(x => x.IdSubOr != null && x.IdSubOr.StartsWith(idSubOrPrefix));
                }

                if (!await queryable.AnyAsync())
                {
                    return Ok(new List<OrdefabrIndexDTO>());
                }

                // Proyección explícita de la entidad Ordefabr al DTO OrdefabrIndexDTO
                // Cuando se filtra por prefijo (ej: DL-), se ordena por IdSubOr descendente ("de mayor a menor").
                var orderedQueryable = !string.IsNullOrWhiteSpace(idSubOrPrefix)
                    ? queryable.OrderByDescending(x => x.IdSubOr)
                    : queryable.OrderByDescending(x => x.FechGen);

                var result = await orderedQueryable
                    .Paginate(pagination)
                    .Select(x => new OrdefabrIndexDTO
                    {
                        Id = x.Id,
                        FechGen = x.FechGen,
                        IdSubOr = x.IdSubOr,
                        CanProy = x.CanProy,
                        CanApro = x.CanApro,
                        Id_Orfa = x.Id_Orfa,
                        StatuOrf = x.StatuOrf,
                        // Agregamos los campos solicitados para mostrar datos del producto
                        Cod_Exte = x.Cod_Exte,
                        Unidad = x.Unidad,
                        FechEnSol = x.FechEnSol,
                        FechLan = x.FechLan,
                        FechCierr = x.FechCierr,
                        Descrip = x.Descrip,
                        Cod_Inte = x.Cod_Inte
                    })
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la lista de órdenes de fabricación");
                return StatusCode(500, $"Error al obtener la lista de órdenes de fabricación: {ex.Message}");
            }
        }

        /// <summary>
        /// Devuelve el total de páginas para la paginación de órdenes de fabricación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, short statuOrf = 0, DateTime? fechaDesde = null, DateTime? fechaHasta = null, string? idSubOrPrefix = null)
        {
            try
            {
                // Validar que RecordsNumber sea mayor que cero
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                var queryable = _context.Ordefabrs.AsQueryable();

                // Filtro por estado
                if (statuOrf == 0)
                {
                    // Mostrar todas
                }
                else if (statuOrf < 3)
                {
                    // Mostrar abiertas (StatuOrf < 3)
                    queryable = queryable.Where(x => x.StatuOrf < 3);
                }
                else
                {
                    // Mostrar cerradas (StatuOrf >= 3)
                    queryable = queryable.Where(x => x.StatuOrf >= 3);
                }

                // Filtro por rango de fechas
                if (fechaDesde.HasValue)
                {
                    queryable = queryable.Where(x => x.FechGen >= fechaDesde);
                }

                if (fechaHasta.HasValue)
                {
                    queryable = queryable.Where(x => x.FechGen <= fechaHasta);
                }

                // Filtro por texto
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    queryable = queryable.Where(x =>
                        (x.IdSubOr != null && x.IdSubOr.ToLower().Contains(pagination.Filter.ToLower())) ||
                        (x.Id_Orfa != null && x.Id_Orfa.ToLower().Contains(pagination.Filter.ToLower())) ||
                        (x.Descrip != null && x.Descrip.ToLower().Contains(pagination.Filter.ToLower()))
                    );
                }

                // Filtro por prefijo de suborden (ej: "DL-" para limpieza)
                if (!string.IsNullOrWhiteSpace(idSubOrPrefix))
                {
                    queryable = queryable.Where(x => x.IdSubOr != null && x.IdSubOr.StartsWith(idSubOrPrefix));
                }

                // Obtener el conteo total de registros que cumplen con el filtro
                double count = await queryable.CountAsync();

                // Si no hay registros, devolver 1 página (para mostrar el mensaje "No hay registros")
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
                // Log del error para debugging
                _logger.LogError(ex, "Error al calcular el total de páginas");
                return StatusCode(500, $"Error al calcular el total de páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene una orden de fabricación individual por Id
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var ordenFabr = await _context.Ordefabrs.FirstOrDefaultAsync(x => x.Id == id);
            if (ordenFabr == null)
            {
                return NotFound();
            }
            return Ok(ordenFabr);
        }

        /// <summary>
        /// Obtiene un listado de secciones productivas para el combo
        /// </summary>
        [HttpGet("secprod/combo")]
        public async Task<IActionResult> GetComboSecProd()
        {
            try
            {
                var secProd = await _context.SecProds
                    .Where(x => x.Status >= 0)
                    .Select(x => new { Id = x.CodigoSecProd, Codigo = x.CodigoSecProd, x.Descripcion })
                    .OrderBy(x => x.Codigo)
                    .ToListAsync();

                return Ok(secProd);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener el combo de secciones productivas");
                return StatusCode(500, $"Error al obtener el combo de secciones productivas: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea una nueva orden de fabricación y genera las reservas de componentes según la fórmula
        /// </summary>
        [HttpPost("crearOrdenFabricacion")]
        public async Task<IActionResult> CrearOrdenFabricacion([FromBody] OrdeFabr nuevaOrden)
        {   // TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Uso del servicio centralizado para obtener información del usuario
            int? idFlexoft = _userContextService.GetCurrentUserId();
            short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();

            // Verificar si el usuario está autenticado usando el servicio centralizado
            if (!_userContextService.IsAuthenticated())
            {
                return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
            }

            // Iniciamos una transacción para asegurar la integridad de los datos
            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);
            try
            {
                //Asignar estado por defecto
                nuevaOrden.StatuOrf = 1;

                // Obtener el último número de orden
                var ultimoNumero = await _context.Ordefabrs
                    .OrderByDescending(o => o.NumeOrFa)
                    .Select(o => o.NumeOrFa)
                    .FirstOrDefaultAsync();

                // Calcular el nuevo número
                nuevaOrden.NumeOrFa = ultimoNumero + 1;

                // Asignar fechas por defecto
                var fechaPorDefecto = new DateTime(1980, 1, 1, 0, 0, 0);
                nuevaOrden.FechProyEntre = fechaPorDefecto;
                nuevaOrden.FechTempEntre = fechaPorDefecto;
                nuevaOrden.CodiEmpr = codigoEmpresa;
                nuevaOrden.NumUsuar = idFlexoft;
                // Formatear y asignar IdSubOr correctamente
                nuevaOrden.IdSubOr = $"OF-{nuevaOrden.NumeOrFa:000000}-1";
                
                // Guardar en el campo Referen la concatenación de cod_exte + " - " + Descrip + " - " + Unidad
                nuevaOrden.Referen = $"{nuevaOrden.Cod_Exte} - {nuevaOrden.Descrip} - {nuevaOrden.Unidad}";

                // Obtener el nombre del usuario y generar el texto para AnotaOrf
                string nombreUsuario = ObtenerNombreUsuario();
                string idUsuario = idFlexoft.HasValue ? idFlexoft.Value.ToString() : "";
                string userTer = $"{idUsuario} - {nombreUsuario}";
                nuevaOrden.AnotaOrf = GenerarTextoAnotaOrf(nuevaOrden, userTer);

                // Guardar la orden de fabricación en base de datos
                _context.Ordefabrs.Add(nuevaOrden);
                await _context.SaveChangesAsync();

                // Generar las reservas de componentes según la fórmula del producto
                if (nuevaOrden.Cod_Inte.HasValue)
                {
                    await GenerarReservasDeFormula(nuevaOrden.Cod_Inte.Value, nuevaOrden.IdSubOr, nuevaOrden.CanProy, codigoEmpresa.Value);
                }

                // Confirmar la transacción
                await transaction.CommitAsync();

                return Ok(nuevaOrden);
            }
            catch (Exception ex)
            {
                // Si hay algún error, revertimos la transacción
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al crear la orden de fabricación y sus reservas");
                return StatusCode(500, $"Error al crear la orden de fabricación: {ex.Message}");
            }
        }

        [HttpPost("generar-limpieza")]
        public async Task<IActionResult> GenerarOrdenFabricacionLimpieza([FromBody] GenerarOFLimpiezaRequestDTO request)
        {
            int? idFlexoft = _userContextService.GetCurrentUserId();
            short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();

            if (!_userContextService.IsAuthenticated())
            {
                return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
            }

            if (request == null || request.NumeroReparacion <= 0)
            {
                return BadRequest("Primero debe seleccionar una nota de reparación.");
            }

            if (request.Items == null || request.Items.Count == 0)
            {
                return BadRequest("Debe agregar al menos un producto.");
            }

            foreach (var item in request.Items)
            {
                if (item == null || item.CodInte <= 0)
                {
                    return BadRequest("Hay productos inválidos (CODINT).");
                }

                if (item.Cantidad <= 0)
                {
                    return BadRequest("Hay productos con cantidad inválida.");
                }
            }

            // Formato requerido: DL- + NNNNNN + -1
            string idSubOr = $"DL-{request.NumeroReparacion:000000}-1";

            // Validar que no exista
            bool existe = await _context.Ordefabrs.AsNoTracking().AnyAsync(x => x.IdSubOr == idSubOr);
            if (existe)
            {
                return BadRequest("O/F de Limpieza Ya Generada para la O/R Elegida");
            }

            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);
            try
            {
                // Leer código genérico (VB6: Control$("REPARA", "CODGENER"))
                var configCodGener = await _context.Set<FLEXCRL>()
                    .AsNoTracking()
                    .FirstOrDefaultAsync(f => f.CLABUS == "REPARA" && f.PROBUS == "CODGENER");

                string codigoGenerico = configCodGener?.VALCONTROL ?? string.Empty;
                if (string.IsNullOrWhiteSpace(codigoGenerico))
                {
                    return BadRequest("No se encontró configurado el Código genérico (FLEXCRL: REPARA/CODGENER).");
                }

                // Obtener CODINT desde MASTER (VB6: CODINT(CODI$))
                var masterGenerico = await _context.masters
                    .AsNoTracking()
                    .FirstOrDefaultAsync(m => m.Codigo == codigoGenerico);

                if (masterGenerico == null || masterGenerico.Codint <= 0)
                {
                    return BadRequest($"El Código genérico '{codigoGenerico}' no existe en MASTER o no tiene CODINT válido.");
                }

                var fecha = DateTime.Now.Date;

                var nuevaOrden = new OrdeFabr
                {
                    NumeOrFa = 0,
                    CodiEmpr = 0,
                    NumUsuar = idFlexoft,

                    IdSubOr = idSubOr,
                    Referen = $"O/F de Limpieza para O/Rep: {request.NumeroReparacion}",

                    FechGen = fecha,
                    FechEnSol = fecha,
                    FechLan = fecha,

                    Cod_Inte = masterGenerico.Codint,
                    Cod_Exte = masterGenerico.Codigo,
                    Descrip = "",
                    Unidad = "U.",
                    CanProy = 1,
                    Destino = "STOCKS",
                    Priorid = 0,
                    StatuOrf = 0,
                    NuOrProd = "",
                    SecciPro = 0,
                    CanApro = 0,
                    CanRech = 0,

                    Observa = $"O/R={request.NumeroReparacion}; MO={request.ManoDeObra:0.####}; Items={string.Join(", ", request.Items.Select(i => $"{i.CodInte} x {i.Cantidad}"))}"
                };

                var operacion = new Operfab
                {
                    IdSubOr = idSubOr,
                    CodiEmpr = 0,
                    Cod_Inte = 0,
                    Cod_Exte = "",
                    NumeOper = 1,
                    CantProg = 1,
                    TiemVari = request.ManoDeObra,
                    Descrip = "",
                    DescOper = "",
                    Maquina = "",
                    DescMaq = "",
                    EquiAux = "",
                    Observa = "",
                    HoBatchRec = "",
                    rowguid = Guid.NewGuid(),
                };

                var codsInte = request.Items.Select(x => x.CodInte).Distinct().ToList();
                var mastersByCodint = await _context.masters
                    .AsNoTracking()
                    .Where(m => codsInte.Contains(m.Codint))
                    .ToDictionaryAsync(m => m.Codint);

                short nuOrItem = 0;
                var reservas = new List<Reserva>();
                foreach (var item in request.Items)
                {
                    nuOrItem++;

                    mastersByCodint.TryGetValue(item.CodInte, out var masterItem);

                    reservas.Add(new Reserva
                    {
                        CodiEmpr = 0,
                        IdSubOr = idSubOr,
                        Cod_Inte = item.CodInte,
                        Cod_Exte = masterItem?.Codigo ?? string.Empty,
                        Descrip = masterItem?.Descripcion ?? string.Empty,
                        CantRes = item.Cantidad,
                        UsoUnit = item.Cantidad,
                        FechRes = DateTime.Now,
                        CantCons = 0,
                        Saldo_Entre = 0,
                        IdenLote = " ",
                        Deposito = 0,
                        Status = 0,
                        NuOrItem = nuOrItem
                    });
                }

                _context.Ordefabrs.Add(nuevaOrden);
                _context.Operfabs.Add(operacion);
                _context.Reservas.AddRange(reservas);
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return Ok(nuevaOrden);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al generar O/F de Limpieza");
                return StatusCode(500, $"Error al generar O/F de Limpieza: {ex.Message}");
            }
        }

        public class GenerarOFLimpiezaRequestDTO
        {
            public int NumeroReparacion { get; set; }
            public decimal ManoDeObra { get; set; }
            public string CodigoProductoSeleccionado { get; set; } = string.Empty;
            public decimal Cantidad { get; set; } = 1m;

            public List<GenerarOFLimpiezaItemDTO> Items { get; set; } = new List<GenerarOFLimpiezaItemDTO>();
        }

        public class GenerarOFLimpiezaItemDTO
        {
            public int CodInte { get; set; }
            public decimal Cantidad { get; set; }
        }

        public class OperfabOperacionDTO
        {
            public short? NumeOper { get; set; }
            public string? Descrip { get; set; } = string.Empty;
            public decimal? CantProg { get; set; }
            public decimal? TiemVari { get; set; }
        }

        public class EditarOFLimpiezaItemDTO
        {
            public int CodInte { get; set; }
            public string Cod_Exte { get; set; } = string.Empty;
            public string Descrip { get; set; } = string.Empty;
            public decimal Cantidad { get; set; }
        }

        public class EditarOFLimpiezaResponseDTO
        {
            public string IdSubOr { get; set; } = string.Empty;
            public decimal ManoDeObra { get; set; }
            public List<EditarOFLimpiezaItemDTO> Items { get; set; } = new List<EditarOFLimpiezaItemDTO>();
        }

        public class EditarOFLimpiezaRequestDTO
        {
            public decimal ManoDeObra { get; set; }
            public List<GenerarOFLimpiezaItemDTO> Items { get; set; } = new List<GenerarOFLimpiezaItemDTO>();
        }

        [HttpGet("limpieza/{idSubOr}/edit")]
        public async Task<IActionResult> GetEdicionOFLimpieza(string idSubOr)
        {
            try
            {
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                if (string.IsNullOrWhiteSpace(idSubOr))
                {
                    return BadRequest("El identificador de la orden de fabricación es requerido.");
                }

                // Obtener mano de obra (TiemVari) usando proyección EF para evitar cast Double->Decimal
                var manoDeObra = await _context.Operfabs
                    .AsNoTracking()
                    .Where(o => o.IdSubOr == idSubOr)
                    .Select(o => o.TiemVari ?? 0m)
                    .FirstOrDefaultAsync();

                // Obtener items de reserva usando proyección EF directa a DTO
                var items = await _context.Reservas
                    .AsNoTracking()
                    .Where(r => r.IdSubOr == idSubOr && r.Cod_Inte > 0 && (r.CantRes ?? 0) > 0)
                    .OrderBy(r => r.NuOrItem)
                    .Select(r => new EditarOFLimpiezaItemDTO
                    {
                        CodInte = r.Cod_Inte,
                        Cod_Exte = r.Cod_Exte ?? string.Empty,
                        Descrip = r.Descrip ?? string.Empty,
                        Cantidad = r.CantRes ?? 0m
                    })
                    .ToListAsync();

                var response = new EditarOFLimpiezaResponseDTO
                {
                    IdSubOr = idSubOr,
                    ManoDeObra = manoDeObra,
                    Items = items
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener edición O/F Limpieza");
                return StatusCode(500, $"Error al obtener edición: {ex.Message}");
            }
        }

        [HttpPut("limpieza/{idSubOr}/edit")]
        public async Task<IActionResult> PutEdicionOFLimpieza(string idSubOr, [FromBody] EditarOFLimpiezaRequestDTO request)
        {
            if (!_userContextService.IsAuthenticated())
            {
                return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
            }

            if (string.IsNullOrWhiteSpace(idSubOr))
            {
                return BadRequest("El identificador de la orden de fabricación es requerido.");
            }

            if (request == null)
            {
                return BadRequest("Datos inválidos.");
            }

            if (request.Items == null || request.Items.Count == 0)
            {
                return BadRequest("Debe agregar al menos un producto.");
            }

            foreach (var item in request.Items)
            {
                if (item == null || item.CodInte <= 0)
                {
                    return BadRequest("Hay productos inválidos (CODINT).");
                }

                if (item.Cantidad <= 0)
                {
                    return BadRequest("Hay productos con cantidad inválida.");
                }
            }

            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);
            try
            {
                // Actualizar TiemVari usando ExecuteUpdateAsync (EF Core 7+) para evitar materialización
                var filasOperfab = await _context.Operfabs
                    .Where(o => o.IdSubOr == idSubOr)
                    .ExecuteUpdateAsync(s => s.SetProperty(o => o.TiemVari, request.ManoDeObra));

                if (filasOperfab == 0)
                {
                    return NotFound($"No se encontró OPERFAB para la O/F {idSubOr}.");
                }

                var reservasExistentes = await _context.Reservas
                    .Where(r => r.IdSubOr == idSubOr)
                    .ToListAsync();

                var existentesPorCodInte = reservasExistentes
                    .Where(r => r.Cod_Inte > 0)
                    .GroupBy(r => r.Cod_Inte)
                    .ToDictionary(g => g.Key, g => g.OrderBy(x => x.NuOrItem).First());

                var codsNuevos = request.Items.Select(i => i.CodInte).Distinct().ToList();
                var mastersByCodint = await _context.masters
                    .AsNoTracking()
                    .Where(m => codsNuevos.Contains(m.Codint))
                    .ToDictionaryAsync(m => m.Codint);

                var maxNuOrItem = reservasExistentes.Count == 0 ? (short)0 : (short)(reservasExistentes.Max(r => r.NuOrItem) ?? 0);
                var codsEnRequest = new HashSet<int>(codsNuevos);

                foreach (var item in request.Items)
                {
                    if (existentesPorCodInte.TryGetValue(item.CodInte, out var reservaExistente))
                    {
                        if ((reservaExistente.CantCons ?? 0m) > item.Cantidad)
                        {
                            return BadRequest($"La cantidad reservada no puede ser menor a lo consumido. CODINT={item.CodInte} Cons={reservaExistente.CantCons:0.####} Res={item.Cantidad:0.####}");
                        }

                        reservaExistente.CantRes = item.Cantidad;
                        reservaExistente.UsoUnit = item.Cantidad;

                        if (mastersByCodint.TryGetValue(item.CodInte, out var masterItem))
                        {
                            reservaExistente.Cod_Exte = masterItem.Codigo;
                            reservaExistente.Descrip = masterItem.Descripcion;
                        }
                    }
                    else
                    {
                        maxNuOrItem++;
                        mastersByCodint.TryGetValue(item.CodInte, out var masterItem);

                        _context.Reservas.Add(new Reserva
                        {
                            CodiEmpr = 0,
                            IdSubOr = idSubOr,
                            Cod_Inte = item.CodInte,
                            Cod_Exte = masterItem?.Codigo ?? string.Empty,
                            Descrip = masterItem?.Descripcion ?? string.Empty,
                            CantRes = item.Cantidad,
                            UsoUnit = item.Cantidad,
                            FechRes = DateTime.Now,
                            CantCons = 0,
                            Saldo_Entre = 0,
                            IdenLote = " ",
                            Deposito = 0,
                            Status = 0,
                            NuOrItem = maxNuOrItem
                        });
                    }
                }

                foreach (var r in reservasExistentes)
                {
                    if (r.Cod_Inte > 0 && !codsEnRequest.Contains(r.Cod_Inte))
                    {
                        _context.Reservas.Remove(r);
                    }
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return Ok(new { mensaje = "O/F Limpieza actualizada", idSubOr });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al editar O/F Limpieza");
                return StatusCode(500, $"Error al editar O/F Limpieza: {ex.Message}");
            }
        }

        [HttpGet("operaciones/por-of/{idSubOr}")]
        public async Task<IActionResult> GetOperacionesPorOf(string idSubOr)
        {
            try
            {
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                if (string.IsNullOrWhiteSpace(idSubOr))
                {
                    return BadRequest("El identificador de la orden de fabricación es requerido.");
                }

                // Obtener operaciones usando proyección EF directa a DTO para evitar cast Double->Decimal
                var resultados = await _context.Operfabs
                    .AsNoTracking()
                    .Where(o => o.IdSubOr == idSubOr)
                    .OrderBy(o => o.NumeOper)
                    .Select(o => new OperfabOperacionDTO
                    {
                        NumeOper = o.NumeOper,
                        Descrip = o.Descrip ?? string.Empty,
                        CantProg = o.CantProg,
                        TiemVari = o.TiemVari
                    })
                    .ToListAsync();

                return Ok(resultados);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener operaciones por O/F");
                return StatusCode(500, $"Error al obtener operaciones: {ex.Message}");
            }
        }

        public async Task<IActionResult> Put(int id, OrdeFabr ordenFabr)
        {
            try
            {
                // Verificamos que el ID sea válido
                if (id != ordenFabr.Id)
                {
                    return BadRequest("El ID de la orden de fabricación no coincide");
                }

                // Verificamos que la orden exista
                var ordenExistente = await _context.Ordefabrs.AsNoTracking().FirstOrDefaultAsync(x => x.Id == id);
                if (ordenExistente == null)
                {
                    return NotFound($"No se encontró la orden de fabricación con ID {id}");
                }

                // Iniciamos una transacción para asegurar la integridad de los datos
                await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);
                try
                {
                    // Verificamos si existe un producto con el código interno especificado
                    if (ordenFabr.Cod_Inte.HasValue)
                    {
                        var master = await _context.masters.FirstOrDefaultAsync(x => x.Codint == ordenFabr.Cod_Inte.Value);
                        if (master == null)
                        {
                            return BadRequest($"No existe un producto con el código interno {ordenFabr.Cod_Inte.Value}");
                        }

                        // Completamos datos del producto si han cambiado
                        ordenFabr.Cod_Exte = master.Codigo;
                        ordenFabr.Descrip = master.Descripcion;
                        ordenFabr.Unidad = master.Umedida;
                    }

                    // Actualizamos la orden de fabricación
                    _context.Update(ordenFabr);
                    await _context.SaveChangesAsync();

                    // Confirmamos la transacción
                    await transaction.CommitAsync();

                    return Ok(ordenFabr);
                }
                catch (Exception ex)
                {
                    // Si hay algún error, revertimos la transacción
                    await transaction.RollbackAsync();
                    throw new Exception($"Error al actualizar la orden de fabricación: {ex.Message}", ex);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar la orden de fabricación");
                return StatusCode(500, $"Error al actualizar la orden de fabricación: {ex.Message}");
            }
        }

        /// <summary>
        /// Procesa el cierre avanzado de una orden de fabricación con validaciones y persistencia completa
        /// Replica la lógica del sistema VB6 original con transacciones y anotaciones
        /// </summary>
        /// <param name="datosCierre">Datos del cierre incluyendo cantidades, depósitos y componentes</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPost("procesarCierreAvanzado")]
        public async Task<IActionResult> ProcesarCierreAvanzado([FromBody] DatosCierreAvanzadoDTO datosCierre)
        {
            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);
            try
            {
                // Verificar autenticación
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Buscar la orden de fabricación con AsNoTracking para evitar conflictos de seguimiento
                var ordenOriginal = await _context.Ordefabrs.AsNoTracking().FirstOrDefaultAsync(o => o.Id == datosCierre.IdOrden);
                if (ordenOriginal == null)
                {
                    return NotFound($"No se encontró la orden de fabricación con ID {datosCierre.IdOrden}.");
                }

                // Verificar que la orden esté abierta
                //puede estar cerrada
                //if (ordenOriginal.StatuOrf >= 3)
                //{
                //    return BadRequest($"La orden de fabricación {ordenOriginal.IdSubOr} ya está cerrada.");
                //}

                // Obtener el parámetro CIERCALI de la base de datos para determinar el modo de cierre
                int modoCierreCalidad = await ObtenerParametroCierreCalidad();

                // Actualizar el DTO con el valor correcto obtenido de la base de datos
                datosCierre.ModoCierreCalidad = modoCierreCalidad;

                // Crear una nueva instancia para evitar conflictos de seguimiento
                var orden = new OrdeFabr
                {
                    Id = ordenOriginal.Id,
                    IdSubOr = ordenOriginal.IdSubOr,
                    Cod_Inte = ordenOriginal.Cod_Inte,
                    Cod_Exte = ordenOriginal.Cod_Exte,
                    Descrip = ordenOriginal.Descrip,
                    Unidad = ordenOriginal.Unidad,
                    IdenLote = ordenOriginal.IdenLote,
                    Referen = ordenOriginal.Referen,
                    Nume_Cli = ordenOriginal.Nume_Cli,
                    CanApro = ordenOriginal.CanApro,
                    CanRech = ordenOriginal.CanRech,
                    Canti_Fin_Proceso = ordenOriginal.Canti_Fin_Proceso,
                    Status_Proceso = ordenOriginal.Status_Proceso,
                    StatuOrf = ordenOriginal.StatuOrf,
                    FechCierr = ordenOriginal.FechCierr,
                    AnotaOrf = ordenOriginal.AnotaOrf
                };

                // Aplicar lógica de cierre según el código VB6 original
                await ProcesarCierreOrdenFabricacion(orden, datosCierre);

                // Generar y agregar anotaciones
                await GenerarAnotacionesCierre(orden, datosCierre);

                // Actualizar depósitos en Master si el código del producto fabricado es válido
                await ActualizarDepositosMaster(orden, datosCierre);

                // Actualizar la orden en el contexto y guardar cambios dentro de la transacción
                var ordenExistente = await _context.Ordefabrs.FirstOrDefaultAsync(o => o.Id == orden.Id);
                if (ordenExistente != null)
                {
                    ordenExistente.CanApro = orden.CanApro;
                    ordenExistente.CanRech = orden.CanRech;
                    ordenExistente.Canti_Fin_Proceso = orden.Canti_Fin_Proceso;
                    ordenExistente.Status_Proceso = orden.Status_Proceso;
                    ordenExistente.StatuOrf = orden.StatuOrf;
                    ordenExistente.FechCierr = orden.FechCierr;
                    ordenExistente.AnotaOrf = orden.AnotaOrf;
                    
                    _context.Ordefabrs.Update(ordenExistente);
                }

                // Guardar todos los cambios dentro de la transacción (incluye ORFAPEN, ORDEFABR, MASTER, etc.)
                await _context.SaveChangesAsync();

                // Confirmar transacción
                await transaction.CommitAsync();

                return Ok(new { mensaje = "Cierre procesado exitosamente", orden = orden.IdSubOr });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al procesar cierre avanzado de orden de fabricación");
                return StatusCode(500, $"Error al procesar el cierre: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza los depósitos de entrada y consumo en la tabla Master para el producto fabricado
        /// Equivalente a la lógica VB6: If CICONJ% > 0 Then UPDATE MASTER SET DepoEntra = DEPX%, DepoCon = DP2% WHERE CODINT = CICONJ%
        /// </summary>
        /// <param name="orden">Orden de fabricación con el código del producto fabricado</param>
        /// <param name="datosCierre">Datos del cierre con los depósitos configurados</param>
        private async Task ActualizarDepositosMaster(OrdeFabr orden, DatosCierreAvanzadoDTO datosCierre)
        {
            // If CICONJ% > 0 Then (CICONJ% = código del producto fabricado = Cod_Inte)
            if (orden.Cod_Inte.HasValue && orden.Cod_Inte.Value > 0)
            {
                // Buscar el registro en Master por código interno
                var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == orden.Cod_Inte.Value);
                
                if (master != null)
                {
                    // UPDATE MASTER SET DepoEntra = DEPX%, DepoCon = DP2% WHERE CODINT = CICONJ%
                    // DEPX% = depósito de entrada del producto fabricado
                    // DP2% = depósito de consumo de los componentes
                    master.DepoEntra = (short?)(datosCierre.DepositoEntrada ?? 0);
                    master.DepoCon = (short?)(datosCierre.DepositoSalida ?? 0);

                    // Actualizar el registro en el contexto
                    _context.masters.Update(master);

                    _logger.LogInformation($"Depósitos actualizados en Master para código interno {orden.Cod_Inte.Value}: DepoEntra={master.DepoEntra}, DepoCon={master.DepoCon}");
                }
                else
                {
                    _logger.LogWarning($"No se encontró registro en Master para código interno {orden.Cod_Inte.Value}");
                }
            }
            else
            {
                _logger.LogInformation("No se actualiza Master: código del producto fabricado no válido o es 0");
            }
        }

        /// <summary>
        /// Cierra una orden de fabricación cambiando su estado a cerrada (statuorf = 3)
        /// </summary>
        /// <param name="id">ID de la orden de fabricación a cerrar</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut("cerrar/{id:int}")]
        public async Task<IActionResult> CerrarOrdenFabricacion(int id)
        {
            try
            {
                // Verificar si el usuario está autenticado usando el servicio centralizado
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Buscar la orden de fabricación
                var orden = await _context.Ordefabrs.FirstOrDefaultAsync(o => o.Id == id);
                if (orden == null)
                {
                    return NotFound($"No se encontró la orden de fabricación con ID {id}.");
                }

                // Verificar que la orden esté abierta (statuorf < 3)
                if (orden.StatuOrf >= 3)
                {
                    return BadRequest($"La orden de fabricación {orden.IdSubOr} ya está cerrada.");
                }

                // Cerrar la orden cambiando el estado a 3
                orden.StatuOrf = 3;

                // Registrar anotación del cierre
                var nombreUsuario = ObtenerNombreUsuario();
                var anotacionCierre = FormatearNuevaAnotacion($"ORDEN CERRADA POR {nombreUsuario}", nombreUsuario);
                
                if (string.IsNullOrEmpty(orden.AnotaOrf))
                {
                    orden.AnotaOrf = anotacionCierre;
                }
                else
                {
                    orden.AnotaOrf += "\r\n" + anotacionCierre;
                }

                // Guardar cambios
                _context.Ordefabrs.Update(orden);
                await _context.SaveChangesAsync();

                _logger.LogInformation($"Orden de fabricación {orden.IdSubOr} (ID: {id}) cerrada exitosamente por {nombreUsuario}");
                
                return Ok(new { mensaje = $"La orden de fabricación {orden.IdSubOr} ha sido cerrada exitosamente." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al cerrar la orden de fabricación con ID {id}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el nombre del usuario actual desde el ClaimsPrincipal
        /// </summary>
        /// <returns>Nombre del usuario o "Usuario" si no se puede obtener</returns>
        private string ObtenerNombreUsuario()
        {
            try
            {
                // Obtener el ClaimsPrincipal del usuario actual
                var user = _userContextService.GetCurrentUser();
                if (user == null) return "Usuario";

                // Intentar obtener el nombre del usuario desde los claims
                var nombreClaim = user.FindFirst(ClaimTypes.Name)?.Value;
                if (!string.IsNullOrEmpty(nombreClaim))
                    return nombreClaim;

                // Si no hay claim de nombre, intentar con el nombre de usuario
                var usernameClaim = user.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (!string.IsNullOrEmpty(usernameClaim))
                    return usernameClaim;

                // Si no se puede obtener ningún identificador, devolver un valor por defecto
                return "Usuario";
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error al obtener el nombre del usuario desde los claims");
                return "Usuario";
            }
        }

        /// <summary>
        /// Genera el texto con formato especial para el campo AnotaOrf
        /// Replica la lógica de concatenación y formato del código VB6
        /// </summary>
        /// <param name="orden">Orden de fabricación</param>
        /// <param name="userTer">Nombre del usuario que crea la orden (ID + Nombre)</param>
        /// <returns>Texto formateado para el campo AnotaOrf</returns>
        private string GenerarTextoAnotaOrf(OrdeFabr orden, string userTer)
        {
            // Replicar el formato exacto del código VB6
            string ttxyz = new string(' ', 76); // Space$(76)
            string hos = DateTime.Now.ToString("dd/MM/yy"); // Equivalente a HOY(HOS$)

            // Call REPLA(TTXYZ$, HOS$, 1, 8)
            ttxyz = Repla(ttxyz, hos, 1, 8);

            // Call REPLA(TTXYZ$, "Generación de de O/Fabricación", 11, 56)
            ttxyz = Repla(ttxyz, "Generación de de O/Fabricación", 11, 56);

            // Call REPLA(TTXYZ$, USERTER$, 69, 24)
            ttxyz = Repla(ttxyz, userTer, 69, 24);

            // TTXYZ$ = TTXYZ$ + Chr$(10) + Chr$(13)
            ttxyz += "\r\n";

            // TTXYZ1$ = Space$(76)
            string ttxyz1 = new string(' ', 76);

            // CA12$ = FORMATNUM(CALOTE#(II%), "F12.1")
            string ca12 = orden.CanProy.HasValue ? orden.CanProy.Value.ToString("F1") : "0.0";

            // Call REPLA(TTXYZ1$, "     - Cantidad =  " & CA12$, 11, 56)
            ttxyz1 = Repla(ttxyz1, "     - Cantidad =  " + ca12, 11, 56);

            // TTXYZ1$ = TTXYZ1$ + Chr$(10) + Chr$(13)
            ttxyz1 += "\r\n";

            // Separa$ = String$(88, "=") + Chr$(10) + Chr$(13)
            string separa = new string('=', 88) + "\r\n";

            // ANOTA$ = TTXYZ$ + TTXYZ1$ + Separa$
            string anota = ttxyz + ttxyz1 + separa;

            return anota;
        }

        /// <summary>
        /// Función auxiliar que replica la función REPLA de VB6
        /// Reemplaza una parte de una cadena en una posición específica
        /// </summary>
        /// <param name="tx">Cadena original (se modifica)</param>
        /// <param name="sbtx">Texto de reemplazo</param>
        /// <param name="pini">Posición inicial (base 1)</param>
        /// <param name="lon">Longitud máxima</param>
        /// <returns>Cadena con el reemplazo realizado</returns>
        private string Repla(string tx, string sbtx, int pini, int lon)
        {
            // ST$ = SBTX$
            string st = sbtx;

            // PX% = PINI%: If PX% < 1 Then PX% = 1
            int px = pini;
            if (px < 1) px = 1;

            // Lx% = LON%: If Lx% = 0 Then Lx% = Len(ST$)
            int lx = lon;
            if (lx == 0) lx = st.Length;

            // While Len(TX$) < PX% + Lx% - 1
            //     TX$ = TX$ + " "
            // Wend
            while (tx.Length < px + lx - 1)
            {
                tx += " ";
            }

            // ST$ = Left$(ST$ + Space$(Lx%), Lx%)
            st = (st + new string(' ', lx)).Substring(0, lx);

            // If Lx% > 0 Then
            //   Mid$(TX$, PX%, Lx%) = ST$
            // End If
            if (lx > 0)
            {
                // En C# no podemos modificar directamente una subcadena como en VB6
                // así que creamos un array de caracteres y lo modificamos
                char[] chars = tx.ToCharArray();
                for (int i = 0; i < lx; i++)
                {
                    if (px - 1 + i < chars.Length) // Ajustamos a base 0 para C#
                    {
                        chars[px - 1 + i] = st[i];
                    }
                }
                tx = new string(chars);
            }

            return tx;
        }

        /// <summary>
        /// Procesa el cierre de la orden de fabricación aplicando la lógica del código VB6 original
        /// Actualiza cantidades, estados y fechas según las reglas de negocio
        /// </summary>
        /// <param name="orden">Orden de fabricación a procesar</param>
        /// <param name="datosCierre">Datos del cierre con cantidades y configuraciones</param>
        private async Task ProcesarCierreOrdenFabricacion(OrdeFabr orden, DatosCierreAvanzadoDTO datosCierre)
        {
            // Aplicar lógica del código VB6 original
            decimal canti = datosCierre.CantidadAprobada;
            decimal cantiApro = 0;
            decimal scrap = datosCierre.CantidadRechazada;

            // If ModoCierreCalidad% = 1 Then
            //     CANTIAPRO = CANTI
            //     CANTI = 0
            //     Call MOVISTOCK_APROF(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, "PESOS", NC%, DEPX%, CANTI, DP2%, CLA0$)
            // End If
            if (datosCierre.ModoCierreCalidad == 1)
            {
                cantiApro = canti;
                canti = 0;

                // Procesar movimiento de stock para cierre con control de calidad
                // Equivalente a: Call MOVISTOCK_APROF(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, "PESOS", NC%, DEPX%, CANTI, DP2%, CLA0$)
                bool resultadoMovimiento = await MoviStockAprof(
                    fecha: datosCierre.FechaCierre,
                    codigoInterno: orden.Cod_Inte ?? 0,
                    lote: orden.IdenLote ?? "",
                    codigoMovimiento: "OF", // Código de movimiento para orden de fabricación
                    dopri: orden.IdSubOr ?? "",
                    dosec: orden.Referen ?? "",
                    referencia: $"Cierre OF {orden.IdSubOr}",
                    valorUnitario: 0, // Se puede parametrizar si es necesario
                    numeroCliente: orden.Nume_Cli ?? 0,
                    deposito: datosCierre.DepositoEntrada ?? 0, // Usar DepositoEntrada como destino
                    cantidad: cantiApro, // Usar la cantidad aprobada en modo calidad
                    depositoConsumo: datosCierre.DepositoSalida ?? 0, // Usar DepositoSalida como consumo
                    idTexto: "" // Se generará automáticamente
                );

                if (!resultadoMovimiento)
                {
                    _logger.LogError($"Error al procesar movimiento de stock para orden {orden.IdSubOr} en modo calidad");
                    throw new Exception("Error al procesar el movimiento de stock en modo control de calidad");
                }

                _logger.LogInformation($"Movimiento de stock procesado exitosamente para orden {orden.IdSubOr} en modo calidad. Cantidad: {cantiApro}");
            }
            else
            {
                // Procesar movimiento de stock estándar (sin control de calidad)
                // Equivalente a: Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, DEPX%, CANTI, , , , , , , TIPOTRAN%)
                var movimientoStock = new MoviStockDTO
                {
                    Fecha = datosCierre.FechaCierre,
                    CodigoInterno = orden.Cod_Inte ?? 0,
                    Lote = orden.IdenLote ?? "",
                    CodigoMovimiento = "OF", // Código de movimiento para orden de fabricación
                    DocumentoPrimarioLargo = orden.IdSubOr ?? "",
                    DocumentoPrimarioCorto = orden.IdSubOr ?? "",
                    DocumentoSecundarioLargo = orden.Referen ?? "",
                    DocumentoSecundarioCorto = orden.Referen ?? "",
                    Referencia = $"Cierre OF {orden.IdSubOr}",
                    NumeroOrdenItem = 0,
                    ValorUnitario = 0, // Se puede parametrizar si es necesario
                    Moneda = 0, // Moneda por defecto
                    NumeroCliente = orden.Nume_Cli ?? 0,
                    Deposito = datosCierre.DepositoEntrada ?? 0,
                    Cantidad = canti, // Usar la cantidad estándar en modo sin calidad
                    TipoTransferencia = null // Se puede parametrizar si es necesario
                };
                //llama al servicio movistok
                var resultadoMovimiento = await _moviStockService.RegistrarMovimientoAsync(movimientoStock, User);
                if (resultadoMovimiento is BadRequestObjectResult)
                {
                    _logger.LogError($"Error al procesar movimiento de stock para orden {orden.IdSubOr} en modo estándar");
                    throw new Exception("Error al procesar el movimiento de stock en modo estándar");
                }

                _logger.LogInformation($"Movimiento de stock procesado exitosamente para orden {orden.IdSubOr} en modo estándar. Cantidad: {canti}");
            }

            // Procesar consumo de componentes
            await ProcesarConsumoComponentes(orden, datosCierre);

            // Actualizar campos de la orden según la lógica VB6
            decimal canFinProceso = orden.Canti_Fin_Proceso ?? 0;
            decimal aprobadoTotal = (orden.CanApro ?? 0) + canti;

            // !CanApro = APROBADO_TOTAL
            orden.CanApro = aprobadoTotal;

            // !CanRech = XVALO(!CanRech) + Scrap
            orden.CanRech = (orden.CanRech ?? 0) + scrap;

            // If APROBADO_TOTAL > CANFINPROCESO Then
            //     !canti_fin_proceso = APROBADO_TOTAL
            //     !Status_Proceso = 2
            // End If
            if (aprobadoTotal > canFinProceso)
            {
                orden.Canti_Fin_Proceso = aprobadoTotal;
                orden.Status_Proceso = 2;
            }

            // If CIERRA% = 1 Then
            //     !StatuOrf = 3
            // End If
            if (datosCierre.CierraCompleto == 1)
            {
                orden.StatuOrf = 3;
            }

            // !Fechcierr = CVDAT(FECHANUM(FECHACIOF$))
            orden.FechCierr = datosCierre.FechaCierre;

            // If CIERRA% = 1 Then 'SI SE ESTA CERRANDO LA ORDEN DE FABRICACION ES DECIR QUE LA MISMA NO VA A QUEDAR ABIERTA Y SU STATUORF = 3 SE DEBERA BORRAR REGISTROS DE LA TABLA OPERFAB Y RESERVA
            if (datosCierre.CierraCompleto == 1)
            {
                await EliminarOperacionesYReservas(orden.IdSubOr ?? "");
            }

            // No necesitamos marcar la entidad como modificada aquí
            // La actualización se hará con ExecuteUpdateAsync en el método principal
            
            // No necesitamos return Task.CompletedTask ya que el método es async Task
        }

        /// <summary>
        /// Elimina las operaciones programadas y reservas de componentes cuando se cierra completamente una orden de fabricación
        /// Equivalente a la lógica VB6: DELETE FROM OperFab WHERE IdSubOr = NORFA$ y DELETE FROM Reserva WHERE IdSubOr = NORFA$
        /// </summary>
        /// <param name="idSubOr">ID de la suborden de fabricación</param>
        private async Task EliminarOperacionesYReservas(string idSubOr)
        {
            try
            {
                if (string.IsNullOrEmpty(idSubOr))
                {
                    _logger.LogWarning("No se puede eliminar operaciones y reservas: IdSubOr está vacío");
                    return;
                }

                // Cancela Operaciones Programadas
                // SQLX$ = "DELETE FROM OperFab WHERE IdSubOr = '" & NORFA$ & "'"
                var operacionesEliminadas = await _context.Database.ExecuteSqlRawAsync(
                    "DELETE FROM OperFab WHERE IdSubOr = {0}", idSubOr);

                _logger.LogInformation($"Eliminadas {operacionesEliminadas} operaciones programadas para orden {idSubOr}");

                // Cancela Reservas de Componentes  
                // SQLX$ = "DELETE FROM Reserva WHERE IdSubOr = '" & NORFA$ & "'"
                var reservasEliminadas = await _context.Database.ExecuteSqlRawAsync(
                    "DELETE FROM Reserva WHERE IdSubOr = {0}", idSubOr);

                _logger.LogInformation($"Eliminadas {reservasEliminadas} reservas de componentes para orden {idSubOr}");

                _logger.LogInformation($"Eliminación completada para orden {idSubOr}: {operacionesEliminadas} operaciones y {reservasEliminadas} reservas");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al eliminar operaciones y reservas para orden {idSubOr}");
                throw;
            }
        }

        /// <summary>
        /// Genera las anotaciones del cierre replicando la lógica del código VB6 original
        /// Crea las anotaciones con el formato específico usando las funciones REPLA
        /// </summary>
        /// <param name="orden">Orden de fabricación</param>
        /// <param name="datosCierre">Datos del cierre</param>
        private async Task GenerarAnotacionesCierre(OrdeFabr orden, DatosCierreAvanzadoDTO datosCierre)
        {
            // Obtener información del usuario
            string nombreUsuario = ObtenerNombreUsuario();
            int? idFlexoft = _userContextService.GetCurrentUserId();
            string idUsuario = idFlexoft.HasValue ? idFlexoft.Value.ToString() : "";
            string userTer = $"{idUsuario} - {nombreUsuario}";

            // Calcular la cantidad final según el modo de cierre
            decimal cantiFinal = datosCierre.CantidadAprobada;
            if (datosCierre.ModoCierreCalidad == 1)
            {
                cantiFinal = datosCierre.CantidadAprobada; // En modo calidad se usa la cantidad aprobada
            }

            // Generar las anotaciones siguiendo el formato VB6 original
            string anotacionesCierre = GenerarAnotacionesCierreFormateadas(
                orden.IdSubOr ?? string.Empty, 
                userTer, 
                cantiFinal, 
                datosCierre.CantidadRechazada
            );

            // Agregar las anotaciones al campo AnotaOrf
            if (string.IsNullOrEmpty(orden.AnotaOrf))
            {
                orden.AnotaOrf = anotacionesCierre;
            }
            else
            {
                orden.AnotaOrf += anotacionesCierre;
            }
        }

        /// <summary>
        /// Genera las anotaciones formateadas del cierre siguiendo el patrón del código VB6
        /// Replica exactamente la lógica de AGREGA_ANOTAORFAB
        /// </summary>
        /// <param name="norfa">Número de la orden de fabricación</param>
        /// <param name="userTer">Usuario que realiza el cierre</param>
        /// <param name="cantidadAprobada">Cantidad aprobada en el cierre</param>
        /// <param name="cantidadScrap">Cantidad de scrap/rechazada</param>
        /// <returns>Texto formateado de las anotaciones</returns>
        private string GenerarAnotacionesCierreFormateadas(string norfa, string userTer, decimal cantidadAprobada, decimal cantidadScrap)
        {
            string anotaciones = "";

            // TTXYZ$ = Space$(76)
            // HOII% = HOY(HOS$)
            // Call REPLA(TTXYZ$, HOS$, 1, 8)
            // Call REPLA(TTXYZ$, "Cierre de O/Fabricación", 11, 56)
            // Call REPLA(TTXYZ$, USERTER$, 69, 24)
            // Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
            string ttxyz = new string(' ', 76);
            string hos = DateTime.Now.ToString("dd/MM/yy");
            ttxyz = Repla(ttxyz, hos, 1, 8);
            ttxyz = Repla(ttxyz, "Cierre de O/Fabricación", 11, 56);
            ttxyz = Repla(ttxyz, userTer, 69, 24);
            anotaciones += ttxyz + "\r\n";

            // TTXYZ$ = Space$(76)
            // Call REPLA(TTXYZ$, "     - Cantidad Aprob.=  " & FORMATNUM(CANTI, "F12.1"), 11, 56)
            // Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
            ttxyz = new string(' ', 76);
            string cantidadFormateada = cantidadAprobada.ToString("F1");
            ttxyz = Repla(ttxyz, $"     - Cantidad Aprob.=  {cantidadFormateada}", 11, 56);
            anotaciones += ttxyz + "\r\n";

            // If Scrap > 0 Then
            //     TTXYZ$ = Space$(76)
            //     Call REPLA(TTXYZ$, "     - Cantidad Scrap=   " & FORMATNUM(Scrap, "F12.1"), 11, 56)
            //     Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
            // End If
            if (cantidadScrap > 0)
            {
                ttxyz = new string(' ', 76);
                string scrapFormateado = cantidadScrap.ToString("F1");
                ttxyz = Repla(ttxyz, $"     - Cantidad Scrap=   {scrapFormateado}", 11, 56);
                anotaciones += ttxyz + "\r\n";
            }

            // Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
            anotaciones += new string('=', 88) + "\r\n";

            return anotaciones;
        }

        /// <summary>
        /// Registra una nueva anotación en una orden de fabricación existente
        /// </summary>
        /// <param name="id">ID de la orden de fabricación</param>
        /// <param name="textoAnotacion">Texto de la anotación a registrar</param>
        /// <returns>Orden de fabricación actualizada</returns>
        [HttpPost("{id:int}/anotacion")]
        public async Task<IActionResult> RegistrarAnotacion(int id, [FromBody] string textoAnotacion)
        {
            try
            {
                // Verificamos que el ID sea válido y que la orden exista
                var ordenExistente = await _context.Ordefabrs.AsNoTracking().FirstOrDefaultAsync(x => x.Id == id);
                if (ordenExistente == null)
                {
                    return NotFound($"No se encontró la orden de fabricación con ID {id}");
                }

                if (string.IsNullOrWhiteSpace(textoAnotacion))
                {
                    return BadRequest("El texto de la anotación no puede estar vacío");
                }

                // Obtenemos el nombre del usuario actual
                string nombreUsuario = ObtenerNombreUsuario();

                // Formateamos la anotación
                string nuevaAnotacion = FormatearNuevaAnotacion(textoAnotacion, nombreUsuario);

                // Iniciamos una transacción para asegurar la integridad de los datos
                await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);
                try
                {
                    // Agregamos la nueva anotación al campo AnotaOrf
                    if (string.IsNullOrEmpty(ordenExistente.AnotaOrf))
                    {
                        // Si no hay anotaciones previas, inicializamos con la nueva
                        ordenExistente.AnotaOrf = nuevaAnotacion;
                    }
                    else
                    {
                        // Si ya hay anotaciones, agregamos la nueva al inicio
                        ordenExistente.AnotaOrf = nuevaAnotacion + ordenExistente.AnotaOrf;
                    }

                    // Actualizamos la orden de fabricación
                    _context.Update(ordenExistente);
                    await _context.SaveChangesAsync();

                    // Confirmamos la transacción
                    await transaction.CommitAsync();

                    return Ok(ordenExistente);
                }
                catch (Exception ex)
                {
                    // Si hay algún error, revertimos la transacción
                    await transaction.RollbackAsync();
                    throw new Exception($"Error al registrar la anotación: {ex.Message}", ex);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar la anotación en la orden de fabricación");
                return StatusCode(500, $"Error al registrar la anotación: {ex.Message}");
            }
        }

        /// <summary>
        /// Procesa el consumo de componentes durante el cierre de orden de fabricación
        /// Replica la lógica del bucle VB6 que recorre los componentes y registra su consumo
        /// </summary>
        /// <param name="orden">Orden de fabricación</param>
        /// <param name="datosCierre">Datos del cierre con los componentes a consumir</param>
        private async Task ProcesarConsumoComponentes(OrdeFabr orden, DatosCierreAvanzadoDTO datosCierre)
        {
            try
            {
                if (datosCierre.Componentes == null || !datosCierre.Componentes.Any())
                {
                    _logger.LogInformation($"No hay componentes para procesar en la orden {orden.IdSubOr}");
                    return;
                }

                int numeroItem = 0;
                string norfa = orden.IdSubOr ?? "";

                // For U& = 1 To FIMATER& (recorrer todos los componentes)
                foreach (var componente in datosCierre.Componentes)
                {
                    // CI2% = CVI(Left$(ZZ$, 2)) 'codigo interno
                    int codigoInterno = componente.CodIElem;
                    
                    // CAN = (-1) * CVS(Mid$(ZZ$, 5, 4)) 'cantidad a consumir pasar a negativo
                    decimal cantidadConsumo = (-1) * componente.CantidadAUtilizar;
                    
                    // USOUNI# = CVD(Mid$(ZZ$, 9, 8))' uso unitario
                    decimal usoUnitario = componente.UsoBruto;

                    // Verificar si el tipo de material es "GK" para ajustar cantidad
                    var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == codigoInterno);
                    if (master != null && master.Timate == 1) // Asumiendo que GK corresponde a Timate = 1
                    {
                        cantidadConsumo = cantidadConsumo / 1000;
                    }

                    // If Abs(CAN) >= 0.05 Then (solo procesar si la cantidad es significativa)
                    if (Math.Abs(cantidadConsumo) >= 0.05m)
                    {
                        // Obtener el lote reservado
                        // CONDIX$ = " IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2% & ""
                        string lote = await ObtenerLoteReservado(norfa, codigoInterno);

                        // CMOV$ = "CF" (código de movimiento para consumo de fabricación)
                        string codigoMovimiento = "CF";

                        // DOPRI$ = NORFA$
                        string dopri = norfa;

                        // DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
                        string dosec = "PF-" + norfa.Substring(Math.Min(3, norfa.Length - 1));

                        // REFE$ = "CONSUMO " + CODICONJ$
                        string referencia = $"CONSUMO {orden.Cod_Exte}";

                        // Determinar el depósito
                        // DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                        int deposito = datosCierre.DepositoSalida ?? 0;
                        if (deposito == 99)
                        {
                            deposito = await ObtenerDepositoPorTipoMaterial(codigoInterno);
                        }

                        numeroItem++;

                        // Procesar según el modo de cierre
                        if (datosCierre.ModoCierreCalidad == 1)
                        {
                            // Modo con control de calidad - usar MoviStockAprof
                            bool resultado = await MoviStockAprof(
                                fecha: datosCierre.FechaCierre,
                                codigoInterno: codigoInterno,
                                lote: lote,
                                codigoMovimiento: codigoMovimiento,
                                dopri: dopri,
                                dosec: dosec,
                                referencia: referencia,
                                valorUnitario: 0,
                                numeroCliente: orden.Nume_Cli ?? 0,
                                deposito: deposito,
                                cantidad: cantidadConsumo,
                                depositoConsumo: datosCierre.DepositoSalida ?? 0,
                                idTexto: ""
                            );

                            if (!resultado)
                            {
                                _logger.LogError($"Error al procesar consumo del componente {componente.CodXElem} en modo calidad");
                                throw new Exception($"Error al procesar el consumo del componente {componente.CodXElem} en modo control de calidad");
                            }
                        }
                        else
                        {
                            // Modo sin control de calidad - usar MoviStockService
                            var movimientoConsumo = new MoviStockDTO
                            {
                                Fecha = datosCierre.FechaCierre,
                                CodigoInterno = codigoInterno,
                                Lote = lote,
                                CodigoMovimiento = codigoMovimiento,
                                DocumentoPrimarioLargo = dopri,
                                DocumentoPrimarioCorto = dopri,
                                DocumentoSecundarioLargo = dosec,
                                DocumentoSecundarioCorto = dosec,
                                Referencia = referencia,
                                NumeroOrdenItem = numeroItem,
                                ValorUnitario = 0,
                                Moneda = 0,
                                NumeroCliente = orden.Nume_Cli ?? 0,
                                Deposito = deposito,
                                Cantidad = cantidadConsumo,
                                TipoTransferencia = null
                            };

                            var resultadoConsumo = await _moviStockService.RegistrarMovimientoAsync(movimientoConsumo, User);
                            if (resultadoConsumo is BadRequestObjectResult)
                            {
                                _logger.LogError($"Error al procesar consumo del componente {componente.CodXElem} en modo estándar");
                                throw new Exception($"Error al procesar el consumo del componente {componente.CodXElem} en modo estándar");
                            }
                        }

                        // Actualizar las reservas (equivalente a la lógica de actualización de CANTCONS)
                        await ActualizarReservasComponente(norfa, codigoInterno, Math.Abs(cantidadConsumo));

                        _logger.LogInformation($"Consumo procesado exitosamente para componente {componente.CodXElem}. Cantidad: {cantidadConsumo}");
                    }
                }

                _logger.LogInformation($"Procesamiento de consumo de componentes completado para orden {orden.IdSubOr}. Total componentes: {numeroItem}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al procesar consumo de componentes para orden {orden.IdSubOr}");
                throw;
            }
        }

        /// <summary>
        /// Obtiene el lote reservado para un componente específico de una orden
        /// Equivalente a: LOTE$ = VALOBADA("Reserva", "IdenLote", CONDIX$, "NOMESS")
        /// </summary>
        /// <param name="idSubOr">ID de la suborden</param>
        /// <param name="codigoInterno">Código interno del componente</param>
        /// <returns>Lote reservado o cadena vacía si no se encuentra</returns>
        private async Task<string> ObtenerLoteReservado(string idSubOr, int codigoInterno)
        {
            try
            {
                var reserva = await _context.Reservas
                    .Where(r => r.IdSubOr == idSubOr && r.Cod_Inte == codigoInterno)
                    .FirstOrDefaultAsync();

                return reserva?.IdenLote ?? "";
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"Error al obtener lote reservado para componente {codigoInterno} en orden {idSubOr}");
                return "";
            }
        }

        /// <summary>
        /// Obtiene el depósito según el tipo de material del componente
        /// Equivalente a: DEPX% = DEPOPRE%(CI2%) que consulta TATIMAT.DEPOSITO según MASTER.TIMATE
        /// </summary>
        /// <param name="codigoInterno">Código interno del componente</param>
        /// <returns>Depósito correspondiente al tipo de material</returns>
        private async Task<int> ObtenerDepositoPorTipoMaterial(int codigoInterno)
        {
            try
            {
                // Obtener el tipo de material del componente
                var master = await _context.masters
                    .Where(m => m.Codint == codigoInterno)
                    .FirstOrDefaultAsync();

                if (master?.Timate == null)
                {
                    return 0; // Depósito por defecto si no se encuentra el tipo
                }

                // Buscar el depósito en la tabla de tipos de material
                var tipoMaterial = await _context.Tatimat
                    .Where(t => t.CodigoTatimat == master.Timate)
                    .FirstOrDefaultAsync();

                return tipoMaterial?.Deposito ?? 0;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"Error al obtener depósito por tipo de material para componente {codigoInterno}");
                return 0;
            }
        }

        /// <summary>
        /// Actualiza las reservas del componente incrementando la cantidad consumida
        /// Equivalente a la lógica VB6 que actualiza CANTCONS en la tabla Reserva
        /// </summary>
        /// <param name="idSubOr">ID de la suborden</param>
        /// <param name="codigoInterno">Código interno del componente</param>
        /// <param name="cantidadConsumida">Cantidad consumida a agregar</param>
        private async Task ActualizarReservasComponente(string idSubOr, int codigoInterno, decimal cantidadConsumida)
        {
            try
            {
                var reservas = await _context.Reservas
                    .Where(r => r.IdSubOr == idSubOr && r.Cod_Inte == codigoInterno)
                    .ToListAsync();

                decimal salcon = cantidadConsumida; // SALCON# = cantidad pendiente de consumir

                foreach (var reserva in reservas)
                {
                    if (salcon <= 0) break;

                    decimal cantidadReservada = reserva.CantRes ?? 0;
                    decimal cantidadConsumidaAnterior = reserva.CantCons ?? 0;
                    decimal disponible = cantidadReservada - cantidadConsumidaAnterior;

                    if (disponible >= salcon)
                    {
                        // La reserva actual puede cubrir todo lo pendiente
                        reserva.CantCons = cantidadConsumidaAnterior + salcon;
                        salcon = 0;
                    }
                    else
                    {
                        // La reserva actual se consume completamente
                        salcon = salcon - disponible;
                        reserva.CantCons = cantidadReservada;
                    }

                    _context.Reservas.Update(reserva);
                }

                if (salcon > 0.05m)
                {
                    _logger.LogWarning($"Quedó cantidad sin asignar a reservas: {salcon} para componente {codigoInterno} en orden {idSubOr}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al actualizar reservas para componente {codigoInterno} en orden {idSubOr}");
                throw;
            }
        }

        /// <summary>
        /// Formatea una nueva anotación para agregarla al campo AnotaOrf
        /// </summary>
        /// <param name="textoAnotacion">Texto de la anotación</param>
        /// <param name="nombreUsuario">Nombre del usuario que registra la anotación</param>
        /// <returns>Texto formateado para el campo AnotaOrf</returns>
        private string FormatearNuevaAnotacion(string textoAnotacion, string nombreUsuario)
        {
            // Creamos una línea de anotación similar a la generada en GenerarTextoAnotaOrf
            string ttxyz = new string(' ', 76); // Space$(76)
            string hos = DateTime.Now.ToString("dd/MM/yy"); // Equivalente a HOY(HOS$)

            // Formateamos la primera línea con fecha, detalle y usuario
            ttxyz = Repla(ttxyz, hos, 1, 8);
            ttxyz = Repla(ttxyz, textoAnotacion, 11, 56);
            ttxyz = Repla(ttxyz, nombreUsuario, 69, 24);

            // Agregamos salto de línea
            ttxyz += "\r\n";

            // Agregamos una línea de separación
            string separa = new string('=', 88) + "\r\n";

            // Devolvemos el texto formateado
            return ttxyz + separa;
        }

        /// <summary>
        /// Genera reservas para los componentes de la fórmula de un producto a fabricar
        /// </summary>
        /// <param name="codigoInterno">Código interno del producto a fabricar</param>
        /// <param name="numeroOrdenFabricacion">Número de la orden de fabricación</param>
        /// <param name="cantidadAFabricar">Cantidad a fabricar</param>
        /// <param name="codigoEmpresa">Código de la empresa</param>
        /// <returns>Task</returns>
        private async Task GenerarReservasDeFormula(int codigoInterno, string numeroOrdenFabricacion, decimal? cantidadAFabricar, short codigoEmpresa)
        {
            // Verificar parámetros
            if (string.IsNullOrEmpty(numeroOrdenFabricacion) || !cantidadAFabricar.HasValue || cantidadAFabricar.Value <= 0)
            {
                _logger.LogWarning("No se generaron reservas: parámetros inválidos");
                return;
            }

            try
            {
                // Eliminar reservas existentes para esta orden (equivalente a DELETE FROM Reserva WHERE IdSubOr = NORFA$)
                var reservasExistentes = await _context.Reservas.Where(r => r.IdSubOr == numeroOrdenFabricacion).ToListAsync();
                if (reservasExistentes.Any())
                {
                    _context.Reservas.RemoveRange(reservasExistentes);
                    await _context.SaveChangesAsync();
                }

                // Obtener los componentes de la fórmula del producto (equivalente a SELECT * FROM FORMULAS WHERE CODICONJ = CI1%)
                var componentes = await _context.Formulas
                    .Where(f => f.CodIConj == codigoInterno)
                    .ToListAsync();

                if (!componentes.Any())
                {
                    _logger.LogWarning($"No se encontraron componentes en la fórmula para el producto {codigoInterno}");
                    return;
                }

                // Contador para el número de ítem (equivalente a NORIT%)
                short numeroItem = 0;

                // Procesar cada componente de la fórmula
                foreach (var componente in componentes)
                {
                    // Verificar que el componente tenga código interno válido
                    if (!componente.CodiElem.HasValue || componente.CodiElem.Value <= 0)
                    {
                        continue;
                    }

                    // Calcular la cantidad necesaria (equivalente a CANTX# * USOII)
                    decimal cantidadReserva = cantidadAFabricar.Value * (componente.UsoBruto ?? 0);

                    // Solo crear reserva si la cantidad es significativa (equivalente a If Abs(CANTX#) >= 0.05 Then)
                    if (Math.Abs(cantidadReserva) >= 0.05m)
                    {
                        // Incrementar el contador de ítems
                        numeroItem++;

                        // Obtener información del componente desde Master
                        var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == componente.CodiElem.Value);
                        string codigoExterno = master?.Codigo ?? string.Empty;
                        string descripcion = master?.Descripcion ?? string.Empty;

                        // Crear la nueva reserva (equivalente a la sección With RESERVVVA)
                        var nuevaReserva = new Reserva
                        {
                            CodiEmpr = codigoEmpresa,
                            Cod_Inte = componente.CodiElem.Value,
                            Cod_Exte = codigoExterno,
                            Descrip = descripcion,
                            IdSubOr = numeroOrdenFabricacion,
                            FechRes = DateTime.Now,
                            CantRes = cantidadReserva,
                            CantCons = 0,
                            Saldo_Entre = 0,
                            IdenLote = " ",
                            Deposito = 0,
                            UsoUnit = componente.UsoBruto,
                            NuOrItem = numeroItem,
                            Nume_Solint = 0
                        };

                        // Agregar la reserva a la base de datos
                        _context.Reservas.Add(nuevaReserva);
                    }
                }

                // Guardar todos los cambios
                await _context.SaveChangesAsync();
                _logger.LogInformation($"Se generaron {numeroItem} reservas para la orden {numeroOrdenFabricacion}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al generar reservas para la orden {numeroOrdenFabricacion}");
                throw; // Relanzamos la excepción para que sea manejada por el método que llamó a este
            }
        }

        /// <summary>
        /// Procesa el movimiento de stock para cierre de orden de fabricación con control de calidad
        /// Equivalente a MOVISTOCK_APROF en VB6
        /// </summary>
        /// <param name="fecha">Fecha del movimiento</param>
        /// <param name="codigoInterno">Código interno del producto</param>
        /// <param name="lote">Identificador del lote</param>
        /// <param name="codigoMovimiento">Código del tipo de movimiento</param>
        /// <param name="dopri">Código primario</param>
        /// <param name="dosec">Código secundario</param>
        /// <param name="referencia">Referencia del movimiento</param>
        /// <param name="valorUnitario">Valor unitario</param>
        /// <param name="numeroCliente">Número de cliente</param>
        /// <param name="deposito">Depósito de origen</param>
        /// <param name="cantidad">Cantidad a mover</param>
        /// <param name="depositoConsumo">Depósito de consumo</param>
        /// <param name="idTexto">Identificador de texto opcional</param>
        /// <returns>Resultado de la operación</returns>
        public async Task<bool> MoviStockAprof(DateTime fecha, int codigoInterno, string lote, string codigoMovimiento, 
            string dopri, string dosec, string referencia, decimal valorUnitario, int numeroCliente, 
            int deposito, decimal cantidad, int depositoConsumo, string idTexto = "")
        {
            try
            {
                // Obtener información del usuario actual
                var idFlexoft = _userContextService.GetCurrentUserId();
                if (!idFlexoft.HasValue)
                {
                    _logger.LogError("No se pudo obtener el contexto del usuario para el movimiento de stock");
                    return false;
                }

                // Procesar códigos dopri y dosec (eliminar "-0" como en VB6)
                string dopriProcesado = ProcesarCodigo(dopri);
                string dosecProcesado = ProcesarCodigo(dosec);

                // Ajustar códigos de movimiento especiales
                if (codigoMovimiento == "*R") codigoMovimiento = "RP";
                if (codigoMovimiento == "*D") codigoMovimiento = "DR";

                // Separar cliente y proveedor
                int numeroClienteFinal = numeroCliente >= 0 ? numeroCliente : 0;
                int numeroProveedor = numeroCliente < 0 ? Math.Abs(numeroCliente) : 0;

                // Llamar al método que agrega el registro
                return await AgregarRegistroOrfapen(fecha, codigoInterno, lote, codigoMovimiento, 
                    dopri, dosec, referencia, valorUnitario, 1, numeroClienteFinal, deposito, 
                    cantidad, depositoConsumo, idTexto, "GenCierre");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en MoviStockAprof");
                return false;
            }
        }

        /// <summary>
        /// Agrega un registro de movimiento en la tabla ORFAPEN
        /// Equivalente a AGREGA_REGISTRO_NORFAPEN en VB6
        /// </summary>
        /// <param name="fecha">Fecha del movimiento</param>
        /// <param name="codigoInterno">Código interno del producto</param>
        /// <param name="lote">Identificador del lote</param>
        /// <param name="codigoMovimiento">Código del tipo de movimiento</param>
        /// <param name="dopri">Código primario</param>
        /// <param name="dosec">Código secundario</param>
        /// <param name="referencia">Referencia del movimiento</param>
        /// <param name="valorUnitario">Valor unitario</param>
        /// <param name="moneda">Código de moneda</param>cierre-avanzado
        /// <param name="numeroCliente">Número de cliente</param>
        /// <param name="deposito">Depósito</param>
        /// <param name="cantidad">Cantidad</param>
        /// <param name="depositoConsumo">Depósito de consumo</param>
        /// <param name="idTexto">Identificador de texto</param>
        /// <param name="idSituacion">Identificador de situación</param>
        /// <returns>Resultado de la operación</returns>
        private async Task<bool> AgregarRegistroOrfapen(DateTime fecha, int codigoInterno, string lote, 
            string codigoMovimiento, string dopri, string dosec, string referencia, decimal valorUnitario, 
            int moneda, int numeroCliente, int deposito, decimal cantidad, int depositoConsumo, 
            string idTexto, string idSituacion)
        {
            try
            {
                // Obtener información del usuario actual
                var idFlexoft = _userContextService.GetCurrentUserId();
                if (!idFlexoft.HasValue)
                {
                    _logger.LogError("No se pudo obtener el contexto del usuario para agregar registro ORFAPEN");
                    return false;
                }

                // Generar ID de texto si está vacío (equivalente a RANDSTRING$(-10) en VB6)
                if (string.IsNullOrWhiteSpace(idTexto))
                {
                    idTexto = GenerarIdTextoAleatorio();
                }
                
                // Verificar duplicados para movimientos OF (equivalente a la validación en VB6)
                if (codigoMovimiento.Trim().ToUpper() == "OF")
                {
                    var existeRegistro = await _context.Orfapens
                        .AnyAsync(o => o.IDTEXT == idTexto);
                    
                    if (existeRegistro)
                    {
                        // Generar nuevo ID si ya existe
                        idTexto = GenerarIdTextoAleatorio();
                    }
                }

                // Validar moneda
                if (moneda <= 0) moneda = 1;

                // Procesar códigos para asegurar longitud de 11 caracteres (como en VB6)
                dopri = ProcesarCodigoLongitud(dopri);
                dosec = ProcesarCodigoLongitud(dosec);

                // Obtener código externo del producto
                var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == codigoInterno);
                string codigoExterno = master?.Codigo ?? "";

                // Determinar cantidades de ingreso y salida
                decimal cantidadIngreso = 0;
                decimal cantidadSalida = 0;

                // Ajustar cantidad si es cierre de aprobación (equivalente a la lógica VB6)
                if (idSituacion.Trim().ToUpper() == "APROCIERRE")
                {
                    cantidad = cantidad * -1; // Para generar movimiento de salida y compensar con el cierre original
                }

                if (cantidad > 0)
                {
                    cantidadIngreso = Math.Abs(cantidad);
                    cantidadSalida = 0;
                }
                else
                {
                    cantidadSalida = Math.Abs(cantidad);
                    cantidadIngreso = 0;
                }

                // Crear el nuevo registro ORFAPEN
                var nuevoRegistro = new Orfapen
                {
                    FechMov = fecha,
                    Cod_Inte = (short)codigoInterno,
                    Cod_Exte = codigoExterno,
                    IdenLote = lote,
                    CodiMovi = codigoMovimiento,
                    DoPriCor = dopri,
                    DoPriLar = dopri,
                    DoSecCor = dosec,
                    DoSecLar = dosec,
                    ReferCor = referencia,
                    ValoUnit = valorUnitario,
                    MoneVal = (short)moneda,
                    Nume_Clie = numeroCliente,
                    DepoMovi = deposito,
                    CantIngre = cantidadIngreso,
                    CantSalid = cantidadSalida,
                    FeReMovi = DateTime.Now,
                    NumUsuar = idFlexoft.Value % 100, // Equivalente a USNBR% Mod 100 en VB6
                    STATUS = 0,
                    IDTEXT = idTexto,
                    DepoConsu = (short)depositoConsumo,
                    IDSITUACION = idSituacion
                };

                // Agregar el registro a la base de datos
                _context.Orfapens.Add(nuevoRegistro);
                // NO llamar SaveChangesAsync aquí - se guardará en la transacción principal

                _logger.LogInformation($"Registro ORFAPEN preparado para agregar. Situación: {idSituacion}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al agregar registro ORFAPEN");
                return false;
            }
        }

        /// <summary>
        /// Procesa un código eliminando las secuencias "-0" (equivalente a la lógica VB6)
        /// </summary>
        /// <param name="codigo">Código a procesar</param>
        /// <returns>Código procesado</returns>
        private string ProcesarCodigo(string codigo)
        {
            if (string.IsNullOrEmpty(codigo)) return "";

            string resultado = codigo;
            int posicion;
            
            // Eliminar todas las ocurrencias de "-0" (equivalente al bucle GoTo en VB6)
            while ((posicion = resultado.IndexOf("-0")) > 0)
            {
                resultado = resultado.Substring(0, posicion) + resultado.Substring(posicion + 2);
            }
            
            return resultado;
        }

        /// <summary>
        /// Procesa un código para asegurar que tenga 11 caracteres de longitud
        /// Agrega "0" después del tercer carácter hasta alcanzar la longitud requerida
        /// </summary>
        /// <param name="codigo">Código a procesar</param>
        /// <returns>Código con longitud de 11 caracteres</returns>
        private string ProcesarCodigoLongitud(string codigo)
        {
            if (string.IsNullOrEmpty(codigo)) return new string('0', 11);

            // Equivalente a: While Len(DOPRI$) < 11: DOPRI$ = Mid$(DOPRI$, 1, 3) & "0" & Mid$(DOPRI$, 4): Wend
            while (codigo.Length < 11)
            {
                if (codigo.Length >= 3)
                {
                    codigo = codigo.Substring(0, 3) + "0" + codigo.Substring(3);
                }
                else
                {
                    codigo = codigo + "0";
                }
            }

            return codigo.Substring(0, 11); // Asegurar que no exceda 11 caracteres
        }

        /// <summary>
        /// Genera un identificador de texto aleatorio de 10 caracteres
        /// Equivalente a RANDSTRING$(-10) en VB6
        /// </summary>
        /// <returns>Cadena aleatoria de 10 caracteres</returns>
        private string GenerarIdTextoAleatorio()
        {
            const string caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            var random = new Random();
            var resultado = new char[10];
            
            for (int i = 0; i < 10; i++)
            {
                resultado[i] = caracteres[random.Next(caracteres.Length)];
            }
            
            return new string(resultado);
        }

        /// <summary>
        /// Endpoint para procesar el cierre de orden de fabricación con control de calidad
        /// Se ejecuta cuando ModoCierreCalidad = 1
        /// </summary>
        /// <param name="datosCierre">Datos del cierre incluyendo parámetros del movimiento</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPost("cierre-calidad")]
        public async Task<IActionResult> ProcesarCierreCalidad([FromBody] CierreCalidadDTO datosCierre)
        {
            try
            {
                // Validar datos de entrada
                if (datosCierre == null)
                {
                    return BadRequest("Los datos del cierre son requeridos");
                }

                // Validar que el modo de cierre de calidad esté activado
                if (datosCierre.ModoCierreCalidad != 1)
                {
                    return BadRequest("El modo de cierre de calidad debe estar activado (ModoCierreCalidad = 1)");
                }

                // Procesar el movimiento de stock
                var resultado = await MoviStockAprof(
                    datosCierre.Fecha,
                    datosCierre.CodigoInterno,
                    datosCierre.Lote,
                    datosCierre.CodigoMovimiento,
                    datosCierre.Dopri,
                    datosCierre.Dosec,
                    datosCierre.Referencia,
                    datosCierre.ValorUnitario,
                    datosCierre.NumeroCliente,
                    datosCierre.Deposito,
                    datosCierre.Cantidad,
                    datosCierre.DepositoConsumo,
                    datosCierre.IdTexto
                );

                if (resultado)
                {
                    return Ok("Cierre de calidad procesado exitosamente");
                }
                else
                {
                    return StatusCode(500, "Error al procesar el cierre de calidad");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en ProcesarCierreCalidad");
                return StatusCode(500, "Error interno del servidor al procesar el cierre de calidad");
            }
        }


       
        /// <summary>
        /// Evalúa si el modo de cierre de calidad está activo
        /// Acepta tanto valores numéricos (1) como texto ("SI", "Si", "si")
        /// </summary>
        
        private async Task<int> ObtenerParametroCierreCalidad()
        {
            try
            {
                // Consultar el parámetro CIERCALI de la tabla FLEXCRL
                // Equivalente a la consulta del sistema VB6 original
                var parametro = await _context.Flexcrls
                    .Where(f => f.CLABUS == "" && f.PROBUS == "CIERCALI")
                    .FirstOrDefaultAsync();

                if (parametro != null && !string.IsNullOrEmpty(parametro.VALCONTROL))
                {
                    string valorControl = parametro.VALCONTROL.Trim().ToUpper();

                    _logger.LogInformation($"Parámetro CIERCALI encontrado con valor original: '{parametro.VALCONTROL}'");

                    // Primero intentar convertir directamente a número
                    if (int.TryParse(valorControl, out int valorNumerico))
                    {
                        _logger.LogInformation($"CIERCALI interpretado como número: {valorNumerico}");
                        return valorNumerico > 0 ? 1 : 0; // Cualquier número positivo se considera activado
                    }

                    // Evaluar valores de texto que indican activación (positivos)
                    string[] valoresPositivos = { "SI", "SÍ", "YES", "Y", "TRUE", "VERDADERO", "1", "ON", "ACTIVO", "ACTIVADO", "HABILITADO" };
                    if (valoresPositivos.Contains(valorControl))
                    {
                        _logger.LogInformation($"CIERCALI interpretado como valor positivo: '{valorControl}' -> 1");
                        return 1;
                    }

                    // Evaluar valores de texto que indican desactivación (negativos)
                    string[] valoresNegativos = { "NO", "N", "FALSE", "FALSO", "0", "OFF", "INACTIVO", "DESACTIVADO", "DESHABILITADO" };
                    if (valoresNegativos.Contains(valorControl))
                    {
                        _logger.LogInformation($"CIERCALI interpretado como valor negativo: '{valorControl}' -> 0");
                        return 0;
                    }

                    // Si es una ruta de archivo o cualquier otro texto, considerar como activado si no está vacío
                    if (valorControl.Length > 0)
                    {
                        // Verificar si parece ser una ruta de archivo
                        if (valorControl.Contains("\\") || valorControl.Contains("/") || valorControl.Contains("."))
                        {
                            _logger.LogInformation($"CIERCALI interpretado como ruta de archivo: '{valorControl}' -> 1 (activado)");
                            return 1;
                        }

                        // Para cualquier otro texto no vacío, considerar como activado
                        _logger.LogInformation($"CIERCALI interpretado como texto no vacío: '{valorControl}' -> 1 (activado)");
                        return 1;
                    }
                }

                // Si no se encuentra el parámetro o está vacío, usar valor por defecto (0 = sin control de calidad)
                _logger.LogWarning("Parámetro CIERCALI no encontrado o vacío, usando valor por defecto: 0");
                return 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener el parámetro CIERCALI");
                // En caso de error, usar valor por defecto (0 = sin control de calidad)
                return 0;
            }
        }
    }

}

// ESTO SE DEBE AGREGAR EN EL ENDPOINT DONDE SE DEBE HACER UNA TRANSACCION ******
//int? idFlexoft = _UsuarioFlexoftService.InterObtenerIdFlexoft(User);
//short? codigoEmpresa = _Codiempresa.IObtenerCodiempresa(User);
//        if (!idFlexoft.HasValue)
//        {
//            return BadRequest("(Sesion Caducada - Debe cerrar sesión y loguearse nuevamente.");

//        }
// ESTA ES LA FORMA DE LLAMADA EN .RAZOR PARA QUE MUESTRE EL MENSAJE QUE RETORNA DESDE EL ENDPOINT  ************
//var response = await repository.Post("api/stockTransferencia/transferencia", transferencia);
//    if (response.Error)
//    {
//        var mensaje = await response.HttpResponseMessage.Content.ReadAsStringAsync();
//await sweetAlertService.FireAsync("Error", mensaje, SweetAlertIcon.Error);
//        return;
//    }