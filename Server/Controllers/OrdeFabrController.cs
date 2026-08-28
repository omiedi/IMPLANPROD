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

        [HttpGet("pendientes-por-producto/{codint:int}")]
        public async Task<ActionResult<List<StockFabricacionPendienteDTO>>> GetPendientesPorProducto(int codint)
        {
            try
            {
                var result = await _context.Ordefabrs
                    .AsNoTracking()
                    .Where(x => x.Cod_Inte == codint
                                && (x.StatuOrf ?? 0) < 3
                                && (x.CanApro ?? 0m) < ((x.CanProy ?? 0m) - 0.005m))
                    .OrderBy(x => x.NumeOrFa)
                    .Select(x => new StockFabricacionPendienteDTO
                    {
                        NumeOrfa = x.NumeOrFa ?? 0,
                        Referencia = x.Referen ?? string.Empty,
                        CantPendiente = ((x.CanProy ?? 0m) - (x.CanApro ?? 0m)) < 0m ? 0m : ((x.CanProy ?? 0m) - (x.CanApro ?? 0m)),
                        FechaEntregaSolicitada = x.FechEnSol,
                        FechaProyectadaEntrega = x.FechProyEntre,
                        FechaRecalculo = x.FechRecal
                    })
                    .Where(x => x.CantPendiente > 0m)
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener órdenes de fabricación pendientes por producto {Codint}", codint);
                return StatusCode(500, "Error interno del servidor");
            }
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
                    // Mostrar todas > 0
                    queryable = queryable.Where(x => x.StatuOrf > 0);
                }
                else if (statuOrf == 1)
                {
                    // Mostrar abiertas (StatuOrf = 0 o 1)
                    queryable = queryable.Where(x => x.StatuOrf == 0 || x.StatuOrf == 1);
                }
                else if (statuOrf == 2)
                {
                    // Mostrar abiertas (StatuOrf < 3)
                    queryable = queryable.Where(x => x.StatuOrf < 3);
                }
                else if (statuOrf == 3)
                {
                    // Mostrar cerradas (StatuOrf = 3)
                    queryable = queryable.Where(x => x.StatuOrf == 3);
                }
                else if (statuOrf == 9)
                {
                    // Mostrar anuladas (StatuOrf = 9)
                    queryable = queryable.Where(x => x.StatuOrf == 9);
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

                // Filtro por texto (solo por IdSubOr)
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    queryable = queryable.Where(x =>
                        x.IdSubOr != null && x.IdSubOr.Contains(pagination.Filter)
                    );
                }

                // Filtro por prefijo de suborden (ej: "DL-" para limpieza)
                if (!string.IsNullOrWhiteSpace(idSubOrPrefix))
                {
                    queryable = queryable.Where(x => x.IdSubOr != null && x.IdSubOr.StartsWith(idSubOrPrefix));
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

                // Calcular valores desde ORFAPEN para todas las órdenes en una sola consulta (evitar N+1)
                var idSubOrs = result.Where(x => !string.IsNullOrWhiteSpace(x.IdSubOr)).Select(x => x.IdSubOr).Distinct().ToList();
                
                if (idSubOrs.Any())
                {
                    var orfapenData = await _context.Orfapens
                        .AsNoTracking()
                        .Where(o => idSubOrs.Contains(o.DoPriCor) && o.CodiMovi == "OF")
                        .GroupBy(o => o.DoPriCor)
                        .Select(g => new
                        {
                            DoPriCor = g.Key,
                            CantidadCerrada = g.Sum(o => o.CantIngre ?? 0m),
                            CantidadAprobada = g.Sum(o => o.CantSalid ?? 0m)
                        })
                        .ToListAsync();

                    var orfapenDict = orfapenData.ToDictionary(x => x.DoPriCor);

                    foreach (var dto in result)
                    {
                        if (string.IsNullOrWhiteSpace(dto.IdSubOr) || !orfapenDict.TryGetValue(dto.IdSubOr, out var data))
                            continue;

                        dto.CantidadCerrada = data.CantidadCerrada;
                        dto.CantidadAprobadaOrfapen = data.CantidadAprobada;
                        dto.PendienteAprobacion = data.CantidadCerrada - data.CantidadAprobada;
                    }
                }

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

                // Filtro por texto (solo por IdSubOr)
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    queryable = queryable.Where(x =>
                        x.IdSubOr != null && x.IdSubOr.Contains(pagination.Filter)
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

        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var ordenFabr = await _context.Ordefabrs.FirstOrDefaultAsync(x => x.Id == id);
            if (ordenFabr == null)
            {
                return NotFound();
            }

            // Proyectar a DTO para incluir cálculos de ORFAPEN
            var dto = new OrdefabrIndexDTO
            {
                Id = ordenFabr.Id,
                FechGen = ordenFabr.FechGen,
                IdSubOr = ordenFabr.IdSubOr,
                CanProy = ordenFabr.CanProy,
                CanApro = ordenFabr.CanApro,
                Id_Orfa = ordenFabr.Id_Orfa,
                StatuOrf = ordenFabr.StatuOrf,
                Cod_Exte = ordenFabr.Cod_Exte,
                Unidad = ordenFabr.Unidad,
                FechEnSol = ordenFabr.FechEnSol,
                FechLan = ordenFabr.FechLan,
                FechCierr = ordenFabr.FechCierr,
                Descrip = ordenFabr.Descrip,
                Cod_Inte = ordenFabr.Cod_Inte,
                AnotaOrf = ordenFabr.AnotaOrf
            };

            // Calcular valores desde ORFAPEN para esta orden
            if (!string.IsNullOrWhiteSpace(dto.IdSubOr))
            {
                var orfapenData = await _context.Orfapens
                    .AsNoTracking()
                    .Where(o => o.DoPriCor == dto.IdSubOr
                               && o.CodiMovi == "OF")
                    .GroupBy(o => o.DoPriCor)
                    .Select(g => new
                    {
                        CantidadCerrada = g.Sum(o => o.CantIngre ?? 0m),
                        CantidadAprobada = g.Sum(o => o.CantSalid ?? 0m)
                    })
                    .FirstOrDefaultAsync();

                if (orfapenData != null)
                {
                    dto.CantidadCerrada = orfapenData.CantidadCerrada;
                    dto.CantidadAprobadaOrfapen = orfapenData.CantidadAprobada;
                    dto.PendienteAprobacion = orfapenData.CantidadCerrada - orfapenData.CantidadAprobada;
                }
            }

            return Ok(dto);
        }

        // Endpoint usado por la UI de Trazabilidad para completar el campo "N° Serie".
        // En VB6, NORFA equivale a IdSubOr.
        // Regla única: buscar por IdSubOr (igual que en VB6).
        [HttpGet("numero-serie/{idSubOr}")]
        public async Task<IActionResult> GetNumeroSeriePorIdSubOr(string idSubOr)
        {
            try
            {
                // Si no viene IdSubOr, no hay forma de buscar la serie principal.
                if (string.IsNullOrWhiteSpace(idSubOr))
                {
                    return BadRequest("Debe indicar el valor de IdSubOr.");
                }

                // Normalizamos el valor recibido para evitar problemas de espacios.
                // NORFA (VB6) = IdSubOr (API actual).
                var norfaBuscada = idSubOr.Trim();

                // Consulta principal estilo VB6:
                // SELECT O.IDSUBOR, C.NumeroSerie
                // FROM CONTROSERIE C INNER JOIN ORDEFABR O ON O.IDSUBOR = C.IDSUBOR
                // WHERE O.IDSUBOR = @NORFA
                var numeroSerie = await (
                    from o in _context.Ordefabrs.AsNoTracking()
                    join c in _context.Controseries.AsNoTracking()
                        on o.IdSubOr equals c.IdSubOr
                    where o.IdSubOr == norfaBuscada
                          && !string.IsNullOrWhiteSpace(c.NumeroSerie)
                    orderby c.FeReMovi descending, c.Id descending
                    select c.NumeroSerie
                ).FirstOrDefaultAsync();

                // Si no hay coincidencias, devolvemos vacío.
                return Ok(numeroSerie?.Trim() ?? string.Empty);
            }
            catch (Exception ex)
            {
                // Log técnico para diagnóstico cuando la UI no puede completar la serie.
                _logger.LogError(ex, "Error al obtener número de serie desde CONTROSERIE para IdSubOr {IdSubOr}", idSubOr);
                return StatusCode(500, $"Error al obtener número de serie: {ex.Message}");
            }
        }

        [HttpGet("depositos-posibles-con-saldo")]
        public async Task<IActionResult> GetDepositosPosiblesConSaldo()
        {
            try
            {
                const int depositoMinimo = 60;
                const int depositoMaximo = 200;

                var depositosSerie = await _context.Tadeposi
                    .AsNoTracking()
                    .Where(x => x.CodigoDepo >= depositoMinimo && x.CodigoDepo <= depositoMaximo)
                    .Select(x => new
                    {
                        x.CodigoDepo,
                        x.Descripcion
                    })
                    .ToListAsync();

                if (depositosSerie.Count == 0)
                {
                    return Ok(new List<DepositosPosiblesConSaldoDTO>());
                }

                var ofPorDeposito = depositosSerie
                    .Select(x => new
                    {
                        x.CodigoDepo,
                        Of = ExtraerOfDesdeDescripcion(x.Descripcion)
                    })
                    .Where(x => !string.IsNullOrWhiteSpace(x.Of))
                    .ToList();

                if (ofPorDeposito.Count == 0)
                {
                    return Ok(new List<DepositosPosiblesConSaldoDTO>());
                }

                var ofsCerradas = await _context.Ordefabrs
                    .AsNoTracking()
                    .Where(x => x.IdSubOr != null && (x.StatuOrf ?? 0) == 3)
                    .Select(x => x.IdSubOr!.Trim())
                    .ToListAsync();

                if (ofsCerradas.Count == 0)
                {
                    return Ok(new List<DepositosPosiblesConSaldoDTO>());
                }

                var ofsCerradasSet = new HashSet<string>(ofsCerradas, StringComparer.OrdinalIgnoreCase);
                var depositosCandidatos = ofPorDeposito
                    .Where(x => ofsCerradasSet.Contains(x.Of))
                    .Select(x => x.CodigoDepo)
                    .Distinct()
                    .ToList();

                if (depositosCandidatos.Count == 0)
                {
                    return Ok(new List<DepositosPosiblesConSaldoDTO>());
                }

                var saldos = await _context.Movistos
                    .AsNoTracking()
                    .Where(x => x.DepoMovi.HasValue
                                && depositosCandidatos.Contains(x.DepoMovi.Value)
                                && x.Cod_Inte.HasValue
                                && x.Cod_Inte.Value > 0)
                    .GroupBy(x => new
                    {
                        Deposito = x.DepoMovi!.Value,
                        CodInte = x.Cod_Inte!.Value
                    })
                    .Select(g => new
                    {
                        g.Key.Deposito,
                        g.Key.CodInte,
                        Saldo = g.Sum(y => (y.CantIngre ?? 0m) - (y.CantSalid ?? 0m))
                    })
                    .Where(x => Math.Abs(x.Saldo) > 0.5m)
                    .OrderBy(x => x.Deposito)
                    .ThenBy(x => x.CodInte)
                    .ToListAsync();

                if (saldos.Count == 0)
                {
                    return Ok(new List<DepositosPosiblesConSaldoDTO>());
                }

                var codigos = saldos
                    .Select(x => x.CodInte)
                    .Distinct()
                    .ToList();

                var masters = await _context.masters
                    .AsNoTracking()
                    .Where(x => codigos.Contains(x.Codint))
                    .ToDictionaryAsync(x => x.Codint, x => x);

                var resultado = saldos
                    .Select(x =>
                    {
                        masters.TryGetValue(x.CodInte, out var master);
                        return new DepositosPosiblesConSaldoDTO
                        {
                            Codigo = master?.Codigo?.Trim() ?? string.Empty,
                            Descripcion = master?.Descripcion?.Trim() ?? string.Empty,
                            SaldoStock = Math.Round(x.Saldo, 1),
                            Deposito = x.Deposito
                        };
                    })
                    .ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener depósitos posibles con saldo");
                return StatusCode(500, $"Error al obtener depósitos posibles con saldo: {ex.Message}");
            }
        }

        // Equivalente a la lógica VB6 de Text1_Change + CargarDatosControSerie + CargaEstructuraMultinivel.
        // 1) Si existe CONTROSERIE original (CantIngre<=0 y CantSalid<=0) para NORFA+NUSERIE, devuelve esas filas.
        // 2) Si no existe, devuelve la estructura multinivel desde FORMULAS.
        [HttpGet("trazabilidad-detalle/{idSubOr}/{codInte:int}/{numeroSerie}")]
        public async Task<IActionResult> GetTrazabilidadDetalle(string idSubOr, int codInte, string numeroSerie, [FromQuery] string? depositos = null)
        {
            try
            {
                var norfa = idSubOr?.Trim();
                var nuserie = numeroSerie?.Trim();

                if (string.IsNullOrWhiteSpace(norfa) || string.IsNullOrWhiteSpace(nuserie))
                {
                    return Ok(new TrazabilidadOFSerieResultadoDTO());
                }

                var filasControserie = await _context.Controseries
                    .AsNoTracking()
                    .Where(x => x.NumeroSerie != null && x.NumeroSerie.Trim() == nuserie
                                && x.IdSubOr != null && x.IdSubOr.Trim() == norfa
                                && (x.CantIngre ?? 0m) <= 0m
                                && (x.CantSalid ?? 0m) <= 0m)
                    .OrderBy(x => x.Orden)
                    .ThenBy(x => x.Id)
                    .ToListAsync();

                if (filasControserie.Count > 0)
                {
                    var codigos = filasControserie
                        .Where(x => (x.CodInte ?? 0) > 0)
                        .Select(x => x.CodInte!.Value)
                        .Distinct()
                        .ToList();

                    var unidadesPorCodInte = await _context.masters
                        .AsNoTracking()
                        .Where(x => codigos.Contains(x.Codint))
                        .ToDictionaryAsync(x => x.Codint, x => x.Umedida ?? string.Empty);

                    var idsOriginales = filasControserie
                        .Where(x => x.ControserieId > 0)
                        .Select(x => x.ControserieId)
                        .Distinct()
                        .ToList();

                    var cantidadesAsignadas = await _context.Controseries
                        .AsNoTracking()
                        .Where(x => x.IdCodAsignado.HasValue && idsOriginales.Contains(x.IdCodAsignado.Value))
                        .GroupBy(x => x.IdCodAsignado!.Value)
                        .Select(g => new
                        {
                            IdCodAsignado = g.Key,
                            Cantidad = g.Sum(y => (y.CantIngre ?? 0m) - (y.CantSalid ?? 0m))
                        })
                        .ToDictionaryAsync(x => x.IdCodAsignado, x => x.Cantidad);

                    var resultadoControserie = filasControserie.Select(x => new TrazabilidadOFSerieDetalleDTO
                    {
                        CodInte = x.CodInte ?? 0,
                        CodExte = x.CodExte?.Trim() ?? string.Empty,
                        Descrip = x.Descrip?.Trim() ?? string.Empty,
                        Unidad = ((x.CodInte ?? 0) > 0 && unidadesPorCodInte.TryGetValue(x.CodInte!.Value, out var um)) ? um : string.Empty,
                        UsoBruto = x.UsoBruto ?? 0m,
                        CantidadAsignada = cantidadesAsignadas.TryGetValue(x.ControserieId, out var cantAsig) ? cantAsig : 0m,
                        ControserieId = x.ControserieId,
                        Orden = x.Orden ?? 0,
                        NivelIndentado = x.EspIdentado ?? 0
                    }).ToList();

                    await CompletarStockPorDepositosAsync(resultadoControserie, depositos);

                    return Ok(new TrazabilidadOFSerieResultadoDTO
                    {
                        TieneDatosControserie = true,
                        Items = resultadoControserie
                    });
                }

                if (codInte < 1)
                {
                    return Ok(new TrazabilidadOFSerieResultadoDTO());
                }

                var itemsEstructura = new List<TrazabilidadOFSerieDetalleDTO>();
                var cacheMasters = new Dictionary<int, Master>();
                await CargarEstructuraFormulaRecursivaAsync(itemsEstructura, codInte, nivel: 0, nivelMaximo: 100, cantidadPadre: 1m, cacheMasters, new HashSet<int>());
                await CompletarStockPorDepositosAsync(itemsEstructura, depositos);

                return Ok(new TrazabilidadOFSerieResultadoDTO
                {
                    TieneDatosControserie = false,
                    Items = itemsEstructura
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener detalle de trazabilidad para IdSubOr {IdSubOr} y NumeroSerie {NumeroSerie}", idSubOr, numeroSerie);
                return StatusCode(500, $"Error al obtener detalle de trazabilidad: {ex.Message}");
            }
        }

        [HttpPost("grabar-numero-serie")]
        public async Task<IActionResult> GrabarNumeroSerie([FromBody] GrabarNumeroSerieRequestDTO request)
        {
            try
            {
                // 1) Normalizamos datos principales recibidos desde UI.
                //    NORFA = IdSubOr de la orden de fabricación.
                //    NUSERIE = número de serie a registrar.
                //    CodInte = código interno del producto principal.
                var norfa = request.IdSubOr?.Trim();
                var nuserie = request.NumeroSerie?.Trim();
                var codInte = request.CodInte;

                // 2) Validaciones mínimas de entrada.
                if (string.IsNullOrWhiteSpace(norfa))
                {
                    return BadRequest("Falta seleccionar Orden de Fabricación");
                }

                if (string.IsNullOrWhiteSpace(nuserie))
                {
                    return BadRequest("Falta seleccionar N° de Serie");
                }

                if (codInte < 1)
                {
                    return BadRequest("Falta Código");
                }

                // 3) Tomamos los componentes válidos (CodInte > 0) en orden lógico.
                var items = request.Items
                    ?.Where(x => x.CodInte > 0)
                    .OrderBy(x => x.Orden)
                    .ToList() ?? new List<TrazabilidadOFSerieDetalleDTO>();

                if (items.Count == 0)
                {
                    return BadRequest("No hay componentes para grabar.");
                }

                // 4) Si ya existe NORFA+NUSERIE en CONTROSERIE, no se vuelve a insertar.
                var yaExiste = await _context.Controseries
                    .AsNoTracking()
                    .AnyAsync(x => x.NumeroSerie != null && x.NumeroSerie.Trim() == nuserie
                                   && x.IdSubOr != null && x.IdSubOr.Trim() == norfa);

                if (yaExiste)
                {
                    return Ok(new GrabarNumeroSerieResultadoDTO
                    {
                        Ok = true,
                        YaExistia = true,
                        Mensaje = "El N° de Serie ya fue grabado para la O.F. seleccionada.",
                        RegistrosGrabados = 0,
                        DepositoMovi = 0
                    });
                }

                // 5) Se resuelve el depósito (creación/reuso) con la lógica 60..200.
                var resultadoDeposito = await CrearDepositoSerieAsync(nuserie, norfa);
                if (!resultadoDeposito.Ok)
                {
                    return BadRequest(resultadoDeposito.Mensaje);
                }

                // 6) Base para generar ControserieId correlativo de las filas nuevas.
                var baseControserieId = await _context.Controseries
                    .AsNoTracking()
                    .Select(x => (int?)x.ControserieId)
                    .MaxAsync() ?? 0;

                // 7) Datos comunes para las nuevas filas.
                var depositoMovi = resultadoDeposito.CodigoDeposito;
                var ahora = DateTime.Now;

                // 8) Armado de filas CONTROSERIE.
                //    - Respeta longitudes de tabla (Descrip 64, CodExte 24).
                //    - CodiEmpr en 0 por regla del proyecto.
                //    - DepoMovi con el depósito determinado en paso 5.
                var nuevos = new List<Controserie>(items.Count);
                for (var i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    var descrip = (item.Descrip ?? string.Empty).Trim();
                    if (descrip.Length > 64)
                    {
                        descrip = descrip[..64];
                    }

                    var codExte = (item.CodExte ?? string.Empty).Trim();
                    if (codExte.Length > 24)
                    {
                        codExte = codExte[..24];
                    }

                    nuevos.Add(new Controserie
                    {
                        ControserieId = baseControserieId + i + 1,
                        CodiEmpr = 0,
                        Orden = (item.Orden > 0 ? item.Orden : i + 1),
                        CodInte = item.CodInte,
                        CodExte = codExte,
                        Descrip = descrip,
                        UsoBruto = item.UsoBruto,
                        NumeroSerie = nuserie,
                        IdSubOr = norfa,
                        CantIngre = 0m,
                        CantSalid = 0m,
                        IdenLote = string.Empty,
                        Marborra = 0,
                        FeReMovi = ahora,
                        Status = 0,
                        DepoMovi = depositoMovi,
                        EspIdentado = item.NivelIndentado
                    });
                }

                // 9) Persistencia de datos.
                _context.Controseries.AddRange(nuevos);
                await _context.SaveChangesAsync();

                // 10) Respuesta final con resumen de grabación.
                return Ok(new GrabarNumeroSerieResultadoDTO
                {
                    Ok = true,
                    YaExistia = false,
                    Mensaje = "Grabación realizada",
                    RegistrosGrabados = nuevos.Count,
                    DepositoMovi = depositoMovi
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al grabar número de serie en CONTROSERIE");
                return StatusCode(500, $"Error al grabar número de serie: {ex.Message}");
            }
        }

        private async Task<(bool Ok, int CodigoDeposito, string Mensaje)> CrearDepositoSerieAsync(string numeroSerie, string idSubOr)
        {
            const int depoMinimo = 60;
            const int depoMaximo = 200;

            // A) Normalización de entrada para comparar y formatear.
            //    - Serie sin espacios.
            //    - OF completa (ej: OF-000123-1) trim.
            var numeroSerieSinEspacios = QuitarEspacios(numeroSerie);
            var idSubOrNormalizado = (idSubOr ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(numeroSerieSinEspacios) || string.IsNullOrWhiteSpace(idSubOrNormalizado))
            {
                return (false, 0, "Datos inválidos para crear depósito de serie.");
            }

            // B) Armado del contenido de TADEPOSI:
            //    Descripcion: [1..16]=Serie, [17]=espacio, [18..31]=OF (14 chars).
            //    DescripAbreviada: NUSERIE-OF_sin_ceros_izquierda (máximo 24).
            var descripcionNueva = FormatearDescripcionDeposito(numeroSerieSinEspacios, idSubOrNormalizado);
            var descripAbreviadaNueva = FormatearDescripAbreviadaDeposito(numeroSerieSinEspacios, idSubOrNormalizado);

            // C) Se trabaja exclusivamente en rango de depósitos de serie (60..200).
            var depositosSerie = await _context.Tadeposi
                .Where(x => x.CodigoDepo >= depoMinimo && x.CodigoDepo <= depoMaximo)
                .OrderBy(x => x.CodigoDepo)
                .ToListAsync();

            // D) Validación: el N° de serie no puede estar ya usado en otro depósito del rango.
            var numeroSerieYaUsado = depositosSerie.Any(x => ExtraerNumeroSerieDesdeDescripcion(x.Descripcion) == numeroSerieSinEspacios);
            if (numeroSerieYaUsado)
            {
                return (false, 0, "El N° de Serie ya se ha utilizado en otro depósito.");
            }

            // E) Validación: la O.F. no puede estar ya asignada en otro depósito del rango.
            var ofYaUsada = depositosSerie.Any(x => string.Equals(ExtraerOfDesdeDescripcion(x.Descripcion), idSubOrNormalizado, StringComparison.OrdinalIgnoreCase));
            if (ofYaUsada)
            {
                return (false, 0, "La Orden de Fabricación ya posee depósito de serie.");
            }

            var codigoMaximoActual = depositosSerie.Count > 0 ? depositosSerie.Max(x => x.CodigoDepo) : 0;
            var codigoDepositoSeleccionado = 0;
            Tadeposi? depositoAReutilizar = null;

            // F) Estrategia de selección de depósito:
            //    1) Si 60 no existe, usar 60.
            //    2) Si aún no se llegó a 200, usar max+1.
            //    3) Si ya se alcanzó 200, buscar uno reutilizable.
            if (!depositosSerie.Any(x => x.CodigoDepo == depoMinimo))
            {
                codigoDepositoSeleccionado = depoMinimo;
            }
            else if (codigoMaximoActual < depoMaximo)
            {
                codigoDepositoSeleccionado = codigoMaximoActual + 1;
            }
            else
            {
                // Reutilización: solo depósitos cuyo contenido actual pertenezca a OF cerrada
                // y que además no tengan saldo en MOVISTO.
                for (var codigo = depoMinimo; codigo <= depoMaximo; codigo++)
                {
                    var depositoExistente = depositosSerie.FirstOrDefault(x => x.CodigoDepo == codigo);
                    if (depositoExistente is null)
                    {
                        codigoDepositoSeleccionado = codigo;
                        break;
                    }

                    var ofDeposito = ExtraerOfDesdeDescripcion(depositoExistente.Descripcion);
                    if (string.IsNullOrWhiteSpace(ofDeposito))
                    {
                        continue;
                    }

                    var ofCerrada = await _context.Ordefabrs
                        .AsNoTracking()
                        .AnyAsync(x => x.IdSubOr != null && x.IdSubOr.Trim() == ofDeposito && (x.StatuOrf ?? 0) == 3);

                    if (!ofCerrada)
                    {
                        continue;
                    }

                    // Si el depósito tiene saldo, NO es reutilizable.
                    if (await DepositoTieneSaldosAsync(codigo))
                    {
                        continue;
                    }

                    codigoDepositoSeleccionado = codigo;
                    depositoAReutilizar = depositoExistente;
                    break;
                }
            }

            if (codigoDepositoSeleccionado < depoMinimo || codigoDepositoSeleccionado > depoMaximo)
            {
                return (false, 0, "No se detectaron depósitos disponibles entre 60 y 200 (libres o reutilizables sin saldo con O.F. cerrada).");
            }

            // G) Persistencia en TADEPOSI:
            //    - Si es reuso, se actualiza registro existente.
            //    - Si es nuevo, se crea alta con status/disponible activos.
            if (depositoAReutilizar is not null)
            {
                depositoAReutilizar.Descripcion = descripcionNueva;
                depositoAReutilizar.DescripAbreviada = descripAbreviadaNueva;
                depositoAReutilizar.Status = 0;
                depositoAReutilizar.disponible = true;
                _context.Tadeposi.Update(depositoAReutilizar);
            }
            else
            {
                var nuevoDeposito = new Tadeposi
                {
                    CodigoDepo = codigoDepositoSeleccionado,
                    Descripcion = descripcionNueva,
                    DescripAbreviada = descripAbreviadaNueva,
                    Status = 0,
                    disponible = true
                };

                _context.Tadeposi.Add(nuevoDeposito);
            }

            await _context.SaveChangesAsync();
            return (true, codigoDepositoSeleccionado, string.Empty);
        }

        private async Task<bool> DepositoTieneSaldosAsync(int nroDeposito)
        {
            // Saldo por producto dentro del depósito:
            // SUM(CantIngre - CantSalid). Si algún producto supera |0.5|,
            // el depósito se considera "con saldo".
            var saldosPorProducto = await _context.Movistos
                .AsNoTracking()
                .Where(x => x.DepoMovi == nroDeposito)
                .GroupBy(x => x.Cod_Exte)
                .Select(g => g.Sum(y => (y.CantIngre ?? 0m) - (y.CantSalid ?? 0m)))
                .ToListAsync();

            return saldosPorProducto.Any(x => Math.Abs(x) > 0.5m);
        }

        private static string FormatearDescripcionDeposito(string numeroSerieSinEspacios, string idSubOr)
        {
            // Formato legado VB6:
            // - Primeras 16 posiciones: N° Serie.
            // - Posición 17: separador espacio.
            // - Desde posición 18 por 14 posiciones: O.F.
            var serie16 = AjustarLongitudFija(numeroSerieSinEspacios, 16);
            var of14 = AjustarLongitudFija((idSubOr ?? string.Empty).Trim(), 14);
            return $"{serie16} {of14}";
        }

        private static string FormatearDescripAbreviadaDeposito(string numeroSerieSinEspacios, string idSubOr)
        {
            var serie = QuitarEspacios(numeroSerieSinEspacios);
            var of = (idSubOr ?? string.Empty).Trim();

            var ofSinCeros = ExtraerNumeroCentralOfSinCeros(of);

            var abreviada = string.IsNullOrWhiteSpace(ofSinCeros)
                ? serie
                : $"{serie}-{ofSinCeros}";

            if (abreviada.Length > 24)
            {
                abreviada = abreviada[..24];
            }

            return abreviada;
        }

        private static string ExtraerNumeroCentralOfSinCeros(string of)
        {
            if (string.IsNullOrWhiteSpace(of))
            {
                return string.Empty;
            }

            var partes = of.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            // Formato esperado: OF-000123-1 => el número central es partes[1].
            if (partes.Length >= 2)
            {
                var bloqueCentral = partes[1];
                if (int.TryParse(bloqueCentral, out var numeroCentral))
                {
                    return numeroCentral.ToString();
                }

                var digitosCentral = new string(bloqueCentral.Where(char.IsDigit).ToArray());
                return int.TryParse(digitosCentral, out var numeroCentralFallback)
                    ? numeroCentralFallback.ToString()
                    : digitosCentral;
            }

            // Fallback para formatos no estándar sin guiones: usar todos los dígitos.
            var soloDigitos = new string(of.Where(char.IsDigit).ToArray());
            return int.TryParse(soloDigitos, out var numeroFallback)
                ? numeroFallback.ToString()
                : soloDigitos;
        }

        private static string AjustarLongitudFija(string valor, int longitud)
        {
            var limpio = valor ?? string.Empty;
            if (limpio.Length >= longitud)
            {
                return limpio[..longitud];
            }

            return limpio.PadRight(longitud);
        }

        private static string ExtraerNumeroSerieDesdeDescripcion(string? descripcion)
        {
            // Extrae el bloque [1..16] y quita espacios para comparar series.
            if (string.IsNullOrEmpty(descripcion))
            {
                return string.Empty;
            }

            var serie = descripcion.Length >= 16 ? descripcion[..16] : descripcion;
            return QuitarEspacios(serie);
        }

        private static string ExtraerOfDesdeDescripcion(string? descripcion)
        {
            // Extrae el bloque de O.F. desde índice 17 (posición 18 humana),
            // con longitud máxima de 14.
            if (string.IsNullOrEmpty(descripcion) || descripcion.Length < 18)
            {
                return string.Empty;
            }

            var largo = Math.Min(14, descripcion.Length - 17);
            return descripcion.Substring(17, largo).Trim();
        }

        private static string QuitarEspacios(string? valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
            {
                return string.Empty;
            }

            return new string(valor.Where(c => !char.IsWhiteSpace(c)).ToArray());
        }

        private async Task CargarEstructuraFormulaRecursivaAsync(
            List<TrazabilidadOFSerieDetalleDTO> resultado,
            int codIConj,
            int nivel,
            int nivelMaximo,
            decimal cantidadPadre,
            Dictionary<int, Master> cacheMasters,
            HashSet<int> ramaActual)
        {
            if (nivel > nivelMaximo || ramaActual.Contains(codIConj))
            {
                return;
            }

            ramaActual.Add(codIConj);

            var formulas = await _context.Formulas
                .AsNoTracking()
                .Where(x => x.CodIConj == codIConj)
                .OrderBy(x => x.NuorIt)
                .ThenBy(x => x.Id)
                .ToListAsync();

            foreach (var formula in formulas)
            {
                if ((formula.CodiElem ?? 0) <= 0)
                {
                    continue;
                }

                var codElem = formula.CodiElem!.Value;
                if (!cacheMasters.TryGetValue(codElem, out var masterElem))
                {
                    masterElem = await _context.masters.AsNoTracking().FirstOrDefaultAsync(x => x.Codint == codElem) ?? new Master { Codint = codElem };
                    cacheMasters[codElem] = masterElem;
                }

                var uso = Math.Round(formula.UsoBruto ?? 0m, 4);
                var cantidadAsignada = Math.Round(uso * Math.Round(cantidadPadre, 2), 4);

                resultado.Add(new TrazabilidadOFSerieDetalleDTO
                {
                    CodInte = codElem,
                    CodExte = formula.CodXElem?.Trim() ?? masterElem.Codigo?.Trim() ?? string.Empty,
                    Descrip = masterElem.Descripcion?.Trim() ?? string.Empty,
                    Unidad = masterElem.Umedida?.Trim() ?? string.Empty,
                    UsoBruto = uso,
                    CantidadAsignada = cantidadAsignada,
                    ControserieId = 0,
                    Orden = formula.NuorIt ?? 0,
                    NivelIndentado = nivel * 2
                });

                await CargarEstructuraFormulaRecursivaAsync(resultado, codElem, nivel + 1, nivelMaximo, cantidadAsignada, cacheMasters, ramaActual);
            }

            ramaActual.Remove(codIConj);
        }

        private async Task CompletarStockPorDepositosAsync(List<TrazabilidadOFSerieDetalleDTO> items, string? depositosRaw)
        {
            if (items.Count == 0)
            {
                return;
            }

            var depositos = ParsearDepositos(depositosRaw);
            if (depositos.Count == 0)
            {
                return;
            }

            var codigos = items
                .Where(x => x.CodInte > 0)
                .Select(x => x.CodInte)
                .Distinct()
                .ToList();

            if (codigos.Count == 0)
            {
                return;
            }

            var stockPorClave = new Dictionary<string, decimal>();
            var depositosNormales = depositos.Where(x => x != 5).Distinct().ToList();

            if (depositosNormales.Count > 0)
            {
                var stockMovisto = await _context.Movistos
                    .AsNoTracking()
                    .Where(x => x.Cod_Inte.HasValue
                                && codigos.Contains(x.Cod_Inte.Value)
                                && x.DepoMovi.HasValue
                                && depositosNormales.Contains(x.DepoMovi.Value))
                    .GroupBy(x => new { CodInte = x.Cod_Inte!.Value, Deposito = x.DepoMovi!.Value })
                    .Select(g => new
                    {
                        g.Key.CodInte,
                        g.Key.Deposito,
                        Stock = g.Sum(y => (y.CantIngre ?? 0m) - (y.CantSalid ?? 0m))
                    })
                    .ToListAsync();

                foreach (var stock in stockMovisto)
                {
                    stockPorClave[GenerarClaveStock(stock.CodInte, stock.Deposito)] = stock.Stock;
                }
            }

            if (depositos.Contains(5))
            {
                var pendientesBolRecep = await _context.Bolreceps
                    .AsNoTracking()
                    .Where(x => x.CodInte.HasValue
                                && codigos.Contains(x.CodInte.Value)
                                && (x.Status ?? 0) == 0
                                && ((x.CantNom ?? 0m) - (x.CanToApro ?? 0m)) > 0m)
                    .GroupBy(x => x.CodInte!.Value)
                    .Select(g => new
                    {
                        CodInte = g.Key,
                        Stock = g.Sum(y => (y.CantNom ?? 0m) - (y.CanToApro ?? 0m))
                    })
                    .ToDictionaryAsync(x => x.CodInte, x => x.Stock);

                var pendientesBrecepen = await _context.Brecepens
                    .AsNoTracking()
                    .Where(x => codigos.Contains(x.COD_INTE) && (x.Depomovi ?? 0) == 5)
                    .GroupBy(x => x.COD_INTE)
                    .Select(g => new
                    {
                        CodInte = g.Key,
                        Stock = g.Sum(y => (y.Cantingre ?? 0m) - (y.Cantsalid ?? 0m))
                    })
                    .ToDictionaryAsync(x => x.CodInte, x => x.Stock);

                foreach (var codigo in codigos)
                {
                    var pendienteRecepcion = pendientesBolRecep.TryGetValue(codigo, out var stockBolrecep) ? stockBolrecep : 0m;
                    var pendienteAprobacion = pendientesBrecepen.TryGetValue(codigo, out var stockBrecepen) ? stockBrecepen : 0m;
                    var stockDeposito5 = pendienteRecepcion + pendienteAprobacion;
                    stockPorClave[GenerarClaveStock(codigo, 5)] = stockDeposito5;
                }
            }

            foreach (var item in items)
            {
                item.StockPorDeposito ??= new Dictionary<int, decimal>();

                foreach (var deposito in depositos)
                {
                    var stock = stockPorClave.TryGetValue(GenerarClaveStock(item.CodInte, deposito), out var valor)
                        ? valor
                        : 0m;

                    item.StockPorDeposito[deposito] = Math.Round(stock, 4);
                }
            }
        }

        private static List<int> ParsearDepositos(string? depositosRaw)
        {
            if (string.IsNullOrWhiteSpace(depositosRaw))
            {
                return new List<int>();
            }

            var resultado = new List<int>();
            var vistos = new HashSet<int>();
            var tokens = depositosRaw.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var token in tokens)
            {
                if (!int.TryParse(token.Trim(), out var deposito) || deposito <= 0)
                {
                    continue;
                }

                if (vistos.Add(deposito))
                {
                    resultado.Add(deposito);
                }
            }

            return resultado;
        }

        private static string GenerarClaveStock(int codInte, int deposito)
        {
            return $"{codInte}|{deposito}";
        }

        [HttpPut("{id:int}")]
        public async Task<IActionResult> PutAsync(int id, [FromBody] OrdeFabr orden)
        {
            try
            {
                int? idFlexoft = _userContextService.GetCurrentUserId();
                short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();

                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión Caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                if (orden == null)
                {
                    return BadRequest("Datos inválidos.");
                }

                var ordenDb = await _context.Ordefabrs.FirstOrDefaultAsync(x => x.Id == id);
                if (ordenDb == null)
                {
                    return NotFound("Orden de fabricación no encontrada");
                }

                ordenDb.Destino = orden.Destino;
                ordenDb.CanProy = orden.CanProy;
                ordenDb.Priorid = orden.Priorid;
                ordenDb.SecciPro = orden.SecciPro;
                ordenDb.Observa = orden.Observa;
                ordenDb.AnotaOrf = orden.AnotaOrf;

                ordenDb.LastUpdate = DateTime.Now;
                ordenDb.NumUsuar = idFlexoft;
                ordenDb.CodiEmpr = codigoEmpresa;

                await _context.SaveChangesAsync();

                return Ok(ordenDb);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar orden de fabricación");
                return StatusCode(500, $"Error al actualizar la orden de fabricación: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene un listado de secciones productivas para el combo
        /// </summary>
        [HttpGet("secprod/combo")]
        public async Task<IActionResult> GetSecprodCombo()
        {
            try
            {
                var secProd = await _context.SecProds
                    .AsNoTracking()
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

                    // Generar operaciones de fabricación (OPERFAB) a partir de la secuencia estándar (SECOPER)
                    await GenerarOperacionesDesdeSecoper(
                        nuevaOrden.Cod_Inte.Value,
                        nuevaOrden.IdSubOr,
                        nuevaOrden.CanProy,
                        nuevaOrden.Cod_Exte,
                        nuevaOrden.Descrip,
                        codigoEmpresa,
                        idFlexoft);
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

        /// <summary>
        /// Recupera los datos completos necesarios para imprimir una Orden de Fabricación recién creada.
        /// Combina la orden, el nombre del usuario, la fecha de revisión del master y las operaciones (SECOPER).
        /// Ruta: GET /api/ordefabr/{id}/datos-impresion-orden-fabricacion
        /// </summary>
        /// <param name="id">Id interno de la orden de fabricación (PK Id de la tabla Ordefabr).</param>
        /// <returns>DTO OrdenFabricacionImpresionDTO listo para consumir desde el cliente.</returns>
        [HttpGet("{id:int}/datos-impresion-orden-fabricacion")]
        public async Task<IActionResult> ObtenerDatosImpresionOrdenFabricacion(int id)
        {
            // Buscar la orden de fabricación por su Id interno. No se modifica, por eso AsNoTracking.
            var orden = await _context.Ordefabrs
                .AsNoTracking()
                .FirstOrDefaultAsync(o => o.Id == id);

            if (orden == null)
            {
                return NotFound("No se encontró la orden de fabricación solicitada.");
            }

            // Nombre del usuario que lanzó la orden (Usuarios.IdFlexoft == Ordefabr.NumUsuar).
            string? nombreUsuario = null;
            if (orden.NumUsuar.HasValue)
            {
                nombreUsuario = await _context.Usuarios
                    .AsNoTracking()
                    .Where(u => u.IdFlexoft == orden.NumUsuar.Value)
                    .Select(u => u.NombreUsuario)
                    .FirstOrDefaultAsync();
            }

            // Fecha de revisión del master para el Cod_Inte de la orden.
            // Master.Codint equivale al Id lógico del producto.
            DateTime? ferevis = null;
            if (orden.Cod_Inte.HasValue)
            {
                ferevis = await _context.masters
                    .AsNoTracking()
                    .Where(m => m.Codint == orden.Cod_Inte.Value)
                    .Select(m => m.Ferevis)
                    .FirstOrDefaultAsync();
            }

            // Operaciones de la ruta de fabricación: SECOPER del conjunto (CodIConj = Cod_Inte), activas (StatOper = 1), ordenadas por NumeOper.
            var operaciones = new List<OrdenFabricacionOperacionDTO>();
            if (orden.Cod_Inte.HasValue)
            {
                operaciones = await _context.Secopers
                    .AsNoTracking()
                    .Where(s => s.CodIConj == orden.Cod_Inte.Value && s.StatOper == 1)
                    .OrderBy(s => s.NumeOper)
                    .Select(s => new OrdenFabricacionOperacionDTO
                    {
                        NumeOper = s.NumeOper,
                        DescOper = s.DescOper,
                        ObseOper = s.ObseOper
                    })
                    .ToListAsync();
            }

            // Armar el DTO con todos los datos que necesita el formulario A4 en el cliente.
            // Las propiedades "...Formateada" son textos listos para insertar en el HTML,
            // de modo que JavaScript no tenga que formatear fechas/decimal y se eviten
            // problemas de zona horaria o cultura en el navegador.
            var dto = new OrdenFabricacionImpresionDTO
            {
                NumeOrFa = orden.NumeOrFa,
                IdSubOr = orden.IdSubOr,
                FechGen = orden.FechGen,
                FechGenFormateada = orden.FechGen?.ToString("dd/MM/yyyy") ?? string.Empty,
                Cod_Exte = orden.Cod_Exte,
                Descrip = orden.Descrip,
                CanProy = orden.CanProy,
                CanProyFormateada = orden.CanProy?.ToString("0.00") ?? string.Empty,
                Cod_Inte = orden.Cod_Inte,
                Unidad = orden.Unidad,
                NumUsuar = orden.NumUsuar,
                NombreUsuario = nombreUsuario ?? orden.NumUsuar?.ToString() ?? "Usuario",
                Ferevis = ferevis,
                FerevisFormateada = ferevis?.ToString("dd/MM/yyyy") ?? string.Empty,
                Operaciones = operaciones
            };

            return Ok(dto);
        }

        /// <summary>
        /// Recupera los datos necesarios para imprimir el Vale de Materiales de una Orden de Fabricación.
        /// Mantiene el mismo encabezado que la OF y lista los materiales reservados con CantRes > 0.
        /// La unidad de medida se obtiene de Master.Umedida según Master.Codint = Reserva.Cod_Inte.
        /// Ruta: GET /api/ordefabr/{id}/datos-impresion-consumo-materiales
        /// </summary>
        /// <param name="id">Id interno de la orden de fabricación (PK Id de la tabla Ordefabr).</param>
        /// <returns>DTO OrdenFabricacionConsumoImpresionDTO listo para consumir desde el cliente.</returns>
        [HttpGet("{id:int}/datos-impresion-consumo-materiales")]
        public async Task<IActionResult> ObtenerDatosImpresionConsumoMateriales(int id)
        {
            // Buscar la orden de fabricación por su Id interno. No se modifica, por eso AsNoTracking.
            var orden = await _context.Ordefabrs
                .AsNoTracking()
                .FirstOrDefaultAsync(o => o.Id == id);

            if (orden == null)
            {
                return NotFound("No se encontró la orden de fabricación solicitada.");
            }

            // Nombre del usuario que lanzó la orden (Usuarios.IdFlexoft == Ordefabr.NumUsuar).
            string? nombreUsuario = null;
            if (orden.NumUsuar.HasValue)
            {
                nombreUsuario = await _context.Usuarios
                    .AsNoTracking()
                    .Where(u => u.IdFlexoft == orden.NumUsuar.Value)
                    .Select(u => u.NombreUsuario)
                    .FirstOrDefaultAsync();
            }

            // Fecha de revisión del master para el Cod_Inte de la orden.
            DateTime? ferevis = null;
            if (orden.Cod_Inte.HasValue)
            {
                ferevis = await _context.masters
                    .AsNoTracking()
                    .Where(m => m.Codint == orden.Cod_Inte.Value)
                    .Select(m => m.Ferevis)
                    .FirstOrDefaultAsync();
            }

            // Materiales reservados para la OF: Reserva.IdSubOr = orden.IdSubOr, CantRes > 0.
            // Se une con Master para traer la unidad de medida (Master.Umedida).
            var materiales = new List<OrdenFabricacionMaterialConsumoDTO>();
            if (!string.IsNullOrWhiteSpace(orden.IdSubOr))
            {
                materiales = await (from r in _context.Reservas
                                    join m in _context.masters on r.Cod_Inte equals m.Codint
                                    where r.IdSubOr == orden.IdSubOr && r.CantRes > 0
                                    orderby r.Cod_Exte
                                    select new OrdenFabricacionMaterialConsumoDTO
                                    {
                                        Cod_Exte = r.Cod_Exte,
                                        Descrip = r.Descrip,
                                        Umedida = m.Umedida,
                                        UsoUnit = r.UsoUnit,
                                        UsoUnitFormateada = r.UsoUnit.HasValue ? r.UsoUnit.Value.ToString("0.0000") : string.Empty,
                                        CantRes = r.CantRes,
                                        CantResFormateada = r.CantRes.HasValue ? r.CantRes.Value.ToString("0.00") : string.Empty
                                    })
                                    .ToListAsync();
            }

            var dto = new OrdenFabricacionConsumoImpresionDTO
            {
                NumeOrFa = orden.NumeOrFa,
                IdSubOr = orden.IdSubOr,
                FechGen = orden.FechGen,
                FechGenFormateada = orden.FechGen?.ToString("dd/MM/yyyy") ?? string.Empty,
                Cod_Exte = orden.Cod_Exte,
                Descrip = orden.Descrip,
                CanProy = orden.CanProy,
                CanProyFormateada = orden.CanProy?.ToString("0.00") ?? string.Empty,
                Cod_Inte = orden.Cod_Inte,
                Unidad = orden.Unidad,
                NumUsuar = orden.NumUsuar,
                NombreUsuario = nombreUsuario ?? orden.NumUsuar?.ToString() ?? "Usuario",
                Ferevis = ferevis,
                FerevisFormateada = ferevis?.ToString("dd/MM/yyyy") ?? string.Empty,
                Materiales = materiales
            };

            return Ok(dto);
        }

        private static string Safe(string? value) => value ?? string.Empty;

        private static string Limit(string value, int maxLen)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value.Length <= maxLen ? value : value.Substring(0, maxLen);
        }

        private static string PadRightExact(string value, int len)
        {
            value = value ?? string.Empty;
            if (value.Length >= len)
            {
                return value.Substring(0, len);
            }

            return value.PadRight(len, ' ');
        }

        private async Task<string> ObtenerDescripcionMaquinaAsync(string codigoMaquina)
        {
            if (string.IsNullOrWhiteSpace(codigoMaquina))
            {
                return string.Empty;
            }

            try
            {
                // Tabla PADMAQ
                var desc = await _context.padmaqs
                    .AsNoTracking()
                    .Where(x => x.CodigoPadMaq == codigoMaquina)
                    .Select(x => x.Descripcion)
                    .FirstOrDefaultAsync();

                return desc ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Anula una orden de fabricación completa con todas sus dependencias
        /// Replica la lógica VB6 de anulación en una sola transacción optimizada
        /// </summary>
        /// <param name="idSubOr">ID de la suborden a anular</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut("anular/{idSubOr}")]
        public async Task<IActionResult> AnularOrdenFabricacionAsync(string idSubOr)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                Console.WriteLine($"[OrdeFabrController] Iniciando anulación de O.F. {idSubOr}");

                // 1. Buscar la orden de fabricación
                var ordenFabricacion = await _context.Ordefabrs
                    .FirstOrDefaultAsync(o => o.IdSubOr == idSubOr);

                if (ordenFabricacion == null)
                {
                    return NotFound($"No se encontró la orden de fabricación con IdSubOr {idSubOr}");
                }

                Console.WriteLine($"[OrdeFabrController] Orden encontrada: NumeOrFa={ordenFabricacion.NumeOrFa}, IdSubOr={ordenFabricacion.IdSubOr}");

                // 2. Actualizar StatuOrf = 9 en OrdeFabr y agregar anotación de anulación
                ordenFabricacion.StatuOrf = 9;
                ordenFabricacion.LastUpdate = DateTime.Now;
                ordenFabricacion.CodiEmpr = 0;
                
                var usuarioActivo = _userContextService.GetCurrentUserId();
                var nombreUsuario = ObtenerNombreUsuario();
                string idUsuario = usuarioActivo.HasValue ? usuarioActivo.Value.ToString() : "";
                string userTer = $"{idUsuario} - {nombreUsuario}";
                
                // Agregar anotación de anulación al campo AnotaOrf
                string anotacionAnulacion = GenerarTextoAnotaOrfAnulacion(ordenFabricacion, userTer);
                ordenFabricacion.AnotaOrf = (ordenFabricacion.AnotaOrf ?? "") + anotacionAnulacion;
                
                if (usuarioActivo.HasValue)
                {
                    ordenFabricacion.NumUsuar = usuarioActivo.Value;
                }

                Console.WriteLine($"[OrdeFabrController] StatuOrf actualizado a 9 y anotación agregada");

                // 3. Actualizar reservas (CANTRES = CANTCONS, Saldo_Entre = 0) - Operación bulk optimizada
                var reservasActualizadas = await _context.Reservas
                    .Where(r => r.IdSubOr == idSubOr)
                    .ExecuteUpdateAsync(setters => setters
                        .SetProperty(r => r.CantRes, r => r.CantCons)
                        .SetProperty(r => r.Saldo_Entre, 0)
                        .SetProperty(r => r.LastUpdate, DateTime.Now)
                        .SetProperty(r => r.NumUsuar, usuarioActivo ?? 0));

                Console.WriteLine($"[OrdeFabrController] Reservas actualizadas: {reservasActualizadas} registros");

                // 4. Eliminar OperFab asociadas - Operación bulk optimizada
                var operFabEliminadas = await _context.Operfabs
                    .Where(o => o.IdSubOr == idSubOr)
                    .ExecuteDeleteAsync();

                Console.WriteLine($"[OrdeFabrController] OperFab eliminadas: {operFabEliminadas} registros");

                // 5. Anular OCOMDETA si Doc_Orig = IdSubOr y N_OrServ > 0 (STATUS = 9)
                var ocomdetaActualizadas = await _context.Ocomdetas
                    .Where(o => o.DocOrig == idSubOr && (o.NOrServ ?? 0) > 0)
                    .ExecuteUpdateAsync(setters => setters
                        .SetProperty(o => o.StatOcom, o => (short?)9)
                        .SetProperty(o => o.LastUpdate, DateTime.Now)
                        .SetProperty(o => o.NumUsuar, o => usuarioActivo.HasValue ? usuarioActivo.Value : (int?)null));

                Console.WriteLine($"[OrdeFabrController] OCOMDETA anuladas: {ocomdetaActualizadas} registros");

                // Guardar cambios
                await _context.SaveChangesAsync();

                // Confirmar transacción
                await transaction.CommitAsync();

                Console.WriteLine($"[OrdeFabrController] Anulación completada exitosamente para O.F. {idSubOr}");

                return Ok(new { 
                    Message = $"Orden de fabricación {idSubOr} anulada exitosamente",
                    ReservasActualizadas = reservasActualizadas,
                    OperFabEliminadas = operFabEliminadas,
                    OcomdetaActualizadas = ocomdetaActualizadas
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al anular la orden de fabricación {IdSubOr}", idSubOr);
                Console.WriteLine($"[OrdeFabrController] Error al anular: {ex.Message}");
                return StatusCode(500, $"Error al anular la orden de fabricación: {ex.Message}");
            }
        }

        private async Task GenerarOperacionesDesdeSecoper(
            int codigoInternoProducto,
            string idSubOr,
            decimal? cantidadProyectada,
            string? codExteProducto,
            string? descripProducto,
            short? codigoEmpresa,
            int? idFlexoft)
        {
            if (string.IsNullOrWhiteSpace(idSubOr))
            {
                return;
            }

            // Equivalente VB6: CAPROYE#
            var caproye = cantidadProyectada ?? 0m;
            if (caproye < 0)
            {
                caproye = 0m;
            }

            // Si existieran operaciones previas para la OF, se limpian.
            var existentes = await _context.Operfabs.Where(o => o.IdSubOr == idSubOr).ToListAsync();
            if (existentes.Count > 0)
            {
                _context.Operfabs.RemoveRange(existentes);
                await _context.SaveChangesAsync();
            }

            var secopers = await _context.Secopers
                .AsNoTracking()
                .Where(x => x.CodIConj == codigoInternoProducto && x.StatOper == 1)
                .OrderBy(x => x.NumeOper)
                .ToListAsync();

            if (secopers.Count == 0)
            {
                return;
            }

            var operaciones = new List<Operfab>(secopers.Count);
            foreach (var s in secopers)
            {
                // VB6: If Not IsNull(!AltOper) Then ATO% = !AltOper; If ATO% < 1 Then...
                var altOper = (int)(s.AltOper ?? 0);
                if (altOper >= 1)
                {
                    continue;
                }

                var numeOper = s.NumeOper ?? 0;

                // VB6: CAREALI# (real) depende de CANREALI().
                // PENDIENTE: CANREALI proviene de la tabla INFOFAB, la cual todavía no está implementada.
                // Por ahora en el alta de la O/F se asigna 0.
                var cantReal = 0m;

                // VB6: CAPENDE# = CAPROYE# - CAREALI#
                var cantPend = caproye - cantReal;
                if (cantPend < 0)
                {
                    cantPend = 0m;
                }

                // Defaults y conversiones VB6
                var caOps = s.CantiOps ?? 1m;
                if (caOps < 0.01m)
                {
                    caOps = 1m;
                }

                var cicloHe = s.PieCiclo ?? 1m;
                if (cicloHe < 1m)
                {
                    cicloHe = 1m;
                }

                const decimal calotes = 1m;

                var minPrep = 60m * (s.HorsPrep ?? 0m) * calotes;
                if (cantReal > 0.05m)
                {
                    minPrep = 0m;
                }

                var minFijo = 60m * (s.HorsFijo ?? 0m) * calotes;
                var minVari = cantPend * (s.MinuStan ?? 0m);
                var minHerr = cantPend * 60m * (s.HorsMHer ?? 0m) / cicloHe;
                var minDemo = (minFijo + minVari + minHerr) * (s.PorDemor ?? 0m) / 100m;
                var minLimp = 60m * (s.HorsLimp ?? 0m) * calotes;
                var minLote = minPrep + minFijo + minVari + minHerr + minDemo + minLimp;

                var maquina = Safe(s.CodMaq0);
                var descMaq = await ObtenerDescripcionMaquinaAsync(maquina);

                var equiAux = string.Empty;
                var m1 = Safe(s.CodMaq1);
                var m2 = Safe(s.CodMaq2);
                var m3 = Safe(s.CodMaq3);
                if (!string.IsNullOrWhiteSpace(m1)) equiAux += PadRightExact(m1, 6);
                if (!string.IsNullOrWhiteSpace(m2)) equiAux += PadRightExact(m2, 6);
                if (!string.IsNullOrWhiteSpace(m3)) equiAux += PadRightExact(m3, 6);
                equiAux = Limit(equiAux, 18);

                operaciones.Add(new Operfab
                {
                    // Nota: en VB se asignaba CODIEMPR = CodiEmp%. En el sistema actual se usa 0 en varios flujos.
                    CodiEmpr = 0,
                    NumUsuar = idFlexoft,

                    IdSubOr = Limit(idSubOr, 12),
                    Cod_Inte = codigoInternoProducto,
                    Cod_Exte = Limit(Safe(codExteProducto), 24),
                    Descrip = Limit(Safe(descripProducto), 64),

                    NumeOper = numeOper,
                    DescOper = Limit(Safe(s.DescOper), 48),
                    Maquina = Limit(maquina, 6),
                    DescMaq = Limit(descMaq, 64),
                    EquiAux = Limit(equiAux, 18),
                    CantiOps = caOps,

                    CantProg = caproye,
                    CantReal = cantReal,
                    CantPend = cantPend,

                    TiemPrep = minPrep,
                    TiemFijo = minFijo,
                    TiemVari = minVari,
                    TiemDemo = minDemo,
                    TiemMHer = minHerr,
                    TiemLimp = minLimp,
                    TiemOtro = 0m,
                    TiemTota = minLote,
                    TiemEjec = 0m,
                    TiemPend = minLote,

                    StatOper = 1,
                    ComiePro = null,
                    TermiPro = null,
                    TermiTemp = null,

                    Observa = " ",
                    HoBatchRec = "",
                    SecciPro = 0m,
                    rowguid = Guid.NewGuid(),
                    NuOrdOpe = null,
                });
            }

            if (operaciones.Count > 0)
            {
                _context.Operfabs.AddRange(operaciones);
                await _context.SaveChangesAsync();
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

        [NonAction]
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

                return Ok(ordenFabr);
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

                // Actualizar la orden directamente en DB (evita DbUpdateConcurrencyException por entidades desincronizadas)
                // y permite validar si la fila existe / fue modificada por otro proceso.
                var filasOrdefabr = await _context.Ordefabrs
                    .Where(o => o.Id == orden.Id)
                    .ExecuteUpdateAsync(s => s
                        .SetProperty(o => o.CanApro, orden.CanApro)
                        .SetProperty(o => o.CanRech, orden.CanRech)
                        .SetProperty(o => o.Canti_Fin_Proceso, orden.Canti_Fin_Proceso)
                        .SetProperty(o => o.Status_Proceso, orden.Status_Proceso)
                        .SetProperty(o => o.StatuOrf, orden.StatuOrf)
                        .SetProperty(o => o.FechCierr, orden.FechCierr)
                        .SetProperty(o => o.AnotaOrf, orden.AnotaOrf));

                if (filasOrdefabr == 0)
                {
                    await transaction.RollbackAsync();
                    return Conflict("La orden no pudo actualizarse (puede haber sido modificada/eliminada por otro proceso). Reintente.");
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
        /// Genera el texto de anotación para la anulación de orden de fabricación
        /// Replica el formato exacto del código VB6 para anulación
        /// </summary>
        /// <param name="orden">Orden de fabricación</param>
        /// <param name="userTer">Nombre del usuario que anula la orden (ID + Nombre)</param>
        /// <returns>Texto formateado para el campo AnotaOrf</returns>
        private string GenerarTextoAnotaOrfAnulacion(OrdeFabr orden, string userTer)
        {
            // Replicar el formato exacto del código VB6 para anulación
            string ttxyz = new string(' ', 76); // Space$(76)
            string hos = DateTime.Now.ToString("dd/MM/yy"); // Equivalente a HOY(HOS$)

            // Call REPLA(TTXYZ$, HOS$, 1, 8)
            ttxyz = Repla(ttxyz, hos, 1, 8);

            // Call REPLA(TTXYZ$, "Anulación de O/Fabricación", 11, 56)
            ttxyz = Repla(ttxyz, "Anulación de O/Fabricación", 11, 56);

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

            // Generar un IDTEXT único para toda la transacción de cierre
            // Este IDTEXT se usará para vincular todos los registros ORFAPEN de esta transacción
            string idTextoTransaccion = GenerarIdTextoAleatorio();

            // Declarar variable para capturar el ID del movimiento OF generado
            // Este ID se usará para vincular los movimientos CF con el OF mediante IdenMovi
            // En modo estándar se captura el ID, en modo control de calidad queda null
            int? idMovimientoOF = null;

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
                //
                // NOTA IMPORTANTE SOBRE VINCULACIÓN DE MOVIMIENTOS EN MODO CONTROL DE CALIDAD:
                // En modo de control de calidad, los movimientos OF se guardan en la tabla ORFAPEN (no en MOVISTO).
                // Por lo tanto, no es posible vincular los movimientos CF con el OF mediante IdenMovi de la misma forma
                // que en modo estándar (donde ambos están en MOVISTO).
                //
                // Los movimientos CF (consumos) siempre se guardan en MOVISTO, pero en modo de control de calidad
                // no tienen un ID de OF para vincular porque el OF está en ORFAPEN.
                //
                // Esto significa que la anulación de cierres parciales en modo de control de calidad
                // no puede usar el campo IdenMovi para anular automáticamente los CF asociados.
                // Sería necesario implementar una lógica diferente para este caso.

                // Generar DOSEC cambiando la 'O' por 'I' en la orden de fabricación
                string ordenFabricacion = orden.IdSubOr ?? "";
                string dosec = ordenFabricacion.Replace("OF-", "IF-");

                bool resultadoMovimiento = await MoviStockAprof(
                    fecha: datosCierre.FechaCierre,
                    codigoInterno: orden.Cod_Inte ?? 0,
                    lote: orden.IdenLote ?? "",
                    codigoMovimiento: "OF", // Código de movimiento para orden de fabricación
                    dopri: orden.IdSubOr ?? "",
                    dosec: dosec, // Cambiar PF- por IF-
                    referencia: "Fabricación Propia", // Referencia fija para cierre de OF
                    valorUnitario: 0, // Se puede parametrizar si es necesario
                    numeroCliente: orden.Nume_Cli ?? 0,
                    deposito: datosCierre.DepositoEntrada ?? 0, // Usar DepositoEntrada como destino
                    cantidad: cantiApro, // Usar la cantidad aprobada en modo calidad
                    depositoConsumo: datosCierre.DepositoSalida ?? 0, // Usar DepositoSalida como consumo
                    idTexto: idTextoTransaccion // Usar el IDTEXT único de la transacción
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

                // Generar DOSEC cambiando la 'O' por 'I' en la orden de fabricación
                string ordenFabricacion = orden.IdSubOr ?? "";
                string dosec = ordenFabricacion.Replace("OF-", "IF-");

                var movimientoStock = new MoviStockDTO
                {
                    Fecha = datosCierre.FechaCierre,
                    CodigoInterno = orden.Cod_Inte ?? 0,
                    Lote = orden.IdenLote ?? "",
                    CodigoMovimiento = "OF", // Código de movimiento para orden de fabricación
                    DocumentoPrimarioLargo = orden.IdSubOr ?? "",
                    DocumentoPrimarioCorto = orden.IdSubOr ?? "",
                    DocumentoSecundarioLargo = dosec, // Cambiar PF- por IF-
                    DocumentoSecundarioCorto = dosec, // Cambiar PF- por IF-
                    Referencia = "Fabricación Propia", // Referencia fija para cierre de OF
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

                // CAPTURA DEL ID DEL MOVIMIENTO OF GENERADO
                // Este ID es crucial para vincular los movimientos de consumo (CF) con el cierre (OF)
                // 
                // Flujo:
                // 1. El servicio RegistrarMovimientoAsync retorna un OkObjectResult con el ID del movimiento creado
                // 2. Extraemos ese ID del resultado dinámico
                // 3. Pasamos este ID a ProcesarConsumoComponentes
                // 4. En ProcesarConsumoComponentes, cada movimiento CF recibe este ID en IdentificadorMovimiento
                // 5. El servicio guarda este valor en el campo MOVISTO.IdenMovi
                // 
                // Uso en anulación de cierres:
                // Cuando se anula un cierre (movimiento OF), el sistema busca todos los movimientos CF
                // que tienen IdenMovi igual al ID del movimiento OF a anular, y también los anula.
                // Esto permite revertir completamente un cierre: al anular el OF se anulan
                // automáticamente todos los consumos (CF) que se generaron en ese mismo cierre.
                if (resultadoMovimiento is OkObjectResult okResult && okResult.Value != null)
                {
                    var resultDict = okResult.Value as dynamic;
                    idMovimientoOF = resultDict?.Id;
                    _logger.LogInformation($"ID del movimiento OF generado: {idMovimientoOF}");
                }

                _logger.LogInformation($"Movimiento de stock procesado exitosamente para orden {orden.IdSubOr} en modo estándar. Cantidad: {canti}");
            }

            // Procesar consumo de componentes, pasando el ID del movimiento OF y el IDTEXT de la transacción
            // El IDTEXT se usará para vincular todos los registros ORFAPEN de esta transacción
            await ProcesarConsumoComponentes(orden, datosCierre, idMovimientoOF, idTextoTransaccion);

            // Actualizar campos de la orden según la lógica VB6
            decimal canFinProceso = orden.Canti_Fin_Proceso ?? 0;
            decimal canAproActual = orden.CanApro ?? 0;

            // En modo de control de calidad, usar cantiApro (que tiene la cantidad real)
            // En modo estándar, usar canti
            decimal cantidadParaCanApro = datosCierre.ModoCierreCalidad == 1 ? cantiApro : canti;
            decimal aprobadoTotal = canAproActual + cantidadParaCanApro;

            _logger.LogInformation($"Actualizando CanApro: valor actual={canAproActual}, cantidad a agregar={cantidadParaCanApro}, nuevo total={aprobadoTotal}, modoCierreCalidad={datosCierre.ModoCierreCalidad}");

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

                // IMPORTANTE: si previamente se cargaron/actualizaron Reservas u Operfab en este mismo DbContext,
                // quedan trackeadas como Modified. Si luego se borran por SQL directo, al final del proceso
                // SaveChanges intentará hacer UPDATE sobre filas ya eliminadas y dispara DbUpdateConcurrencyException.
                // Por eso, antes de borrar, desvinculamos del ChangeTracker cualquier entidad relacionada a la O/F.
                foreach (var entry in _context.ChangeTracker.Entries<Reserva>()
                             .Where(e => e.Entity.IdSubOr == idSubOr)
                             .ToList())
                {
                    entry.State = EntityState.Detached;
                }

                foreach (var entry in _context.ChangeTracker.Entries<Operfab>()
                             .Where(e => e.Entity.IdSubOr == idSubOr)
                             .ToList())
                {
                    entry.State = EntityState.Detached;
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
            // Call REPLA(TTXYZ$, "     - Cierre Pendiente de Aprobacion = " & FORMATNUM(CANTI, "F12.1"), 11, 56)
            // Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
            ttxyz = new string(' ', 76);
            string cantidadFormateada = cantidadAprobada.ToString("F1");
            ttxyz = Repla(ttxyz, $"     - Cierre Pendiente de Aprobacion = {cantidadFormateada}", 11, 56);
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
        /// <param name="idMovimientoOF">ID del movimiento OF generado para asignar a IdenMovi de los movimientos CF</param>
        /// <param name="idTextoTransaccion">IDTEXT único para todos los registros ORFAPEN de esta transacción</param>
        private async Task ProcesarConsumoComponentes(OrdeFabr orden, DatosCierreAvanzadoDTO datosCierre, int? idMovimientoOF = null, string idTextoTransaccion = "")
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
                            // 
                            // NOTA IMPORTANTE: En modo de control de calidad, los movimientos CF (consumos)
                            // se guardan usando MoviStockAprof que guarda en ORFAPEN, no en MOVISTO.
                            // Por lo tanto, estos movimientos CF no tendrán el campo IdenMovi vinculado
                            // al movimiento OF porque el OF también está en ORFAPEN y no hay un ID común.
                            // 
                            // Esto significa que la anulación de cierres parciales en modo de control de calidad
                            // no puede usar el campo IdenMovi para anular automáticamente los CF asociados.
                            //
                            // Para GenCierre + CF: cantidad positiva en CantSalid
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
                                cantidad: cantidadConsumo, // Pasar cantidad negativa para que se guarde en CantSalid
                                depositoConsumo: datosCierre.DepositoSalida ?? 0,
                                idTexto: idTextoTransaccion // Usar el IDTEXT único de la transacción
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
                                TipoTransferencia = null,
                                
                                // IMPORTANTE: Vinculación entre movimiento OF y movimientos CF
                                // IdentificadorMovimiento se guarda en el campo MOVISTO.IdenMovi
                                // Este campo contiene el ID del movimiento OF que generó este consumo CF
                                // 
                                // Uso principal: Anulación de cierres parciales
                                // Cuando se anula un cierre (movimiento OF), el sistema busca todos los movimientos CF
                                // que tienen IdenMovi igual al ID del movimiento OF a anular, y también los anula.
                                // 
                                // Esto permite revertir completamente un cierre: al anular el OF se anulan
                                // automáticamente todos los consumos (CF) que se generaron en ese mismo cierre.
                                // 
                                // Ejemplo:
                                // - Cierre OF con ID 12345 genera 3 consumos CF con IdenMovi = "12345"
                                // - Si se anula el cierre OF 12345, se buscan los CF con IdenMovi = "12345"
                                // - Se anulan los 3 CF encontrados (cantidades a 0)
                                IdentificadorMovimiento = idMovimientoOF.HasValue ? idMovimientoOF.Value.ToString() : null
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
        /// Regenera las reservas de una orden de fabricación tras la anulación de un cierre parcial
        /// Equivalente al código VB6 GENERES_ANUCIEOF
        /// 
        /// Lógica VB6:
        /// - DELETE FROM RESERVA WHERE IdSubOr = NORFA$
        /// - CP = CantiProyectadaOF(NORFA$)
        /// - CA = CantiAprobadaOF(NORFA$)
        /// - CAN = CantiProyectadaOF(NORFA$) - CantiAprobadaOF(NORFA$)
        /// - Si CAN > 0, regenerar reservas para la cantidad pendiente
        /// 
        /// Esto es necesario porque al anular un cierre, se reduce la cantidad aprobada
        /// y las reservas deben ajustarse a la nueva cantidad pendiente (proyectada - aprobada)
        /// </summary>
        /// <param name="orden">Orden de fabricación que sufrió la anulación de cierre</param>
        private async Task RegenerarReservasTrasAnulacionCierre(OrdeFabr orden)
        {
            try
            {
                // Paso 1: Eliminar todas las reservas existentes para esta orden
                // Equivalente VB6: DELETE FROM RESERVA WHERE IdSubOr = NORFA$
                var reservasExistentes = await _context.Reservas
                    .Where(r => r.IdSubOr == orden.IdSubOr)
                    .ToListAsync();

                if (reservasExistentes.Any())
                {
                    _context.Reservas.RemoveRange(reservasExistentes);
                    await _context.SaveChangesAsync();
                    _logger.LogInformation($"Eliminadas {reservasExistentes.Count} reservas existentes para orden {orden.IdSubOr}");
                }

                // Paso 2: Calcular la cantidad pendiente (proyectada - aprobada)
                // Equivalente VB6: CAN = CantiProyectadaOF(NORFA$) - CantiAprobadaOF(NORFA$)
                decimal cantidadProyectada = orden.CanProy ?? 0;
                decimal cantidadAprobada = orden.CanApro ?? 0;
                decimal cantidadPendiente = cantidadProyectada - cantidadAprobada;

                // Paso 3: Si la cantidad pendiente es <= 0, no hay nada que reservar
                // Equivalente VB6: If CAN <= 0 Then Exit Sub
                if (cantidadPendiente <= 0)
                {
                    _logger.LogInformation($"Cantidad pendiente ({cantidadPendiente}) <= 0, no se regeneran reservas para orden {orden.IdSubOr}");
                    return;
                }

                // Paso 4: Verificar que la orden tenga código interno válido
                if (!orden.Cod_Inte.HasValue || orden.Cod_Inte.Value <= 0)
                {
                    _logger.LogWarning($"Orden {orden.IdSubOr} no tiene Cod_Inte válido, no se regeneran reservas");
                    return;
                }

                // Paso 5: Obtener los componentes de la fórmula del producto
                // Equivalente VB6: Call LEEEXPLO(CODD$, 1, 1) - lee fórmula desde tabla FORMULAS
                var componentes = await _context.Formulas
                    .Where(f => f.CodIConj == orden.Cod_Inte.Value)
                    .ToListAsync();

                if (!componentes.Any())
                {
                    _logger.LogWarning($"No se encontraron componentes en la fórmula para producto {orden.Cod_Inte.Value}");
                    return;
                }

                // Paso 6: Regenerar reservas para la cantidad pendiente
                // Equivalente VB6: For U& = 1 To CAIT% ... genero nuevamente la reserva
                short numeroItem = 0;
                short codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa() ?? 0;

                foreach (var componente in componentes)
                {
                    // Verificar que el componente tenga código interno válido
                    if (!componente.CodiElem.HasValue || componente.CodiElem.Value <= 0)
                    {
                        continue;
                    }

                    // Calcular la cantidad necesaria (equivalente a USO * CAN)
                    // En VB6: CANTRX# = USO * CAN
                    decimal cantidadReserva = cantidadPendiente * (componente.UsoBruto ?? 0);

                    // Solo crear reserva si la cantidad es significativa
                    if (Math.Abs(cantidadReserva) >= 0.05m)
                    {
                        // Incrementar el contador de ítems (equivalente a NORIT%)
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
                            IdSubOr = orden.IdSubOr,
                            FechRes = DateTime.Now,
                            CantRes = cantidadReserva,
                            CantCons = 0,
                            Saldo_Entre = cantidadReserva, // Saldo inicial es igual a la cantidad reservada
                            IdenLote = " ",
                            Deposito = 0,
                            UsoUnit = componente.UsoBruto,
                            NuOrItem = numeroItem,
                            Nume_Solint = 0,
                            Marborra = 0
                        };

                        // Agregar la reserva a la base de datos
                        _context.Reservas.Add(nuevaReserva);
                    }
                }

                // Guardar todos los cambios
                await _context.SaveChangesAsync();
                _logger.LogInformation($"Se regeneraron {numeroItem} reservas para orden {orden.IdSubOr} (cantidad pendiente: {cantidadPendiente})");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al regenerar reservas para orden {orden.IdSubOr}");
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
        /// <returns>Resultado de la operación (true = éxito, false = error)</returns>
        /// 
        /// NOTA IMPORTANTE SOBRE VINCULACIÓN DE MOVIMIENTOS:
        /// Este método guarda movimientos en la tabla ORFAPEN (no en MOVISTO).
        /// Por lo tanto, retorna un bool (éxito/error) y no el ID del movimiento generado.
        /// 
        /// Esto significa que no es posible vincular movimientos CF con el OF mediante IdenMovi
        /// cuando se usa este método (modo de control de calidad).
        /// 
        /// En modo estándar, los movimientos OF se guardan en MOVISTO y se captura el ID
        /// para vincularlo con los movimientos CF mediante IdenMovi.
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

        /// <summary>
        /// Obtiene los cierres reversibles de una orden de fabricación
        /// 
        /// Este endpoint busca en la tabla MOVISTO los movimientos de tipo 'OF' (Orden de Fabricación)
        /// que tienen cantidad ingresada mayor a cero. Estos son los cierres que pueden ser anulados.
        /// 
        /// Equivalente al código VB6:
        /// SQLX$ = " SELECT DoPriCor, FECHMOV, CANTINGRE, FeReMovi,IDENLOTE, ID FROM MOVISTO WITH(NOLOCK)"
        /// SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
        /// SQLX$ = SQLX$ + "AND CodiMovi = 'OF' "
        /// SQLX$ = SQLX$ + "AND CantIngre > 0 "
        /// SQLX$ = SQLX$ + "ORDER BY FeReMovi ASC "
        /// 
        /// Flujo:
        /// 1. Verifica que el usuario esté autenticado
        /// 2. Consulta MOVISTO filtrando por código interno, tipo 'OF' y CantIngre > 0
        /// 3. Ordena por fecha de registro ascendente (más antiguos primero)
        /// 4. Retorna lista de DTOs para mostrar en el modal de anulación
        /// </summary>
        /// <param name="codInte">Código interno del producto (OrdeFabr.Cod_Inte)</param>
        /// <param name="idSubOr">ID de la suborden de fabricación (OrdeFabr.IdSubOr)</param>
        /// <returns>Lista de cierres reversibles</returns>
        [HttpGet("cierres-reversibles/{codInte:int}/{idSubOr}")]
        public async Task<ActionResult<List<CierreReversibleDTO>>> GetCierresReversibles(int codInte, string idSubOr)
        {
            try
            {
                // Verificar que el usuario tenga una sesión válida
                // Si no está autenticado, no puede realizar ninguna operación
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Consulta LINQ equivalente a la consulta SQL del código VB6 original
                // SELECT * FROM MOVISTO WHERE CODIMOVI = 'OF' AND COD_INTE = ORDEFABR.COD_INTE AND CANTINGRE > 0
                // 
                // Agregamos filtro por DoPriCor (IdSubOr) para mostrar solo los cierres de esta orden específica
                // DoPriCor contiene el IdSubOr de la orden de fabricación que generó el movimiento
                // 
                // AsNoTracking: mejora rendimiento porque no vamos a modificar estos registros
                // Where: filtra por código interno, tipo de movimiento 'OF', cantidad ingresada > 0, y orden específica
                // OrderBy: ordena por fecha de registro para mostrar los cierres más antiguos primero
                var cierres = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.Cod_Inte == codInte
                               && m.CodiMovi == "OF"
                               && m.CantIngre > 0
                               && m.DoPriCor == idSubOr)
                    .OrderBy(m => m.FeReMovi)
                    .Select(m => new CierreReversibleDTO
                    {
                        Id = m.Id,
                        DoPriCor = m.DoPriCor ?? "",
                        FechMov = m.FechMov,
                        CantIngre = m.CantIngre ?? 0,
                        FeReMovi = m.FeReMovi,
                        IdenLote = m.IdenLote
                    })
                    .ToListAsync();

                _logger.LogInformation($"Se encontraron {cierres.Count} cierres reversibles para código interno {codInte}");

                return Ok(cierres);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener cierres reversibles para código interno {codInte}");
                return StatusCode(500, $"Error al obtener cierres reversibles: {ex.Message}");
            }
        }

        /// <summary>
        /// Anula un cierre específico de orden de fabricación
        /// 
        /// Este endpoint realiza la anulación de un cierre parcial de orden de fabricación siguiendo la lógica del código VB6.
        /// La anulación implica varios pasos que deben ejecutarse en una transacción para mantener la integridad de datos.
        /// 
        /// Flujo completo de anulación:
        /// 1. Inicia una transacción de base de datos (todo o nada)
        /// 2. Verifica autenticación del usuario
        /// 3. Obtiene el movimiento OF a anular y valida que tenga cantidad > 0
        /// 4. Actualiza el movimiento OF: pone CantIngre=0, CantSalid=0, ReferCor="Anul. "+original
        /// 5. Anula movimientos CF/SC asociados: busca por IdenMovi que coincide con el ID del OF
        /// 6. Actualiza OrdeFabr: resta la cantidad anulada de CanApro, pone StatuOrf=1 (en curso)
        /// 7. Genera anotación en OrdeFabr.AnotaOrf con formato "Anulación de Cierre x {cantidad}"
        /// 8. Confirma la transacción si todo fue exitoso
        /// 
        /// Equivalente al código VB6:
        /// - UPDATE MOVISTO SET ReferCor='Anul.'+ReferCor, CANTINGRE=0, CANTSALID=0 WHERE ID=idMov& AND CODIMOVI='OF'
        /// - UPDATE MOVISTO SET CANTINGRE=0, CANTSALID=0 WHERE IDENMOVI=idMov& AND (CODIMOVI='CF' OR CODIMOVI='SC')
        /// - UPDATE OrdeFabr SET CanApro=CanApro-CANUL, StatuOrf=1 WHERE IdSubOr=NOFASE$
        /// - Genera anotación en AnotaOrf
        /// </summary>
        /// <param name="idMovimiento">ID del movimiento a anular (ID en tabla MOVISTO)</param>
        /// <param name="idSubOr">ID de la suborden de fabricación (opcional, para validación)</param>
        /// <param name="regenerarReservas">Indica si se deben regenerar las reservas (true) o mantenerlas (false)</param>
        /// <returns>Resultado de la operación con cantidad anulada y movimientos CF afectados</returns>
        [HttpPut("anular-cierre/{idMovimiento:int}")]
        public async Task<IActionResult> AnularCierre(int idMovimiento, [FromQuery] string idSubOr, [FromQuery] bool regenerarReservas = true)
        {
            // Iniciar transacción de base de datos
            // Usamos ReadCommitted para balance entre consistencia y rendimiento
            // Si algo falla, se hace rollback automático en el catch
            await using var transaction = await _context.Database.BeginTransactionAsync(IsolationLevel.ReadCommitted);
            try
            {
                // Paso 1: Verificar que el usuario tenga una sesión válida
                // Sin autenticación, no se permite ninguna operación
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Paso 2: Obtener el movimiento OF a anular
                // AsNoTracking: no necesitamos tracking porque vamos a hacer update directo con ExecuteUpdateAsync
                // Filtramos por ID y por CodiMovi='OF' para asegurar que es un movimiento de orden de fabricación
                var movimientoOF = await _context.Movistos
                    .AsNoTracking()
                    .FirstOrDefaultAsync(m => m.Id == idMovimiento && m.CodiMovi == "OF");

                if (movimientoOF == null)
                {
                    return NotFound($"No se encontró el movimiento OF con ID {idMovimiento}");
                }

                // Paso 3: Validar que el movimiento tenga cantidad ingresada
                // Solo se pueden anular cierres que realmente agregaron stock (CantIngre > 0)
                if (movimientoOF.CantIngre <= 0)
                {
                    return BadRequest("El movimiento no tiene cantidad ingresada para anular");
                }

                // Guardar datos importantes para usar en las actualizaciones
                decimal cantidadAnular = movimientoOF.CantIngre ?? 0;
                string dopriCor = movimientoOF.DoPriCor ?? "";

                // Paso 4: Anular el movimiento OF del producto fabricado
                // Esto pone las cantidades en cero y marca la referencia como "Anul."
                // ExecuteUpdateAsync: hace UPDATE directo en SQL sin cargar entidades en memoria (más eficiente)
                // Equivalente VB6: UPDATE MOVISTO SET ReferCor='Anul.'+ReferCor, CANTINGRE=0, CANTSALID=0, PESOKG=0, METROS=0 WHERE ID=idMov& AND CODIMOVI='OF'
                var filasOF = await _context.Movistos
                    .Where(m => m.DoPriCor == dopriCor && m.Id == idMovimiento && m.CodiMovi == "OF")
                    .ExecuteUpdateAsync(s => s
                        .SetProperty(m => m.ReferCor, "Anul. " + (movimientoOF.ReferCor ?? ""))
                        .SetProperty(m => m.CantIngre, 0)
                        .SetProperty(m => m.CantSalid, 0)
                        .SetProperty(m => m.PESOKG, (decimal?)null)
                        .SetProperty(m => m.METROS, (decimal?)null));

                // Si no se actualizó ninguna fila, algo salió mal
                if (filasOF == 0)
                {
                    await transaction.RollbackAsync();
                    return Conflict("No se pudo actualizar el movimiento OF");
                }

                _logger.LogInformation($"Movimiento OF {idMovimiento} actualizado a cero");

                // Paso 5: Anular los movimientos de consumo (CF) asociados
                // Estos movimientos tienen IdenMovi que coincide con el ID del movimiento OF
                // Esto es importante porque los consumos se generaron durante el mismo cierre
                // CF = Consumo de Fabricación, SC = Salida de Consumo
                // Equivalente VB6: UPDATE MOVISTO SET CANTINGRE=0, CANTSALID=0, METROS=0, PESOKG=0 WHERE IDENMOVI=idMov& AND (CODIMOVI='CF' OR CODIMOVI='SC')
                var filasCF = await _context.Movistos
                    .Where(m => m.DoPriCor == dopriCor 
                               && m.IdenMovi == idMovimiento.ToString()
                               && (m.CodiMovi == "CF" || m.CodiMovi == "SC"))
                    .ExecuteUpdateAsync(s => s
                        .SetProperty(m => m.CantIngre, 0)
                        .SetProperty(m => m.CantSalid, 0)
                        .SetProperty(m => m.METROS, (decimal?)null)
                        .SetProperty(m => m.PESOKG, (decimal?)null));

                _logger.LogInformation($"Se anularon {filasCF} movimientos CF/SC asociados al movimiento OF {idMovimiento}");

                // Paso 6: Actualizar la orden de fabricación
                // Restar la cantidad anulada de la cantidad aprobada
                // Cambiar el estado a 1 (en curso) porque al anular un cierre, la orden ya no está completamente cerrada
                // Equivalente VB6: SELECT * FROM OrdeFabr WHERE IdSubOr=NOFASE$; !CanApro=!CanApro-CANUL; !StatuOrf=1
                var orden = await _context.Ordefabrs
                    .FirstOrDefaultAsync(o => o.IdSubOr == dopriCor);

                if (orden != null)
                {
                    // Calcular nueva cantidad aprobada restando la cantidad anulada
                    // Si el resultado es negativo, poner en 0 (no puede haber cantidad negativa)
                    decimal nuevaCantidadAprobada = (orden.CanApro ?? 0) - cantidadAnular;
                    if (nuevaCantidadAprobada < 0) nuevaCantidadAprobada = 0;

                    orden.CanApro = nuevaCantidadAprobada;
                    orden.StatuOrf = 1; // Marca como en curso (ya no está cerrada)

                    // Paso 7: Regenerar reservas según código VB6 GENERES_ANUCIEOF (solo si el usuario lo solicitó)
                    // Lógica VB6:
                    // - DELETE FROM RESERVA WHERE IdSubOr = NORFA$
                    // - CAN = CantiProyectadaOF(NORFA$) - CantiAprobadaOF(NORFA$)
                    // - Si CAN > 0, regenerar reservas para la cantidad pendiente
                    // 
                    // Esto es necesario porque al anular un cierre, se reduce la cantidad aprobada
                    // y las reservas deben ajustarse a la nueva cantidad pendiente (proyectada - aprobada)
                    // 
                    // Si el usuario elige mantener las reservas, no se regeneran y podrá editarlas manualmente
                    if (regenerarReservas)
                    {
                        await RegenerarReservasTrasAnulacionCierre(orden);
                    }

                    // Paso 8: Generar anotación en la orden
                    // El formato sigue el patrón existente de anotaciones del sistema
                    // Incluye fecha, descripción del evento y nombre del usuario
                    // Informa si se regeneraron las reservas o se mantuvieron
                    string nombreUsuario = ObtenerNombreUsuario();
                    string anotacionAnulacion = GenerarTextoAnotaOrfAnulacionCierre(orden, cantidadAnular, nombreUsuario, regenerarReservas);

                    // Si ya hay anotaciones, agregar la nueva al final con salto de línea
                    // Si no hay anotaciones, asignar directamente
                    if (string.IsNullOrEmpty(orden.AnotaOrf))
                    {
                        orden.AnotaOrf = anotacionAnulacion;
                    }
                    else
                    {
                        orden.AnotaOrf += "\r\n" + anotacionAnulacion;
                    }

                    // Guardar cambios en la orden
                    _context.Ordefabrs.Update(orden);
                    await _context.SaveChangesAsync();

                    string logReservas = regenerarReservas ? "Reservas regeneradas" : "Reservas mantenidas";
                    _logger.LogInformation($"Orden {orden.IdSubOr} actualizada: CanApro={nuevaCantidadAprobada}, StatuOrf=1, {logReservas}");
                }
                else
                {
                    _logger.LogWarning($"No se encontró la orden de fabricación con IdSubOr {dopriCor}");
                }

                // Paso 8: Confirmar la transacción
                // Solo llega aquí si no hubo errores en ningún paso anterior
                await transaction.CommitAsync();

                // Retornar respuesta con información de lo que se anuló
                return Ok(new { 
                    mensaje = $"Cierre anulado exitosamente. Cantidad: {cantidadAnular}",
                    cantidadAnulada = cantidadAnular,
                    movimientosCFAnulados = filasCF
                });
            }
            catch (Exception ex)
            {
                // Si hubo error, hacer rollback de la transacción
                // Esto asegura que no quede datos parcialmente actualizados
                await transaction.RollbackAsync();
                _logger.LogError(ex, $"Error al anular cierre {idMovimiento}");
                return StatusCode(500, $"Error al anular cierre: {ex.Message}");
            }
        }

        /// <summary>
        /// Genera el texto de anotación para anulación de cierre
        /// 
        /// Este método genera una anotación en formato posicional estilo VB6 que se agrega
        /// al campo OrdeFabr.AnotaOrf. El formato sigue el patrón existente de anotaciones
        /// del sistema para mantener consistencia.
        /// 
        /// Formato de la anotación:
        /// - Posición 1-8: Fecha en formato dd/MM/yy
        /// - Posición 11-66: Descripción del evento ("Anulación de Cierre x {cantidad}")
        /// - Posición 69-92: Nombre del usuario que realizó la anulación
        /// - Línea separadora: 88 caracteres '='
        /// 
        /// Ejemplo de resultado:
        /// "10/07/26  Anulación de Cierre x 100.0                    JUAN PEREZ"
        /// "========================================================================================"
        /// 
        /// Este formato posicional se usa para compatibilidad con el sistema VB6 original
        /// y para mantener un formato consistente en todas las anotaciones del sistema.
        /// </summary>
        /// <param name="orden">Orden de fabricación (no se usa directamente pero se mantiene por consistencia)</param>
        /// <param name="cantidadAnulada">Cantidad que se anuló del cierre</param>
        /// <param name="nombreUsuario">Nombre del usuario que realizó la anulación</param>
        /// <param name="regeneroReservas">Indica si se regeneraron las reservas (true por defecto)</param>
        /// <returns>Texto de anotación formateado para agregar a OrdeFabr.AnotaOrf</returns>
        private string GenerarTextoAnotaOrfAnulacionCierre(OrdeFabr orden, decimal cantidadAnulada, string nombreUsuario, bool regeneroReservas = false)
        {
            // Crear una línea de 76 espacios en blanco como base
            // Esta línea se va llenando con datos en posiciones específicas
            string ttxyz = new string(' ', 76);
            
            // Obtener la fecha actual en formato dd/MM/yy (ej: 10/07/26)
            string hos = DateTime.Now.ToString("dd/MM/yy");
            
            // Insertar la fecha en la posición 1 (longitud 8 caracteres)
            // Repla es un método auxiliar que reemplaza texto en una posición específica
            ttxyz = Repla(ttxyz, hos, 1, 8);
            
            // Insertar la descripción en la posición 11 (longitud 56 caracteres)
            // La descripción incluye la cantidad anulada con 1 decimal
            // Informa si se regeneraron las reservas o se mantuvieron
            string descripcion;
            if (regeneroReservas)
            {
                descripcion = $"Anulación de Cierre x {cantidadAnulada:F1} - Reservas regeneradas";
            }
            else
            {
                descripcion = $"Anulación de Cierre x {cantidadAnulada:F1} - Reservas mantenidas";
            }
            ttxyz = Repla(ttxyz, descripcion, 11, 56);
            
            // Insertar el nombre del usuario en la posición 69 (longitud 24 caracteres)
            ttxyz = Repla(ttxyz, nombreUsuario, 69, 24);
            
            // Agregar salto de línea al final de la línea de datos
            ttxyz += "\r\n";
            
            // Crear línea separadora de 88 caracteres '='
            // Esta línea separa las anotaciones para mejor legibilidad
            string separa = new string('=', 88) + "\r\n";
            
            // Retornar la línea de datos + la línea separadora
            return ttxyz + separa;
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