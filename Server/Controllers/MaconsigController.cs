using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using System.Data;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class MaconsigController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<MaconsigController> _logger;
        private readonly ConsignacionService _consignacionService;
        private readonly IUsuarioFlexoftService _usuarioFlexoftService;

        public MaconsigController(
            DataContext context,
            ILogger<MaconsigController> logger,
            ConsignacionService consignacionService,
            IUsuarioFlexoftService usuarioFlexoftService)
        {
            _context = context;
            _logger = logger;
            _consignacionService = consignacionService;
            _usuarioFlexoftService = usuarioFlexoftService;
        }

        [HttpGet("terceros-pendientes-por-producto/{codint:int}")]
        public async Task<ActionResult<List<StockTercerosPendienteDTO>>> GetTercerosPendientesPorProducto(int codint)
        {
            try
            {
                var agrupado = await (from m in _context.Maconsigs.AsNoTracking()
                                      where (m.CodInte ?? 0) == codint
                                            && m.NroOcom != null
                                      group m by new
                                      {
                                          Ocom = m.NroOcom.Length >= 9 ? m.NroOcom.Substring(0, 9) : m.NroOcom,
                                          NproOcom = (int)(m.NproOcom ?? 0)
                                      }
                    into g
                                      select new
                                      {
                                          g.Key.Ocom,
                                          g.Key.NproOcom,
                                          Entregado = g.Sum(x => x.CantEntreg ?? 0m),
                                          Rendido = g.Sum(x => x.CantRendid ?? 0m),
                                          Devuelto = g.Sum(x => x.CantDevu ?? 0m),
                                          Asignado = g.Sum(x => x.CantAsig ?? 0m)
                                      })
                    .ToListAsync();

                var numerosOcom = agrupado
                    .Select(x =>
                    {
                        if (string.IsNullOrWhiteSpace(x.Ocom) || x.Ocom.Length < 9)
                        {
                            return (int?)null;
                        }

                        if (int.TryParse(x.Ocom.Substring(3, 6), out var num))
                        {
                            return (int?)num;
                        }

                        return (int?)null;
                    })
                    .Where(x => x.HasValue)
                    .Select(x => x!.Value)
                    .Distinct()
                    .ToList();

                var ocomEncas = await _context.Ocomencas
                    .AsNoTracking()
                    .Where(x => x.NumeOcom.HasValue && numerosOcom.Contains(x.NumeOcom.Value))
                    .Select(x => new
                    {
                        NumeOcom = x.NumeOcom ?? 0,
                        x.StatOcom,
                        x.FemiOcom
                    })
                    .ToListAsync();

                var ocomEncaByNumero = ocomEncas
                    .GroupBy(x => x.NumeOcom)
                    .ToDictionary(g => g.Key, g => g.First());

                var proveedores = await _context.Proveedores
                    .AsNoTracking()
                    .Select(p => new { p.Nume_Cli, p.Raso_Cli })
                    .ToListAsync();

                var provByNumero = proveedores
                    .GroupBy(p => p.Nume_Cli)
                    .ToDictionary(g => g.Key, g => g.First().Raso_Cli ?? string.Empty);

                var result = new List<StockTercerosPendienteDTO>();

                foreach (var x in agrupado)
                {
                    var capen = x.Entregado - x.Rendido - x.Devuelto;
                    var capenEnt = x.Asignado - x.Entregado;

                    short? statOcom = null;
                    DateTime? femiOcom = null;

                    if (!string.IsNullOrWhiteSpace(x.Ocom) && x.Ocom.Length >= 9 && int.TryParse(x.Ocom.Substring(3, 6), out var numOcom))
                    {
                        if (ocomEncaByNumero.TryGetValue(numOcom, out var enca))
                        {
                            statOcom = enca.StatOcom;
                            femiOcom = enca.FemiOcom;
                        }
                    }

                    result.Add(new StockTercerosPendienteDTO
                    {
                        Ocom = x.Ocom ?? string.Empty,
                        NproOcom = x.NproOcom,
                        Proveedor = provByNumero.TryGetValue(x.NproOcom, out var nombre) ? nombre : string.Empty,
                        Entregado = x.Entregado,
                        Rendido = x.Rendido,
                        Devuelto = x.Devuelto,
                        Asignado = x.Asignado,
                        CapenRendir = capen < 0m ? 0m : capen,
                        CapenEntregar = capenEnt,
                        StatOcom = statOcom,
                        FechaEmisionOcom = femiOcom
                    });
                }

                return Ok(result
                    .OrderByDescending(x => x.CapenRendir)
                    .ThenBy(x => x.Ocom)
                    .ToList());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetTercerosPendientesPorProducto] Error al obtener terceros pendientes por producto {Codint}", codint);
                return StatusCode(500, $"Error al obtener terceros pendientes: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene materiales en consignación con saldo pendiente
        /// Filtro: Cant_Rendid < Cant_Entreg - 0.001
        /// </summary>
        [HttpGet("materiales-consignacion")]
        public async Task<ActionResult<List<Maconsig>>> GetMaterialesConsignacionAsync(
            [FromQuery] PaginationDTO pagination,
            string? filtroProveedor = null,
            string? filtroOrdenCompra = null,
            string? filtroArticulo = null,
            string? filtroOrdenFab = null)
        {
            try
            {
                _logger.LogInformation("[GetMaterialesConsignacion] Obteniendo materiales en consignación");

                var page = pagination.Page <= 0 ? 1 : pagination.Page;
                var recordsNumber = pagination.RecordsNumber <= 0 ? 10 : pagination.RecordsNumber;

                var queryable = ConstruirQueryMaterialesConsignacionFiltrada(
                    filtroProveedor,
                    filtroOrdenCompra,
                    filtroArticulo,
                    filtroOrdenFab);

                var materiales = await queryable
                    .OrderByDescending(m => m.Id)
                    .Skip((page - 1) * recordsNumber)
                    .Take(recordsNumber)
                    .ToListAsync();

                _logger.LogInformation($"[GetMaterialesConsignacion] Se encontraron {materiales.Count} materiales en consignación");

                return Ok(materiales);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[GetMaterialesConsignacion] Error al obtener materiales en consignación. Se devolverá lista vacía.");
                return Ok(new List<Maconsig>());
            }
        }

        [HttpGet("materiales-consignacion/totalPages")]
        public async Task<ActionResult> GetMaterialesConsignacionTotalPages(
            [FromQuery] PaginationDTO pagination,
            string? filtroProveedor = null,
            string? filtroOrdenCompra = null,
            string? filtroArticulo = null,
            string? filtroOrdenFab = null)
        {
            try
            {
                var recordsNumber = pagination.RecordsNumber <= 0 ? 10 : pagination.RecordsNumber;

                var queryable = ConstruirQueryMaterialesConsignacionFiltrada(
                    filtroProveedor,
                    filtroOrdenCompra,
                    filtroArticulo,
                    filtroOrdenFab);

                var count = await queryable.CountAsync();
                var totalPages = Math.Ceiling((double)count / recordsNumber);

                return Ok(totalPages <= 0 ? 1 : totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[GetMaterialesConsignacionTotalPages] Error obteniendo total de páginas. Se devolverá 1.");
                return Ok(1);
            }
        }

        private IQueryable<Maconsig> ConstruirQueryMaterialesConsignacionFiltrada(
            string? filtroProveedor,
            string? filtroOrdenCompra,
            string? filtroArticulo,
            string? filtroOrdenFab)
        {
            var queryable = _context.Maconsigs
                .AsNoTracking()
                .Where(m => m.CantRendid < m.CantEntreg - 0.001m)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(filtroProveedor))
            {
                var filtro = filtroProveedor.Trim();
                queryable = queryable.Where(m => m.RaSoProv != null && m.RaSoProv.Contains(filtro));
            }

            if (!string.IsNullOrWhiteSpace(filtroOrdenCompra))
            {
                var filtro = filtroOrdenCompra.Trim();
                queryable = queryable.Where(m => m.NroOcom != null && m.NroOcom.Contains(filtro));
            }

            if (!string.IsNullOrWhiteSpace(filtroArticulo))
            {
                var filtro = filtroArticulo.Trim();
                queryable = queryable.Where(m =>
                    (m.CodExte != null && m.CodExte.Contains(filtro)) ||
                    (m.Descritem != null && m.Descritem.Contains(filtro)));
            }

            if (!string.IsNullOrWhiteSpace(filtroOrdenFab))
            {
                var filtro = filtroOrdenFab.Trim();
                queryable = queryable.Where(m => m.NroOrfa != null && m.NroOrfa.Contains(filtro));
            }

            return queryable;
        }

        /// <summary>
        /// Borra consignaciones sin entregas de MACONSIG
        /// Equivalente a: DELETE FROM MACONSIG WHERE NRO_OCOM LIKE 'OC-000001%' AND Cant_Entreg = 0
        /// </summary>
        [HttpPost("borrar-consignaciones")]
        public async Task<IActionResult> BorrarConsignaciones([FromBody] BorrarConsignacionesRequest request)
        {
            try
            {
                _logger.LogInformation($"[BorrarConsignaciones] Procesando orden: {request.NumeroOrden}");
                _logger.LogInformation($"[BorrarConsignaciones] Condición: {request.Condicion}");

                // Construir la consulta SQL para borrar consignaciones
                var sql = $"DELETE FROM MACONSIG WHERE NRO_OCOM LIKE '{request.CodigoFormateado}%' AND Cant_Entreg = 0";

                _logger.LogInformation($"[BorrarConsignaciones] Ejecutando SQL: {sql}");

                // Ejecutar la consulta
                var registrosAfectados = await _context.Database.ExecuteSqlRawAsync(sql);

                _logger.LogInformation($"[BorrarConsignaciones] Registros eliminados: {registrosAfectados}");

                return Ok(new
                {
                    Success = true,
                    Message = $"Consignaciones eliminadas exitosamente. Registros afectados: {registrosAfectados}",
                    RegistrosAfectados = registrosAfectados
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[BorrarConsignaciones] Error al borrar consignaciones para orden {request.NumeroOrden}");

                // Si no hay registros para borrar, no es un error
                if (ex.Message.Contains("no rows affected") || ex.Message.Contains("0 registros"))
                {
                    return Ok(new
                    {
                        Success = true,
                        Message = "No hay consignaciones para borrar",
                        RegistrosAfectados = 0
                    });
                }

                return StatusCode(500, new
                {
                    Success = false,
                    Message = "Ha ocurrido un error inesperado, tener en cuenta que puede ser que no haya consignaciones para dicha orden de compra",
                    Error = ex.Message
                });
            }
        }

        /// <summary>
        /// Procesa las consignaciones para una orden de compra según la lógica VB6
        /// Implementa la lógica de remanente: CANTIASIGN# = CANTIASIGN# - CANYASIG#
        /// 
        /// USO: TANTO EN ALTA COMO EN EDICIÓN DE ÓRDENES DE COMPRA
        /// - ALTA: Se ejecuta desde OcomencaController.PostAsync() dentro de la transacción
        /// - EDICIÓN: Se ejecuta en el PASO 6 del flujo de edición (OrdenesCompraEditForm.razor)
        /// - Calcula remanente considerando consignaciones ya existentes
        /// - Solo crea registros si remanente > 0.05 y componente tiene CodXElem
        /// </summary>
        [HttpPost("procesar-orden")]
        public async Task<IActionResult> ProcesarConsignacionesOrdenAsync([FromBody] ProcesarConsignacionesRequest request)
        {
            try
            {
                _logger.LogInformation($"[ProcesarConsignaciones] 🔄 Procesando consignaciones para O.C. {request.NumeroOrden}");
                _logger.LogInformation($"[ProcesarConsignaciones] Proveedor: {request.NumeroProveedor} - {request.RazonSocialProveedor}");
                _logger.LogInformation($"[ProcesarConsignaciones] Fecha: {request.FechaEmision:yyyy-MM-dd}, Moneda: {request.MonedaEmision}");

                // Usar transacción para garantizar integridad
                using var transaction = await _context.Database.BeginTransactionAsync();
                try
                {
                    var resultado = await _consignacionService.ProcesarConsignacionesOrdenCompraAsync(
                        request.NumeroOrden,
                        request.NumeroProveedor,
                        request.RazonSocialProveedor,
                        request.FechaEmision,
                        request.MonedaEmision
                    );

                    // ===== GUARDAR CAMBIOS EN LA BASE DE DATOS =====
                    if (resultado)
                    {
                        await _context.SaveChangesAsync();
                        await transaction.CommitAsync();
                        _logger.LogInformation($"[ProcesarConsignaciones] 💾 Consignaciones guardadas y confirmadas en BD para O.C. {request.NumeroOrden}");
                    }
                    else
                    {
                        await transaction.CommitAsync(); // Commit aunque no haya consignaciones
                    }

                    var response = new
                    {
                        Success = resultado,
                        Message = resultado
                            ? "Consignaciones procesadas y guardadas exitosamente"
                            : "No se encontraron items con estructura para procesar",
                        NumeroOrden = request.NumeroOrden
                    };

                    _logger.LogInformation($"[ProcesarConsignaciones] ✅ Resultado: {response.Message}");
                    return Ok(response);
                }
                catch (Exception)
                {
                    await transaction.RollbackAsync();
                    throw; // Re-lanzar para que lo capture el catch externo
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[ProcesarConsignaciones] ❌ Error procesando consignaciones para O.C. {request.NumeroOrden}");
                return StatusCode(500, new
                {
                    Success = false,
                    Message = $"Error al procesar consignaciones: {ex.Message}",
                    NumeroOrden = request.NumeroOrden
                });
            }
        }

        [HttpPost("validar-edicion-orden")]
        public async Task<ActionResult<ValidacionEdicionConsignacionResponse>> ValidarEdicionOrden([FromBody] ValidarEdicionConsignacionRequest request)
        {
            try
            {
                if (request.NumeroOrden <= 0)
                {
                    return BadRequest("El número de orden debe ser mayor a 0.");
                }

                if (request.Items == null || !request.Items.Any())
                {
                    return Ok(new ValidacionEdicionConsignacionResponse
                    {
                        PuedeContinuar = true,
                        Mensaje = "No hay items válidos para validar."
                    });
                }

                var itemsValidos = request.Items
                    .Where(i => i.CodIntePadre > 0 && i.Cantidad > 0.0001m)
                    .ToList();

                if (!itemsValidos.Any())
                {
                    return Ok(new ValidacionEdicionConsignacionResponse
                    {
                        PuedeContinuar = true,
                        Mensaje = "No hay items activos con cantidad mayor a 0 para validar."
                    });
                }

                var codigosPadre = itemsValidos.Select(i => i.CodIntePadre).Distinct().ToList();
                var formulas = await _context.Formulas
                    .AsNoTracking()
                    .Where(f => f.CodIConj.HasValue
                                && codigosPadre.Contains(f.CodIConj.Value)
                                && f.CodiElem.HasValue
                                && (f.Marborra ?? 0) != 1)
                    .Select(f => new
                    {
                        CodPadre = f.CodIConj!.Value,
                        CodComponente = f.CodiElem!.Value,
                        UsoBruto = f.UsoBruto ?? 0m
                    })
                    .ToListAsync();

                if (!formulas.Any())
                {
                    return Ok(new ValidacionEdicionConsignacionResponse
                    {
                        PuedeContinuar = true,
                        Mensaje = "Los items no poseen estructura de consignación."
                    });
                }

                var necesidadPorComponente = new Dictionary<int, decimal>();
                var itemsPorComponente = new Dictionary<int, HashSet<short>>();

                foreach (var item in itemsValidos)
                {
                    var componentes = formulas.Where(f => f.CodPadre == item.CodIntePadre).ToList();
                    foreach (var comp in componentes)
                    {
                        var necesidad = Math.Round(item.Cantidad * comp.UsoBruto, 4);
                        if (necesidad <= 0.0001m)
                        {
                            continue;
                        }

                        if (!necesidadPorComponente.ContainsKey(comp.CodComponente))
                        {
                            necesidadPorComponente[comp.CodComponente] = 0m;
                        }

                        necesidadPorComponente[comp.CodComponente] += necesidad;

                        if (!itemsPorComponente.ContainsKey(comp.CodComponente))
                        {
                            itemsPorComponente[comp.CodComponente] = new HashSet<short>();
                        }

                        itemsPorComponente[comp.CodComponente].Add(item.NumeroItem);
                    }
                }

                if (!necesidadPorComponente.Any())
                {
                    return Ok(new ValidacionEdicionConsignacionResponse
                    {
                        PuedeContinuar = true,
                        Mensaje = "No hay componentes de consignación para validar."
                    });
                }

                var prefijoOrden = $"OC-{request.NumeroOrden:000000}-";

                var entregasPorComponente = await _context.Maconsigs
                    .AsNoTracking()
                    .Where(m => m.NroOcom != null
                                && m.NroOcom.StartsWith(prefijoOrden)
                                && m.CodInte.HasValue
                                && necesidadPorComponente.Keys.Contains(m.CodInte.Value))
                    .GroupBy(m => m.CodInte!.Value)
                    .Select(g => new
                    {
                        CodComponente = g.Key,
                        Entregado = g.Sum(x => x.CantEntreg ?? 0m),
                        Rendido = g.Sum(x => x.CantRendid ?? 0m)
                    })
                    .ToListAsync();

                if (!entregasPorComponente.Any())
                {
                    return Ok(new ValidacionEdicionConsignacionResponse
                    {
                        PuedeContinuar = true,
                        Mensaje = "No existen entregas registradas en consignación para los componentes a validar."
                    });
                }

                var codigosComponentes = entregasPorComponente.Select(x => x.CodComponente).Distinct().ToList();
                var masterByCod = await _context.masters
                    .AsNoTracking()
                    .Where(m => codigosComponentes.Contains(m.Codint))
                    .Select(m => new { m.Codint, m.Codigo, m.Descripcion })
                    .ToDictionaryAsync(x => x.Codint, x => x);

                var conflictos = new List<ConflictoEdicionConsignacionItemDTO>();

                foreach (var entrega in entregasPorComponente)
                {
                    var cantidadNecesaria = Math.Round(necesidadPorComponente.GetValueOrDefault(entrega.CodComponente, 0m), 4);
                    var cantidadEntregada = Math.Round(entrega.Entregado, 4);
                    var cantidadRendida = Math.Round(entrega.Rendido, 4);

                    if (cantidadEntregada <= cantidadNecesaria + 0.0001m)
                    {
                        continue;
                    }

                    var exceso = Math.Round(cantidadEntregada - cantidadNecesaria, 4);
                    var itemsAfectados = itemsPorComponente.TryGetValue(entrega.CodComponente, out var items)
                        ? string.Join(", ", items.OrderBy(x => x).Select(x => x.ToString("00")))
                        : "-";
                    var itemsAfectadosNumeros = itemsPorComponente.TryGetValue(entrega.CodComponente, out var itemsNumeros)
                        ? itemsNumeros.OrderBy(x => x).ToList()
                        : new List<short>();

                    masterByCod.TryGetValue(entrega.CodComponente, out var masterComp);

                    conflictos.Add(new ConflictoEdicionConsignacionItemDTO
                    {
                        CodInte = entrega.CodComponente,
                        CodExte = masterComp?.Codigo ?? string.Empty,
                        Descripcion = masterComp?.Descripcion ?? string.Empty,
                        CantidadNecesaria = cantidadNecesaria,
                        CantidadEntregada = cantidadEntregada,
                        CantidadRendida = cantidadRendida,
                        ExcesoEntregado = exceso,
                        ItemsAfectados = itemsAfectados,
                        ItemsAfectadosNumeros = itemsAfectadosNumeros
                    });
                }

                if (!conflictos.Any())
                {
                    return Ok(new ValidacionEdicionConsignacionResponse
                    {
                        PuedeContinuar = true,
                        Mensaje = "Validación correcta: las entregas actuales no superan las cantidades requeridas por la nueva edición.",
                        Conflictos = new List<ConflictoEdicionConsignacionItemDTO>()
                    });
                }

                return Ok(new ValidacionEdicionConsignacionResponse
                {
                    PuedeContinuar = false,
                    Mensaje = "No se puede continuar: existen componentes con entrega mayor a la necesaria para la nueva cantidad de la orden.",
                    Conflictos = conflictos
                        .OrderByDescending(x => x.ExcesoEntregado)
                        .ThenBy(x => x.CodExte)
                        .ThenBy(x => x.CodInte)
                        .ToList()
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[ValidarEdicionOrden] Error al validar edición de consignaciones para O.C. {NumeroOrden}", request?.NumeroOrden);
                return StatusCode(500, "Error al validar consignaciones para la edición de la orden.");
            }
        }

        [HttpPost("sincronizar-edicion-orden")]
        public async Task<ActionResult<SincronizarEdicionConsignacionResponse>> SincronizarEdicionOrden([FromBody] ValidarEdicionConsignacionRequest request)
        {
            try
            {
                if (request.NumeroOrden <= 0)
                {
                    return BadRequest("El número de orden debe ser mayor a 0.");
                }

                if (request.Items == null || !request.Items.Any())
                {
                    return Ok(new SincronizarEdicionConsignacionResponse
                    {
                        Success = true,
                        Mensaje = "Sin items para sincronizar.",
                        RegistrosAjustadosRendidos = 0,
                        RegistrosActualizadosIndice = 0
                    });
                }

                var itemsValidos = request.Items
                    .Where(i => i.CodIntePadre > 0 && i.Cantidad > 0.0001m)
                    .ToList();

                if (!itemsValidos.Any())
                {
                    return Ok(new SincronizarEdicionConsignacionResponse
                    {
                        Success = true,
                        Mensaje = "Sin items válidos para sincronizar.",
                        RegistrosAjustadosRendidos = 0,
                        RegistrosActualizadosIndice = 0
                    });
                }

                var codigosPadre = itemsValidos.Select(i => i.CodIntePadre).Distinct().ToList();
                var formulas = await _context.Formulas
                    .AsNoTracking()
                    .Where(f => f.CodIConj.HasValue
                                && codigosPadre.Contains(f.CodIConj.Value)
                                && f.CodiElem.HasValue
                                && (f.Marborra ?? 0) != 1)
                    .Select(f => new
                    {
                        CodPadre = f.CodIConj!.Value,
                        CodComponente = f.CodiElem!.Value
                    })
                    .ToListAsync();

                var componenteItems = new Dictionary<int, HashSet<short>>();
                foreach (var item in itemsValidos)
                {
                    var componentesPadre = formulas.Where(f => f.CodPadre == item.CodIntePadre).ToList();
                    foreach (var comp in componentesPadre)
                    {
                        if (!componenteItems.ContainsKey(comp.CodComponente))
                        {
                            componenteItems[comp.CodComponente] = new HashSet<short>();
                        }

                        componenteItems[comp.CodComponente].Add(item.NumeroItem);
                    }
                }

                if (!componenteItems.Any())
                {
                    return Ok(new SincronizarEdicionConsignacionResponse
                    {
                        Success = true,
                        Mensaje = "No hay componentes de consignación para sincronizar.",
                        RegistrosAjustadosRendidos = 0,
                        RegistrosActualizadosIndice = 0
                    });
                }

                var prefijoOrden = $"OC-{request.NumeroOrden:000000}-";
                var componentes = componenteItems.Keys.ToList();

                var registrosConEntrega = await _context.Maconsigs
                    .Where(m => m.NroOcom != null
                                && m.NroOcom.StartsWith(prefijoOrden)
                                && m.CodInte.HasValue
                                && componentes.Contains(m.CodInte.Value)
                                && (m.CantEntreg ?? 0m) > 0m)
                    .ToListAsync();

                int ajustesRendidos = 0;
                int indicesActualizados = 0;

                foreach (var reg in registrosConEntrega)
                {
                    var cantEntreg = reg.CantEntreg ?? 0m;
                    var cantRendid = reg.CantRendid ?? 0m;

                    if (cantRendid > cantEntreg)
                    {
                        reg.CantRendid = cantEntreg;
                        ajustesRendidos++;
                    }

                    if ((reg.CantAsig ?? 0m) < cantEntreg)
                    {
                        reg.CantAsig = cantEntreg;
                    }

                    if (!reg.CodInte.HasValue || !componenteItems.TryGetValue(reg.CodInte.Value, out var itemsComp))
                    {
                        continue;
                    }

                    if (itemsComp.Count == 1)
                    {
                        var item = itemsComp.First();
                        var nuevoNroOcom = $"OC-{request.NumeroOrden:000000}-{item:00}";
                        if (!string.Equals(reg.NroOcom, nuevoNroOcom, StringComparison.OrdinalIgnoreCase))
                        {
                            reg.NroOcom = nuevoNroOcom;
                            reg.NuOrdItem = item;
                            indicesActualizados++;
                        }
                    }
                }

                await _context.SaveChangesAsync();

                return Ok(new SincronizarEdicionConsignacionResponse
                {
                    Success = true,
                    Mensaje = "Sincronización de consignaciones completada.",
                    RegistrosAjustadosRendidos = ajustesRendidos,
                    RegistrosActualizadosIndice = indicesActualizados
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[SincronizarEdicionOrden] Error al sincronizar consignaciones para O.C. {NumeroOrden}", request?.NumeroOrden);
                return StatusCode(500, "Error al sincronizar consignaciones de la orden editada.");
            }
        }

        [HttpPost("pendientes-remitir")]
        public async Task<ActionResult<List<ConsignacionPendienteRemisionDTO>>> GetPendientesRemitir([FromBody] PendientesRemisionRequest request)
        {
            try
            {
                if (request.NumeroProveedor <= 0)
                {
                    return BadRequest("Debe indicar un proveedor válido.");
                }

                if (request.Deposito <= 0)
                {
                    return BadRequest("Debe seleccionar un depósito válido.");
                }

                var query = _context.Maconsigs
                    .AsNoTracking()
                    .Where(m => (m.NproOcom ?? 0) == request.NumeroProveedor
                                && (m.CantAsig ?? 0) > ((m.CantEntreg ?? 0) + (m.CantRendid ?? 0))
                                && (m.CantAsig ?? 0) >= 0.05m);

                var ordenesSeleccionadas = (request.NumerosOrdenCompra ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (!string.IsNullOrWhiteSpace(request.NumeroOrdenCompra) && !ordenesSeleccionadas.Contains(request.NumeroOrdenCompra.Trim(), StringComparer.OrdinalIgnoreCase))
                {
                    ordenesSeleccionadas.Add(request.NumeroOrdenCompra.Trim());
                }

                var registros = (await query
                    .OrderBy(m => m.NroOcom)
                    .ThenBy(m => m.CodInte)
                    .ThenBy(m => m.Id)
                    .ToListAsync())
                    .Where(m => ordenesSeleccionadas.Count == 0
                        || (!string.IsNullOrWhiteSpace(m.NroOcom)
                            && ordenesSeleccionadas.Any(oc => m.NroOcom!.StartsWith(oc, StringComparison.OrdinalIgnoreCase))))
                    .ToList();

                if (!registros.Any())
                {
                    return Ok(new List<ConsignacionPendienteRemisionDTO>());
                }

                var numerosOrden = registros
                    .Select(r => ExtraerNumeroOrden(r.NroOcom))
                    .Where(n => n.HasValue)
                    .Select(n => n!.Value)
                    .Distinct()
                    .ToList();

                var estadosOcom = await _context.Ocomencas
                    .AsNoTracking()
                    .Where(o => o.NumeOcom.HasValue && numerosOrden.Contains(o.NumeOcom.Value))
                    .Select(o => new { Numero = o.NumeOcom!.Value, Estado = o.StatOcom })
                    .ToDictionaryAsync(x => x.Numero, x => x.Estado);

                var registrosValidos = registros.Where(r =>
                {
                    var nuoc = ExtraerNumeroOrden(r.NroOcom);
                    if (!nuoc.HasValue)
                    {
                        return true;
                    }

                    if (!estadosOcom.TryGetValue(nuoc.Value, out var estado))
                    {
                        return true;
                    }

                    return estado >= 0 && estado < 8;
                }).ToList();

                if (!registrosValidos.Any())
                {
                    return Ok(new List<ConsignacionPendienteRemisionDTO>());
                }

                var codigos = registrosValidos
                    .Where(x => x.CodInte.HasValue)
                    .Select(x => x.CodInte!.Value)
                    .Distinct()
                    .ToList();

                var masters = await _context.masters
                    .AsNoTracking()
                    .Where(m => codigos.Contains(m.Codint))
                    .Select(m => new { m.Codint, m.Trazable, m.Umedida, m.Codigo, m.Descripcion })
                    .ToDictionaryAsync(x => x.Codint, x => x);

                var stockByCod = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.DepoMovi == request.Deposito && m.Cod_Inte.HasValue && codigos.Contains(m.Cod_Inte.Value))
                    .GroupBy(m => m.Cod_Inte!.Value)
                    .Select(g => new
                    {
                        CodInte = g.Key,
                        Stock = g.Sum(x => (x.CantIngre ?? 0m) - (x.CantSalid ?? 0m))
                    })
                    .ToDictionaryAsync(x => x.CodInte, x => x.Stock);

                var resultado = new List<ConsignacionPendienteRemisionDTO>();

                foreach (var r in registrosValidos)
                {
                    if (!r.CodInte.HasValue || r.CodInte.Value <= 0)
                    {
                        continue;
                    }

                    var codInte = r.CodInte.Value;
                    var cantAsig = Math.Round(r.CantAsig ?? 0m, 4);
                    var cantEntreg = Math.Round(r.CantEntreg ?? 0m, 4);
                    var cantRendid = Math.Round(r.CantRendid ?? 0m, 4);
                    var faltaEntregar = Math.Round(cantAsig - cantEntreg - cantRendid, 4);

                    if (faltaEntregar <= 0.005m)
                    {
                        continue;
                    }

                    var stockDeposito = stockByCod.TryGetValue(codInte, out var sto) ? sto : 0m;
                    var cantidadPendiente = faltaEntregar;
                    var quedaPendiente = faltaEntregar - cantidadPendiente;

                    masters.TryGetValue(codInte, out var master);

                    resultado.Add(new ConsignacionPendienteRemisionDTO
                    {
                        IdMaconsig = r.Id,
                        CodInte = codInte,
                        CodExte = r.CodExte ?? master?.Codigo ?? string.Empty,
                        Descripcion = r.Descritem ?? master?.Descripcion ?? string.Empty,
                        Trazabilidad = master?.Trazable,
                        Unidad = r.Unidad ?? master?.Umedida ?? string.Empty,
                        CantAsig = cantAsig,
                        CantEntreg = cantEntreg,
                        CantRendid = cantRendid,
                        CantidadPendiente = Math.Round(cantidadPendiente, 4),
                        OCompra = r.NroOcom ?? string.Empty,
                        Numero = ExtraerNumeroOrden(r.NroOcom) ?? 0,
                        Pendiente = Math.Round(quedaPendiente, 4),
                        FaltaEntregar = Math.Round(faltaEntregar, 4),
                        StockDeposito = Math.Round(stockDeposito, 4),
                        CantidadEntregar = Math.Round(cantidadPendiente, 4),
                        Lote = r.LoteOcom ?? string.Empty
                    });
                }

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetPendientesRemitir] Error al obtener pendientes de remisión.");
                return StatusCode(500, $"Error al obtener pendientes de remisión: {ex.Message}");
            }
        }

        [HttpGet("config-remision-inicial")]
        public async Task<ActionResult<ConfigRemisionInicialDTO>> GetConfigRemisionInicial()
        {
            try
            {
                List<ProveedorPendienteRemisionDTO> proveedores;
                try
                {
                    proveedores = await ObtenerProveedoresPendientesRemisionAsync();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[GetConfigRemisionInicial] Error parcial cargando proveedores pendientes. Se devolverá lista vacía.");
                    proveedores = new List<ProveedorPendienteRemisionDTO>();
                }

                List<DepositoRemisionDTO> depositos;
                try
                {
                    depositos = await _context.Tadeposi
                        .AsNoTracking()
                        .Where(d => d.Status >= 0)
                        .OrderBy(d => d.CodigoDepo)
                        .Select(d => new DepositoRemisionDTO
                        {
                            CodigoDepo = d.CodigoDepo,
                            Descripcion = d.Descripcion ?? string.Empty
                        })
                        .ToListAsync();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[GetConfigRemisionInicial] Error parcial cargando depósitos. Se devolverá lista vacía.");
                    depositos = new List<DepositoRemisionDTO>();
                }

                return Ok(new ConfigRemisionInicialDTO
                {
                    Proveedores = proveedores,
                    Depositos = depositos
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetConfigRemisionInicial] Error al obtener la configuración inicial de remisión.");
                return StatusCode(500, "Error al obtener configuración inicial de remisión.");
            }
        }

        [HttpGet("ordenes-compra-pendientes-remision/{numeroProveedor:int}")]
        public async Task<ActionResult<List<OrdenCompraPendienteRemisionDTO>>> GetOrdenesCompraPendientesRemision(int numeroProveedor)
        {
            try
            {
                if (numeroProveedor <= 0)
                {
                    return BadRequest("Debe indicar un proveedor válido.");
                }

                var pendientesRaw = await _context.Maconsigs
                    .AsNoTracking()
                    .Where(m => (m.NproOcom ?? 0) == numeroProveedor
                                && (m.CantAsig ?? 0m) > (m.CantEntreg ?? 0m)
                                && (m.CantAsig ?? 0m) >= 0.05m
                                && m.NroOcom != null
                                && m.NroOcom.Length >= 9)
                    .Select(m => new
                    {
                        NroOcomBase = m.NroOcom!.Substring(0, 9),
                        NumeroOrdenTexto = m.NroOcom.Substring(3, 6),
                        FechaEmision = m.FemiOcom
                    })
                    .ToListAsync();

                var pendientes = pendientesRaw
                    .Select(x => new
                    {
                        x.NroOcomBase,
                        NumeroOrden = int.TryParse(x.NumeroOrdenTexto, out var nro) ? nro : 0,
                        x.FechaEmision
                    })
                    .Where(x => x.NumeroOrden > 0)
                    .ToList();

                if (!pendientes.Any())
                {
                    return Ok(new List<OrdenCompraPendienteRemisionDTO>());
                }

                var numerosOrden = pendientes
                    .Select(x => x.NumeroOrden)
                    .Distinct()
                    .ToList();

                var estadosOrdenes = await _context.Ocomencas
                    .AsNoTracking()
                    .Where(o => o.NumeOcom.HasValue && numerosOrden.Contains(o.NumeOcom.Value))
                    .Select(o => new { NumeroOrden = o.NumeOcom!.Value, Estado = o.StatOcom })
                    .ToDictionaryAsync(x => x.NumeroOrden, x => x.Estado);

                var resultado = pendientes
                    .Where(x => !estadosOrdenes.TryGetValue(x.NumeroOrden, out var estado) || (estado >= 0 && estado < 8))
                    .GroupBy(x => new { x.NumeroOrden, x.NroOcomBase })
                    .Select(g => new OrdenCompraPendienteRemisionDTO
                    {
                        NumeroOrden = g.Key.NumeroOrden,
                        NroOcom = g.Key.NroOcomBase,
                        FechaEmision = g.Max(x => x.FechaEmision)
                    })
                    .OrderBy(x => x.NumeroOrden)
                    .ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetOrdenesCompraPendientesRemision] Error al obtener órdenes pendientes para proveedor {NumeroProveedor}.", numeroProveedor);
                return StatusCode(500, "Error al obtener órdenes de compra pendientes.");
            }
        }

        private async Task<List<ProveedorPendienteRemisionDTO>> ObtenerProveedoresPendientesRemisionAsync()
        {
            // Cargar todos los proveedores activos (no dados de baja) desde PROVED12.
            // Stat_Cli >= 0 significa que el proveedor está activo.
            var proveedoresActivos = await _context.Proveedores
                .AsNoTracking()
                .Where(p => p.Stat_Cli >= 0)
                .Select(p => new
                {
                    p.Nume_Cli,
                    Raso_Cli = p.Raso_Cli ?? string.Empty
                })
                .ToListAsync();

            if (!proveedoresActivos.Any())
            {
                return new List<ProveedorPendienteRemisionDTO>();
            }

            // Contar órdenes pendientes de remisión por proveedor desde MACONSIG.
            // Un proveedor puede tener 0 órdenes pendientes y aun así aparecer en la lista.
            var pendientesPorProveedor = await _context.Maconsigs
                .AsNoTracking()
                .Where(m => (m.NproOcom ?? 0) > 0
                            && (m.CantAsig ?? 0m) > ((m.CantEntreg ?? 0m) + (m.CantRendid ?? 0m))
                            && (m.CantAsig ?? 0m) >= 0.05m
                            && m.NroOcom != null)
                .Select(m => new
                {
                    NumeroProveedor = (int)(m.NproOcom ?? 0),
                    NumeroOrden = m.NroOcom!.Length >= 9 ? m.NroOcom.Substring(0, 9) : m.NroOcom
                })
                .GroupBy(x => x.NumeroProveedor)
                .Select(g => new
                {
                    NumeroProveedor = g.Key,
                    CantidadOrdenes = g.Select(x => x.NumeroOrden).Distinct().Count()
                })
                .ToDictionaryAsync(x => x.NumeroProveedor, x => x.CantidadOrdenes);

            return proveedoresActivos
                .Select(p => new ProveedorPendienteRemisionDTO
                {
                    NumeroProveedor = p.Nume_Cli,
                    RazonSocial = p.Raso_Cli,
                    CantidadOrdenesPendientes = pendientesPorProveedor.TryGetValue(p.Nume_Cli, out var cant) ? cant : 0
                })
                .OrderBy(x => x.RazonSocial)
                .ThenBy(x => x.NumeroProveedor)
                .ToList();
        }

        [HttpGet("proximo-comprobante-rc")]
        public async Task<ActionResult<int>> GetProximoComprobanteRc()
        {
            try
            {
                var proximo = await ObtenerProximoNumeroComprobanteAsync("RC");
                return Ok(proximo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetProximoComprobanteRc] Error al obtener próximo comprobante RC.");
                return StatusCode(500, $"Error al obtener próximo comprobante RC: {ex.Message}");
            }
        }

        // Devuelve los remitos de consignación (RC) de un mes/año para mostrarlos en el modal del cliente.
        [HttpGet("remitos-rc/{anio:int}/{mes:int}")]
        public async Task<ActionResult<List<RemitoConsignacionListadoDTO>>> GetRemitosRcPorPeriodo(int anio, int mes)
        {
            try
            {
                // Validación básica de período para evitar consultas inválidas.
                if (anio < 2000 || anio > 2100 || mes < 1 || mes > 12)
                {
                    return BadRequest("Período inválido.");
                }

                // Armamos el rango [desde, hasta) del mes solicitado.
                var desde = new DateTime(anio, mes, 1);
                var hasta = desde.AddMonths(1);

                // Tomamos movimientos RC del período y nos quedamos con los datos mínimos de cabecera.
                var remitosRaw = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.CodiMovi == "RC"
                                && m.DoPriCor != null
                                && m.FechMov.HasValue
                                && m.FechMov.Value >= desde
                                && m.FechMov.Value < hasta)
                    .Select(m => new
                    {
                        DoPriCor = m.DoPriCor!,
                        Fecha = m.FechMov!.Value,
                        NumeroProveedor = m.Nume_Prov ?? 0
                    })
                    .ToListAsync();

                if (!remitosRaw.Any())
                {
                    return Ok(new List<RemitoConsignacionListadoDTO>());
                }

                // Agrupamos por número de remito para devolver una fila por remito (y no una por ítem).
                var remitosAgrupados = remitosRaw
                    .GroupBy(x => x.DoPriCor)
                    .Select(g => new
                    {
                        DoPriCor = g.Key,
                        Fecha = g.Max(x => x.Fecha),
                        NumeroProveedor = g.Select(x => x.NumeroProveedor).FirstOrDefault(p => p > 0)
                    })
                    .OrderByDescending(x => x.Fecha)
                    .ThenByDescending(x => x.DoPriCor)
                    .ToList();

                // Buscamos los proveedores necesarios para mostrar razón social en la grilla.
                var proveedoresIds = remitosAgrupados
                    .Where(x => x.NumeroProveedor > 0)
                    .Select(x => x.NumeroProveedor)
                    .Distinct()
                    .ToList();

                var proveedores = await _context.Proveedores
                    .AsNoTracking()
                    .Where(p => proveedoresIds.Contains(p.Nume_Cli))
                    .Select(p => new { p.Nume_Cli, p.Raso_Cli })
                    .ToDictionaryAsync(p => p.Nume_Cli, p => p.Raso_Cli ?? string.Empty);

                // Mapeo final al DTO que consume la UI.
                var result = remitosAgrupados
                    .Select(x => new RemitoConsignacionListadoDTO
                    {
                        DoPriCor = x.DoPriCor,
                        Fecha = x.Fecha,
                        NumeroProveedor = x.NumeroProveedor,
                        RazonSocialProveedor = x.NumeroProveedor > 0 && proveedores.TryGetValue(x.NumeroProveedor, out var razon)
                            ? razon
                            : string.Empty
                    })
                    .ToList();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetRemitosRcPorPeriodo] Error al obtener remitos RC del período {Anio}-{Mes}", anio, mes);
                return StatusCode(500, "Error al obtener remitos RC.");
            }
        }

        // Devuelve el detalle completo de un remito RC para imprimirlo con el formato estándar.
        [HttpGet("remitos-rc/detalle")]
        public async Task<ActionResult<RemitoConsignacionDetalleDTO>> GetRemitoRcDetalle([FromQuery] string doPriCor)
        {
            try
            {
                // Validación de parámetro de entrada.
                if (string.IsNullOrWhiteSpace(doPriCor))
                {
                    return BadRequest("Debe indicar un número de remito válido.");
                }

                var remito = doPriCor.Trim().ToUpperInvariant();

                // Recupera todos los ítems del remito RC.
                var movimientos = await _context.Movistos
                    .AsNoTracking()
                    .Where(m => m.CodiMovi == "RC" && m.DoPriCor == remito)
                    .OrderBy(m => m.Id)
                    .ToListAsync();

                if (!movimientos.Any())
                {
                    return NotFound();
                }

                // La cabecera se toma del primer movimiento del remito.
                var cabecera = movimientos.First();
                var numeroProveedor = cabecera.Nume_Prov ?? 0;

                Proved12? proveedor = null;
                if (numeroProveedor > 0)
                {
                    // Cargamos datos del proveedor para completar el encabezado de impresión.
                    proveedor = await _context.Proveedores
                        .AsNoTracking()
                        .FirstOrDefaultAsync(p => p.Nume_Cli == numeroProveedor);
                }

                // Armado final del DTO con cabecera + detalle de líneas.
                var detalle = new RemitoConsignacionDetalleDTO
                {
                    DoPriCor = remito,
                    Fecha = cabecera.FechMov,
                    NumeroProveedor = numeroProveedor,
                    RazonSocialProveedor = proveedor?.Raso_Cli ?? string.Empty,
                    DomicilioProveedor = proveedor?.Domi_Cli ?? string.Empty,
                    LocalidadProveedor = proveedor?.Loca_Cli ?? string.Empty,
                    CuitProveedor = proveedor?.Cuit_Cli ?? string.Empty,
                    PivaProveedor = proveedor?.Piva_Cli,
                    OrdenCompra = cabecera.DoSecLar ?? string.Empty,
                    Items = movimientos
                        .Where(m => (m.CantSalid ?? 0m) > 0.0001m)
                        .Select(m => new RemitoConsignacionDetalleItemDTO
                        {
                            Cantidad = Math.Round(m.CantSalid ?? 0m, 4),
                            Codigo = m.Cod_Exte ?? string.Empty,
                            Descripcion = m.Descrip ?? string.Empty
                        })
                        .ToList()
                };

                return Ok(detalle);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetRemitoRcDetalle] Error al obtener detalle de remito RC {DoPriCor}", doPriCor);
                return StatusCode(500, "Error al obtener detalle de remito RC.");
            }
        }

        [HttpGet("trazabilidad-oc/{numeroOrden:int}")]
        public async Task<ActionResult<List<ConsignacionTrazabilidadPadreDTO>>> GetTrazabilidadPorOrdenCompra(int numeroOrden)
        {
            try
            {
                if (numeroOrden <= 0)
                {
                    return BadRequest("El número de orden debe ser mayor a 0.");
                }

                var prefijoOrden = $"OC-{numeroOrden:000000}-";

                var consignaciones = await _context.Maconsigs
                    .AsNoTracking()
                    .Where(m => m.NroOcom != null && m.NroOcom.StartsWith(prefijoOrden))
                    .OrderBy(m => m.NroOcom)
                    .ThenBy(m => m.CodInte)
                    .ThenBy(m => m.Id)
                    .ToListAsync();

                if (!consignaciones.Any())
                {
                    return Ok(new List<ConsignacionTrazabilidadPadreDTO>());
                }

                var itemNumbers = consignaciones
                    .Select(m => ExtraerNumeroItem(m.NroOcom))
                    .Where(x => x.HasValue)
                    .Select(x => x!.Value)
                    .Distinct()
                    .ToList();

                var detalleOrden = await _context.Ocomdetas
                    .AsNoTracking()
                    .Where(d => d.NumeOcom == numeroOrden && d.NuOrdItem.HasValue && itemNumbers.Contains(d.NuOrdItem.Value))
                    .Select(d => new
                    {
                        Item = d.NuOrdItem!.Value,
                        CodPadre = (int)(d.CodInte ?? 0),
                        d.CodExte,
                        d.Descritem,
                        d.CantOcom
                    })
                    .ToListAsync();

                var codigosPadre = detalleOrden
                    .Where(x => x.CodPadre > 0)
                    .Select(x => x.CodPadre)
                    .Distinct()
                    .ToList();

                var mastersPadre = await _context.masters
                    .AsNoTracking()
                    .Where(m => codigosPadre.Contains(m.Codint))
                    .Select(m => new { m.Codint, m.Codigo, m.Descripcion })
                    .ToDictionaryAsync(x => x.Codint, x => x);

                var brData = await _context.Bolreceps
                    .AsNoTracking()
                    .Where(b => b.NumOC == numeroOrden && b.NroItemOC.HasValue && itemNumbers.Contains(b.NroItemOC.Value))
                    .Select(b => new
                    {
                        Item = b.NroItemOC!.Value,
                        NumeroBoleta = b.NumeBoRe,
                        FechaBr = b.FechRecep ?? b.FEYHORecep,
                        CantidadAprobada = b.CanToApro ?? 0m
                    })
                    .ToListAsync();

                var brPorItem = brData
                    .GroupBy(x => x.Item)
                    .ToDictionary(g => g.Key, g => g
                        .OrderByDescending(x => x.FechaBr)
                        .ThenByDescending(x => x.NumeroBoleta)
                        .ToList());

                var detalleByItem = detalleOrden
                    .GroupBy(x => x.Item)
                    .ToDictionary(g => g.Key, g => g.First());

                var grupos = consignaciones
                    .GroupBy(m => new
                    {
                        Item = ExtraerNumeroItem(m.NroOcom) ?? (short)0,
                        NroOcom = m.NroOcom ?? string.Empty
                    })
                    .OrderBy(g => g.Key.Item)
                    .ToList();

                var resultado = new List<ConsignacionTrazabilidadPadreDTO>();

                foreach (var grupo in grupos)
                {
                    var item = grupo.Key.Item;
                    detalleByItem.TryGetValue(item, out var padre);

                    var codPadre = padre?.CodPadre ?? 0;
                    var codigoPadre = codPadre > 0 && mastersPadre.TryGetValue(codPadre, out var mp)
                        ? (mp.Codigo ?? string.Empty)
                        : (padre?.CodExte ?? string.Empty);
                    var descripcionPadre = codPadre > 0 && mastersPadre.TryGetValue(codPadre, out var mp2)
                        ? (mp2.Descripcion ?? string.Empty)
                        : (padre?.Descritem ?? string.Empty);

                    var detalleConsignaciones = grupo
                        .Select(m =>
                        {
                            var cantAsig = Math.Round(m.CantAsig ?? 0m, 4);
                            var cantEntreg = Math.Round(m.CantEntreg ?? 0m, 4);
                            var cantRendid = Math.Round(m.CantRendid ?? 0m, 4);
                            var saldo = Math.Round(cantAsig - cantEntreg - cantRendid, 4);

                            return new ConsignacionTrazabilidadDetalleDTO
                            {
                                IdMaconsig = m.Id,
                                NroOcom = m.NroOcom ?? string.Empty,
                                CodInte = m.CodInte ?? 0,
                                CodExte = m.CodExte ?? string.Empty,
                                Descripcion = m.Descritem ?? string.Empty,
                                Unidad = m.Unidad ?? string.Empty,
                                CantAsig = cantAsig,
                                CantEntreg = cantEntreg,
                                CantRendid = cantRendid,
                                Saldo = saldo
                            };
                        })
                        .OrderBy(x => x.CodExte)
                        .ThenBy(x => x.CodInte)
                        .ThenBy(x => x.IdMaconsig)
                        .ToList();

                    var brDelPadre = brPorItem.TryGetValue(item, out var brList)
                        ? brList.Select(br =>
                        {
                            var numeroBoleta = (int)Math.Floor(br.NumeroBoleta ?? 0m);
                            return new ConsignacionTrazabilidadBrDTO
                            {
                                NroBr = numeroBoleta > 0 ? $"BR-{numeroBoleta:000000}-{item:00}" : string.Empty,
                                FechaBr = br.FechaBr,
                                CantidadAprobada = Math.Round(br.CantidadAprobada, 4)
                            };
                        }).ToList()
                        : new List<ConsignacionTrazabilidadBrDTO>();

                    resultado.Add(new ConsignacionTrazabilidadPadreDTO
                    {
                        NumeroOrden = numeroOrden,
                        NumeroItem = item,
                        NroOcom = grupo.Key.NroOcom,
                        CodIntePadre = codPadre,
                        CodigoPadre = codigoPadre,
                        DescripcionPadre = descripcionPadre,
                        CantidadPadreOrden = Math.Round(padre?.CantOcom ?? 0m, 4),
                        TotalCantAsig = Math.Round(detalleConsignaciones.Sum(x => x.CantAsig), 4),
                        TotalCantEntreg = Math.Round(detalleConsignaciones.Sum(x => x.CantEntreg), 4),
                        TotalCantRendid = Math.Round(detalleConsignaciones.Sum(x => x.CantRendid), 4),
                        TotalSaldo = Math.Round(detalleConsignaciones.Sum(x => x.Saldo), 4),
                        Brs = brDelPadre,
                        Detalle = detalleConsignaciones
                    });
                }

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[GetTrazabilidadPorOrdenCompra] Error al obtener trazabilidad de consignaciones para O.C. {NumeroOrden}", numeroOrden);
                return StatusCode(500, $"Error al obtener trazabilidad de consignaciones: {ex.Message}");
            }
        }

        [HttpPost("remitir")]
        public async Task<IActionResult> RemitirConsignaciones([FromBody] RemitirConsignacionesRequest request)
        {
            if (request.Deposito <= 0)
            {
                return BadRequest("Debe seleccionar un depósito válido.");
            }

            if (request.NumeroProveedor <= 0)
            {
                return BadRequest("Debe indicar un proveedor válido.");
            }

            if (request.Items == null || !request.Items.Any(x => x.CantidadEntregar > 0.005m))
            {
                return BadRequest("No hay ítems con cantidad a remitir.");
            }

            var codigoMovimiento = string.IsNullOrWhiteSpace(request.CodigoMovimiento)
                ? "RC"
                : request.CodigoMovimiento.Trim().ToUpper();

            if (codigoMovimiento != "RC")
            {
                return BadRequest("El código de movimiento para remisión de consignación debe ser 'RC'.");
            }

            if (request.CodiTra.HasValue
                && (request.CodiTra.Value < short.MinValue || request.CodiTra.Value > short.MaxValue))
            {
                // MOVISTO.TipoTran es short?, por eso validamos rango antes de persistir.
                return BadRequest("El transporte seleccionado tiene un Codi_Tra fuera de rango para MOVISTO.TipoTran.");
            }

            using var tx = await _context.Database.BeginTransactionAsync(IsolationLevel.Serializable);
            try
            {
                var numeroComprobante = request.NumeroComprobante ?? await ObtenerProximoNumeroComprobanteAsync(codigoMovimiento);
                var doPriCor = $"{codigoMovimiento}-{numeroComprobante}";

                var comprobanteLibre = !await _context.Movistos
                    .AnyAsync(m => m.CodiMovi == codigoMovimiento && m.DoPriCor == doPriCor);

                if (!comprobanteLibre)
                {
                    return BadRequest($"El comprobante {doPriCor} ya existe. Debe generar uno nuevo.");
                }

                var fechaRemito = request.FechaRemito ?? DateTime.Now;
                var usuarioActivo = _usuarioFlexoftService.InterObtenerIdFlexoft(User);

                var codigos = request.Items
                    .Where(i => i.CodInte > 0)
                    .Select(i => i.CodInte)
                    .Distinct()
                    .ToList();

                var maestros = await _context.masters
                    .Where(m => codigos.Contains(m.Codint))
                    .ToDictionaryAsync(m => m.Codint, m => m);

                decimal totalRemitido = 0m;
                int itemsProcesados = 0;

                foreach (var item in request.Items.Where(x => x.CantidadEntregar > 0.005m && x.CodInte > 0))
                {
                    var restante = Math.Round(item.CantidadEntregar, 4);

                    while (restante > 0.005m)
                    {
                        var queryMaconsig = _context.Maconsigs
                            .Where(m => (m.NproOcom ?? 0) == request.NumeroProveedor
                                        && (m.CodInte ?? 0) == item.CodInte
                                        && (m.CantAsig ?? 0) > ((m.CantEntreg ?? 0) + (m.CantRendid ?? 0)));

                        if (!string.IsNullOrWhiteSpace(item.NroOcom))
                        {
                            queryMaconsig = queryMaconsig.Where(m => m.NroOcom == item.NroOcom);
                        }
                        else if (!string.IsNullOrWhiteSpace(request.NumeroOrdenCompra))
                        {
                            queryMaconsig = queryMaconsig.Where(m => m.NroOcom != null && m.NroOcom.StartsWith(request.NumeroOrdenCompra));
                        }

                        var registro = await queryMaconsig
                            .OrderBy(m => m.Id)
                            .FirstOrDefaultAsync();

                        if (registro == null)
                        {
                            if (!request.PermitirCrearRegistroSiNoExiste)
                            {
                                break;
                            }

                            var nroOcomNuevo = !string.IsNullOrWhiteSpace(item.NroOcom)
                                ? item.NroOcom
                                : (string.IsNullOrWhiteSpace(request.NumeroOrdenCompra) ? string.Empty : $"{request.NumeroOrdenCompra}-00");

                            var master = maestros.TryGetValue(item.CodInte, out var m) ? m : null;

                            var nuevoMaconsig = new Maconsig
                            {
                                CodiEmpr = 0,
                                FemiOcom = fechaRemito,
                                NproOcom = request.NumeroProveedor,
                                RaSoProv = request.RazonSocialProveedor ?? string.Empty,
                                NroOcom = nroOcomNuevo,
                                CodInte = item.CodInte,
                                CodExte = item.CodExte ?? master?.Codigo ?? string.Empty,
                                Descritem = item.Descripcion ?? master?.Descripcion ?? string.Empty,
                                Unidad = item.Unidad ?? master?.Umedida ?? string.Empty,
                                CantAsig = restante,
                                CantEntreg = restante,
                                CantRendid = 0,
                                CantDevu = 0,
                                SaldoPend = 0,
                                LoteOcom = item.Lote,
                                FechaEntreg = fechaRemito,
                                FechEntreg = fechaRemito,
                                NRemiConsig = numeroComprobante.ToString(),
                                NuOrdItem = item.NuOrdItem ?? 0
                            };

                            _context.Maconsigs.Add(nuevoMaconsig);

                            var movNuevo = CrearMovimientoSalidaRc(
                                codigoMovimiento,
                                doPriCor,
                                request,
                                item,
                                fechaRemito,
                                restante,
                                usuarioActivo,
                                master);

                            _context.Movistos.Add(movNuevo);

                            totalRemitido += restante;
                            itemsProcesados++;
                            restante = 0m;
                            continue;
                        }

                        var cantPendiente = Math.Round((registro.CantAsig ?? 0m) - (registro.CantEntreg ?? 0m) - (registro.CantRendid ?? 0m), 4);
                        if (cantPendiente <= 0.005m)
                        {
                            break;
                        }

                        var entregar = Math.Min(restante, cantPendiente);

                        registro.CantEntreg = Math.Round((registro.CantEntreg ?? 0m) + entregar, 4);
                        registro.FechaEntreg = fechaRemito;
                        registro.FechEntreg = fechaRemito;
                        registro.NRemiConsig = numeroComprobante.ToString();

                        var masterExist = maestros.TryGetValue(item.CodInte, out var masterReg) ? masterReg : null;

                        var mov = CrearMovimientoSalidaRc(
                            codigoMovimiento,
                            doPriCor,
                            request,
                            item,
                            fechaRemito,
                            entregar,
                            usuarioActivo,
                            masterExist,
                            registro.NroOcom);

                        _context.Movistos.Add(mov);

                        restante = Math.Round(restante - entregar, 4);
                        totalRemitido += entregar;
                        itemsProcesados++;
                    }
                }

                await _context.SaveChangesAsync();
                await tx.CommitAsync();

                return Ok(new
                {
                    Success = true,
                    CodigoMovimiento = codigoMovimiento,
                    NumeroComprobante = numeroComprobante,
                    DoPriCor = doPriCor,
                    Deposito = request.Deposito,
                    ItemsProcesados = itemsProcesados,
                    CantidadTotalRemitida = Math.Round(totalRemitido, 4)
                });
            }
            catch (Exception ex)
            {
                await tx.RollbackAsync();
                _logger.LogError(ex, "[RemitirConsignaciones] Error al remitir consignaciones.");
                return StatusCode(500, $"Error al remitir consignaciones: {ex.Message}");
            }
        }

        private Movisto CrearMovimientoSalidaRc(
            string codigoMovimiento,
            string doPriCor,
            RemitirConsignacionesRequest request,
            RemitirConsignacionItemRequest item,
            DateTime fechaRemito,
            decimal cantidad,
            int? usuarioActivo,
            Master? master,
            string? nroOcom = null)
        {
            return new Movisto
            {
                CodiEmpr = 0,
                NumUsuar = usuarioActivo,
                FeReMovi = DateTime.Now,
                FechMov = fechaRemito,
                CodiMovi = codigoMovimiento,
                DoPriCor = doPriCor,
                DoPriLar = doPriCor,
                DoSecCor = (nroOcom ?? item.NroOcom ?? request.NumeroOrdenCompra ?? string.Empty).Substring(0, Math.Min(12, (nroOcom ?? item.NroOcom ?? request.NumeroOrdenCompra ?? string.Empty).Length)),
                DoSecLar = (nroOcom ?? item.NroOcom ?? request.NumeroOrdenCompra ?? string.Empty).Substring(0, Math.Min(24, (nroOcom ?? item.NroOcom ?? request.NumeroOrdenCompra ?? string.Empty).Length)),
                DocuOrig = (nroOcom ?? item.NroOcom ?? request.NumeroOrdenCompra ?? string.Empty).Substring(0, Math.Min(24, (nroOcom ?? item.NroOcom ?? request.NumeroOrdenCompra ?? string.Empty).Length)),
                Cod_Inte = item.CodInte,
                Cod_Exte = (item.CodExte ?? master?.Codigo ?? string.Empty).Substring(0, Math.Min(24, (item.CodExte ?? master?.Codigo ?? string.Empty).Length)),
                Descrip = (item.Descripcion ?? master?.Descripcion ?? string.Empty).Substring(0, Math.Min(64, (item.Descripcion ?? master?.Descripcion ?? string.Empty).Length)),
                Descri_Lar = item.Descripcion ?? master?.Descripcion ?? string.Empty,
                IdenLote = string.IsNullOrWhiteSpace(item.Lote) ? null : item.Lote.Substring(0, Math.Min(16, item.Lote.Length)),
                DepoMovi = request.Deposito,
                Nume_Prov = request.NumeroProveedor,
                Nume_Clie = 0,
                CantIngre = 0,
                CantSalid = Math.Round(Math.Abs(cantidad), 4),
                CantSaldo = 0,
                NuorItem = item.NuOrdItem ?? 0,
                // Se persiste el Codi_Tra seleccionado en UI dentro de MOVISTO.TipoTran.
                TipoTran = request.CodiTra.HasValue ? (short?)request.CodiTra.Value : null,
                ReferCor = "REMISION CONSIGNACION",
                VerificadoPor = 0,
                FHor_Verif = new DateTime(1900, 1, 1)
            };
        }

        private async Task<int> ObtenerProximoNumeroComprobanteAsync(string codigoMovimiento)
        {
            var codMov = codigoMovimiento.Trim().ToUpper();

            var correlativos = await _context.Movistos
                .Where(m => m.CodiMovi == codMov
                            && m.DoPriCor != null
                            && m.DoPriCor.StartsWith(codMov + "-"))
                .Select(m => m.DoPriCor!)
                .ToListAsync();

            var maximo = correlativos
                .Select(x =>
                {
                    var partes = x.Split('-', StringSplitOptions.RemoveEmptyEntries);
                    var parteNumerica = partes.Length >= 2 ? partes[1].Trim() : string.Empty;
                    return int.TryParse(parteNumerica, out var nro) ? (int?)nro : null;
                })
                .Where(x => x.HasValue)
                .Max();

            return (maximo ?? 0) + 1;
        }

        private static int? ExtraerNumeroOrden(string? nroOcom)
        {
            if (string.IsNullOrWhiteSpace(nroOcom) || nroOcom.Length < 9)
            {
                return null;
            }

            return int.TryParse(nroOcom.Substring(3, 6), out var numero) ? numero : null;
        }

        private static short? ExtraerNumeroItem(string? nroOcom)
        {
            if (string.IsNullOrWhiteSpace(nroOcom) || nroOcom.Length < 12)
            {
                return null;
            }

            var ultimoGuion = nroOcom.LastIndexOf('-');
            if (ultimoGuion < 0 || ultimoGuion + 1 >= nroOcom.Length)
            {
                return null;
            }

            var itemStr = nroOcom[(ultimoGuion + 1)..];
            return short.TryParse(itemStr, out var item) ? item : null;
        }
    }

    /// <summary>
    /// DTO para la solicitud de borrar consignaciones
    /// USO: SOLO EN EDICIÓN - PASO 1 del flujo de edición
    /// </summary>
    public class BorrarConsignacionesRequest
    {
        public int NumeroOrden { get; set; }
        public string CodigoFormateado { get; set; } = string.Empty;
        public string Condicion { get; set; } = string.Empty;
    }

    /// <summary>
    /// DTO para la solicitud de procesamiento de consignaciones
    /// USO: TANTO EN ALTA COMO EN EDICIÓN
    /// - ALTA: Desde OcomencaController.PostAsync()
    /// - EDICIÓN: PASO 6 del flujo de edición
    /// </summary>
    public class ProcesarConsignacionesRequest
    {
        public int NumeroOrden { get; set; }
        public int NumeroProveedor { get; set; }
        public string RazonSocialProveedor { get; set; } = string.Empty;
        public DateTime FechaEmision { get; set; }
        public short MonedaEmision { get; set; }
    }

    public class PendientesRemisionRequest
    {
        public int NumeroProveedor { get; set; }
        public string? NumeroOrdenCompra { get; set; }
        public List<string>? NumerosOrdenCompra { get; set; }
        public int Deposito { get; set; }
    }

    public class ConfigRemisionInicialDTO
    {
        public List<ProveedorPendienteRemisionDTO> Proveedores { get; set; } = new();
        public List<DepositoRemisionDTO> Depositos { get; set; } = new();
    }

    public class ProveedorPendienteRemisionDTO
    {
        public int NumeroProveedor { get; set; }
        public string RazonSocial { get; set; } = string.Empty;
        public int CantidadOrdenesPendientes { get; set; }
    }

    public class DepositoRemisionDTO
    {
        public int CodigoDepo { get; set; }
        public string Descripcion { get; set; } = string.Empty;
    }

    public class OrdenCompraPendienteRemisionDTO
    {
        public int NumeroOrden { get; set; }
        public string NroOcom { get; set; } = string.Empty;
        public DateTime? FechaEmision { get; set; }
    }

    public class ConsignacionPendienteRemisionDTO
    {
        public int IdMaconsig { get; set; }
        public int CodInte { get; set; }
        public string CodExte { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public short? Trazabilidad { get; set; }
        public string Unidad { get; set; } = string.Empty;
        public decimal CantAsig { get; set; }
        public decimal CantEntreg { get; set; }
        public decimal CantRendid { get; set; }
        public decimal CantidadPendiente { get; set; }
        public string OCompra { get; set; } = string.Empty;
        public int Numero { get; set; }
        public decimal Pendiente { get; set; }
        public decimal FaltaEntregar { get; set; }
        public decimal StockDeposito { get; set; }
        public decimal CantidadEntregar { get; set; }
        public string Lote { get; set; } = string.Empty;
    }

    public class RemitirConsignacionesRequest
    {
        public int NumeroProveedor { get; set; }
        public string? RazonSocialProveedor { get; set; }
        public string? NumeroOrdenCompra { get; set; }
        public int Deposito { get; set; }
        public DateTime? FechaRemito { get; set; }
        public string CodigoMovimiento { get; set; } = "RC";
        public int? NumeroComprobante { get; set; }
        // Codi_Tra de TRANSPORTES (viene del modal) y se mapea a MOVISTO.TipoTran.
        public int? CodiTra { get; set; }
        public bool PermitirCrearRegistroSiNoExiste { get; set; } = true;
        public List<RemitirConsignacionItemRequest> Items { get; set; } = new();
    }

    public class RemitirConsignacionItemRequest
    {
        public int IdMaconsig { get; set; }
        public int CodInte { get; set; }
        public string? CodExte { get; set; }
        public string? Descripcion { get; set; }
        public string? Unidad { get; set; }
        public string? NroOcom { get; set; }
        public short? NuOrdItem { get; set; }
        public string? Lote { get; set; }
        public decimal CantidadEntregar { get; set; }
    }

    // Fila de grilla para el modal "Ver Remitos en Consignación".
    public class RemitoConsignacionListadoDTO
    {
        public string DoPriCor { get; set; } = string.Empty;
        public DateTime Fecha { get; set; }
        public int NumeroProveedor { get; set; }
        public string RazonSocialProveedor { get; set; } = string.Empty;
    }

    // Estructura completa para imprimir un remito RC.
    public class RemitoConsignacionDetalleDTO
    {
        public string DoPriCor { get; set; } = string.Empty;
        public DateTime? Fecha { get; set; }
        public int NumeroProveedor { get; set; }
        public string RazonSocialProveedor { get; set; } = string.Empty;
        public string DomicilioProveedor { get; set; } = string.Empty;
        public string LocalidadProveedor { get; set; } = string.Empty;
        public string CuitProveedor { get; set; } = string.Empty;
        public short? PivaProveedor { get; set; }
        public string OrdenCompra { get; set; } = string.Empty;
        public List<RemitoConsignacionDetalleItemDTO> Items { get; set; } = new();
    }

    // Línea de detalle (ítem) para la impresión de remito RC.
    public class RemitoConsignacionDetalleItemDTO
    {
        public decimal Cantidad { get; set; }
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
    }

    public class ValidarEdicionConsignacionRequest
    {
        public int NumeroOrden { get; set; }
        public List<ValidarEdicionConsignacionItemRequest> Items { get; set; } = new();
    }

    public class ValidarEdicionConsignacionItemRequest
    {
        public short NumeroItem { get; set; }
        public int CodIntePadre { get; set; }
        public string CodExtePadre { get; set; } = string.Empty;
        public string DescripcionPadre { get; set; } = string.Empty;
        public decimal Cantidad { get; set; }
    }

    public class ValidacionEdicionConsignacionResponse
    {
        public bool PuedeContinuar { get; set; }
        public string Mensaje { get; set; } = string.Empty;
        public List<ConflictoEdicionConsignacionItemDTO> Conflictos { get; set; } = new();
    }

    public class ConflictoEdicionConsignacionItemDTO
    {
        public int CodInte { get; set; }
        public string CodExte { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public decimal CantidadNecesaria { get; set; }
        public decimal CantidadEntregada { get; set; }
        public decimal CantidadRendida { get; set; }
        public decimal ExcesoEntregado { get; set; }
        public string ItemsAfectados { get; set; } = string.Empty;
        public List<short> ItemsAfectadosNumeros { get; set; } = new();
    }

    public class SincronizarEdicionConsignacionResponse
    {
        public bool Success { get; set; }
        public string Mensaje { get; set; } = string.Empty;
        public int RegistrosAjustadosRendidos { get; set; }
        public int RegistrosActualizadosIndice { get; set; }
    }

    public class ConsignacionTrazabilidadPadreDTO
    {
        public int NumeroOrden { get; set; }
        public short NumeroItem { get; set; }
        public string NroOcom { get; set; } = string.Empty;
        public int CodIntePadre { get; set; }
        public string CodigoPadre { get; set; } = string.Empty;
        public string DescripcionPadre { get; set; } = string.Empty;
        public decimal CantidadPadreOrden { get; set; }
        public decimal TotalCantAsig { get; set; }
        public decimal TotalCantEntreg { get; set; }
        public decimal TotalCantRendid { get; set; }
        public decimal TotalSaldo { get; set; }
        public List<ConsignacionTrazabilidadBrDTO> Brs { get; set; } = new();
        public List<ConsignacionTrazabilidadDetalleDTO> Detalle { get; set; } = new();
    }

    public class ConsignacionTrazabilidadDetalleDTO
    {
        public int IdMaconsig { get; set; }
        public string NroOcom { get; set; } = string.Empty;
        public int CodInte { get; set; }
        public string CodExte { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public string Unidad { get; set; } = string.Empty;
        public decimal CantAsig { get; set; }
        public decimal CantEntreg { get; set; }
        public decimal CantRendid { get; set; }
        public decimal Saldo { get; set; }
    }

    public class ConsignacionTrazabilidadBrDTO
    {
        public string NroBr { get; set; } = string.Empty;
        public DateTime? FechaBr { get; set; }
        public decimal CantidadAprobada { get; set; }
    }
}
