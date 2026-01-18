using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la funcionalidad de Carga Horaria de Selección Múltiple
    /// Implementa la lógica del formulario VB6 CARHORAR
    /// </summary>
    [ApiController]
    [Route("/api/cargahoraria")]
    public class CargaHorariaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<CargaHorariaController> _logger;

        public CargaHorariaController(DataContext context, ILogger<CargaHorariaController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Procesa los programas seleccionados y retorna la información de piezas y carga horaria
        /// Implementa la lógica del Command4_Click del VB6 CARHORAR
        /// </summary>
        /// <param name="request">Request con los números de programas seleccionados</param>
        /// <returns>Response con programas, piezas y carga horaria</returns>
        [HttpPost("procesar-programas-seleccionados")]
        public async Task<ActionResult<CargaHorariaResponseDto>> ProcesarProgramasSeleccionados(
            [FromBody] CargaHorariaRequestDto request)
        {
            try
            {
                // Validación: debe haber al menos un programa seleccionado
                if (request.NumerosProgramas == null || !request.NumerosProgramas.Any())
                {
                    return BadRequest("Debe seleccionar al menos un programa");
                }

                var response = new CargaHorariaResponseDto();

                // 1. Obtener información de los programas seleccionados
                response.Programas = await _context.Profabris
                    .Where(p => request.NumerosProgramas.Contains(p.NumeroPrograma))
                    .OrderBy(p => p.NumeroPrograma)
                    .Select(p => new ProgramaSeleccionDto
                    {
                        NumeroPrograma = p.NumeroPrograma,
                        Descripcion = p.Descripcion,
                        FechaCreacion = p.FechaCreacion
                    })
                    .ToListAsync();

                if (!response.Programas.Any())
                {
                    return NotFound("No se encontraron programas con los números especificados");
                }

                // 2. Procesar piezas por programa (LISTARPROYECTO en VB6)
                response.Piezas = await ProcessPiezasPorPrograma(request.NumerosProgramas);

                // 3. Procesar carga horaria por máquina (CARGARHORASMAQUINA en VB6)
                response.CargaHoraria = await ProcessCargaHorariaPorMaquina(request.NumerosProgramas);

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar programas seleccionados");
                return StatusCode(500, $"Error al procesar programas: {ex.Message}");
            }
        }

        /// <summary>
        /// Procesa las piezas de los programas seleccionados, agrupándolas y calculando promedios
        /// Implementa la lógica de LISTARPROYECTO del VB6
        /// </summary>
        private async Task<List<PiezaProgramaDto>> ProcessPiezasPorPrograma(List<int> numerosProgramas)
        {
            var piezasPorPrograma = new Dictionary<int, Dictionary<int, decimal>>(); // Key: CodigoProducto, Value: Dict<NumPrograma, Cantidad>
            var descripcionesPiezas = new Dictionary<int, string>(); // Key: CodigoProducto, Value: Descripción + Código

            // Obtener detalles de cada programa
            foreach (var numeroPrograma in numerosProgramas)
            {
                var detalles = await (from d in _context.Deprofabs
                                      join m in _context.masters on d.CodigoProducto equals m.Codint
                                      where d.NumeroPrograma == numeroPrograma
                                      select new
                                      {
                                          d.CodigoProducto,
                                          Codigo = m.Codigo ?? "",
                                          Descripcion = m.Descripcion ?? "",
                                          d.Cantidad,
                                          NumeroPrograma = numeroPrograma
                                      }).ToListAsync();

                foreach (var detalle in detalles)
                {
                    // Agregar/actualizar descripción de la pieza (fuera del query SQL)
                    if (!descripcionesPiezas.ContainsKey(detalle.CodigoProducto))
                    {
                        // Concatenación en memoria, no en SQL
                        descripcionesPiezas[detalle.CodigoProducto] = 
                            detalle.Codigo + " - " + detalle.Descripcion;
                    }

                    // Agregar cantidad al diccionario
                    if (!piezasPorPrograma.ContainsKey(detalle.CodigoProducto))
                    {
                        piezasPorPrograma[detalle.CodigoProducto] = new Dictionary<int, decimal>();
                    }

                    piezasPorPrograma[detalle.CodigoProducto][numeroPrograma] = detalle.Cantidad;
                }
            }

            // Convertir a lista de DTOs y calcular promedios
            var resultado = new List<PiezaProgramaDto>();

            foreach (var kvp in piezasPorPrograma)
            {
                var codigoProducto = kvp.Key;
                var cantidadesPorPrograma = kvp.Value;

                // Calcular promedio
                var promedio = cantidadesPorPrograma.Values.Any() 
                    ? cantidadesPorPrograma.Values.Average() 
                    : 0;

                resultado.Add(new PiezaProgramaDto
                {
                    CodigoProducto = codigoProducto.ToString(),
                    DescripcionPieza = descripcionesPiezas.ContainsKey(codigoProducto) 
                        ? descripcionesPiezas[codigoProducto] 
                        : "Sin descripción",
                    CantidadesPorPrograma = cantidadesPorPrograma,
                    Promedio = promedio
                });
            }

            return resultado.OrderBy(p => p.CodigoProducto).ToList();
        }

        /// <summary>
        /// Procesa la carga horaria por máquina de los programas seleccionados
        /// Implementa la lógica de CARGARHORASMAQUINA del VB6
        /// </summary>
        private async Task<List<CargaHorariaMaquinaDto>> ProcessCargaHorariaPorMaquina(List<int> numerosProgramas)
        {
            var cargaPorMaquina = new Dictionary<string, CargaHorariaMaquinaDto>(); // Key: CodigoMaquina-CodigoOperacion-Categoria

            foreach (var numeroPrograma in numerosProgramas)
            {
                // Obtener productos del programa
                var productos = await _context.Deprofabs
                    .Where(d => d.NumeroPrograma == numeroPrograma)
                    .Select(d => new { d.CodigoProducto, d.Cantidad })
                    .ToListAsync();

                if (!productos.Any()) continue;

                var codigosProductos = productos.Select(p => p.CodigoProducto).ToList();

                // Obtener operaciones de estos productos (desde SECOPER)
                var operaciones = await _context.Secopers
                    .Where(s => codigosProductos.Contains(s.CodIConj)
                                && (s.StatOper == null || s.StatOper > 0)
                                && (s.AltOper == null || s.AltOper == 0))
                    .OrderBy(s => s.NumeOper)
                    .ToListAsync();

                foreach (var producto in productos)
                {
                    var operacionesProducto = operaciones.Where(o => o.CodIConj == producto.CodigoProducto);

                    foreach (var operacion in operacionesProducto)
                    {
                        var codigoMaquina = operacion.CodMaq0?.Trim();
                        if (string.IsNullOrEmpty(codigoMaquina)) continue;

                        // Obtener información de la máquina
                        var maquina = await _context.padmaqs
                            .AsNoTracking()
                            .FirstOrDefaultAsync(m => m.CodigoPadMaq == codigoMaquina);

                        if (maquina == null) continue;

                        // Calcular tiempos según lógica VB6
                        // MINPREP = 60 * HorsPrep
                        // MINFIJO = 60 * HorsFijo
                        // MINVARI = LoteStan * MinuStan
                        // MINHERR = LoteStan * 60 * HorsMHer / PieCiclo

                        decimal loteStan = operacion.LoteStan ?? 1;
                        if (loteStan < 1) loteStan = 1;

                        decimal pieCiclo = operacion.PieCiclo ?? 1;
                        if (pieCiclo < 1) pieCiclo = 1;

                        decimal minPrep = 60 * (operacion.HorsPrep ?? 0);
                        decimal minFijo = 60 * (operacion.HorsFijo ?? 0);
                        decimal minVari = loteStan * (operacion.MinuStan ?? 0);
                        decimal minHerr = loteStan * 60 * (operacion.HorsMHer ?? 0) / pieCiclo;

                        // Total en segundos = (MINPREP + MINFIJO + MINVARI + MINHERR) * 60
                        decimal tiempoTotalSegundos = (minPrep + minFijo + minVari + minHerr) * 60;

                        // Multiplicar por cantidad del producto
                        tiempoTotalSegundos = (tiempoTotalSegundos * producto.Cantidad) / loteStan;

                        // Obtener categoría y operario
                        int categoria = operacion.CategOps ?? 1;
                        var catoper = await _context.Catopers
                            .AsNoTracking()
                            .FirstOrDefaultAsync(c => c.CodigoCategoria == categoria);

                        // Concatenación en memoria, fuera de query SQL
                        string nombreOperario = catoper?.Descripcion ?? ("Categoría " + categoria.ToString());

                        // Crear clave única por máquina + operación + categoría (en memoria)
                        string key = codigoMaquina + "-" + (operacion.NumeOper ?? 0).ToString() + "-" + categoria.ToString();

                        if (!cargaPorMaquina.ContainsKey(key))
                        {
                            cargaPorMaquina[key] = new CargaHorariaMaquinaDto
                            {
                                CodigoMaquina = int.TryParse(codigoMaquina, out int codMaq) ? codMaq : 0,
                                DescripcionMaquina = maquina.Descripcion ?? codigoMaquina,
                                CodigoOperacion = operacion.NumeOper ?? 0,
                                DescripcionOperacion = operacion.DescOper ?? "Sin descripción",
                                CategoriaOperario = categoria,
                                NombreOperario = nombreOperario,
                                TiemposPorPrograma = new Dictionary<int, decimal>()
                            };
                        }

                        // Agregar/acumular tiempo para este programa
                        if (cargaPorMaquina[key].TiemposPorPrograma.ContainsKey(numeroPrograma))
                        {
                            cargaPorMaquina[key].TiemposPorPrograma[numeroPrograma] += tiempoTotalSegundos;
                        }
                        else
                        {
                            cargaPorMaquina[key].TiemposPorPrograma[numeroPrograma] = tiempoTotalSegundos;
                        }
                    }
                }
            }

            // Calcular promedios
            foreach (var kvp in cargaPorMaquina)
            {
                var carga = kvp.Value;
                if (carga.TiemposPorPrograma.Any())
                {
                    carga.Promedio = carga.TiemposPorPrograma.Values.Average();
                }
            }

            return cargaPorMaquina.Values
                .OrderBy(c => c.DescripcionMaquina)
                .ThenBy(c => c.CodigoOperacion)
                .ToList();
        }
    }
}
