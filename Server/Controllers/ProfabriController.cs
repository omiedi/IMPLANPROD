using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la gestión de Programas de Fabricación (Profabri/Deprofab)
    /// y cálculo de capacidad de producción
    /// </summary>
    [ApiController]
    [Route("/api/profabri")]
    public class ProfabriController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<ProfabriController> _logger;

        public ProfabriController(DataContext context, ILogger<ProfabriController> logger)
        {
            _context = context;
            _logger = logger;
        }

        #region CRUD Básico

        /// <summary>
        /// Obtiene todos los programas de fabricación
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<List<ProfabriDTO>>> GetAllAsync()
        {
            try
            {
                var programas = await _context.Profabris
                    .AsNoTracking()
                    .OrderByDescending(p => p.NumeroPrograma)
                    .Select(p => new ProfabriDTO
                    {
                        Id = p.Id,
                        NumeroPrograma = p.NumeroPrograma,
                        Descripcion = p.Descripcion,
                        FechaCreacion = p.FechaCreacion,
                        CodiEmpr = p.CodiEmpr,
                        Status = p.Status
                    })
                    .ToListAsync();

                return Ok(programas);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener programas de fabricación");
                return StatusCode(500, $"Error al obtener programas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene un programa de fabricación por Id con sus detalles
        /// y los resúmenes de capacidad si ya fueron calculados (RECALMAQ/RECALMOD)
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<ActionResult<ProgramaCapacidadDTO>> GetByIdAsync(int id)
        {
            try
            {
                var programa = await _context.Profabris
                    .FirstOrDefaultAsync(p => p.Id == id);

                if (programa == null)
                {
                    return NotFound($"No se encontró el programa con ID {id}");
                }

                // Obtener detalles con información de productos desde Master
                var detalles = await (from d in _context.Deprofabs
                                      join m in _context.masters on d.CodigoProducto equals m.Codint
                                      where d.NumeroPrograma == programa.NumeroPrograma
                                      select new DeprofabDTO
                                      {
                                          Id = d.Id,
                                          NumeroPrograma = d.NumeroPrograma,
                                          CodigoProducto = d.CodigoProducto,
                                          CodigoExterno = m.Codigo,
                                          Descripcion = m.Descripcion ?? d.Descripcion,
                                          Cantidad = d.Cantidad,
                                          Observaciones = d.Observaciones
                                      }).ToListAsync();

                // Cargar resúmenes de máquinas (RECALMAQ) si existen
                var resumenMaquinas = await _context.CapacidadResumenMaquinas
                    .Where(r => r.NumeroPrograma == programa.NumeroPrograma)
                    .GroupBy(r => new { r.Equipo, r.DescripcionEquipo })
                    .Select(g => new CapacidadMaquinaDTO
                    {
                        CodigoMaquina = g.Key.Equipo,
                        DescripcionMaquina = g.Key.DescripcionEquipo,
                        TiempoRequerido = TimeSpan.FromHours((double)g.Sum(r => r.HorasNecesarias))
                    })
                    .OrderBy(m => m.DescripcionMaquina)
                    .ToListAsync();

                // Cargar resúmenes de M.O.D. (RECALMOD) si existen
                var resumenManoObra = await _context.CapacidadResumenManoObras
                    .Where(r => r.NumeroPrograma == programa.NumeroPrograma)
                    .GroupBy(r => new { r.Categoria, r.DescripcionCategoria })
                    .Select(g => new CapacidadManoObraDTO
                    {
                        Categoria = g.Key.Categoria,
                        DescripcionCategoria = g.Key.DescripcionCategoria,
                        TiempoRequerido = TimeSpan.FromHours((double)g.Sum(r => r.HorasNecesarias))
                    })
                    .OrderBy(m => m.Categoria)
                    .ToListAsync();

                // Calcular totales
                var totalHorasMaquinas = resumenMaquinas.Sum(m => m.TiempoRequerido.TotalHours);
                var totalHorasManoObra = resumenManoObra.Sum(m => m.TiempoRequerido.TotalHours);

                var result = new ProgramaCapacidadDTO
                {
                    Programa = new ProfabriDTO
                    {
                        Id = programa.Id,
                        NumeroPrograma = programa.NumeroPrograma,
                        Descripcion = programa.Descripcion,
                        FechaCreacion = programa.FechaCreacion,
                        CodiEmpr = programa.CodiEmpr,
                        Status = programa.Status
                    },
                    Detalles = detalles,
                    // Si hay resúmenes guardados, cargarlos
                    CapacidadMaquinas = resumenMaquinas,
                    CapacidadManoObra = resumenManoObra,
                    TotalTiempoMaquinas = TimeSpan.FromHours(totalHorasMaquinas),
                    TotalTiempoManoObra = TimeSpan.FromHours(totalHorasManoObra)
                };

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener programa {id}");
                return StatusCode(500, $"Error al obtener programa: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea un nuevo programa de fabricación
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<ProfabriDTO>> CreateAsync([FromBody] ProfabriDTO profabriDTO)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // Generar el siguiente número de programa
                int maxNumero = await _context.Profabris.MaxAsync(p => (int?)p.NumeroPrograma) ?? 0;
                int nuevoNumero = maxNumero + 1;

                var programa = new Profabri
                {
                    NumeroPrograma = nuevoNumero,
                    Descripcion = profabriDTO.Descripcion,
                    FechaCreacion = DateTime.Now,
                    CodiEmpr = profabriDTO.CodiEmpr ?? 1,
                    Status = profabriDTO.Status ?? 0
                };

                _context.Profabris.Add(programa);
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                profabriDTO.Id = programa.Id;
                profabriDTO.NumeroPrograma = programa.NumeroPrograma;
                profabriDTO.FechaCreacion = programa.FechaCreacion;

                return Ok(profabriDTO);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al crear programa");
                return StatusCode(500, $"Error al crear programa: {ex.Message}");
            }
        }

        /// <summary>
        /// Agrega un producto al detalle de un programa
        /// </summary>
        [HttpPost("{programaId:int}/detalle")]
        public async Task<ActionResult<DeprofabDTO>> AddDetalleAsync(int programaId, [FromBody] DeprofabDTO detalleDTO)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // Verificar que el programa existe
                var programa = await _context.Profabris
                    .FirstOrDefaultAsync(p => p.Id == programaId);
                if (programa == null)
                {
                    return NotFound($"No se encontró el programa con ID {programaId}");
                }

                // Verificar que el producto existe
                var producto = await _context.masters
                    .FirstOrDefaultAsync(m => m.Codint == detalleDTO.CodigoProducto);
                if (producto == null)
                {
                    return NotFound($"No se encontró el producto con código {detalleDTO.CodigoProducto}");
                }

                var detalle = new Deprofab
                {
                    NumeroPrograma = programa.NumeroPrograma,
                    CodigoProducto = detalleDTO.CodigoProducto,
                    Descripcion = producto.Descripcion,
                    Cantidad = detalleDTO.Cantidad,
                    Observaciones = detalleDTO.Observaciones
                };

                _context.Deprofabs.Add(detalle);
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                detalleDTO.Id = detalle.Id;
                detalleDTO.CodigoExterno = producto.Codigo;
                detalleDTO.Descripcion = producto.Descripcion;

                return Ok(detalleDTO);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, $"Error al agregar detalle al programa {programaId}");
                return StatusCode(500, $"Error al agregar detalle: {ex.Message}");
            }
        }

        /// <summary>
        /// Elimina un detalle de un programa
        /// </summary>
        [HttpDelete("detalle/{detalleId:int}")]
        public async Task<ActionResult> DeleteDetalleAsync(int detalleId)
        {
            try
            {
                var detalle = await _context.Deprofabs.FindAsync(detalleId);
                if (detalle == null)
                {
                    return NotFound($"No se encontró el detalle con ID {detalleId}");
                }

                _context.Deprofabs.Remove(detalle);
                await _context.SaveChangesAsync();

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al eliminar detalle {detalleId}");
                return StatusCode(500, $"Error al eliminar detalle: {ex.Message}");
            }
        }

        /// <summary>
        /// Limpia los datos de capacidad de un programa (CAPANECE, RECALMAQ, RECALMOD)
        /// Se usa cuando se elimina el último producto de un programa
        /// </summary>
        [HttpDelete("limpiar-capacidad/{numeroPrograma:int}")]
        public async Task<ActionResult> LimpiarCapacidadProgramaAsync(int numeroPrograma)
        {
            try
            {
                // Eliminar registros de CAPANECE
                var registrosCapanece = await _context.Capaneces
                    .Where(c => c.NumeroPrograma == numeroPrograma)
                    .ToListAsync();
                if (registrosCapanece.Any())
                {
                    _context.Capaneces.RemoveRange(registrosCapanece);
                    _logger.LogInformation($"Eliminados {registrosCapanece.Count} registros de CAPANECE para programa {numeroPrograma}");
                }

                // Eliminar registros de RECALMAQ
                var registrosRecalmaq = await _context.CapacidadResumenMaquinas
                    .Where(r => r.NumeroPrograma == numeroPrograma)
                    .ToListAsync();
                if (registrosRecalmaq.Any())
                {
                    _context.CapacidadResumenMaquinas.RemoveRange(registrosRecalmaq);
                    _logger.LogInformation($"Eliminados {registrosRecalmaq.Count} registros de RECALMAQ para programa {numeroPrograma}");
                }

                // Eliminar registros de RECALMOD
                var registrosRecalmod = await _context.CapacidadResumenManoObras
                    .Where(r => r.NumeroPrograma == numeroPrograma)
                    .ToListAsync();
                if (registrosRecalmod.Any())
                {
                    _context.CapacidadResumenManoObras.RemoveRange(registrosRecalmod);
                    _logger.LogInformation($"Eliminados {registrosRecalmod.Count} registros de RECALMOD para programa {numeroPrograma}");
                }

                await _context.SaveChangesAsync();

                _logger.LogInformation($"Datos de capacidad limpiados para programa {numeroPrograma}");
                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al limpiar datos de capacidad del programa {numeroPrograma}");
                return StatusCode(500, $"Error al limpiar datos de capacidad: {ex.Message}");
            }
        }

        /// <summary>
        /// Elimina un programa completo (con CASCADE se eliminan los detalles automáticamente)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<ActionResult> DeleteAsync(int id)
        {
            try
            {
                var programa = await _context.Profabris.FindAsync(id);
                if (programa == null)
                {
                    return NotFound($"No se encontró el programa con ID {id}");
                }

                _context.Profabris.Remove(programa);
                await _context.SaveChangesAsync();

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al eliminar programa {id}");
                return StatusCode(500, $"Error al eliminar programa: {ex.Message}");
            }
        }

        #endregion

        #region Cálculo de Capacidad

        /// <summary>
        /// Calcula la capacidad de producción (máquinas y mano de obra) para un programa
        /// Basado en la lógica del VB6 CALCAPA07 - CalcuCap()
        /// Persiste los resultados en la tabla CAPANECE para auditoría
        /// </summary>
        [HttpPost("calcular-capacidad")]
        public async Task<ActionResult<ProgramaCapacidadDTO>> CalcularCapacidadAsync([FromBody] CalculoCapacidadRequest request)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // VALIDACIÓN 1: Verificar que existe el programa
                var programa = await _context.Profabris
                    .FirstOrDefaultAsync(p => p.Id == request.ProgramaId);

                if (programa == null)
                {
                    return NotFound($"No se encontró el programa de fabricación con ID {request.ProgramaId}. Por favor, verifique que el programa existe.");
                }

                // VALIDACIÓN 2: Verificar que el programa tiene productos
                var detalles = await (from d in _context.Deprofabs
                                      join m in _context.masters on d.CodigoProducto equals m.Codint
                                      where d.NumeroPrograma == programa.NumeroPrograma
                                      select new
                                      {
                                          DetallDTO = new DeprofabDTO
                                          {
                                              Id = d.Id,
                                              NumeroPrograma = d.NumeroPrograma,
                                              CodigoProducto = d.CodigoProducto,
                                              CodigoExterno = m.Codigo,
                                              Descripcion = m.Descripcion ?? d.Descripcion,
                                              Cantidad = d.Cantidad,
                                              Observaciones = d.Observaciones
                                          },
                                          d.Cantidad,
                                          d.CodigoProducto,
                                          CodigoExternoProducto = m.Codigo,
                                          DescripcionProducto = m.Descripcion
                                      }).ToListAsync();

                if (!detalles.Any())
                {
                    return BadRequest($"El programa '{programa.Descripcion}' no tiene productos agregados. Debe agregar al menos un producto antes de calcular la capacidad.");
                }

                // Diccionarios para acumular tiempos por máquina y categoría
                var tiemposPorMaquina = new Dictionary<string, (string descripcion, double horas)>();
                var tiemposPorCategoria = new Dictionary<int, (string descripcion, double horas)>();
                
                // Lista para almacenar registros CAPANECE
                var registrosCapanece = new List<Capanece>();
                
                // Listas para almacenar registros de resúmenes
                var registrosResumenMaquina = new List<CapacidadResumenMaquina>();
                var registrosResumenManoObra = new List<CapacidadResumenManoObra>();

                int productosConProcesos = 0;
                int productosSinProcesos = 0;
                var productosSinProcesosList = new List<string>();

                // Por cada producto en el detalle
                foreach (var detalle in detalles)
                {
                    // VALIDACIÓN 3: Verificar que el producto tiene procesos definidos en SECOPER
                    var procesos = await _context.Secopers
                        .Where(s => s.CodIConj == detalle.CodigoProducto)
                        .ToListAsync();

                    if (!procesos.Any())
                    {
                        productosSinProcesos++;
                        productosSinProcesosList.Add($"{detalle.CodigoExternoProducto} - {detalle.DescripcionProducto}");
                        _logger.LogWarning($"Producto {detalle.CodigoExternoProducto} no tiene procesos definidos en SECOPER");
                        continue; // Saltar este producto pero continuar con los demás
                    }

                    productosConProcesos++;

                    foreach (var proceso in procesos)
                    {
                        // VALIDACIÓN 4: Verificar que el proceso tiene tiempo estándar
                        if (!proceso.MinuStan.HasValue || proceso.MinuStan.Value <= 0)
                        {
                            _logger.LogWarning($"Operación {proceso.NumeOper} del producto {detalle.CodigoExternoProducto} no tiene tiempo estándar definido");
                            continue; // Saltar esta operación pero continuar con las demás
                        }

                        // Calcular horas de producción
                        decimal horasDeProduccion = (detalle.Cantidad * proceso.MinuStan.Value) / 60;

                        // Solo procesar si las horas son significativas (> 0.01 como en VB6)
                        if (horasDeProduccion <= 0.01m)
                        {
                            continue;
                        }

                        // Procesar máquinas (CodMaq0 a CodMaq5)
                        string[] codigosMaquinas = new[] 
                        { 
                            proceso.CodMaq0, proceso.CodMaq1, proceso.CodMaq2, 
                            proceso.CodMaq3, proceso.CodMaq4, proceso.CodMaq5 
                        };

                        foreach (var codMaq in codigosMaquinas)
                        {
                            if (string.IsNullOrWhiteSpace(codMaq)) continue;

                            // Obtener información de la máquina desde PADMAQ
                            var maquina = await _context.padmaqs
                                .AsNoTracking()
                                .FirstOrDefaultAsync(p => p.CodigoPadMaq == codMaq);

                            if (maquina == null)
                            {
                                _logger.LogWarning($"Máquina {codMaq} no encontrada en PADMAQ para operación {proceso.NumeOper}");
                                continue;
                            }

                            // Acumular tiempos por máquina
                            string codigoMaquina = maquina.CodigoPadMaq.Trim();
                            double horas = (double)horasDeProduccion;
                            
                            if (tiemposPorMaquina.ContainsKey(codigoMaquina))
                            {
                                var actual = tiemposPorMaquina[codigoMaquina];
                                tiemposPorMaquina[codigoMaquina] = (actual.descripcion, actual.horas + horas);
                            }
                            else
                            {
                                tiemposPorMaquina[codigoMaquina] = (maquina.Descripcion ?? codigoMaquina, horas);
                            }

                            // Crear registro CAPANECE (como lo hace el VB6)
                            var registroCapanece = new Capanece
                            {
                                Equipo = codigoMaquina.PadRight(6).Substring(0, 6),
                                DescripcionEquipo = (maquina.Descripcion ?? "").PadRight(48).Substring(0, 48),
                                CodigoPieza = detalle.CodigoProducto,
                                DescripcionPieza = (detalle.DescripcionProducto ?? "").PadRight(64).Substring(0, 64),
                                NumeroOperacion = proceso.NumeOper,
                                DescripcionOperacion = (proceso.DescOper ?? "").PadRight(48).Substring(0, 48),
                                CantidadProgramada = detalle.Cantidad,
                                CodigoProductoDestino = detalle.CodigoProducto,
                                HorasNecesarias = (decimal)horas,
                                TotalHorasPorEquipo = (decimal)horas,
                                NumeroPrograma = programa.NumeroPrograma
                            };

                            registrosCapanece.Add(registroCapanece);
                            
                            // Crear registro para RECALMAQ (máquina)
                            var registroResumenMaquina = new CapacidadResumenMaquina
                            {
                                Equipo = codigoMaquina,
                                DescripcionEquipo = maquina.Descripcion ?? string.Empty,
                                CodigoPieza = detalle.CodigoProducto,
                                DescripcionPieza = detalle.DescripcionProducto ?? string.Empty,
                                NumeroOperacion = (int)proceso.NumeOper,
                                DescripcionOperacion = proceso.DescOper ?? string.Empty,
                                CantidadProgramada = detalle.Cantidad,
                                CodigoProductoDestino = detalle.CodigoProducto,
                                HorasNecesarias = (decimal)horas,
                                TotalHorasPorEquipo = (decimal)horas,
                                NumeroPrograma = programa.NumeroPrograma
                            };
                            registrosResumenMaquina.Add(registroResumenMaquina);
                        }

                        // Calcular tiempo de mano de obra (MOD)
                        int categoria = proceso.CategOps ?? 1; // Si no tiene categoría, usar 1 por defecto
                        if (categoria < 1) categoria = 1;

                        double cantidadOperarios = (double)(proceso.CantiOps ?? 1);
                        double horasMOD = (double)horasDeProduccion * cantidadOperarios;

                        // Obtener descripción de la categoría
                        var catoper = await _context.Catopers
                            .AsNoTracking()
                            .FirstOrDefaultAsync(c => c.CodigoCategoria == categoria);

                        string descripcionCategoria = catoper?.Descripcion ?? $"Categoría {categoria}";

                        if (tiemposPorCategoria.ContainsKey(categoria))
                        {
                            var actual = tiemposPorCategoria[categoria];
                            tiemposPorCategoria[categoria] = (actual.descripcion, actual.horas + horasMOD);
                        }
                        else
                        {
                            tiemposPorCategoria[categoria] = (descripcionCategoria, horasMOD);
                        }
                        
                        // Crear registro para RECALMOD (mano de obra)
                        var registroResumenManoObra = new CapacidadResumenManoObra
                        {
                            Categoria = categoria,
                            DescripcionCategoria = descripcionCategoria,
                            HorasNecesarias = (decimal)horasMOD,
                            TotalHorasPorCategoria = (decimal)horasMOD,
                            NumeroPrograma = programa.NumeroPrograma
                        };
                        registrosResumenManoObra.Add(registroResumenManoObra);
                    }
                }

                // VALIDACIÓN 5: Verificar que se calculó capacidad para al menos un producto
                if (productosConProcesos == 0)
                {
                    string mensaje = "No se pudo calcular la capacidad porque ninguno de los productos tiene procesos definidos en el sistema.\n\n";
                    mensaje += "Productos sin procesos:\n" + string.Join("\n", productosSinProcesosList.Take(10));
                    
                    if (productosSinProcesosList.Count > 10)
                    {
                        mensaje += $"\n... y {productosSinProcesosList.Count - 10} más.";
                    }
                    
                    mensaje += "\n\nPor favor, configure los procesos de fabricación (SECOPER) para estos productos.";
                    
                    return BadRequest(mensaje);
                }

                // VALIDACIÓN 6: Verificar que se calcularon máquinas o mano de obra
                if (!tiemposPorMaquina.Any() && !tiemposPorCategoria.Any())
                {
                    return BadRequest($"No se pudo calcular capacidad. Verifique que los productos tengan:\n" +
                        "1. Procesos definidos en SECOPER\n" +
                        "2. Máquinas asignadas en cada operación\n" +
                        "3. Tiempos estándar configurados");
                }

                // PERSISTENCIA: Limpiar registros anteriores de este programa en CAPANECE
                var registrosAnteriores = await _context.Capaneces
                    .Where(c => c.NumeroPrograma == programa.NumeroPrograma)
                    .ToListAsync();

                if (registrosAnteriores.Any())
                {
                    _context.Capaneces.RemoveRange(registrosAnteriores);
                    _logger.LogInformation($"Se eliminaron {registrosAnteriores.Count} registros anteriores de CAPANECE para el programa {programa.NumeroPrograma}");
                }

                // PERSISTENCIA: Limpiar registros anteriores de RECALMAQ (resumen máquinas)
                var resumenMaquinasAnteriores = await _context.CapacidadResumenMaquinas
                    .Where(r => r.NumeroPrograma == programa.NumeroPrograma)
                    .ToListAsync();

                if (resumenMaquinasAnteriores.Any())
                {
                    _context.CapacidadResumenMaquinas.RemoveRange(resumenMaquinasAnteriores);
                    _logger.LogInformation($"Se eliminaron {resumenMaquinasAnteriores.Count} registros anteriores de RECALMAQ para el programa {programa.NumeroPrograma}");
                }

                // PERSISTENCIA: Limpiar registros anteriores de RECALMOD (resumen M.O.D.)
                var resumenManoObraAnteriores = await _context.CapacidadResumenManoObras
                    .Where(r => r.NumeroPrograma == programa.NumeroPrograma)
                    .ToListAsync();

                if (resumenManoObraAnteriores.Any())
                {
                    _context.CapacidadResumenManoObras.RemoveRange(resumenManoObraAnteriores);
                    _logger.LogInformation($"Se eliminaron {resumenManoObraAnteriores.Count} registros anteriores de RECALMOD para el programa {programa.NumeroPrograma}");
                }

                // PERSISTENCIA: Guardar nuevos registros en CAPANECE
                if (registrosCapanece.Any())
                {
                    await _context.Capaneces.AddRangeAsync(registrosCapanece);
                    _logger.LogInformation($"Se crearon {registrosCapanece.Count} registros en CAPANECE para el programa {programa.NumeroPrograma}");
                }
                
                // PERSISTENCIA: Guardar nuevos resúmenes de máquinas (RECALMAQ)
                if (registrosResumenMaquina.Any())
                {
                    await _context.CapacidadResumenMaquinas.AddRangeAsync(registrosResumenMaquina);
                    _logger.LogInformation($"Se guardaron {registrosResumenMaquina.Count} registros en RECALMAQ para el programa {programa.NumeroPrograma}");
                }
                
                // PERSISTENCIA: Guardar nuevos resúmenes de M.O.D. (RECALMOD)
                if (registrosResumenManoObra.Any())
                {
                    await _context.CapacidadResumenManoObras.AddRangeAsync(registrosResumenManoObra);
                    _logger.LogInformation($"Se guardaron {registrosResumenManoObra.Count} registros en RECALMOD para el programa {programa.NumeroPrograma}");
                }

                // Actualizar fecha de último cálculo en PROFABRI
                programa.LastUpdate = DateTime.Now;
                _context.Profabris.Update(programa);

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                // Construir lista de máquinas
                var capacidadMaquinas = tiemposPorMaquina
                    .Select(kvp => new CapacidadMaquinaDTO
                    {
                        CodigoMaquina = kvp.Key,
                        DescripcionMaquina = kvp.Value.descripcion,
                        TiempoRequerido = TimeSpan.FromHours(kvp.Value.horas)
                    })
                    .OrderBy(m => m.DescripcionMaquina)
                    .ToList();

                // Construir lista de mano de obra
                var capacidadManoObra = tiemposPorCategoria
                    .Select(kvp => new CapacidadManoObraDTO
                    {
                        Categoria = kvp.Key,
                        DescripcionCategoria = kvp.Value.descripcion,
                        TiempoRequerido = TimeSpan.FromHours(kvp.Value.horas)
                    })
                    .OrderBy(m => m.Categoria)
                    .ToList();

                // Calcular totales
                double totalHorasMaquinas = tiemposPorMaquina.Values.Sum(v => v.horas);
                double totalHorasManoObra = tiemposPorCategoria.Values.Sum(v => v.horas);

                var result = new ProgramaCapacidadDTO
                {
                    Programa = new ProfabriDTO
                    {
                        Id = programa.Id,
                        NumeroPrograma = programa.NumeroPrograma,
                        Descripcion = programa.Descripcion,
                        FechaCreacion = programa.FechaCreacion,
                        CodiEmpr = programa.CodiEmpr,
                        Status = programa.Status
                    },
                    Detalles = detalles.Select(d => d.DetallDTO).ToList(),
                    CapacidadMaquinas = capacidadMaquinas,
                    CapacidadManoObra = capacidadManoObra,
                    TotalTiempoMaquinas = TimeSpan.FromHours(totalHorasMaquinas),
                    TotalTiempoManoObra = TimeSpan.FromHours(totalHorasManoObra)
                };

                // Mensaje informativo si hay productos sin procesos
                if (productosSinProcesos > 0)
                {
                    _logger.LogWarning($"Se calculó capacidad con {productosConProcesos} productos. {productosSinProcesos} productos fueron omitidos por no tener procesos configurados.");
                }

                return Ok(result);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, $"Error al calcular capacidad del programa {request.ProgramaId}");
                return StatusCode(500, $"Error al calcular capacidad: {ex.Message}");
            }
        }

        #endregion

        #region Generación Automática desde Pedidos y Remitos

        /// <summary>
        /// Genera un programa de fabricación desde pedidos pendientes
        /// Implementa la lógica de CALCAPEDCLI del VB6
        /// </summary>
        [HttpPost("generar-desde-pedidos")]
        public async Task<ActionResult<ProfabriDTO>> GenerarDesdePedidosAsync([FromBody] FiltroGeneracionDTO filtro)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // Validar fechas
                if (filtro.FechaDesde > filtro.FechaHasta)
                {
                    return BadRequest("La fecha 'Desde' debe ser menor o igual a 'Hasta'");
                }

                // Consultar pedidos pendientes
                var pedidosPendientes = await _context.Peddetas
                    .Where(p => p.FeSolEnt >= filtro.FechaDesde && 
                                p.FeSolEnt <= filtro.FechaHasta &&
                                p.Status < 8 && // No anulado ni cumplido
                                p.CantEnt < p.CanPed - 0.005m && // Tiene saldo pendiente
                                (filtro.ClienteId == null || p.NroClie == filtro.ClienteId))
                    .GroupBy(p => p.CodiInt)
                    .Select(g => new
                    {
                        CodigoProducto = g.Key ?? 0,
                        CantidadTotal = g.Sum(p => (p.CanPed ?? 0) - (p.CantEnt ?? 0))
                    })
                    .Where(x => x.CantidadTotal > 0)
                    .ToListAsync();

                if (!pedidosPendientes.Any())
                {
                    return BadRequest("No se encontraron pedidos pendientes en el rango de fechas especificado");
                }

                // Crear el programa
                int maxNumero = await _context.Profabris.MaxAsync(p => (int?)p.NumeroPrograma) ?? 0;
                int nuevoNumero = maxNumero + 1;

                string descripcion = string.IsNullOrEmpty(filtro.DescripcionPrograma)
                    ? $"Pedidos {filtro.FechaDesde:dd/MM/yyyy}-{filtro.FechaHasta:dd/MM/yyyy}"
                    : filtro.DescripcionPrograma;

                var programa = new Profabri
                {
                    NumeroPrograma = nuevoNumero,
                    Descripcion = descripcion,
                    FechaCreacion = DateTime.Now,
                    CodiEmpr = 1,
                    Status = 1 // Activo
                };

                _context.Profabris.Add(programa);
                await _context.SaveChangesAsync();

                // Agregar detalles
                foreach (var pedido in pedidosPendientes)
                {
                    var producto = await _context.masters.FirstOrDefaultAsync(m => m.Codint == pedido.CodigoProducto);
                    if (producto != null)
                    {
                        var detalle = new Deprofab
                        {
                            NumeroPrograma = programa.NumeroPrograma,
                            CodigoProducto = pedido.CodigoProducto,
                            Descripcion = producto.Descripcion,
                            Cantidad = pedido.CantidadTotal,
                            Observaciones = "Generado desde pedidos pendientes"
                        };

                        _context.Deprofabs.Add(detalle);
                    }
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return Ok(new ProfabriDTO
                {
                    Id = programa.Id,
                    NumeroPrograma = programa.NumeroPrograma,
                    Descripcion = programa.Descripcion,
                    FechaCreacion = programa.FechaCreacion,
                    CodiEmpr = programa.CodiEmpr,
                    Status = programa.Status
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al generar programa desde pedidos");
                return StatusCode(500, $"Error al generar programa: {ex.Message}");
            }
        }

        /// <summary>
        /// Genera un programa de fabricación desde remitos
        /// Implementa la lógica de CALCAREMITO del VB6
        /// </summary>
        [HttpPost("generar-desde-remitos")]
        public async Task<ActionResult<ProfabriDTO>> GenerarDesdeRemitosAsync([FromBody] FiltroGeneracionDTO filtro)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // Validar fechas
                if (filtro.FechaDesde > filtro.FechaHasta)
                {
                    return BadRequest("La fecha 'Desde' debe ser menor o igual a 'Hasta'");
                }

                // Consultar remitos en el rango de fechas
                var remitos = await _context.Movistos
                    .Where(m => m.FechMov >= filtro.FechaDesde &&
                                m.FechMov <= filtro.FechaHasta &&
                                m.CodiMovi == "RT") // Tipo Remito
                    .GroupBy(m => m.Cod_Inte)
                    .Select(g => new
                    {
                        CodigoProducto = g.Key ?? 0,
                        CantidadTotal = g.Sum(m => (m.CantSalid ?? 0) - (m.CantIngre ?? 0))
                    })
                    .Where(x => x.CantidadTotal > 0)
                    .ToListAsync();

                if (!remitos.Any())
                {
                    return BadRequest("No se encontraron remitos en el rango de fechas especificado");
                }

                // Crear el programa
                int maxNumero = await _context.Profabris.MaxAsync(p => (int?)p.NumeroPrograma) ?? 0;
                int nuevoNumero = maxNumero + 1;

                string descripcion = string.IsNullOrEmpty(filtro.DescripcionPrograma)
                    ? $"Remitos {filtro.FechaDesde:dd/MM/yyyy}-{filtro.FechaHasta:dd/MM/yyyy}"
                    : filtro.DescripcionPrograma;

                var programa = new Profabri
                {
                    NumeroPrograma = nuevoNumero,
                    Descripcion = descripcion,
                    FechaCreacion = DateTime.Now,
                    CodiEmpr = 1,
                    Status = 1
                };

                _context.Profabris.Add(programa);
                await _context.SaveChangesAsync();

                // Agregar detalles
                foreach (var remito in remitos)
                {
                    var producto = await _context.masters.FirstOrDefaultAsync(m => m.Codint == remito.CodigoProducto);
                    if (producto != null)
                    {
                        var detalle = new Deprofab
                        {
                            NumeroPrograma = programa.NumeroPrograma,
                            CodigoProducto = remito.CodigoProducto,
                            Descripcion = producto.Descripcion,
                            Cantidad = remito.CantidadTotal,
                            Observaciones = "Generado desde remitos"
                        };

                        _context.Deprofabs.Add(detalle);
                    }
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return Ok(new ProfabriDTO
                {
                    Id = programa.Id,
                    NumeroPrograma = programa.NumeroPrograma,
                    Descripcion = programa.Descripcion,
                    FechaCreacion = programa.FechaCreacion,
                    CodiEmpr = programa.CodiEmpr,
                    Status = programa.Status
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al generar programa desde remitos");
                return StatusCode(500, $"Error al generar programa: {ex.Message}");
            }
        }

        #endregion
    }
}
