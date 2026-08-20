using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using System.Text;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/repara")]
    public class ReparaController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<ReparaController> _logger;
        private readonly IEmpresaDependenciaService _empresaDependenciaService;

        public ReparaController(
            DataContext context, 
            ILogger<ReparaController> logger,
            IEmpresaDependenciaService empresaDependenciaService)
        {
            _context = context;
            _logger = logger;
            _empresaDependenciaService = empresaDependenciaService;
        }

        /// <summary>
        /// Agrega una anotación manual al campo memo REPARA.AnotaRep usando formato posicional VB6.
        /// </summary>
        [HttpPost("agregar-anotacion")]
        public async Task<IActionResult> AgregarAnotacionAsync([FromBody] AgregarAnotacionReparacionRequest request)
        {
            try
            {
                if (request.NumRepa <= 0)
                {
                    return BadRequest("NumRepa inválido");
                }

                var texto = (request.Texto ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(texto))
                {
                    return BadRequest("El texto de la anotación es obligatorio");
                }

                var reparacion = await _context.Reparas.FirstOrDefaultAsync(x => x.NumRepa == request.NumRepa);
                if (reparacion == null)
                {
                    return NotFound($"No se encontró la reparación N° {request.NumRepa}");
                }

                AgregarAnotacionPosicional(reparacion, texto, request.Usuario);

                _context.Reparas.Update(reparacion);
                await _context.SaveChangesAsync();

                return Ok(new AgregarAnotacionReparacionResponse
                {
                    NumRepa = reparacion.NumRepa,
                    AnotaRep = reparacion.AnotaRep
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al agregar anotación manual para NumRepa={NumRepa}", request.NumRepa);
                return StatusCode(500, $"Error al agregar anotación: {ex.Message}");
            }
        }

        /// <summary>
        /// Agrega anotación de presupuestación de reparación replicando el formato posicional VB6.
        /// </summary>
        [HttpPost("agregar-anotacion-presupuesto")]
        public async Task<IActionResult> AgregarAnotacionPresupuestoAsync([FromBody] AgregarAnotacionPresupuestoRequest request)
        {
            try
            {
                if (request.NumRepa <= 0)
                {
                    return BadRequest("NumRepa inválido");
                }

                var reparacion = await _context.Reparas.FirstOrDefaultAsync(x => x.NumRepa == request.NumRepa);
                if (reparacion == null)
                {
                    return NotFound($"No se encontró la reparación N° {request.NumRepa}");
                }

                var usuario = ObtenerUsuarioActual();
                AgregarAnotacionPresupuestoPosicional(reparacion, request, usuario);

                _context.Reparas.Update(reparacion);
                await _context.SaveChangesAsync();

                return Ok(new AgregarAnotacionReparacionResponse
                {
                    NumRepa = reparacion.NumRepa,
                    AnotaRep = reparacion.AnotaRep
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al agregar anotación de presupuesto para NumRepa={NumRepa}", request.NumRepa);
                return StatusCode(500, $"Error al agregar anotación de presupuesto: {ex.Message}");
            }
        }

        private static void AgregarAnotacionPosicional(Repara reparacion, string texto, string? usuario = null)
        {
            var textoLimpio = (texto ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(textoLimpio))
            {
                return;
            }

            var fechaTexto = DateTime.Now.ToString("dd/MM/yy");
            var usuarioTexto = string.IsNullOrWhiteSpace(usuario) ? "SISTEMA" : usuario.Trim();

            var lineas = FraccionarTexto(textoLimpio, 56);
            var anotacion = new StringBuilder();

            for (var i = 0; i < lineas.Count; i++)
            {
                var caracteres = new string(' ', 96).ToCharArray();

                if (i == 0)
                {
                    ReemplazarEnPosicion(caracteres, fechaTexto, 0, 8);
                    ReemplazarEnPosicion(caracteres, usuarioTexto, 68, 24);
                }

                ReemplazarEnPosicion(caracteres, lineas[i], 10, 56);
                anotacion.AppendLine(new string(caracteres));
            }

            anotacion.AppendLine(new string('=', 88));
            reparacion.AnotaRep = (reparacion.AnotaRep ?? string.Empty) + anotacion;
        }

        private static List<string> FraccionarTexto(string texto, int longitudMaxima)
        {
            var lineas = new List<string>();
            if (string.IsNullOrWhiteSpace(texto))
            {
                return lineas;
            }

            var palabras = texto.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var lineaActual = new StringBuilder();

            foreach (var palabra in palabras)
            {
                if (lineaActual.Length + palabra.Length + 1 <= longitudMaxima)
                {
                    if (lineaActual.Length > 0)
                    {
                        lineaActual.Append(' ');
                    }

                    lineaActual.Append(palabra);
                }
                else
                {
                    if (lineaActual.Length > 0)
                    {
                        lineas.Add(lineaActual.ToString());
                        lineaActual.Clear();
                    }

                    lineaActual.Append(palabra);
                }
            }

            if (lineaActual.Length > 0)
            {
                lineas.Add(lineaActual.ToString());
            }

            return lineas;
        }

        private static void ReemplazarEnPosicion(char[] destino, string texto, int posicionInicio, int longitudMaxima)
        {
            if (string.IsNullOrEmpty(texto) || posicionInicio < 0 || posicionInicio >= destino.Length)
            {
                return;
            }

            var textoAInsertar = texto.Length > longitudMaxima ? texto.Substring(0, longitudMaxima) : texto;
            for (var i = 0; i < textoAInsertar.Length && (posicionInicio + i) < destino.Length; i++)
            {
                destino[posicionInicio + i] = textoAInsertar[i];
            }
        }

        private static void AgregarAnotacionPresupuestoPosicional(Repara reparacion, AgregarAnotacionPresupuestoRequest request, string usuario)
        {
            var fechaTexto = DateTime.Now.ToString("dd/MM/yy");
            var usuarioTexto = string.IsNullOrWhiteSpace(usuario) ? "SISTEMA" : usuario.Trim();

            var lineas = new List<string>
            {
                CrearLineaPosicional92(fechaTexto, usuarioTexto, "Recepción Material para Presupuestación de Reparación:", 10, 56),
                CrearLineaPosicional92(null, null, string.Empty, 10, 56),
                CrearLineaPosicional92(null, null, $"Cliente: {request.Cliente ?? string.Empty}", 10, 56),
                CrearLineaPosicional92(null, null, $"    Tel: {(request.Telefono ?? string.Empty).Trim()}", 10, 56),
                CrearLineaPosicional92(null, null, $"    Fax: {(request.Fax ?? string.Empty).Trim()}", 10, 56),
                CrearLineaPosicional92(null, null, string.Empty, 10, 56),
                CrearLineaPosicional92(null, null, $"  S/Nro: {request.NumeroSerie ?? string.Empty}", 10, 56),
                CrearLineaPosicional92(null, null, $"Referen: {request.Referencia ?? string.Empty}", 10, 56),
                CrearLineaPosicional92(null, null, string.Empty, 10, 56)
            };

            var lineasSitio = FraccionarTexto(request.Sitio ?? string.Empty, 46);
            for (var i = 0; i < lineasSitio.Count; i++)
            {
                var linea = CrearLineaPosicional92(null, null, lineasSitio[i], 19, 46);
                if (i == 0)
                {
                    linea = AplicarTextoEnLinea(linea, "  Sitio: ", 10, 9);
                }

                lineas.Add(linea);
            }

            lineas.Add(CrearLineaPosicional92(null, null, string.Empty, 10, 56));

            var lineasObserv = FraccionarTexto(request.Observaciones ?? string.Empty, 46);
            for (var i = 0; i < lineasObserv.Count; i++)
            {
                var linea = CrearLineaPosicional92(null, null, lineasObserv[i], 19, 46);
                if (i == 0)
                {
                    linea = AplicarTextoEnLinea(linea, " Observ: ", 10, 9);
                }

                lineas.Add(linea);
            }

            var anotacion = new StringBuilder();
            foreach (var linea in lineas)
            {
                anotacion.AppendLine(linea);
            }

            anotacion.AppendLine(new string('=', 88));

            reparacion.AnotaRep = (reparacion.AnotaRep ?? string.Empty) + anotacion;
        }

        private static string CrearLineaPosicional92(string? fecha, string? usuario, string texto, int posicionTexto, int longitudTexto)
        {
            var caracteres = new string(' ', 92).ToCharArray();

            if (!string.IsNullOrWhiteSpace(fecha))
            {
                ReemplazarEnPosicion(caracteres, fecha, 0, 8);
            }

            if (!string.IsNullOrWhiteSpace(usuario))
            {
                ReemplazarEnPosicion(caracteres, usuario, 68, 24);
            }

            ReemplazarEnPosicion(caracteres, texto ?? string.Empty, posicionTexto, longitudTexto);
            return new string(caracteres);
        }

        private static string AplicarTextoEnLinea(string linea, string texto, int posicion, int longitud)
        {
            var caracteres = (linea ?? string.Empty).PadRight(92).ToCharArray();
            ReemplazarEnPosicion(caracteres, texto, posicion, longitud);
            return new string(caracteres);
        }

        private string ObtenerUsuarioActual()
        {
            var usuario = User?.FindFirst("idflexoft")?.Value
                ?? User?.FindFirst(ClaimTypes.NameIdentifier)?.Value
                ?? User?.Identity?.Name;

            return string.IsNullOrWhiteSpace(usuario) ? "SISTEMA" : usuario.Trim();
        }

        /// <summary>
        /// Obtiene todas las reparaciones con paginación y filtros individuales
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync(
            [FromQuery] PaginationDTO pagination,
            [FromQuery] int? filtroId = null,
            [FromQuery] int? filtroNumRepa = null,
            [FromQuery] DateTime? filtroFechaDesde = null,
            [FromQuery] DateTime? filtroFechaHasta = null,
            [FromQuery] int? filtroNumeCli = null,
            [FromQuery] string? filtroRasoCli = null,
            [FromQuery] string? filtroCodRepa = null,
            [FromQuery] string? filtroNumeroSerie = null,
            [FromQuery] string? filtroStatus = null)
        {
            try
            {
                var queryable = _context.Reparas.AsQueryable();

                // Aplicar filtros individuales
                if (filtroId.HasValue && filtroId > 0)
                {
                    queryable = queryable.Where(x => x.Id == filtroId);
                }

                if (filtroNumRepa.HasValue && filtroNumRepa > 0)
                {
                    queryable = queryable.Where(x => x.NumRepa == filtroNumRepa);
                }

                if (filtroFechaDesde.HasValue)
                {
                    queryable = queryable.Where(x => x.FechRecep >= filtroFechaDesde.Value);
                }

                if (filtroFechaHasta.HasValue)
                {
                    queryable = queryable.Where(x => x.FechRecep <= filtroFechaHasta.Value);
                }

                if (filtroNumeCli.HasValue && filtroNumeCli > 0)
                {
                    queryable = queryable.Where(x => x.Nume_Cli == filtroNumeCli);
                }

                if (!string.IsNullOrWhiteSpace(filtroRasoCli))
                {
                    var filtro = filtroRasoCli.ToLower();
                    queryable = queryable.Where(x => x.Raso_Cli!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroCodRepa))
                {
                    var filtro = filtroCodRepa.ToLower();
                    queryable = queryable.Where(x => x.CodRepa!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroNumeroSerie))
                {
                    var filtro = filtroNumeroSerie.ToLower();
                    queryable = queryable.Where(x => x.NumeroSerie!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroStatus))
                {
                    if (short.TryParse(filtroStatus, out var status))
                    {
                        queryable = queryable.Where(x => x.Status == status);
                    }
                }

                // Aplicar filtro general si existe
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var filtro = pagination.Filter.ToLower();
                    queryable = queryable.Where(x =>
                        x.NumRepa.ToString().Contains(filtro) ||
                        x.NumeroSerie!.ToLower().Contains(filtro) ||
                        x.Raso_Cli!.ToLower().Contains(filtro) ||
                        x.CodRepa!.ToLower().Contains(filtro) ||
                        x.Referen!.ToLower().Contains(filtro));
                }

                return Ok(await queryable
                    .OrderByDescending(x => x.FechRecep)
                    .Paginate(pagination)
                    .ToListAsync());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener reparaciones");
                return StatusCode(500, $"Error al obtener reparaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación con filtros individuales
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages(
            [FromQuery] PaginationDTO pagination,
            [FromQuery] int? filtroId = null,
            [FromQuery] int? filtroNumRepa = null,
            [FromQuery] DateTime? filtroFechaDesde = null,
            [FromQuery] DateTime? filtroFechaHasta = null,
            [FromQuery] int? filtroNumeCli = null,
            [FromQuery] string? filtroRasoCli = null,
            [FromQuery] string? filtroCodRepa = null,
            [FromQuery] string? filtroNumeroSerie = null,
            [FromQuery] string? filtroStatus = null)
        {
            try
            {
                var queryable = _context.Reparas.AsQueryable();

                // Aplicar los mismos filtros que en GetAsync
                if (filtroId.HasValue && filtroId > 0)
                {
                    queryable = queryable.Where(x => x.Id == filtroId);
                }

                if (filtroNumRepa.HasValue && filtroNumRepa > 0)
                {
                    queryable = queryable.Where(x => x.NumRepa == filtroNumRepa);
                }

                if (filtroFechaDesde.HasValue)
                {
                    queryable = queryable.Where(x => x.FechRecep >= filtroFechaDesde.Value);
                }

                if (filtroFechaHasta.HasValue)
                {
                    queryable = queryable.Where(x => x.FechRecep <= filtroFechaHasta.Value);
                }

                if (filtroNumeCli.HasValue && filtroNumeCli > 0)
                {
                    queryable = queryable.Where(x => x.Nume_Cli == filtroNumeCli);
                }

                if (!string.IsNullOrWhiteSpace(filtroRasoCli))
                {
                    var filtro = filtroRasoCli.ToLower();
                    queryable = queryable.Where(x => x.Raso_Cli!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroCodRepa))
                {
                    var filtro = filtroCodRepa.ToLower();
                    queryable = queryable.Where(x => x.CodRepa!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroNumeroSerie))
                {
                    var filtro = filtroNumeroSerie.ToLower();
                    queryable = queryable.Where(x => x.NumeroSerie!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroStatus))
                {
                    if (short.TryParse(filtroStatus, out var status))
                    {
                        queryable = queryable.Where(x => x.Status == status);
                    }
                }

                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var filtro = pagination.Filter.ToLower();
                    queryable = queryable.Where(x =>
                        x.NumRepa.ToString().Contains(filtro) ||
                        x.NumeroSerie!.ToLower().Contains(filtro) ||
                        x.Raso_Cli!.ToLower().Contains(filtro) ||
                        x.CodRepa!.ToLower().Contains(filtro) ||
                        x.Referen!.ToLower().Contains(filtro));
                }

                double count = await queryable.CountAsync();
                double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular páginas");
                return StatusCode(500, $"Error al calcular páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene una reparación por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<ActionResult> GetAsync(int id)
        {
            try
            {
                var repara = await _context.Reparas.FirstOrDefaultAsync(x => x.NumRepa == id);
                if (repara == null)
                {
                    return NotFound();
                }
                return Ok(repara);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener reparación {Id}", id);
                return StatusCode(500, $"Error al obtener reparación: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el siguiente número de reparación disponible
        /// </summary>
        [HttpGet("siguiente-numero")]
        public async Task<ActionResult<int>> GetSiguienteNumeroAsync()
        {
            try
            {
                var maxNumRepa = await _context.Reparas.MaxAsync(x => (int?)x.NumRepa) ?? 0;
                var siguienteNumero = maxNumRepa + 1;
                _logger.LogInformation("Siguiente número de reparación: {SiguienteNumero}", siguienteNumero);
                return Ok(siguienteNumero);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener siguiente número de reparación");
                return StatusCode(500, $"Error al obtener siguiente número: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea una nueva reparación
        /// Equivalente a GRAPRESU en VB6
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Repara repara)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // Validaciones según VB6
                if (string.IsNullOrWhiteSpace(repara.NumeroSerie))
                {
                    return BadRequest("Falta Número de Serie o Identificación del Equipo.");
                }

                if (repara.NumeroSerie.Length > 32)
                {
                    return BadRequest("Exceso Longitud Número de Serie (máximo 32 caracteres)");
                }

                if (!string.IsNullOrEmpty(repara.Referen2) && repara.Referen2.Length > 64)
                {
                    return BadRequest("Exceso Longitud Referencia (máximo 64 caracteres)");
                }

                if (string.IsNullOrWhiteSpace(repara.Sitio))
                {
                    return BadRequest("Falta Sitio de Ingreso de la Bomba.");
                }

                // Obtener el siguiente número disponible
                var maxNumRepa = await _context.Reparas.MaxAsync(x => (int?)x.NumRepa) ?? 0;
                var siguienteNumero = maxNumRepa + 1;

                // Guardar el número que venía de pantalla para comparación
                var numeroPantalla = repara.NumRepa;
                bool numeroModificado = false;

                // Si el número de pantalla no coincide con el siguiente, usar el siguiente disponible
                if (numeroPantalla != siguienteNumero)
                {
                    numeroModificado = true;
                    repara.NumRepa = siguienteNumero;
                    _logger.LogWarning("Número de pantalla {NumeroPantalla} diferente al siguiente disponible {SiguienteNumero}. Se usará {SiguienteNumero}",
                        numeroPantalla, siguienteNumero, siguienteNumero);
                }

                // Establecer fecha de recepción si no viene
                if (!repara.FechRecep.HasValue)
                {
                    repara.FechRecep = DateTime.Now;
                }

                // Status inicial
                if (!repara.Status.HasValue)
                {
                    repara.Status = 0; // Activo
                }

                _context.Add(repara);
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();
                
                _logger.LogInformation("Reparación creada: {NumRepa}", repara.NumRepa);
                
                // Devolver información sobre si se modificó el número
                return Ok(new 
                { 
                    repara = repara,
                    numeroModificado = numeroModificado,
                    numeroPantalla = numeroPantalla,
                    numeroAsignado = repara.NumRepa
                });
            }
            catch (DbUpdateException dbUpdateException)
            {
                await transaction.RollbackAsync();
                _logger.LogError(dbUpdateException, "Error de BD al crear reparación");
                return BadRequest($"Error al crear reparación: {dbUpdateException.InnerException?.Message ?? dbUpdateException.Message}");
            }
            catch (Exception exception)
            {
                await transaction.RollbackAsync();
                _logger.LogError(exception, "Error al crear reparación");
                return BadRequest($"Error al crear reparación: {exception.Message}");
            }
        }

        /// <summary>
        /// Actualiza una reparación existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Repara repara)
        {
            try
            {
                // Validaciones
                if (string.IsNullOrWhiteSpace(repara.NumeroSerie))
                {
                    return BadRequest("Falta Número de Serie o Identificación del Equipo.");
                }

                if (repara.NumeroSerie.Length > 32)
                {
                    return BadRequest("Exceso Longitud Número de Serie (máximo 32 caracteres)");
                }

                _context.Update(repara);
                await _context.SaveChangesAsync();
                
                _logger.LogInformation("Reparación actualizada: {NumRepa}", repara.NumRepa);
                return Ok(repara);
            }
            catch (DbUpdateException dbUpdateException)
            {
                _logger.LogError(dbUpdateException, "Error de BD al actualizar reparación");
                return BadRequest($"Error al actualizar reparación: {dbUpdateException.InnerException?.Message ?? dbUpdateException.Message}");
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, "Error al actualizar reparación");
                return BadRequest($"Error al actualizar reparación: {exception.Message}");
            }
        }

        /// <summary>
        /// Elimina una reparación
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            try
            {
                var repara = await _context.Reparas.FirstOrDefaultAsync(x => x.NumRepa == id);
                if (repara == null)
                {
                    return NotFound();
                }

                _context.Remove(repara);
                await _context.SaveChangesAsync();
                
                _logger.LogInformation("Reparación eliminada: {NumRepa}", id);
                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar reparación {Id}", id);
                return StatusCode(500, $"Error al eliminar reparación: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene reparaciones por cliente
        /// </summary>
        [HttpGet("cliente/{numeroCliente:int}")]
        public async Task<ActionResult> GetByClienteAsync(int numeroCliente)
        {
            try
            {
                var reparaciones = await _context.Reparas
                    .Where(x => x.Nume_Cli == numeroCliente)
                    .OrderByDescending(x => x.FechRecep)
                    .ToListAsync();

                return Ok(reparaciones);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener reparaciones del cliente {NumeroCliente}", numeroCliente);
                return StatusCode(500, $"Error al obtener reparaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene reparaciones por código de reparación
        /// </summary>
        [HttpGet("codigo/{codigoReparacion}")]
        public async Task<ActionResult> GetByCodigoAsync(string codigoReparacion)
        {
            try
            {
                var reparaciones = await _context.Reparas
                    .Where(x => x.CodRepa == codigoReparacion)
                    .OrderByDescending(x => x.FechRecep)
                    .ToListAsync();

                return Ok(reparaciones);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener reparaciones del código {CodigoReparacion}", codigoReparacion);
                return StatusCode(500, $"Error al obtener reparaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Cambia el estado de una reparación
        /// </summary>
        [HttpPut("cambiarStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> CambiarStatusAsync(int id, int status)
        {
            try
            {
                var repara = await _context.Reparas.FirstOrDefaultAsync(x => x.NumRepa == id);
                if (repara == null)
                {
                    return NotFound();
                }

                repara.Status = (short)status;
                _context.Update(repara);
                await _context.SaveChangesAsync();
                
                _logger.LogInformation("Estado de reparación {NumRepa} cambiado a {Status}", id, status);
                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al cambiar estado de reparación {Id}", id);
                return StatusCode(500, $"Error al cambiar estado: {ex.Message}");
            }
        }

        /// <summary>
        /// Anula una recepción cambiando su status a -1
        /// </summary>
        [HttpPut("anular/{id:int}")]
        public async Task<IActionResult> AnularAsync(int id)
        {
            try
            {
                var repara = await _context.Reparas.FirstOrDefaultAsync(x => x.Id == id);
                if (repara == null)
                {
                    return NotFound("Recepción no encontrada");
                }

                if (repara.Status == -1)
                {
                    return BadRequest("La recepción ya está anulada");
                }

                repara.Status = -1;
                _context.Update(repara);
                await _context.SaveChangesAsync();
                
                _logger.LogInformation("Recepción {NumRepa} (ID: {Id}) anulada", repara.NumRepa, id);
                return Ok(new { message = "Recepción anulada correctamente" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al anular recepción {Id}", id);
                return StatusCode(500, $"Error al anular recepción: {ex.Message}");
            }
        }

        /// <summary>
        /// Reactiva una recepción cambiando su status a 0
        /// </summary>
        [HttpPut("reactivar/{id:int}")]
        public async Task<IActionResult> ReactivarAsync(int id)
        {
            try
            {
                var repara = await _context.Reparas.FirstOrDefaultAsync(x => x.Id == id);
                if (repara == null)
                {
                    return NotFound("Recepción no encontrada");
                }

                if (repara.Status == 0)
                {
                    return BadRequest("La recepción ya está activa");
                }

                repara.Status = 0;
                _context.Update(repara);
                await _context.SaveChangesAsync();
                
                _logger.LogInformation("Recepción {NumRepa} (ID: {Id}) reactivada", repara.NumRepa, id);
                return Ok(new { message = "Recepción reactivada correctamente" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al reactivar recepción {Id}", id);
                return StatusCode(500, $"Error al reactivar recepción: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene presupuestos de reparación con paginación y filtros
        /// Condición: NroPed > 0.5 AND Status < 9
        /// </summary>
        [HttpGet("presupuestos")]
        public async Task<IActionResult> GetPresupuestosAsync(
            [FromQuery] PaginationDTO pagination,
            [FromQuery] int? filtroId = null,
            [FromQuery] int? filtroNumRepa = null,
            [FromQuery] int? filtroNroPed = null,
            [FromQuery] DateTime? filtroFechaDesde = null,
            [FromQuery] DateTime? filtroFechaHasta = null,
            [FromQuery] int? filtroNumeCli = null,
            [FromQuery] string? filtroRasoCli = null,
            [FromQuery] string? filtroCodRepa = null,
            [FromQuery] string? filtroNuOrdRep = null)
        {
            try
            {
                var queryable = _context.Reparas
                    .Where(x => x.NroPed > 0.5m && x.Status < 9);

                // Aplicar filtros individuales
                if (filtroId.HasValue && filtroId > 0)
                {
                    queryable = queryable.Where(x => x.Id == filtroId);
                }

                if (filtroNumRepa.HasValue && filtroNumRepa > 0)
                {
                    queryable = queryable.Where(x => x.NumRepa == filtroNumRepa);
                }

                if (filtroNroPed.HasValue && filtroNroPed > 0)
                {
                    queryable = queryable.Where(x => x.NroPed == filtroNroPed);
                }

                if (filtroFechaDesde.HasValue)
                {
                    queryable = queryable.Where(x => x.FechRecep >= filtroFechaDesde.Value);
                }

                if (filtroFechaHasta.HasValue)
                {
                    queryable = queryable.Where(x => x.FechRecep <= filtroFechaHasta.Value);
                }

                if (filtroNumeCli.HasValue && filtroNumeCli > 0)
                {
                    queryable = queryable.Where(x => x.Nume_Cli == filtroNumeCli);
                }

                if (!string.IsNullOrWhiteSpace(filtroRasoCli))
                {
                    var filtro = filtroRasoCli.ToLower();
                    queryable = queryable.Where(x => x.Raso_Cli!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroCodRepa))
                {
                    var filtro = filtroCodRepa.ToLower();
                    queryable = queryable.Where(x => x.CodRepa!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroNuOrdRep))
                {
                    if (int.TryParse(filtroNuOrdRep, out var nuOrdRep))
                    {
                        queryable = queryable.Where(x => x.NuOrdRep == nuOrdRep);
                    }
                }

                return Ok(await queryable
                    .OrderByDescending(x => x.FechRecep)
                    .Paginate(pagination)
                    .ToListAsync());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener presupuestos");
                return StatusCode(500, $"Error al obtener presupuestos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para presupuestos
        /// </summary>
        [HttpGet("presupuestos/totalPages")]
        public async Task<ActionResult> GetPresupuestosPagesAsync(
            [FromQuery] PaginationDTO pagination,
            [FromQuery] int? filtroId = null,
            [FromQuery] int? filtroNumRepa = null,
            [FromQuery] int? filtroNroPed = null,
            [FromQuery] DateTime? filtroFechaDesde = null,
            [FromQuery] DateTime? filtroFechaHasta = null,
            [FromQuery] int? filtroNumeCli = null,
            [FromQuery] string? filtroRasoCli = null,
            [FromQuery] string? filtroCodRepa = null,
            [FromQuery] string? filtroNuOrdRep = null)
        {
            try
            {
                var queryable = _context.Reparas
                    .Where(x => x.NroPed > 0.5m && x.Status < 9);

                // Aplicar los mismos filtros
                if (filtroId.HasValue && filtroId > 0)
                {
                    queryable = queryable.Where(x => x.Id == filtroId);
                }

                if (filtroNumRepa.HasValue && filtroNumRepa > 0)
                {
                    queryable = queryable.Where(x => x.NumRepa == filtroNumRepa);
                }

                if (filtroNroPed.HasValue && filtroNroPed > 0)
                {
                    queryable = queryable.Where(x => x.NroPed == filtroNroPed);
                }

                if (filtroFechaDesde.HasValue)
                {
                    queryable = queryable.Where(x => x.FechRecep >= filtroFechaDesde.Value);
                }

                if (filtroFechaHasta.HasValue)
                {
                    queryable = queryable.Where(x => x.FechRecep <= filtroFechaHasta.Value);
                }

                if (filtroNumeCli.HasValue && filtroNumeCli > 0)
                {
                    queryable = queryable.Where(x => x.Nume_Cli == filtroNumeCli);
                }

                if (!string.IsNullOrWhiteSpace(filtroRasoCli))
                {
                    var filtro = filtroRasoCli.ToLower();
                    queryable = queryable.Where(x => x.Raso_Cli!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroCodRepa))
                {
                    var filtro = filtroCodRepa.ToLower();
                    queryable = queryable.Where(x => x.CodRepa!.ToLower().Contains(filtro));
                }

                if (!string.IsNullOrWhiteSpace(filtroNuOrdRep))
                {
                    if (int.TryParse(filtroNuOrdRep, out var nuOrdRep))
                    {
                        queryable = queryable.Where(x => x.NuOrdRep == nuOrdRep);
                    }
                }

                double count = await queryable.CountAsync();
                double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular páginas de presupuestos");
                return StatusCode(500, $"Error al calcular páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene reparaciones sin presupuesto para el selector
        /// Condición: NroPed IS NULL OR NroPed <= 0.5
        /// </summary>
        [HttpGet("sin-presupuesto")]
        public async Task<IActionResult> GetSinPresupuestoAsync([FromQuery] string? filtro = null)
        {
            try
            {
                var queryable = _context.Reparas
                    .Where(x => x.Status >= 0 && (x.NroPed == null || x.NroPed <= 0.5m));

                if (!string.IsNullOrWhiteSpace(filtro))
                {
                    var filtroLower = filtro.ToLower();
                    queryable = queryable.Where(x =>
                        x.NumRepa.ToString().Contains(filtroLower) ||
                        x.NumeroSerie!.ToLower().Contains(filtroLower) ||
                        x.Raso_Cli!.ToLower().Contains(filtroLower) ||
                        x.CodRepa!.ToLower().Contains(filtroLower));
                }

                var reparaciones = await queryable
                    .OrderByDescending(x => x.FechRecep)
                    .Take(50)
                    .ToListAsync();

                return Ok(reparaciones);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener reparaciones sin presupuesto");
                return StatusCode(500, $"Error al obtener reparaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene órdenes de venta de reparación pendientes de generar O.R.
        /// Equivalente a VB6: 
        /// - CONDI$ = "NROPED > 0 AND STATUS < 9 AND NuOrdRep < 0.5" (estándar)
        /// - If InStr(1, EMPREAC$, "NEUQUEN") > 0 Then CONDI$ = "NuOrdRep < 0.5 AND STATUS < 9" (NEUQUEN)
        /// </summary>
        [HttpGet("pendientes-generar-or")]
        public async Task<IActionResult> GetPendientesGenerarOrAsync()
        {
            try
            {
                // Obtener información de diagnóstico
                var dependenciaActual = _empresaDependenciaService.ObtenerNombreDependenciaActual();
                var empresaActual = _empresaDependenciaService.ObtenerNombreEmpresaActual();
                
                _logger.LogInformation("=== DIAGNÓSTICO DEPENDENCIA ===");
                _logger.LogInformation("Empresa actual: {Empresa}", empresaActual ?? "NULL");
                _logger.LogInformation("Dependencia actual: {Dependencia}", dependenciaActual ?? "NULL");
                
                // Validar si la dependencia activa es NEUQUEN
                var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
                _logger.LogInformation("¿Es NEUQUEN?: {EsNeuquen}", esNeuquen);
                
                // Comparación adicional para diagnóstico
                if (!string.IsNullOrEmpty(dependenciaActual))
                {
                    var contieneNeuquen = dependenciaActual.Contains("NEUQUEN", StringComparison.OrdinalIgnoreCase);
                    _logger.LogInformation("¿Contiene 'NEUQUEN'?: {Contiene}", contieneNeuquen);
                }

                IQueryable<Repara> queryable;

                if (esNeuquen)
                {
                    // Condición para NEUQUEN: NuOrdRep < 0.5 AND STATUS < 9
                    // (no requiere NROPED > 0)
                    queryable = _context.Reparas
                        .Where(x => (x.NuOrdRep == null || x.NuOrdRep < 0.5m) && x.Status < 9);
                    
                    _logger.LogInformation("✅ Aplicando filtro NEUQUEN: NuOrdRep < 0.5 AND STATUS < 9");
                }
                else
                {
                    // Condición estándar: NROPED > 0 AND STATUS < 9 AND NuOrdRep < 0.5
                    queryable = _context.Reparas
                        .Where(x => x.NroPed > 0 && x.Status < 9 
                        && (x.NuOrdRep == null || x.NuOrdRep < 0.5m));
                    
                    _logger.LogInformation("📋 Aplicando filtro estándar: NROPED > 0 AND STATUS < 9 AND NuOrdRep < 0.5");
                }

                // Ordenar por NumRepa
                queryable = queryable.OrderBy(x => x.NumRepa);

                var reparaciones = await queryable.ToListAsync();

                _logger.LogInformation("Órdenes pendientes de generar O.R.: {Count}", reparaciones.Count);
                
                return Ok(reparaciones);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener órdenes pendientes de generar O.R.");
                return StatusCode(500, $"Error al obtener órdenes pendientes: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el próximo número de orden de reparación disponible
        /// Equivalente a VB6: Max(NuOrdRep) + 1
        /// </summary>
        [HttpGet("proximo-numero-orden")]
        public async Task<IActionResult> GetProximoNumeroOrdenAsync()
        {
            try
            {
                var maxNumero = await _context.Reparas
                    .Where(x => x.NuOrdRep.HasValue && x.NuOrdRep > 0)
                    .MaxAsync(x => (int?)x.NuOrdRep) ?? 0;

                var proximoNumero = maxNumero + 1;

                _logger.LogInformation("Próximo número de orden de reparación: {Numero}", proximoNumero);
                
                return Ok(proximoNumero);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener próximo número de orden");
                return StatusCode(500, $"Error al obtener próximo número: {ex.Message}");
            }
        }

        /// <summary>
        /// Genera una orden de reparación a partir de una reparación existente.
        /// Replica la lógica VB6 para persistir NuOrdRep, RefOrdRep, Referen2,
        /// NumeroSerie, FechOrdRep y Observa sobre REPARA.
        /// </summary>
        [HttpPost("generar-orden")]
        public async Task<IActionResult> GenerarOrdenReparacionAsync([FromBody] GenerarOrdenReparacionRequest request)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                if (request.NumRepa <= 0)
                {
                    return BadRequest("NumRepa inválido");
                }

                var reparacion = await _context.Reparas.FirstOrDefaultAsync(x => x.NumRepa == request.NumRepa);
                if (reparacion == null)
                {
                    return NotFound($"No se encontró la reparación N° {request.NumRepa}");
                }

                var maxNuOrdRep = await _context.Reparas
                    .Where(x => x.NumRepa > 0 && x.NuOrdRep.HasValue && x.NuOrdRep > 0)
                    .MaxAsync(x => (int?)x.NuOrdRep) ?? 0;

                var nuevoNumeroOrden = maxNuOrdRep + 1;
                var fechaOrden = request.FechOrdRep?.Date ?? DateTime.Today;

                var codRepa = (reparacion.CodRepa ?? string.Empty).Trim();
                var razonSocial = (reparacion.Raso_Cli ?? string.Empty).Trim();
                if (razonSocial.Length > 30)
                {
                    razonSocial = razonSocial.Substring(0, 30);
                }

                var referen2 = (request.Referen2 ?? string.Empty).Trim();
                var referenciaOrden = $"{codRepa} - {razonSocial} - R: {referen2} - {fechaOrden:dd/MM/yyyy}";
                if (referenciaOrden.Length > 128)
                {
                    referenciaOrden = referenciaOrden.Substring(0, 128);
                }

                reparacion.NuOrdRep = nuevoNumeroOrden;
                reparacion.RefOrdRep = referenciaOrden;
                reparacion.Referen2 = request.Referen2;
                reparacion.NumeroSerie = request.NumeroSerie;
                reparacion.FechOrdRep = fechaOrden;
                reparacion.Observa = request.Observa;

                var usuarioActual = ObtenerUsuarioActual();

                AgregarAnotacionPosicional(
                    reparacion,
                    $"Generación O.R. {nuevoNumeroOrden} para reparación {reparacion.NumRepa}",
                    usuarioActual);

                _context.Reparas.Update(reparacion);

                var numeroOrdenFormateado = $"OR-{nuevoNumeroOrden:D6}";

                var reservasExistentes = await _context.Reservas
                    .Where(r => r.IdSubOr == numeroOrdenFormateado)
                    .ToListAsync();

                if (reservasExistentes.Any())
                {
                    _context.Reservas.RemoveRange(reservasExistentes);
                }

                var repuestos = await _context.Repurepas
                    .Where(x => x.NumRepa == request.NumRepa
                        && x.Cod_Inte > 0
                        && x.Cant_Sol.HasValue
                        && Math.Abs(x.Cant_Sol.Value) >= 0.05m)
                    .ToListAsync();

                var codigosInternos = repuestos
                    .Select(x => x.Cod_Inte)
                    .Distinct()
                    .ToList();

                var mastersByCodint = await _context.masters
                    .AsNoTracking()
                    .Where(m => codigosInternos.Contains(m.Codint))
                    .ToDictionaryAsync(m => m.Codint);

                short nuOrItem = 0;
                var nuevasReservas = new List<Reserva>();

                foreach (var repuesto in repuestos)
                {
                    nuOrItem++;
                    mastersByCodint.TryGetValue(repuesto.Cod_Inte, out var masterItem);

                    var cantidad = repuesto.Cant_Sol ?? 0m;
                    var now = DateTime.Now;

                    nuevasReservas.Add(new Reserva
                    {
                        CodiEmpr = 0,
                        CodiEmprNet = 0,
                        Cod_Inte = repuesto.Cod_Inte,
                        Cod_Exte = masterItem?.Codigo ?? string.Empty,
                        Descrip = masterItem?.Descripcion ?? string.Empty,
                        IdSubOr = numeroOrdenFormateado,
                        NuOrItem = nuOrItem,
                        Status = 0,
                        FechRes = now,
                        CantRes = cantidad,
                        CantCons = 0m,
                        Saldo_Entre = cantidad,
                        IdenLote = string.IsNullOrWhiteSpace(repuesto.IdenLote) ? string.Empty : repuesto.IdenLote,
                        Deposito = 0,
                        UsoUnit = cantidad,
                        AddRecord = now,
                        LastUpdate = now
                    });
                }

                if (nuevasReservas.Any())
                {
                    _context.Reservas.AddRange(nuevasReservas);
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                _logger.LogInformation(
                    "Orden de reparación generada. NumRepa={NumRepa}, NuOrdRep={NuOrdRep}, NroPed={NroPed}, ImpTot={ImpTot}, Reservas={ReservasCount}",
                    reparacion.NumRepa,
                    reparacion.NuOrdRep,
                    reparacion.NroPed,
                    reparacion.Imp_Tot,
                    nuevasReservas.Count);

                return Ok(new GenerarOrdenReparacionResponse
                {
                    NumRepa = reparacion.NumRepa,
                    NuOrdRep = reparacion.NuOrdRep ?? 0,
                    RefOrdRep = reparacion.RefOrdRep ?? string.Empty,
                    FechOrdRep = reparacion.FechOrdRep,
                    NroPed = reparacion.NroPed,
                    NumeCli = reparacion.Nume_Cli,
                    RasoCli = reparacion.Raso_Cli,
                    ImpTot = reparacion.Imp_Tot
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al generar orden de reparación para NumRepa={NumRepa}", request.NumRepa);
                return StatusCode(500, $"Error al generar orden de reparación: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene órdenes de reparación en curso que pueden cerrarse
        /// Condición: NuOrdRep > 0 AND Status < 3
        /// </summary>
        [HttpGet("ordenes-en-curso")]
        public async Task<IActionResult> GetOrdenesEnCursoAsync()
        {
            try
            {
                var ordenes = await _context.Reparas
                    .Where(x => x.NuOrdRep.HasValue && x.NuOrdRep > 0 && x.Status < 3)
                    .OrderByDescending(x => x.NuOrdRep)
                    .Select(x => new
                    {
                        x.Id,
                        x.NumRepa,
                        x.NuOrdRep,
                        x.RefOrdRep,
                        x.Status,
                        x.FechRecep,
                        x.Nume_Cli,
                        x.Raso_Cli
                    })
                    .ToListAsync();

                _logger.LogInformation("Órdenes de reparación en curso: {Count}", ordenes.Count);
                
                return Ok(ordenes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener órdenes en curso");
                return StatusCode(500, $"Error al obtener órdenes en curso: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene los repuestos a consumir de una reparación
        /// Filtra por NumRepa en la tabla REPUREPA
        /// </summary>
        [HttpGet("repuestos-consumir/{numRepa}")]
        public async Task<IActionResult> GetRepuestosConsumirAsync(int numRepa)
        {
            try
            {
                var repuestos = await _context.Repurepas
                    .Where(x => x.NumRepa == numRepa)
                    .Select(x => new
                    {
                        x.Id,
                        x.Cod_Inte,
                        x.Cant_Sol,
                        x.PrUn_Neto,
                        x.Impo_Neto,
                        x.IdenLote,
                        x.Origen
                    })
                    .ToListAsync();

                _logger.LogInformation("Repuestos a consumir para NumRepa {NumRepa}: {Count}", numRepa, repuestos.Count);
                
                return Ok(repuestos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener repuestos a consumir para NumRepa {NumRepa}", numRepa);
                return StatusCode(500, $"Error al obtener repuestos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene la lista de depósitos disponibles
        /// </summary>
        [HttpGet("depositos-disponibles")]
        public async Task<IActionResult> GetDepositosDisponiblesAsync()
        {
            try
            {
                var depositos = await _context.Tadeposi
                    .Where(x => x.disponible && x.Status == 0)
                    .OrderBy(x => x.CodigoDepo)
                    .Select(x => new
                    {
                        x.Id,
                        x.CodigoDepo,
                        x.Descripcion,
                        x.DescripAbreviada
                    })
                    .ToListAsync();

                _logger.LogInformation("Depósitos disponibles: {Count}", depositos.Count);
                
                return Ok(depositos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener depósitos disponibles");
                return StatusCode(500, $"Error al obtener depósitos: {ex.Message}");
            }
        }

        /// <summary>
        /// Cierra una orden de reparación
        /// Actualiza REPARA, crea movimientos en MOVISTO y opcionalmente en ENTRECOM (para NEUQUEN)
        /// </summary>
        [HttpPost("cerrar-orden")]
        public async Task<IActionResult> CerrarOrdenReparacionAsync([FromBody] CerrarOrdenRequest request)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            
            try
            {
                // 1. Obtener la reparación
                var reparacion = await _context.Reparas
                    .FirstOrDefaultAsync(x => x.Id == request.ReparacionId);

                if (reparacion == null)
                {
                    return NotFound("Reparación no encontrada");
                }

                if (!reparacion.NuOrdRep.HasValue || reparacion.NuOrdRep <= 0)
                {
                    return BadRequest("La reparación no tiene un número de orden válido");
                }

                // 2. Formatear número de orden: OR-XXXXXX
                var numeroOrden = ((int)reparacion.NuOrdRep.Value).ToString();
                var numeroOrdenFormateado = $"OR-{numeroOrden}";

                // 3. Actualizar REPARA.RefOrdRep y REPARA.Status
                var referenciaActual = reparacion.RefOrdRep ?? "";
                var nuevaReferencia = $"(cerr) {referenciaActual}";
                if (nuevaReferencia.Length > 64)
                {
                    nuevaReferencia = nuevaReferencia.Substring(0, 64);
                }

                reparacion.RefOrdRep = nuevaReferencia;
                reparacion.Status = 3; // Status cerrado

                var usuarioActual = ObtenerUsuarioActual();
                AgregarAnotacionPosicional(reparacion, "Generación de Cierre de Reparación", usuarioActual);

                _context.Reparas.Update(reparacion);

                // 4. Obtener repuestos a consumir
                var repuestos = await _context.Repurepas
                    .Where(x => x.NumRepa == reparacion.NumRepa)
                    .ToListAsync();

                if (!repuestos.Any())
                {
                    _logger.LogWarning("No hay repuestos para consumir en NumRepa {NumRepa}", reparacion.NumRepa);
                }

                // 5. Verificar si es dependencia NEUQUEN
                var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
                _logger.LogInformation("Cerrando orden para dependencia NEUQUEN: {EsNeuquen}", esNeuquen);

                // 6. Calcular IDOCUM
                var fechaHoy = DateTime.Now;
                var ff = (fechaHoy - new DateTime(1900, 1, 1)).Days; // Días desde 1900
                var nroo = (int)reparacion.NuOrdRep.Value;
                var idocum = 10000 * (ff % 1000) + 100 * (nroo % 100);

                // 7. Procesar cada repuesto
                int nuorit = 0;
                foreach (var repuesto in repuestos)
                {
                    if (!repuesto.Cant_Sol.HasValue || repuesto.Cant_Sol <= 0)
                        continue;

                    nuorit++;

                    // Obtener código externo del Master
                    var master = await _context.masters
                        .FirstOrDefaultAsync(m => m.Codint == repuesto.Cod_Inte);
                    
                    var codigoExterno = master?.Codigo ?? "";

                    // 8. Crear movimiento en MOVISTO
                    var movimiento = new Movisto
                    {
                        FechMov = fechaHoy,
                        Cod_Inte = repuesto.Cod_Inte,
                        IdenLote = "",
                        CodiMovi = "RP", // Código movimiento: Reparación
                        DoPriCor = numeroOrdenFormateado.Length > 12 ? numeroOrdenFormateado.Substring(0, 12) : numeroOrdenFormateado,
                        DoPriLar = numeroOrdenFormateado,
                        DoSecCor = numeroOrdenFormateado.Length > 12 ? numeroOrdenFormateado.Substring(0, 12) : numeroOrdenFormateado,
                        DoSecLar = numeroOrdenFormateado,
                        DeLaRemi = numeroOrdenFormateado,
                        ReferCor = $"Reparación {numeroOrdenFormateado}",
                        NuorItem = nuorit,
                        Nume_Clie = reparacion.Nume_Cli,
                        DepoMovi = request.DepositoId,
                        CantSalid = repuesto.Cant_Sol.Value, // Salida de stock
                        CantIngre = 0,
                        Cod_Exte = codigoExterno,
                        Descrip = master?.Descripcion?.Length > 64 ? master.Descripcion.Substring(0, 64) : master?.Descripcion ?? "",
                        FeReMovi = fechaHoy
                    };

                    _context.Movistos.Add(movimiento);

                    // 9. Si es NEUQUEN, crear registro en ENTRECOM
                    if (esNeuquen)
                    {
                        var entrecom = new Entrecom
                        {
                            Fecha = fechaHoy,
                            Documento = numeroOrdenFormateado.Length > 14 ? numeroOrdenFormateado.Substring(0, 14) : numeroOrdenFormateado,
                            TipoDoc = "Cons. Reparacion".Length > 16 ? "Cons. Reparacion".Substring(0, 16) : "Cons. Reparacion",
                            Nume_Cli = reparacion.Nume_Cli,
                            IDocum = idocum,
                            Cod_Inte = repuesto.Cod_Inte,
                            Cod_Exte = codigoExterno.Length > 24 ? codigoExterno.Substring(0, 24) : codigoExterno,
                            Cantidad = repuesto.Cant_Sol.Value,
                            Status = 1,
                            VecesEnviado = 0
                        };

                        _context.Entrecoms.Add(entrecom);
                    }
                }

                // 10. Guardar todos los cambios
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                _logger.LogInformation(
                    "Orden de reparación cerrada exitosamente. NumRepa: {NumRepa}, NuOrdRep: {NuOrdRep}, Repuestos: {Count}, NEUQUEN: {EsNeuquen}",
                    reparacion.NumRepa, reparacion.NuOrdRep, repuestos.Count, esNeuquen);

                return Ok(new
                {
                    mensaje = "Orden de reparación cerrada exitosamente",
                    numeroOrden = numeroOrdenFormateado,
                    numRepa = reparacion.NumRepa,
                    repuestosConsumidos = repuestos.Count,
                    movimientosCreados = nuorit,
                    entrecomCreados = esNeuquen ? nuorit : 0
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al cerrar orden de reparación");
                return StatusCode(500, $"Error al cerrar orden: {ex.Message}");
            }
        }

        /// <summary>
        /// Anula una orden de reparación
        /// Replica la lógica del código VB6:
        /// - Cambia Status a 9 (anulado)
        /// - Elimina registros de REPUREPA (repuestos)
        /// - Elimina registros de RESERVA (reservas)
        /// </summary>
        [HttpPost("anular-orden")]
        public async Task<IActionResult> AnularOrdenAsync([FromBody] AnularOrdenRequest request)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                _logger.LogInformation("Iniciando anulación de orden de reparación. NumRepa: {NumRepa}, NuOrdRep: {NuOrdRep}", 
                    request.NumRepa, request.NuOrdRep);

                // 1. Verificar que la orden existe y está en curso (Status < 3)
                var reparacion = await _context.Reparas
                    .FirstOrDefaultAsync(r => r.NumRepa == request.NumRepa && r.NuOrdRep == request.NuOrdRep);

                if (reparacion == null)
                {
                    return NotFound($"No se encontró la orden de reparación con NumRepa: {request.NumRepa} y NuOrdRep: {request.NuOrdRep}");
                }

                if (reparacion.Status >= 3)
                {
                    return BadRequest($"La orden de reparación ya está cerrada o anulada (Status: {reparacion.Status})");
                }

                // 2. Cambiar Status a 9 (anulado)
                // CONDI$ = "NuOrdRep = " & NOREP&
                // Call ACTUREGISTRO("REPARA", "Status", CONDI$, 9, REG&)
                reparacion.Status = 9;
                _context.Reparas.Update(reparacion);
                _logger.LogInformation("Status cambiado a 9 (anulado) para NumRepa: {NumRepa}", request.NumRepa);

                // 3. Eliminar repuestos de REPUREPA
                // CONDI$ = "NumRepa = " & NOREP&
                // Call BORRAREGISTROS("REPUREPA", CONDI$, REG&)
                var repuestos = await _context.Repurepas
                    .Where(r => r.NumRepa == request.NumRepa)
                    .ToListAsync();

                if (repuestos.Any())
                {
                    _context.Repurepas.RemoveRange(repuestos);
                    _logger.LogInformation("Eliminados {Count} repuestos de REPUREPA para NumRepa: {NumRepa}", 
                        repuestos.Count, request.NumRepa);
                }

                // 4. Eliminar reservas de RESERVA
                // Formatear número de orden (OR-XXXXXX)
                var numeroOrdenFormateado = $"OR-{request.NuOrdRep:D6}";
                // CONDI$ = "IdSubOr ='" & NOREX$ & "'"
                // Call BORRAREGISTROS("RESERVA", CONDI$, REG&)
                var reservas = await _context.Reservas
                    .Where(r => r.IdSubOr == numeroOrdenFormateado)
                    .ToListAsync();

                if (reservas.Any())
                {
                    _context.Reservas.RemoveRange(reservas);
                    _logger.LogInformation("Eliminadas {Count} reservas de RESERVA para IdSubOr: {IdSubOr}", 
                        reservas.Count, numeroOrdenFormateado);
                }

                // 5. Guardar todos los cambios
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                _logger.LogInformation(
                    "Anulación completa. NumRepa: {NumRepa}, NuOrdRep: {NuOrdRep}, Repuestos eliminados: {RepuestosCount}, Reservas eliminadas: {ReservasCount}",
                    request.NumRepa, request.NuOrdRep, repuestos.Count, reservas.Count);

                return Ok(new
                {
                    mensaje = "Anulación Completa",
                    numRepa = request.NumRepa,
                    nuOrdRep = request.NuOrdRep,
                    numeroOrdenFormateado = numeroOrdenFormateado,
                    repuestosEliminados = repuestos.Count,
                    reservasEliminadas = reservas.Count
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al anular orden de reparación. NumRepa: {NumRepa}, NuOrdRep: {NuOrdRep}", 
                    request.NumRepa, request.NuOrdRep);
                return StatusCode(500, $"Error al anular orden: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene los repuestos asociados a una reparación (tabla REPUREPA)
        /// </summary>
        [HttpGet("repuestos-por-reparacion/{numRepa:int}")]
        public async Task<IActionResult> GetRepuestosPorReparacionAsync(int numRepa)
        {
            try
            {
                var repuestos = await _context.Repurepas
                    .Where(r => r.NumRepa == numRepa)
                    .Select(r => new
                    {
                        r.Id,
                        r.Cod_Inte,
                        r.Cant_Sol,
                        r.PrUn_Neto,
                        r.Impo_Neto
                    })
                    .ToListAsync();

                return Ok(repuestos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener repuestos para NumRepa: {NumRepa}", numRepa);
                return StatusCode(500, $"Error al obtener repuestos: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene las reservas asociadas a una orden de reparación (tabla RESERVA)
        /// </summary>
        [HttpGet("reservas-por-orden/{idSubOr}")]
        public async Task<IActionResult> GetReservasPorOrdenAsync(string idSubOr)
        {
            try
            {
                var reservas = await _context.Reservas
                    .Where(r => r.IdSubOr == idSubOr)
                    .Select(r => new
                    {
                        r.Reserva_Id,
                        r.Cod_Inte,
                        r.Descrip,
                        r.CantRes,
                        r.Deposito,
                        r.IdSubOr
                    })
                    .ToListAsync();

                return Ok(reservas);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener reservas para IdSubOr: {IdSubOr}", idSubOr);
                return StatusCode(500, $"Error al obtener reservas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene las órdenes de reparación cerradas (Status = 3) pendientes de entrega
        /// Replica la lógica VB6: CONDI$ = "NuOrdRep > 0 And Status = 3"
        /// </summary>
        [HttpGet("ordenes-cerradas")]
        public async Task<IActionResult> GetOrdenesCerradasAsync()
        {
            try
            {
                var ordenes = await _context.Reparas
                    .Where(r => r.NuOrdRep > 0 && r.Status == 3)
                    .OrderBy(r => r.NuOrdRep)
                    .Select(r => new
                    {
                        r.Id,
                        r.NumRepa,
                        r.NuOrdRep,
                        r.RefOrdRep,
                        r.Status,
                        r.FechRecep,
                        r.Nume_Cli,
                        r.Raso_Cli
                    })
                    .ToListAsync();

                _logger.LogInformation("Órdenes cerradas encontradas: {Count}", ordenes.Count);
                return Ok(ordenes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener órdenes cerradas");
                return StatusCode(500, $"Error al obtener órdenes cerradas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene los datos del cliente desde DEUDOR12
        /// </summary>
        [HttpGet("datos-cliente/{numeCli:int}")]
        public async Task<IActionResult> GetDatosClienteAsync(int numeCli)
        {
            try
            {
                var cliente = await _context.Clientes
                    .Where(d => d.Nume_Cli == numeCli)
                    .Select(d => new
                    {
                        d.Nume_Cli,
                        d.Raso_Cli,
                        d.Domi_Cli,
                        d.Cpos_Cli,
                        d.Loca_Cli
                    })
                    .FirstOrDefaultAsync();

                if (cliente == null)
                {
                    return NotFound($"Cliente {numeCli} no encontrado");
                }

                return Ok(cliente);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener datos del cliente {NumeCli}", numeCli);
                return StatusCode(500, $"Error al obtener datos del cliente: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene los datos completos de una orden de reparación
        /// </summary>
        [HttpGet("orden-completa/{numRepa:int}")]
        public async Task<IActionResult> GetOrdenCompletaAsync(int numRepa)
        {
            try
            {
                var orden = await _context.Reparas
                    .Where(r => r.NumRepa == numRepa)
                    .Select(r => new
                    {
                        r.Observa,
                        r.NumeroSerie,
                        r.Referen2
                    })
                    .FirstOrDefaultAsync();

                if (orden == null)
                {
                    return NotFound($"Orden de reparación {numRepa} no encontrada");
                }

                return Ok(orden);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener orden completa {NumRepa}", numRepa);
                return StatusCode(500, $"Error al obtener orden completa: {ex.Message}");
            }
        }

        /// <summary>
        /// Registra la entrega de una orden de reparación
        /// Crea registro en MOVISTO con CODIMOVI='RT' (igual que remitos de cliente)
        /// </summary>
        [HttpPost("registrar-entrega")]
        public async Task<IActionResult> RegistrarEntregaAsync([FromBody] DespachoDTO despacho)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                _logger.LogInformation(
                    "Registrando entrega - NumRepa: {NumRepa}, NuOrdRep: {NuOrdRep}, Fecha: {Fecha}, NumeroRemito: {NumeroRemito}",
                    despacho.NumRepa, despacho.NuOrdRep, despacho.Fecha, despacho.NumeroDocumento);

                // Validar que se haya proporcionado número de remito
                if (string.IsNullOrEmpty(despacho.NumeroDocumento) || !int.TryParse(despacho.NumeroDocumento, out int numeroRemito))
                {
                    return BadRequest("Número de remito inválido");
                }

                // Obtener datos de la orden de reparación
                var reparacion = await _context.Reparas
                    .FirstOrDefaultAsync(r => r.NumRepa == despacho.NumRepa && r.NuOrdRep == despacho.NuOrdRep);

                if (reparacion == null)
                {
                    return NotFound($"No se encontró la orden de reparación NumRepa: {despacho.NumRepa}, NuOrdRep: {despacho.NuOrdRep}");
                }

                // Obtener código general desde FLEXCRL (PROBUS='CODGENER' en tabla REPARA)
                int codigoGeneral = 0;
                try
                {
                    var configCodGen = await _context.Flexcrls
                        .FirstOrDefaultAsync(f => f.CLABUS == "REPARA" && f.PROBUS == "CODGENER");
                    
                    if (configCodGen != null && int.TryParse(configCodGen.VALCONTROL, out int codGen))
                    {
                        codigoGeneral = codGen;
                        _logger.LogInformation("Código general obtenido de FLEXCRL: {CodigoGeneral}", codigoGeneral);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error al obtener código general desde FLEXCRL");
                    return StatusCode(500, "Error al obtener configuración de código general desde FLEXCRL");
                }

                // Validar que el código general sea válido (obligatorio)
                if (codigoGeneral <= 0)
                {
                    _logger.LogError("Código general no configurado o inválido. CLABUS='REPARA', PROBUS='CODGENER'");
                    return BadRequest("El código general (CODGENER) no está configurado en FLEXCRL. Configure el valor en CLABUS='REPARA', PROBUS='CODGENER'");
                }

                // Construir número de remito formateado: RT-X-0001-00000025
                var prefijoFormateado = despacho.Prefijo?.ToString("D4") ?? "0001";
                var numeroFormateado = numeroRemito.ToString("D8");
                var codigoMovimiento = despacho.CodigoMovimiento ?? "RT";
                var numeroRemitoCompleto = $"{codigoMovimiento}-X-{prefijoFormateado}-{numeroFormateado}";
                var numeroRemitoCorto = $"{codigoMovimiento}.{numeroRemito}"; // RT.25

                // LOTE$ = NumeroSerie de la orden de reparación
                var lote = (reparacion.NumeroSerie ?? "").Trim();
                if (lote.Length > 16)
                {
                    lote = lote.Substring(0, 16);
                }

                // Crear registro en MOVISTO con CODIMOVI='RT' (según lógica VB6)
                var movimiento = new Movisto
                {
                    CodiMovi = codigoMovimiento,
                    DOPRINUM = numeroRemito,
                    FechMov = despacho.Fecha,
                    Cod_Inte = codigoGeneral, // Código general desde FLEXCRL
                    Nume_Clie = reparacion.Nume_Cli ?? 0,
                    DepoMovi = despacho.Deposito ?? 0, // Depósito de salida
                    CantSalid = 1, // Cantidad = 1 (salida)
                    IdenLote = lote, // LOTE$ = NumeroSerie (trimmed y truncado a 16 caracteres)
                    DoPriCor = numeroRemitoCorto, // RT.25
                    DoPriLar = numeroRemitoCompleto, // RT-X-0001-00000025
                    DoSecCor = numeroRemitoCorto, // DOSEC$ = REMI$
                    DoSecLar = numeroRemitoCompleto, // DOSEC$ = NROREMI$
                    ReferCor = "", // Dejar vacío
                    DeLaRemi = $"Entrega O.R. {despacho.NuOrdRep} - {despacho.Observaciones ?? ""}".Trim(),
                    FeReMovi = DateTime.Now, // Fecha y hora de registro
                    CodiEmpr = 0, // Código de empresa
                    NuorItem = 1, // Número de orden de item (inicia en 1)
                    CantIngre = 0 // Cantidad de ingreso
                };

                _context.Movistos.Add(movimiento);
                await _context.SaveChangesAsync();

                _logger.LogInformation(
                    "Registro de entrega creado en MOVISTO - CodiMovi: {CodiMovi}, DOPRINUM: {NroDocum}, Cliente: {NumeClie}",
                    movimiento.CodiMovi, movimiento.DOPRINUM, movimiento.Nume_Clie);

                await transaction.CommitAsync();

                return Ok(new
                {
                    mensaje = "Entrega registrada exitosamente",
                    numRepa = despacho.NumRepa,
                    nuOrdRep = despacho.NuOrdRep,
                    numeroRemito = numeroRemito,
                    codigoMovimiento = movimiento.CodiMovi
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al registrar entrega - NumRepa: {NumRepa}, NuOrdRep: {NuOrdRep}", 
                    despacho.NumRepa, despacho.NuOrdRep);
                return StatusCode(500, $"Error al registrar entrega: {ex.Message}");
            }
        }

        /// <summary>
        /// DTO para cerrar orden de reparación
        /// </summary>
        public class CerrarOrdenRequest
        {
            public int ReparacionId { get; set; }
            public int DepositoId { get; set; }
        }

        /// <summary>
        /// DTO para anular orden de reparación
        /// </summary>
        public class AnularOrdenRequest
        {
            public int NumRepa { get; set; }
            public int NuOrdRep { get; set; }
        }

        /// <summary>
        /// DTO para generar orden de reparación.
        /// </summary>
        public class GenerarOrdenReparacionRequest
        {
            public int NumRepa { get; set; }
            public string? Referen2 { get; set; }
            public string? NumeroSerie { get; set; }
            public DateTime? FechOrdRep { get; set; }
            public string? Observa { get; set; }
        }

        /// <summary>
        /// Respuesta de generación de orden de reparación.
        /// </summary>
        public class GenerarOrdenReparacionResponse
        {
            public int NumRepa { get; set; }
            public int NuOrdRep { get; set; }
            public string RefOrdRep { get; set; } = string.Empty;
            public DateTime? FechOrdRep { get; set; }
            public int? NroPed { get; set; }
            public int? NumeCli { get; set; }
            public string? RasoCli { get; set; }
            public decimal? ImpTot { get; set; }
        }

        public class AgregarAnotacionReparacionRequest
        {
            public int NumRepa { get; set; }
            public string Texto { get; set; } = string.Empty;
            public string? Usuario { get; set; }
        }

        public class AgregarAnotacionPresupuestoRequest
        {
            public int NumRepa { get; set; }
            public string? Cliente { get; set; }
            public string? Telefono { get; set; }
            public string? Fax { get; set; }
            public string? NumeroSerie { get; set; }
            public string? Referencia { get; set; }
            public string? Sitio { get; set; }
            public string? Observaciones { get; set; }
        }

        public class AgregarAnotacionReparacionResponse
        {
            public int NumRepa { get; set; }
            public string? AnotaRep { get; set; }
        }

        /// <summary>
        /// DTO para registrar entrega de orden de reparación
        /// </summary>
        public class DespachoDTO
        {
            public DateTime Fecha { get; set; }
            public string? NumeroDocumento { get; set; }
            public int NumRepa { get; set; }
            public int NuOrdRep { get; set; }
            public string? OrdenCompra { get; set; }
            public int? NumeroBultos { get; set; }
            public string? Conductor { get; set; }
            public string? DNI { get; set; }
            public string? Transportista { get; set; }
            public string? CUIT { get; set; }
            public string? Patente { get; set; }
            public string? DomicilioTransporte { get; set; }
            public decimal? ValorDeclarado { get; set; }
            public int? Deposito { get; set; }
            public string? Observaciones { get; set; }
            public string? CodigoMovimiento { get; set; }
            public int? Prefijo { get; set; }
        }

        /// <summary>
        /// Obtiene lista de reparaciones para el selector de Material Suministrado con filtros y paginación
        /// </summary>
        [HttpGet("lista-reparaciones")]
        public async Task<ActionResult> ObtenerListaReparaciones(
            [FromQuery] int page = 1,
            [FromQuery] int pageSize = 20,
            [FromQuery] int? filtroNumRepa = null,
            [FromQuery] int? filtroNumeCli = null,
            [FromQuery] string? filtroRasoCli = null,
            [FromQuery] string? filtroReferen = null,
            [FromQuery] int? filtroNuOrdRep = null)
        {
            try
            {
                _logger.LogInformation("Obteniendo lista de reparaciones - Página: {Page}, Tamaño: {PageSize}", page, pageSize);

                var query = _context.Reparas.AsQueryable();

                // Aplicar filtros
                if (filtroNumRepa.HasValue && filtroNumRepa > 0)
                {
                    query = query.Where(r => r.NumRepa == filtroNumRepa.Value);
                }

                if (filtroNumeCli.HasValue && filtroNumeCli > 0)
                {
                    query = query.Where(r => r.Nume_Cli == filtroNumeCli.Value);
                }

                if (!string.IsNullOrWhiteSpace(filtroRasoCli))
                {
                    query = query.Where(r => r.Raso_Cli != null && r.Raso_Cli.Contains(filtroRasoCli));
                }

                if (!string.IsNullOrWhiteSpace(filtroReferen))
                {
                    query = query.Where(r => r.Referen != null && r.Referen.Contains(filtroReferen));
                }

                if (filtroNuOrdRep.HasValue && filtroNuOrdRep > 0)
                {
                    query = query.Where(r => r.NuOrdRep == filtroNuOrdRep.Value);
                }

                // Contar total de registros
                var totalCount = await query.CountAsync();

                // Aplicar paginación
                var reparaciones = await query
                    .OrderByDescending(r => r.NumRepa)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .Select(r => new ReparacionListaDTO
                    {
                        NumRepa = r.NumRepa,
                        Referen = r.Referen,
                        Nume_Cli = r.Nume_Cli,
                        Raso_Cli = r.Raso_Cli,
                        NuOrdRep = r.NuOrdRep,
                        CodRepa = r.CodRepa,
                        RefOrdRep = r.RefOrdRep
                    })
                    .ToListAsync();

                var totalPages = (int)Math.Ceiling(totalCount / (double)pageSize);

                var resultado = new
                {
                    data = reparaciones,
                    currentPage = page,
                    pageSize = pageSize,
                    totalCount = totalCount,
                    totalPages = totalPages
                };

                _logger.LogInformation("Se encontraron {Count} reparaciones de {Total} total", reparaciones.Count, totalCount);
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener lista de reparaciones");
                return StatusCode(500, $"Error al obtener lista de reparaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene materiales suministrados por el cliente agrupados por producto
        /// Replica la consulta SQL de CARGAMATSUM en VB6
        /// </summary>
        /// <param name="numRepa">Número de reparación (TALON)</param>
        [HttpGet("materiales-suministrados/{numRepa}")]
        public async Task<ActionResult<List<MaterialSuministradoDTO>>> ObtenerMaterialesSuministrados(int numRepa)
        {
            try
            {
                _logger.LogInformation("Obteniendo materiales suministrados para NumRepa: {NumRepa}", numRepa);

                var materiales = await _context.Detamasus
                    .Where(d => d.Talon == numRepa)
                    .GroupBy(d => new { d.Cod_Inte, d.Cod_Exte })
                    .Select(g => new
                    {
                        Cod_Inte = g.Key.Cod_Inte ?? 0,
                        Cod_Exte = g.Key.Cod_Exte,
                        Saldo = (g.Sum(d => d.CantIngre ?? 0) - g.Sum(d => d.CantSalid ?? 0))
                    })
                    .Where(m => m.Saldo > 0)
                    .ToListAsync();

                var resultado = new List<MaterialSuministradoDTO>();

                foreach (var material in materiales)
                {
                    // Obtener descripción del producto
                    var producto = await _context.masters
                        .Where(m => m.Codint == material.Cod_Inte)
                        .Select(m => m.Descripcion)
                        .FirstOrDefaultAsync();

                    resultado.Add(new MaterialSuministradoDTO
                    {
                        Cod_Inte = material.Cod_Inte,
                        Cod_Exte = material.Cod_Exte,
                        Descripcion = producto ?? "",
                        Saldo = material.Saldo
                    });
                }

                _logger.LogInformation("Se encontraron {Count} materiales con saldo positivo", resultado.Count);
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener materiales suministrados para NumRepa: {NumRepa}", numRepa);
                return StatusCode(500, $"Error al obtener materiales suministrados: {ex.Message}");
            }
        }

        /// <summary>
        /// Registra un movimiento de material suministrado
        /// Replica la rutina MOVI_SUMINISTRADOS de VB6
        /// Códigos de movimiento: BR (Agregado Posterior), DV (Devolución), CF (Consumo de Fábrica)
        /// </summary>
        [HttpPost("movimiento-material-suministrado")]
        public async Task<IActionResult> RegistrarMovimientoMaterial([FromBody] MovimientoMaterialDTO movimiento)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                _logger.LogInformation(
                    "Registrando movimiento de material - CodiMovi: {CodiMovi}, Cod_Inte: {CodInte}, Cantidad: {Cantidad}, Talon: {Talon}",
                    movimiento.CodiMovi, movimiento.Cod_Inte, movimiento.Cantidad, movimiento.Talon);

                // Validaciones
                if (string.IsNullOrEmpty(movimiento.CodiMovi) || 
                    (movimiento.CodiMovi != "BR" && movimiento.CodiMovi != "DV" && movimiento.CodiMovi != "CF"))
                {
                    return BadRequest("Código de movimiento inválido. Debe ser BR, DV o CF");
                }

                if (movimiento.Cod_Inte <= 0)
                {
                    return BadRequest("Código interno de producto inválido");
                }

                if (movimiento.Talon <= 0)
                {
                    return BadRequest("Número de reparación (Talon) inválido");
                }

                // Obtener código externo del producto
                var producto = await _context.masters
                    .Where(m => m.Codint == movimiento.Cod_Inte)
                    .Select(m => m.Codigo)
                    .FirstOrDefaultAsync();

                if (producto == null)
                {
                    return BadRequest($"No se encontró el producto con código interno {movimiento.Cod_Inte}");
                }

                // Crear registro en DETAMASU
                var nuevoMovimiento = new Detamasu
                {
                    CodiEmprNet = 0,
                    CodiMovi = movimiento.CodiMovi,
                    Nume_Clie = movimiento.Nume_Clie,
                    Cod_Inte = movimiento.Cod_Inte,
                    Cod_Exte = producto,
                    Talon = movimiento.Talon,
                    CantIngre = movimiento.Cantidad >= 0 ? Math.Round(movimiento.Cantidad, 4) : 0,
                    CantSalid = movimiento.Cantidad < 0 ? Math.Round(Math.Abs(movimiento.Cantidad), 4) : 0,
                    FechMov = movimiento.FechMov,
                    FeRemovi = DateTime.Now,
                    NumUsuar = 0,
                    NReparación = 0,
                    NumRemito = null,
                    Observaciones = "",
                    Referencia = movimiento.Referencia ?? ""
                };

                _context.Detamasus.Add(nuevoMovimiento);
                await _context.SaveChangesAsync();

                _logger.LogInformation(
                    "Movimiento de material registrado exitosamente - Id: {Id}, CodiMovi: {CodiMovi}",
                    nuevoMovimiento.Id, nuevoMovimiento.CodiMovi);

                await transaction.CommitAsync();

                return Ok(new { 
                    mensaje = "Movimiento registrado exitosamente",
                    id = nuevoMovimiento.Id
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al registrar movimiento de material");
                return StatusCode(500, $"Error al registrar movimiento: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene todos los movimientos de una reparación en formato de cuenta contable
        /// Muestra ingresos y salidas con saldo acumulado
        /// </summary>
        [HttpGet("movimientos-cuenta-contable/{numRepa}")]
        public async Task<IActionResult> GetMovimientosCuentaContable(int numRepa)
        {
            try
            {
                _logger.LogInformation("Obteniendo movimientos cuenta contable para NumRepa: {NumRepa}", numRepa);

                // Obtener todos los movimientos de la reparación ordenados por fecha y código interno
                var movimientos = await _context.Detamasus
                    .Where(d => d.Talon == numRepa)
                    .OrderBy(d => d.FechMov)
                    .ThenBy(d => d.Cod_Inte)
                    .Select(d => new
                    {
                        d.Cod_Inte,
                        d.Cod_Exte,
                        d.CodiMovi,
                        d.CantIngre,
                        d.CantSalid,
                        d.FechMov,
                        d.Referencia
                    })
                    .ToListAsync();

                if (!movimientos.Any())
                {
                    _logger.LogInformation("No se encontraron movimientos para NumRepa: {NumRepa}", numRepa);
                    return Ok(new List<object>());
                }

                // Obtener descripciones de productos desde MASTER
                var codigosInternos = movimientos.Where(m => m.Cod_Inte.HasValue).Select(m => m.Cod_Inte!.Value).Distinct().ToList();
                var productos = await _context.masters
                    .Where(m => codigosInternos.Contains(m.Codint))
                    .Select(m => new { m.Codint, m.Descripcion })
                    .ToDictionaryAsync(m => m.Codint, m => m.Descripcion);

                // Mapear a DTO con descripción
                var resultado = movimientos.Select(m => new
                {
                    Cod_Inte = m.Cod_Inte ?? 0,
                    Cod_Exte = m.Cod_Exte ?? "",
                    Descripcion = m.Cod_Inte.HasValue && productos.ContainsKey(m.Cod_Inte.Value) 
                        ? (productos[m.Cod_Inte.Value] ?? "Sin descripción") 
                        : "Sin descripción",
                    CodiMovi = m.CodiMovi ?? "",
                    CantIngre = m.CantIngre,
                    CantSalid = m.CantSalid,
                    FechMov = m.FechMov,
                    Referencia = m.Referencia ?? ""
                }).ToList();

                _logger.LogInformation("Se encontraron {Count} movimientos para NumRepa {NumRepa}", resultado.Count, numRepa);

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener movimientos cuenta contable para NumRepa: {NumRepa}", numRepa);
                return StatusCode(500, $"Error al obtener movimientos: {ex.Message}");
            }
        }
    }
}
