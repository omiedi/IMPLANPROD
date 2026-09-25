using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la tabla ANOTAPED - anotaciones de pedidos estructuradas.
    /// Permite consultar las anotaciones guardadas en formato tabular (a diferencia
    /// del campo ntext PEDENCA.AnotaPed que requiere parsing).
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class AnotapedController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly UserContextService _userContextService;

        public AnotapedController(DataContext context, UserContextService userContextService)
        {
            _context = context;
            _userContextService = userContextService;
        }

        /// <summary>
        /// Obtiene todas las anotaciones de un pedido ordenadas por IdAnotacion.
        /// Cada grupo de items con el mismo IdAnotacion representa un update completo.
        /// </summary>
        /// <param name="nroPed">Número de pedido</param>
        /// <returns>Lista de registros ANOTAPED ordenados por IdAnotacion, NuOrItem</returns>
        [HttpGet("por-pedido/{nroPed:int}")]
        public async Task<ActionResult<List<Anotaped>>> GetPorPedidoAsync(int nroPed)
        {
            try
            {
                var anotaciones = await _context.Anotapeds
                    .Where(a => a.NroPed == nroPed)
                    .OrderBy(a => a.IdAnotacion)
                    .ThenBy(a => a.Id)
                    .AsNoTracking()
                    .ToListAsync();

                if (anotaciones == null || !anotaciones.Any())
                    return Ok(new List<Anotaped>());

                return Ok(anotaciones);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener anotaciones del pedido: {ex.Message}");
            }
        }

        /// <summary>
        /// Agrega una observación general como un nuevo registro en ANOTAPED.
        /// Crea un nuevo IdAnotacion (incremental por pedido) con un único registro
        /// que solo tiene ObservaGral, FechaCambio y Usuario (sin producto asociado).
        /// </summary>
        /// <param name="request">Objeto con NroPed y ObservaGral</param>
        /// <returns>Ok si se guardó correctamente</returns>
        [HttpPost("agregar-observacion")]
        public async Task<ActionResult> AgregarObservacionAsync([FromBody] AgregarObservacionRequest request)
        {
            try
            {
                if (request == null || request.NroPed <= 0 || string.IsNullOrWhiteSpace(request.ObservaGral))
                {
                    return BadRequest("Debe especificar NroPed y ObservaGral.");
                }

                // Truncar ObservaGral a 1024 caracteres
                var observaGral = request.ObservaGral.Length > 1024
                    ? request.ObservaGral.Substring(0, 1024)
                    : request.ObservaGral;

                // Calcular el siguiente IdAnotacion para este pedido
                var maxIdAnotacion = await _context.Anotapeds
                    .Where(a => a.NroPed == request.NroPed)
                    .MaxAsync(a => (int?)a.IdAnotacion) ?? 0;
                var nuevoIdAnotacion = maxIdAnotacion + 1;

                // Obtener datos del usuario activo
                var numUsuar = _userContextService.GetCurrentUserId() ?? 0;
                var nombreUsuario = await ObtenerNombreUsuarioAsync();

                var ahora = DateTime.Now;

                // Crear un único registro de observación (sin producto)
                var anotaped = new Anotaped
                {
                    IdAnotacion = nuevoIdAnotacion,
                    NroPed = request.NroPed,
                    CodiInt = 0,
                    CodiExt = null,
                    Descrip = null,
                    CanPed = null,
                    CantEnt = null,
                    ObservaItem = null,
                    ObservaGral = observaGral,
                    FechaCambio = ahora,
                    Usuario = string.IsNullOrEmpty(nombreUsuario) ? "SISTEMA" : nombreUsuario,
                    CodiEmpr = 0,
                    AddRecord = ahora,
                    LastUpdate = ahora,
                    CodiEmprNet = 0,
                    NumUsuar = numUsuar
                };

                _context.Anotapeds.Add(anotaped);
                await _context.SaveChangesAsync();

                return Ok(new { IdAnotacion = nuevoIdAnotacion, Mensaje = "Observación guardada correctamente." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al guardar la observación: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el nombre del usuario activo desde el contexto HTTP.
        /// Usa el claim Name (NombreUsuario) del JWT; si no está, busca en la base.
        /// </summary>
        private async Task<string?> ObtenerNombreUsuarioAsync()
        {
            var user = _userContextService.GetCurrentUser();
            if (user?.Identity?.IsAuthenticated ?? false)
            {
                var nombre = user.FindFirst(ClaimTypes.Name)?.Value
                             ?? user.Identity?.Name;
                if (!string.IsNullOrEmpty(nombre))
                    return nombre;

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

    /// <summary>
    /// Request para agregar una observación general a ANOTAPED.
    /// </summary>
    public class AgregarObservacionRequest
    {
        public int NroPed { get; set; }
        public string ObservaGral { get; set; } = string.Empty;
    }
}
