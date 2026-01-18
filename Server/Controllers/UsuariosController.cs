using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Cryptography;
using System.Text;
using IMPLANPROD.Server.Utils;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using System.Security.Claims;
using Microsoft.AspNetCore.Components.Authorization;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Client.Repositories;


namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la gestión de usuarios en el sistema.
    /// Maneja operaciones relacionadas con el registro y gestión de usuarios.
    /// </summary>
    [ApiController]
    [Route("/api/Usuarios")]
    public class UsuariosController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<UsuariosController> _logger;
        private readonly FileStorageService _fileStorage;
        //private readonly IUsuarioFlexoftService _UsuarioFlexoftService;
        private readonly ICodiempresa _codiempresaService;
        private readonly IUserContextService _userContextService;

        /// <summary>
        /// Constructor del controlador de usuarios.
        /// </summary>
        /// <param name="context">Contexto de la base de datos inyectado por DI</param>
        /// <param name="logger">Logger para registro de eventos</param>
        /// <param name="fileStorage">Servicio de almacenamiento de archivos</param>
        /// <param name="UsuarioFlexoftService">Servicio para obtener datos del usuario autenticado</param>
        /// <param name="codiempresaService">Servicio para obtener información de empresas</param>
        public UsuariosController(DataContext context, ILogger<UsuariosController> logger, FileStorageService fileStorage, ICodiempresa codiempresaService, IUserContextService userContextService)
        {
            _context = context;
            _logger = logger;
            _fileStorage = fileStorage;
            _codiempresaService = codiempresaService;
            _userContextService = userContextService;
            //IUsuarioFlexoftService _UsuarioFlexoftService;
        }
        /// <summary>
        /// Retorna el ID Flexoft del usuario autenticado.
        /// </summary>
        [HttpGet("InterObtenerIdFlexoft")]
        public ActionResult<int?> InterObtenerIdFlexoft()
        {
            int? idFlexoft = _userContextService.GetCurrentUserId();
            // Alternativamente: int? idFlexoft = _UsuarioFlexoftService.InterObtenerIdFlexoft(User);

            if (!idFlexoft.HasValue)
            {
                return BadRequest(new
                {
                    mensaje = "No se pudo obtener el ID de usuario Flexoft. Verifique si el usuario está correctamente autenticado."
                });
            }

            return Ok(idFlexoft.Value);
        }

        /// <summary>
        /// Obtiene la lista de usuarios para combo (IdFlexoft, NombreUsuario)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            try
            {
                var usuarios = await _context.Usuarios
                    .Where(u => u.Activo && u.IdFlexoft > 0)
                    .Select(u => new { Key = u.IdFlexoft, Value = u.NombreUsuario })
                    .OrderBy(u => u.Value)
                    .ToDictionaryAsync(u => u.Key, u => u.Value);

                return new OkObjectResult(usuarios);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Obtiene todos los usuarios activos
        /// </summary>
        [HttpGet("activos")]
        public async Task<ActionResult> GetActivos()
        {
            try
            {
                return new OkObjectResult(await _context.Usuarios
                    .Where(u => u.Activo)
                    .OrderBy(u => u.NombreUsuario)
                    .ToListAsync());
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Endpoint para registrar un nuevo usuario en el sistema.
        /// </summary>
        /// <param name="usuarioJson">Datos del usuario a registrar en formato JSON</param>
        /// <param name="foto">Foto del usuario (opcional)</param>
        /// <returns>
        /// - 200 OK si el usuario se registró correctamente
        /// - 400 Bad Request si el email ya existe o si los datos de empresa/dependencia/sector son inválidos
        /// - 500 Internal Server Error si ocurrió un error interno del servidor
        /// </returns>
        /// 
        [HttpPost("registro")]
        [AllowAnonymous]
        public async Task<ActionResult> Registro([FromForm] string usuarioJson, IFormFile? foto)
        {
            try
            {
                // Deserializar el usuario desde el JSON
                var usuario = JsonSerializer.Deserialize<Usuario>(usuarioJson, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (usuario == null)
                {
                    return BadRequest("Datos de usuario inválidos");
                }

                // Verificar si el correo ya está registrado
                if (await _context.Usuarios.AnyAsync(u => u.Email == usuario.Email))
                {
                    return BadRequest("El correo electrónico ya está registrado");
                }

                // Verificar si el IdFlexoft ya está registrado
                if (usuario.IdFlexoft > 0 && await _context.Usuarios.AnyAsync(u => u.IdFlexoft == usuario.IdFlexoft))
                {
                    return BadRequest($"El ID Flexoft {usuario.IdFlexoft} ya está registrado para otro usuario");
                }

                // Verificar si la empresa existe
                var empresa = await _context.Empresas.FindAsync(usuario.EmpresaId);
                if (empresa == null)
                {
                    return BadRequest("La empresa seleccionada no existe");
                }

                // Verificar si la dependencia existe y pertenece a la empresa
                var dependencia = await _context.Dependencias
                    .FirstOrDefaultAsync(d => d.Id == usuario.DependenciaId && d.EmpresaId == usuario.EmpresaId);
                if (dependencia == null)
                {
                    return BadRequest("La dependencia seleccionada no existe o no pertenece a la empresa");
                }

                // Verificar si el sector existe y pertenece a la dependencia
                var sector = await _context.Sectores
                    .FirstOrDefaultAsync(s => s.Id == usuario.SectorId && s.DependenciaId == usuario.DependenciaId);
                if (sector == null)
                {
                    return BadRequest("El sector seleccionado no existe o no pertenece a la dependencia");
                }

                // Procesar la foto si se proporcionó una
                if (foto != null && foto.Length > 0)
                {
                    var fileName = $"{Guid.NewGuid()}{Path.GetExtension(foto.FileName)}";
                    var (filePath, fileUrl) = await _fileStorage.SaveFileAsync(foto, "usuarios", fileName);

                    // Guardar la URL de la foto
                    usuario.FotoUrl = fileUrl;
                }

                // Hashear la contraseña
                using (var sha256 = SHA256.Create())
                {
                    var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(usuario.Password));
                    usuario.Password = Convert.ToBase64String(hashedBytes);
                }

                // Establecer valores predeterminados
                usuario.FechaRegistro = DateTime.Now;
                usuario.Activo = true;

                // Guardar el usuario
                _context.Usuarios.Add(usuario);
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar usuario");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        [HttpGet]
        public async Task<ActionResult<List<Usuario>>> Get()
        {
            try
            {
                // Cargar usuarios con su perfil
                var usuarios = await _context.Usuarios
                    .Include(u => u.Perfil)
                    .ToListAsync();

                _logger.LogInformation($"Se cargaron {usuarios.Count} usuarios");
                
                // Log para diagnosticar perfiles
                foreach (var u in usuarios)
                {
                    _logger.LogInformation($"Usuario: {u.NombreUsuario}, PerfilId: {u.PerfilId}, Perfil: {(u.Perfil != null ? u.Perfil.Nombre : "NULL")}");
                }

                // Proyectar a DTO para enviar al cliente
                var usuariosDto = usuarios.Select(u => new Usuario
                {
                    Id = u.Id,
                    NombreUsuario = u.NombreUsuario,
                    Email = u.Email,
                    Password = "", // Campo requerido pero no lo enviamos por seguridad
                    FechaRegistro = u.FechaRegistro,
                    Activo = u.Activo,
                    FotoUrl = u.FotoUrl,
                    PerfilId = u.PerfilId,
                    EmpresaId = u.EmpresaId,
                    DependenciaId = u.DependenciaId,
                    SectorId = u.SectorId,
                    CountryId = u.CountryId,
                    StateId = u.StateId,
                    CityId = u.CityId,
                    IdFlexoft = u.IdFlexoft,
                    NumUsuar = u.NumUsuar, // Campo necesario para derechos de usuarios
                    Perfil = u.Perfil != null ? new Perfil { Id = u.Perfil.Id, Nombre = u.Perfil.Nombre } : null
                }).ToList();

                return Ok(usuariosDto);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener usuarios");
                return StatusCode(500, "Error interno del servidor");
            }
        }
        // ******   ESTE SE USA PARA TRAER LOS DERECHOS DE ACCESO POR EL IDFLEXOFT   *********
        [HttpGet("{idFlexoft}")]
        public async Task<ActionResult<Usuario>> Get(int idFlexoft)
        {
            try
            {
                var usuario = await _context.Usuarios
                    .Include(u => u.Perfil)
                    .Select(u => new Usuario
                    {
                        Id = u.Id,
                        NombreUsuario = u.NombreUsuario,
                        Email = u.Email,
                        Password = u.Password, // Enviamos la contraseña hasheada
                        FechaRegistro = u.FechaRegistro,
                        Activo = u.Activo,
                        FotoUrl = u.FotoUrl,
                        PerfilId = u.PerfilId,
                        EmpresaId = u.EmpresaId,
                        DependenciaId = u.DependenciaId,
                        SectorId = u.SectorId,
                        CountryId = u.CountryId,
                        StateId = u.StateId,
                        CityId = u.CityId,
                        IdFlexoft = u.IdFlexoft,
                        NumUsuar = u.NumUsuar, // Campo necesario para derechos de usuarios
                        Perfil = u.Perfil != null ? new Perfil { Id = u.Perfil.Id, Nombre = u.Perfil.Nombre } : null
                    })
                    .FirstOrDefaultAsync(u => u.IdFlexoft == idFlexoft);

                if (usuario == null)
                {
                    return NotFound("Usuario no encontrado");
                }

                return Ok(usuario);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener usuario");
                return StatusCode(500, "Error interno del servidor");
            }
        }
        // GET api/Usuarios/ObtenerIdFlexoft?username=xxx&email=yyy
        [HttpGet("ObtenerIdFlexoft")]
        public ActionResult<int?> ObtenerIdFlexoft(string username, string email)
        {
            var usuario = _context.Usuarios.FirstOrDefault(u =>
                u.NombreUsuario == username && u.Email == email);
            return usuario?.IdFlexoft;
        }

        /// <summary>
        /// Obtiene el nombre de la empresa (razón social) según el ID del usuario.
        /// Utiliza el servicio Codiempresa para realizar la consulta.
        /// </summary>
        /// <param name="usuarioId">ID del usuario</param>
        /// <returns>Nombre de la empresa o NotFound si no se encuentra</returns>
        // GET api/Usuarios/NombreEmpresa/5
        [HttpGet("NombreEmpresa/{usuarioId}")]
        public ActionResult<string> ObtenerNombreEmpresa(int usuarioId)
        {
            try
            {
                // Utilizar el servicio Codiempresa para obtener el nombre de la empresa
                var nombreEmpresa = _codiempresaService.ObtenerNombreEmpresa(usuarioId);
                
                if (string.IsNullOrEmpty(nombreEmpresa))
                {
                    return NotFound($"No se encontró empresa para el usuario con ID: {usuarioId}");
                }
                
                return Ok(nombreEmpresa);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener nombre de empresa para usuario {UsuarioId}", usuarioId);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene el nombre de la dependencia del usuario por su ID
        /// </summary>
        /// <param name="usuarioId">ID del usuario</param>
        /// <returns>Nombre de la dependencia o NotFound si no se encuentra</returns>
        // GET api/Usuarios/NombreDependencia/5
        [HttpGet("NombreDependencia/{usuarioId}")]
        public ActionResult<string> ObtenerNombreDependencia(int usuarioId)
        {
            try
            {
                // Utilizar el servicio Codiempresa para obtener el nombre de la dependencia
                var nombreDependencia = _codiempresaService.ObtenerNombreDependencia(usuarioId);
                
                if (string.IsNullOrEmpty(nombreDependencia))
                {
                    return NotFound($"No se encontró dependencia para el usuario con ID: {usuarioId}");
                }
                
                return Ok(nombreDependencia);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener nombre de dependencia para usuario {UsuarioId}", usuarioId);
                return StatusCode(500, "Error interno del servidor");
            }
        }

        /// <summary>
        /// Obtiene la información completa del usuario actual (empresa y dependencia)
        /// Usa usuarios.Id para obtener los datos de empresa y dependencia
        /// </summary>
        /// <returns>Objeto con información de empresa y dependencia del usuario actual</returns>
        // GET api/Usuarios/InfoUsuarioActual
        [HttpGet("InfoUsuarioActual")]
        [Authorize]
        public ActionResult<object> ObtenerInfoUsuarioActual()
        {
            try
            {
                // Obtener el usuarios.Id del usuario actual (no IdFlexoft)
                var usuarioId = _userContextService.GetCurrentUsuarioId();
                
                if (!usuarioId.HasValue)
                {
                    return Unauthorized("No se pudo obtener información del usuario actual");
                }

                // Obtener IdFlexoft del usuario actual
                var idFlexoft = _userContextService.GetCurrentUserId() ?? 0;

                // Obtener nombres de empresa y dependencia usando usuarios.Id
                var nombreEmpresa = _codiempresaService.ObtenerNombreEmpresa(usuarioId.Value);
                var nombreDependencia = _codiempresaService.ObtenerNombreDependencia(usuarioId.Value);
                
                return Ok(new
                {
                    usuarioId = usuarioId.Value,
                    idFlexoft = idFlexoft,
                    empresa = nombreEmpresa ?? "Sin empresa",
                    dependencia = nombreDependencia ?? "Sin dependencia"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener información del usuario actual");
                return StatusCode(500, "Error interno del servidor");
            }
        }
        //------------------------------------------------------

        //-----------------------------------------------------------------
        [HttpPut("{id}")]
        public async Task<ActionResult> Put(int id, [FromForm] string usuarioJson, IFormFile? foto)
        {
            try
            {
                var usuario = await _context.Usuarios.FindAsync(id);
                if (usuario == null)
                {
                    return NotFound("Usuario no encontrado");
                }

                var usuarioActualizado = JsonSerializer.Deserialize<Usuario>(usuarioJson, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (usuarioActualizado == null)
                {
                    return BadRequest("Datos de usuario inválidos");
                }

                // Verificar si el email ya está en uso por otro usuario
                if (await _context.Usuarios.AnyAsync(u => u.Email == usuarioActualizado.Email && u.Id != id))
                {
                    return BadRequest("El correo electrónico ya está registrado");
                }

                // Verificar si el nombre de usuario ya está en uso por otro usuario
                if (await _context.Usuarios.AnyAsync(u => u.NombreUsuario == usuarioActualizado.NombreUsuario && u.Id != id))
                {
                    return BadRequest("El nombre de usuario ya está registrado");
                }

                // Verificar si el IdFlexoft ya está en uso por otro usuario
                if (usuarioActualizado.IdFlexoft > 0 && await _context.Usuarios.AnyAsync(u => u.IdFlexoft == usuarioActualizado.IdFlexoft && u.Id != id))
                {
                    return BadRequest($"El ID Flexoft {usuarioActualizado.IdFlexoft} ya está registrado para otro usuario");
                }

                // Actualizar solo los campos permitidos
                usuario.NombreUsuario = usuarioActualizado.NombreUsuario;
                usuario.Email = usuarioActualizado.Email;
                usuario.PerfilId = usuarioActualizado.PerfilId;
                usuario.EmpresaId = usuarioActualizado.EmpresaId;
                usuario.DependenciaId = usuarioActualizado.DependenciaId;
                usuario.SectorId = usuarioActualizado.SectorId;
                usuario.PerfilId = usuarioActualizado.PerfilId;
                usuario.CountryId = usuarioActualizado.CountryId;
                usuario.StateId = usuarioActualizado.StateId;
                usuario.IdFlexoft = usuarioActualizado.IdFlexoft;
                usuario.CityId = usuarioActualizado.CityId;

                // Actualizar contraseña solo si se proporciona una nueva
                if (!string.IsNullOrEmpty(usuarioActualizado.Password))
                {
                    using (var sha256 = SHA256.Create())
                    {
                        var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(usuarioActualizado.Password));
                        usuario.Password = Convert.ToBase64String(hashedBytes);
                    }
                }

                // Procesar la foto si se proporcionó una nueva
                if (foto != null && foto.Length > 0)
                {
                    // Eliminar la foto anterior si existe
                    if (!string.IsNullOrEmpty(usuario.FotoUrl))
                    {
                        _fileStorage.DeleteFileByUrl(usuario.FotoUrl);
                    }

                    // Guardar la nueva foto
                    var fileName = $"{Guid.NewGuid()}{Path.GetExtension(foto.FileName)}";
                    var (filePath, fileUrl) = await _fileStorage.SaveFileAsync(foto, "usuarios", fileName);

                    // Actualizar la URL de la foto
                    usuario.FotoUrl = fileUrl;
                }

                await _context.SaveChangesAsync();
                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar usuario");
                return StatusCode(500, "Error interno del servidor");
            }
        }

        private string HashPassword(string password)
        {
            return PasswordHasher.HashPassword(password);
        }

        [HttpPut("{id}/password")]
        public async Task<ActionResult<Usuario>> UpdatePassword(int id, [FromBody] PasswordUpdateModel model)
        {
            try
            {
                // Obtener el usuario de la base de datos
                var usuario = await _context.Usuarios.FindAsync(id);
                if (usuario == null)
                {
                    return NotFound("Usuario no encontrado");
                }

                // Hashear la nueva contraseña usando la clase compartida PasswordHasher
                usuario.Password = PasswordHasher.HashPassword(model.Password);
                
                await _context.SaveChangesAsync();

                return Ok(usuario);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al actualizar la contraseña: {ex.Message}");
            }
        }

        [HttpPut("toggle-estado/{id}")]
        public async Task<ActionResult> ToggleEstado(int id)
        {
            try
            {
                var usuario = await _context.Usuarios.FindAsync(id);
                if (usuario == null)
                {
                    return NotFound("Usuario no encontrado");
                }

                usuario.Activo = !usuario.Activo;
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al cambiar estado del usuario");
                return StatusCode(500, "Error interno del servidor");
            }
        }
    }

    public class PasswordUpdateModel
    {
        public int Id { get; set; }
        public string Password { get; set; } = string.Empty;
    }
}
// TODO: OBTENER EMPRESA A PARTIR DE USUARIO ObtenerNombreEmpresa
// Método para obtener el ID del usuario logueado DESDE EL FRONT
//private async Task<int?> ObtenerIdUsuarioLogueado()
//{
//    try
//    {
//        // Llamada al endpoint que obtiene el ID Flexoft del usuario autenticado
//        var response = await repository.Get<int>("api/Usuarios/InterObtenerIdFlexoft");

//        if (!response.Error && response.Response.HasValue)
//        {
//            return response.Response.Value;
//        }
//        else
//        {
//            // Manejar caso donde no se pudo obtener el ID
//            Console.WriteLine("No se pudo obtener el ID del usuario logueado");
//            return null;
//        }
//    }
//    catch (Exception ex)
//    {
//        // Manejar errores de la llamada
//        Console.WriteLine($"Error al obtener ID usuario: {ex.Message}");
//        return null;
//    }
//}

//// Ejemplo de uso en OnInitializedAsync o cualquier método
//protected override async Task OnInitializedAsync()
//{
//    var idUsuario = await ObtenerIdUsuarioLogueado();

//    if (idUsuario.HasValue)
//    {
//        // Usar el ID del usuario para otras operaciones
//        Console.WriteLine($"Usuario logueado ID: {idUsuario.Value}");

//        // Ejemplo: Obtener el nombre de la empresa del usuario
//        var responseEmpresa = await repository.Get<string>($"api/Usuarios/NombreEmpresa/{idUsuario.Value}");
//        if (!responseEmpresa.Error)
//        {
//            var nombreEmpresa = responseEmpresa.Response;
//            Console.WriteLine($"Empresa: {nombreEmpresa}");
//        }
//    }