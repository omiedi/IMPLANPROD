# Guía de Implementación: Servicio Centralizado de Contexto de Usuario

## Introducción

Esta guía explica la implementación del servicio centralizado para acceder al usuario actual y su código de empresa desde cualquier parte de la aplicación. Esta solución evita tener que inyectar múltiples servicios relacionados con el usuario en cada controlador.

## Componentes Implementados

### 1. Servicio de Contexto de Usuario

- **IUserContextService**: Interfaz que define los métodos para acceder a la información del usuario.
- **UserContextService**: Implementación que utiliza los servicios existentes para obtener la información del usuario.

### 2. Middleware de Contexto de Usuario

- **UserContextMiddleware**: Middleware que establece el contexto del usuario en cada solicitud HTTP.
- **MiddlewareExtensions**: Extensiones para registrar el middleware en la pipeline de la aplicación.

## Cómo Utilizar el Servicio

### En Controladores

```csharp
[ApiController]
[Route("api/ejemplo")]
public class EjemploController : ControllerBase
{
    private readonly IUserContextService _userContextService;

    public EjemploController(IUserContextService userContextService)
    {
        _userContextService = userContextService;
    }

    [HttpGet]
    public IActionResult Get()
    {
        // Obtener información del usuario
        int? idUsuario = _userContextService.GetCurrentUserId();
        short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();
        
        // Verificar autenticación
        if (!_userContextService.IsAuthenticated())
        {
            return Unauthorized("Sesión caducada");
        }
        
        // Resto del código...
    }
}
```

### En Servicios

```csharp
public class MiServicio
{
    private readonly IUserContextService _userContextService;

    public MiServicio(IUserContextService userContextService)
    {
        _userContextService = userContextService;
    }

    public void RealizarOperacion()
    {
        int? idUsuario = _userContextService.GetCurrentUserId();
        short? codigoEmpresa = _userContextService.GetCurrentCodigoEmpresa();
        
        // Utilizar la información...
    }
}
```

## Ventajas de Esta Implementación

1. **Centralización**: Toda la lógica relacionada con la obtención del usuario y código de empresa está en un solo lugar.
2. **Reutilización**: No es necesario inyectar múltiples servicios en cada controlador.
3. **Mantenibilidad**: Facilita los cambios futuros en la lógica de obtención de información del usuario.
4. **Consistencia**: Garantiza que la información del usuario se obtenga de la misma manera en toda la aplicación.

## Archivos Modificados

Para encontrar todos los archivos relacionados con esta implementación, busca el comentario:

```csharp
// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO
```

## Próximos Pasos

1. Actualizar los controladores existentes para utilizar el nuevo servicio.
2. Considerar la implementación de un filtro de acción para verificar automáticamente la autenticación.
3. Extender el servicio para proporcionar información adicional del usuario según sea necesario.
