# Guía: Validación por Empresa y Dependencia

## 📋 RESUMEN

Se ha implementado un sistema completo para:
1. **Obtener** el nombre de empresa y dependencia del usuario actual
2. **Validar** en el código si el usuario pertenece a una empresa/dependencia específica
3. **Mostrar** la información en el navbar de la aplicación

---

## 🎯 COMPONENTES IMPLEMENTADOS

### **1. Backend - Servicios**

#### **ICodiempresa.cs** (Extendido)
```csharp
public interface ICodiempresa
{
    short? IObtenerCodiempresa(ClaimsPrincipal user);
    string? ObtenerNombreEmpresa(int usuarioId);
    string? ObtenerNombreDependencia(int usuarioId);  // ✅ NUEVO
}
```

#### **Codiempresa.cs** (Implementación)
```csharp
public string? ObtenerNombreDependencia(int usuarioId)
{
    var usuario = _context.Usuarios.FirstOrDefault(u => u.Id == usuarioId);
    if (usuario?.DependenciaId == null) return null;
    
    var dependencia = _context.Dependencias.FirstOrDefault(d => d.Id == usuario.DependenciaId);
    return dependencia?.dependencia;
}
```

### **2. Backend - Endpoints**

#### **UsuariosController.cs**

**Endpoint 1: Obtener nombre de dependencia**
```csharp
// GET api/Usuarios/NombreDependencia/5
[HttpGet("NombreDependencia/{usuarioId}")]
public ActionResult<string> ObtenerNombreDependencia(int usuarioId)
```

**Endpoint 2: Obtener información completa del usuario actual**
```csharp
// GET api/Usuarios/InfoUsuarioActual
[HttpGet("InfoUsuarioActual")]
[Authorize]
public ActionResult<object> ObtenerInfoUsuarioActual()
{
    var usuarioId = _userContextService.GetCurrentUserId();
    var nombreEmpresa = _codiempresaService.ObtenerNombreEmpresa(usuarioId.Value);
    var nombreDependencia = _codiempresaService.ObtenerNombreDependencia(usuarioId.Value);
    
    return Ok(new {
        usuarioId,
        empresa = nombreEmpresa ?? "Sin empresa",
        dependencia = nombreDependencia ?? "Sin dependencia"
    });
}
```

### **3. Frontend - Componentes**

#### **UserInfoDisplay.razor**
Componente que muestra la información de empresa y dependencia en el navbar.

**Características:**
- ✅ Carga automática al inicializar
- ✅ Se actualiza al cambiar de sesión
- ✅ Diseño responsive
- ✅ Iconos visuales
- ✅ Estilos integrados

#### **NavMenu.razor** (Modificado)
Se agregó el componente `<UserInfoDisplay />` en la parte superior del menú lateral.

---

## 💻 CÓMO USAR EN TU CÓDIGO

### **Opción 1: Validar en un Controlador**

```csharp
using IMPLANPROD.Server.Services;

[ApiController]
[Route("api/[controller]")]
public class MiController : ControllerBase
{
    private readonly IEmpresaDependenciaService _empresaDependenciaService;

    public MiController(IEmpresaDependenciaService empresaDependenciaService)
    {
        _empresaDependenciaService = empresaDependenciaService;
    }

    [HttpPost("procesar-orden")]
    public IActionResult ProcesarOrden()
    {
        // TU CASO: Validar si es ARCON y NEUQUEN
        if (_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
        {
            // Lógica específica para ARCON - NEUQUEN
            return Ok("Procesando con reglas especiales ARCON-NEUQUEN");
        }
        else
        {
            // Lógica estándar
            return Ok("Procesando con reglas estándar");
        }
    }
}
```

### **Opción 2: Validar en un Servicio**

```csharp
public class OrdenCompraService
{
    private readonly IEmpresaDependenciaService _empresaDependenciaService;

    public OrdenCompraService(IEmpresaDependenciaService empresaDependenciaService)
    {
        _empresaDependenciaService = empresaDependenciaService;
    }

    public decimal CalcularDescuento(decimal monto)
    {
        // Validación específica por empresa y dependencia
        if (_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
        {
            return monto * 0.15m; // 15% descuento
        }
        else if (_empresaDependenciaService.EsEmpresa("ARCON"))
        {
            return monto * 0.10m; // 10% descuento
        }
        
        return 0; // Sin descuento
    }
}
```

### **Opción 3: Validaciones Múltiples**

```csharp
public IActionResult ProcesarPedido()
{
    var empresa = _empresaDependenciaService.ObtenerNombreEmpresaActual();
    var dependencia = _empresaDependenciaService.ObtenerNombreDependenciaActual();

    // Validación con switch
    switch (empresa)
    {
        case "ARCON":
            if (dependencia == "NEUQUEN")
            {
                // Lógica ARCON-NEUQUEN
            }
            else if (dependencia == "BUENOS AIRES")
            {
                // Lógica ARCON-BUENOS AIRES
            }
            else
            {
                // Lógica ARCON general
            }
            break;
            
        default:
            // Lógica estándar
            break;
    }
}
```

---

## 🔧 MÉTODOS DISPONIBLES

### **IEmpresaDependenciaService**

| Método | Descripción | Retorno |
|--------|-------------|---------|
| `EsEmpresa(string nombre)` | Valida si el usuario pertenece a la empresa | `bool` |
| `EsDependencia(string nombre)` | Valida si el usuario pertenece a la dependencia | `bool` |
| `EsEmpresaYDependencia(string empresa, string dependencia)` | Valida ambos criterios | `bool` |
| `ObtenerNombreEmpresaActual()` | Obtiene nombre de empresa del usuario actual | `string?` |
| `ObtenerNombreDependenciaActual()` | Obtiene nombre de dependencia del usuario actual | `string?` |
| `ObtenerEmpresaIdActual()` | Obtiene ID de empresa | `int?` |
| `ObtenerDependenciaIdActual()` | Obtiene ID de dependencia | `int?` |

### **ICodiempresa**

| Método | Descripción | Retorno |
|--------|-------------|---------|
| `IObtenerCodiempresa(ClaimsPrincipal user)` | Obtiene código de empresa | `short?` |
| `ObtenerNombreEmpresa(int usuarioId)` | Obtiene nombre de empresa por ID usuario | `string?` |
| `ObtenerNombreDependencia(int usuarioId)` | Obtiene nombre de dependencia por ID usuario | `string?` |

---

## 📱 VISUALIZACIÓN EN LA UI

El componente `UserInfoDisplay` muestra automáticamente en el navbar:

```
┌─────────────────────────────────────┐
│ 🏢 Empresa: ARCON                   │
│ 📍 Dependencia: NEUQUEN             │
└─────────────────────────────────────┘
```

**Características visuales:**
- Fondo con transparencia
- Borde izquierdo amarillo
- Iconos de Font Awesome
- Texto en mayúsculas
- Responsive para móviles

---

## 🚀 EJEMPLOS PRÁCTICOS

### **Ejemplo 1: Restricción de operación**

```csharp
[HttpPost("operacion-critica")]
public IActionResult OperacionCritica()
{
    // Solo permitir para ARCON-NEUQUEN
    if (!_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
    {
        return Forbid("Esta operación solo está disponible para ARCON-NEUQUEN");
    }
    
    // Ejecutar operación
    return Ok("Operación ejecutada");
}
```

### **Ejemplo 2: Cálculo de precios diferenciados**

```csharp
public decimal CalcularPrecio(decimal precioBase)
{
    if (_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
    {
        return precioBase * 0.85m; // 15% descuento
    }
    else if (_empresaDependenciaService.EsEmpresa("ARCON"))
    {
        return precioBase * 0.90m; // 10% descuento
    }
    
    return precioBase; // Precio estándar
}
```

### **Ejemplo 3: Logging con contexto**

```csharp
public void RegistrarOperacion(string operacion)
{
    var empresa = _empresaDependenciaService.ObtenerNombreEmpresaActual();
    var dependencia = _empresaDependenciaService.ObtenerNombreDependenciaActual();
    
    _logger.LogInformation(
        "Operación: {Operacion} | Empresa: {Empresa} | Dependencia: {Dependencia}",
        operacion, empresa, dependencia
    );
}
```

---

## ✅ VENTAJAS DEL SISTEMA

1. **Centralizado**: Un solo lugar para todas las validaciones
2. **Reutilizable**: Úsalo en cualquier controlador o servicio
3. **Type-safe**: Compilación estricta en C#
4. **Mantenible**: Fácil de modificar y extender
5. **Testeable**: Puedes crear tests unitarios
6. **Sin JavaScript**: Todo en C# como prefieres
7. **Automático**: La UI se actualiza automáticamente

---

## 📝 NOTAS IMPORTANTES

- Las comparaciones de nombres son **case-insensitive** (no distinguen mayúsculas/minúsculas)
- Los métodos retornan `null` si el usuario no está autenticado
- El servicio ya está registrado en `Program.cs`
- Solo necesitas inyectarlo donde lo necesites
- La información se actualiza automáticamente al cambiar de sesión

---

## 🔗 ARCHIVOS RELACIONADOS

### **Backend:**
- `Server/Services/ICodiempresa.cs`
- `Server/Services/Codiempresa.cs`
- `Server/Services/IEmpresaDependenciaService.cs`
- `Server/Services/EmpresaDependenciaService.cs`
- `Server/Controllers/UsuariosController.cs`
- `Server/Program.cs`

### **Frontend:**
- `Client/Shared/UserInfoDisplay.razor`
- `Client/Shared/NavMenu.razor`

### **Documentación:**
- `Docs/EjemploValidacionEmpresaDependencia.cs`
- `Docs/GuiaValidacionEmpresaDependencia.md` (este archivo)

---

## 🎓 CONCLUSIÓN

Ahora tienes un sistema completo para:
- ✅ Validar empresa y dependencia en tu código backend
- ✅ Mostrar la información en la UI automáticamente
- ✅ Aplicar lógica de negocio específica por empresa/dependencia
- ✅ Todo sin necesidad de modificar el login

**¡El sistema está listo para usar!** 🚀
