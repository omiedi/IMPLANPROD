# Configuración: Uso de usuarios.Id para Empresa y Dependencia

## 📋 RESUMEN

Se ha configurado el sistema para que la información de **empresa y dependencia** mostrada en el NavMenu se obtenga usando `usuarios.Id` del usuario logueado, mientras que la auditoría sigue usando `usuarios.IdFlexoft`.

---

## 🔑 DIFERENCIACIÓN DE IDs

### **usuarios.Id**
- **Uso:** Obtener datos de empresa y dependencia
- **Tipo:** `int` (Primary Key de la tabla Usuarios)
- **Propósito:** Identificación interna del usuario en la base de datos
- **Dónde se usa:**
  - `ICodiempresa.ObtenerNombreEmpresa(int usuarioId)`
  - `ICodiempresa.ObtenerNombreDependencia(int usuarioId)`
  - `UserContextService.GetCurrentUsuarioId()`

### **usuarios.IdFlexoft**
- **Uso:** Auditoría y sistema Flexoft
- **Tipo:** `int` (puede ser NULL)
- **Propósito:** Identificación del usuario en el sistema Flexoft
- **Dónde se usa:**
  - Campo `NumUsuar` en tablas auditables
  - `IUsuarioFlexoftService.InterObtenerIdFlexoft()`
  - `UserContextService.GetCurrentUserId()`

---

## ✅ CAMBIOS IMPLEMENTADOS

### **1. IUserContextService.cs**

Se agregó un nuevo método para obtener `usuarios.Id`:

```csharp
/// <summary>
/// Obtiene el ID real del usuario actual (usuarios.Id) para obtener datos de empresa/dependencia
/// </summary>
/// <returns>usuarios.Id del usuario o null si no está autenticado</returns>
int? GetCurrentUsuarioId();
```

### **2. UserContextService.cs**

**Agregado `DataContext` al constructor:**

```csharp
private readonly DataContext _context;

public UserContextService(
    IHttpContextAccessor httpContextAccessor,
    IUsuarioFlexoftService usuarioFlexoftService,
    ICodiempresa codiempresa,
    DataContext context)
{
    _httpContextAccessor = httpContextAccessor;
    _usuarioFlexoftService = usuarioFlexoftService;
    _codiempresa = codiempresa;
    _context = context;
}
```

**Implementado método `GetCurrentUsuarioId()`:**

```csharp
public int? GetCurrentUsuarioId()
{
    var user = GetCurrentUser();
    if (user?.Identity?.IsAuthenticated ?? false)
    {
        var username = user.Identity?.Name;
        var email = user.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value;

        if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
        {
            var usuario = _context.Usuarios.FirstOrDefault(u =>
                u.NombreUsuario == username && u.Email == email);

            return usuario?.Id;  // Retorna usuarios.Id
        }
    }

    return null;
}
```

### **3. UsuariosController.cs**

**Endpoint `InfoUsuarioActual` actualizado:**

```csharp
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

        // Obtener nombres de empresa y dependencia usando usuarios.Id
        var nombreEmpresa = _codiempresaService.ObtenerNombreEmpresa(usuarioId.Value);
        var nombreDependencia = _codiempresaService.ObtenerNombreDependencia(usuarioId.Value);
        
        return Ok(new
        {
            usuarioId = usuarioId.Value,
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
```

---

## 🔄 FLUJO COMPLETO

```
┌─────────────────────────────────────────────────────────────┐
│ 1. Usuario se loguea en el sistema                         │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. NavMenu carga componente UserInfoDisplay                │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. UserInfoDisplay llama a:                                 │
│    GET api/Usuarios/InfoUsuarioActual                       │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│ 4. UsuariosController.ObtenerInfoUsuarioActual()            │
│    - Llama a: _userContextService.GetCurrentUsuarioId()    │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│ 5. UserContextService.GetCurrentUsuarioId()                 │
│    - Obtiene username y email del ClaimsPrincipal          │
│    - Busca en BD: WHERE NombreUsuario = X AND Email = Y    │
│    - Retorna: usuario.Id (NO IdFlexoft)                    │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│ 6. Con usuarios.Id, obtiene:                                │
│    - ObtenerNombreEmpresa(usuarioId)                        │
│    - ObtenerNombreDependencia(usuarioId)                    │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│ 7. Retorna JSON:                                            │
│    {                                                        │
│      usuarioId: 123,                                        │
│      empresa: "ARCON",                                      │
│      dependencia: "NEUQUEN"                                 │
│    }                                                        │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│ 8. UserInfoDisplay muestra en NavMenu:                     │
│    🏢 Empresa: ARCON                                        │
│    📍 Dependencia: NEUQUEN                                  │
└─────────────────────────────────────────────────────────────┘
```

---

## 📊 MATRIZ DE USO DE IDs

| Operación | Campo Usado | Método/Servicio |
|-----------|-------------|-----------------|
| Mostrar empresa en NavMenu | `usuarios.Id` | `GetCurrentUsuarioId()` |
| Mostrar dependencia en NavMenu | `usuarios.Id` | `GetCurrentUsuarioId()` |
| Auditoría (NumUsuar) | `usuarios.IdFlexoft` | `GetCurrentUserId()` |
| Código de empresa (CodiEmprNet) | `usuarios.EmpresaId` | `GetCurrentCodigoEmpresa()` |
| Obtener nombre empresa | `usuarios.Id` | `ObtenerNombreEmpresa(usuarioId)` |
| Obtener nombre dependencia | `usuarios.Id` | `ObtenerNombreDependencia(usuarioId)` |

---

## 🎯 VENTAJAS DE ESTA CONFIGURACIÓN

1. ✅ **Separación de responsabilidades:**
   - `usuarios.Id` → Datos del usuario (empresa, dependencia)
   - `usuarios.IdFlexoft` → Auditoría y sistema Flexoft

2. ✅ **Flexibilidad:**
   - Un usuario puede tener `IdFlexoft = NULL` pero aún mostrar empresa/dependencia
   - `IdFlexoft` puede cambiar sin afectar la relación con empresa/dependencia

3. ✅ **Consistencia:**
   - Siempre se usa la misma clave primaria (`usuarios.Id`) para relaciones
   - No hay confusión entre IDs

4. ✅ **Mantenibilidad:**
   - Código más claro y fácil de entender
   - Cada ID tiene un propósito específico

---

## 📝 ARCHIVOS MODIFICADOS

1. ✅ `Server/Services/IUserContextService.cs`
   - Agregado método `GetCurrentUsuarioId()`

2. ✅ `Server/Services/UserContextService.cs`
   - Agregado `DataContext` al constructor
   - Implementado método `GetCurrentUsuarioId()`
   - Agregado `using IMPLANPROD.Server.Data;`

3. ✅ `Server/Controllers/UsuariosController.cs`
   - Actualizado `ObtenerInfoUsuarioActual()` para usar `GetCurrentUsuarioId()`

---

## 🔍 VALIDACIONES ADICIONALES

También se agregaron validaciones para `IdFlexoft` único:

### **Alta de usuarios (Registro):**
```csharp
if (usuario.IdFlexoft > 0 && await _context.Usuarios.AnyAsync(u => u.IdFlexoft == usuario.IdFlexoft))
{
    return BadRequest($"El ID Flexoft {usuario.IdFlexoft} ya está registrado para otro usuario");
}
```

### **Edición de usuarios (Put):**
```csharp
if (usuarioActualizado.IdFlexoft > 0 && await _context.Usuarios.AnyAsync(u => u.IdFlexoft == usuarioActualizado.IdFlexoft && u.Id != id))
{
    return BadRequest($"El ID Flexoft {usuarioActualizado.IdFlexoft} ya está registrado para otro usuario");
}
```

---

## ✅ ESTADO FINAL

- ✅ NavMenu muestra empresa y dependencia usando `usuarios.Id`
- ✅ Auditoría sigue usando `usuarios.IdFlexoft`
- ✅ `IdFlexoft` es único en el sistema
- ✅ Sistema funcional y sin errores

---

**Fecha de implementación:** 2025-12-29  
**Versión:** 1.0  
**Estado:** ✅ COMPLETADO
