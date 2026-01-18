# Validación de Dependencia y Empresa

Este documento explica cómo validar la dependencia y empresa activa del usuario en cualquier parte de la aplicación (frontend o backend).

## 📋 Tabla de Contenidos
- [Frontend (Blazor)](#frontend-blazor)
- [Backend (Controllers)](#backend-controllers)
- [Casos de Uso Comunes](#casos-de-uso-comunes)

---

## Frontend (Blazor)

### 1. Inyectar el Servicio

En cualquier componente `.razor`, inyecta el servicio `DependenciaService`:

```razor
@inject DependenciaService dependenciaService
```

### 2. Ejemplos de Uso

#### Ejemplo 1: Validar si es una dependencia específica

```csharp
@code {
    private bool esNeuquen = false;

    protected override async Task OnInitializedAsync()
    {
        // Validar si la dependencia es "NEUQUEN" (case-insensitive)
        esNeuquen = await dependenciaService.EsDependenciaAsync("NEUQUEN");
        
        if (esNeuquen)
        {
            Console.WriteLine("Usuario pertenece a NEUQUEN");
            // Ejecutar lógica específica para NEUQUEN
        }
        else
        {
            Console.WriteLine("Usuario NO pertenece a NEUQUEN");
            // Ejecutar lógica alternativa
        }
    }
}
```

#### Ejemplo 2: Renderizado condicional según dependencia

```razor
@if (esNeuquen)
{
    <div class="alert alert-info">
        <i class="fas fa-info-circle"></i> Funcionalidad específica para NEUQUEN
    </div>
    
    <button class="btn btn-primary" @onclick="FuncionEspecialNeuquen">
        Acción Especial NEUQUEN
    </button>
}
else
{
    <div class="alert alert-warning">
        <i class="fas fa-exclamation-triangle"></i> Funcionalidad estándar
    </div>
    
    <button class="btn btn-secondary" @onclick="FuncionEstandar">
        Acción Estándar
    </button>
}

@code {
    private bool esNeuquen = false;

    protected override async Task OnInitializedAsync()
    {
        esNeuquen = await dependenciaService.EsDependenciaAsync("NEUQUEN");
    }

    private void FuncionEspecialNeuquen()
    {
        // Lógica específica para NEUQUEN
    }

    private void FuncionEstandar()
    {
        // Lógica estándar
    }
}
```

#### Ejemplo 3: Validar múltiples dependencias

```csharp
@code {
    private bool esDependenciaEspecial = false;

    protected override async Task OnInitializedAsync()
    {
        // Validar si es NEUQUEN, AVELLANEDA o CAPITAL
        esDependenciaEspecial = await dependenciaService.EsAlgunaDependenciaAsync(
            "NEUQUEN", 
            "AVELLANEDA", 
            "CAPITAL"
        );
        
        if (esDependenciaEspecial)
        {
            // Ejecutar lógica para dependencias especiales
        }
    }
}
```

#### Ejemplo 4: Obtener nombre de dependencia y empresa

```csharp
@code {
    private string? dependenciaActual;
    private string? empresaActual;

    protected override async Task OnInitializedAsync()
    {
        dependenciaActual = await dependenciaService.ObtenerDependenciaActivaAsync();
        empresaActual = await dependenciaService.ObtenerEmpresaActivaAsync();
        
        Console.WriteLine($"Dependencia: {dependenciaActual}");
        Console.WriteLine($"Empresa: {empresaActual}");
    }
}
```

#### Ejemplo 5: Validar empresa

```csharp
@code {
    private bool esSiamArcon = false;

    protected override async Task OnInitializedAsync()
    {
        // Validar si la empresa es "SIAM ARCON" (case-insensitive)
        esSiamArcon = await dependenciaService.EsEmpresaAsync("SIAM ARCON");
        
        if (esSiamArcon)
        {
            // Lógica específica para SIAM ARCON
        }
    }
}
```

#### Ejemplo 6: Lógica condicional en métodos

```csharp
@code {
    private async Task GuardarPedido()
    {
        // Validar dependencia antes de guardar
        var esNeuquen = await dependenciaService.EsDependenciaAsync("NEUQUEN");
        
        if (esNeuquen)
        {
            // Aplicar reglas especiales para NEUQUEN
            await GuardarConReglasNeuquen();
        }
        else
        {
            // Aplicar reglas estándar
            await GuardarEstandar();
        }
    }

    private async Task GuardarConReglasNeuquen()
    {
        // Implementación específica para NEUQUEN
    }

    private async Task GuardarEstandar()
    {
        // Implementación estándar
    }
}
```

---

## Backend (Controllers)

### 1. Inyectar el Servicio

En cualquier controller, inyecta el servicio `IEmpresaDependenciaService`:

```csharp
using IMPLANPROD.Server.Services;

public class MiController : ControllerBase
{
    private readonly IEmpresaDependenciaService _empresaDependenciaService;

    public MiController(IEmpresaDependenciaService empresaDependenciaService)
    {
        _empresaDependenciaService = empresaDependenciaService;
    }
}
```

**Nota:** El servicio `IEmpresaDependenciaService` ya está registrado en `Program.cs` del servidor, no necesitas registrarlo nuevamente.

### 2. Ejemplos de Uso

#### Ejemplo 1: Validar dependencia en un endpoint

```csharp
[HttpGet("datos")]
public async Task<ActionResult> GetDatos()
{
    // Validar si el usuario pertenece a NEUQUEN
    var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
    
    if (esNeuquen)
    {
        // Retornar datos específicos para NEUQUEN
        var datosNeuquen = await ObtenerDatosNeuquen();
        return Ok(datosNeuquen);
    }
    else
    {
        // Retornar datos estándar
        var datosEstandar = await ObtenerDatosEstandar();
        return Ok(datosEstandar);
    }
}
```

#### Ejemplo 2: Aplicar lógica de negocio según dependencia

```csharp
[HttpPost("crear-orden")]
public async Task<ActionResult> CrearOrden([FromBody] OrdenDTO orden)
{
    var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
    
    if (esNeuquen)
    {
        // Aplicar descuento especial para NEUQUEN
        orden.Descuento = 15;
        orden.ObservacionesInternas = "Descuento especial NEUQUEN aplicado";
    }
    
    // Continuar con la creación de la orden
    var resultado = await _ordenService.CrearOrden(orden);
    return Ok(resultado);
}
```

#### Ejemplo 3: Validar permisos según dependencia

```csharp
[HttpDelete("eliminar/{id}")]
public async Task<ActionResult> EliminarRegistro(int id)
{
    var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
    
    if (!esNeuquen)
    {
        return Forbid("Solo usuarios de NEUQUEN pueden eliminar registros");
    }
    
    // Proceder con la eliminación
    await _service.Eliminar(id);
    return Ok();
}
```

#### Ejemplo 4: Obtener información de dependencia

```csharp
[HttpGet("info-usuario")]
public ActionResult GetInfoUsuario()
{
    var dependencia = _empresaDependenciaService.ObtenerNombreDependenciaActual();
    var empresa = _empresaDependenciaService.ObtenerNombreEmpresaActual();
    var dependenciaId = _empresaDependenciaService.ObtenerDependenciaIdActual();
    var empresaId = _empresaDependenciaService.ObtenerEmpresaIdActual();
    
    return Ok(new
    {
        Dependencia = dependencia,
        Empresa = empresa,
        DependenciaId = dependenciaId,
        EmpresaId = empresaId
    });
}
```

#### Ejemplo 5: Validar empresa y dependencia juntas

```csharp
[HttpPost("accion-especial")]
public async Task<ActionResult> AccionEspecial()
{
    var esValido = _empresaDependenciaService.EsEmpresaYDependencia(
        "SIAM ARCON", 
        "NEUQUEN"
    );
    
    if (!esValido)
    {
        return BadRequest("Esta acción solo está disponible para SIAM ARCON - NEUQUEN");
    }
    
    // Ejecutar acción especial
    await _service.EjecutarAccionEspecial();
    return Ok();
}
```

#### Ejemplo 6: Filtrar consultas según dependencia

```csharp
[HttpGet("pedidos")]
public async Task<ActionResult<List<Pedido>>> GetPedidos()
{
    var query = _context.Pedidos.AsQueryable();
    
    var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
    
    if (esNeuquen)
    {
        // Para NEUQUEN, mostrar solo pedidos de su región
        var dependenciaId = _empresaDependenciaService.ObtenerDependenciaIdActual();
        query = query.Where(p => p.DependenciaId == dependenciaId);
    }
    
    var pedidos = await query.ToListAsync();
    return Ok(pedidos);
}
```

---

## Casos de Uso Comunes

### 1. Precios Diferentes por Dependencia

**Frontend:**
```csharp
private async Task<decimal> CalcularPrecio(decimal precioBase)
{
    var esNeuquen = await dependenciaService.EsDependenciaAsync("NEUQUEN");
    
    if (esNeuquen)
    {
        return precioBase * 0.85m; // 15% descuento para NEUQUEN
    }
    
    return precioBase;
}
```

**Backend:**
```csharp
[HttpGet("precio/{productoId}")]
public ActionResult<decimal> GetPrecio(int productoId)
{
    var precioBase = _context.Productos
        .Where(p => p.Id == productoId)
        .Select(p => p.Precio)
        .FirstOrDefault();
    
    var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
    
    if (esNeuquen)
    {
        return Ok(precioBase * 0.85m); // 15% descuento
    }
    
    return Ok(precioBase);
}
```

### 2. Flujos de Aprobación Diferentes

**Backend:**
```csharp
[HttpPost("aprobar-orden/{id}")]
public async Task<ActionResult> AprobarOrden(int id)
{
    var orden = await _context.Ordenes.FindAsync(id);
    
    var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
    
    if (esNeuquen)
    {
        // NEUQUEN requiere doble aprobación
        if (!orden.PrimeraAprobacion)
        {
            orden.PrimeraAprobacion = true;
            await _context.SaveChangesAsync();
            return Ok("Primera aprobación registrada. Se requiere segunda aprobación.");
        }
    }
    
    // Aprobar definitivamente
    orden.Estado = "APROBADA";
    await _context.SaveChangesAsync();
    return Ok("Orden aprobada");
}
```

### 3. Reportes Personalizados

**Frontend:**
```csharp
private async Task GenerarReporte()
{
    var esNeuquen = await dependenciaService.EsDependenciaAsync("NEUQUEN");
    
    if (esNeuquen)
    {
        await GenerarReporteNeuquen();
    }
    else
    {
        await GenerarReporteEstandar();
    }
}
```

### 4. Configuraciones Específicas

**Backend:**
```csharp
[HttpGet("configuracion")]
public ActionResult GetConfiguracion()
{
    var dependencia = _empresaDependenciaService.ObtenerNombreDependenciaActual();
    
    var config = dependencia?.ToUpper() switch
    {
        "NEUQUEN" => new { MaxDescuento = 20, RequiereAprobacion = true },
        "AVELLANEDA" => new { MaxDescuento = 15, RequiereAprobacion = false },
        "CAPITAL" => new { MaxDescuento = 10, RequiereAprobacion = true },
        _ => new { MaxDescuento = 5, RequiereAprobacion = false }
    };
    
    return Ok(config);
}
```

---

## 🔑 Características Importantes

### Case-Insensitive
Todas las validaciones son **case-insensitive**, por lo que estas comparaciones son equivalentes:
- `"NEUQUEN"` = `"neuquen"` = `"Neuquen"` = `"NeUqUeN"`

### Caché (Solo Frontend)
El servicio frontend implementa caché de 5 minutos para optimizar rendimiento. Para forzar actualización:
```csharp
dependenciaService.LimpiarCache();
```

### Null Safety
Todos los métodos manejan correctamente valores `null` y retornan `false` o `null` según corresponda.

---

## 📝 Notas Adicionales

1. **Frontend**: El servicio está registrado como `Scoped` en `Program.cs`
2. **Backend**: El servicio está registrado como `Scoped` en `Program.cs` del servidor
3. **Seguridad**: Las validaciones en frontend son solo para UX. Siempre validar en backend para seguridad real.
4. **Performance**: El backend no usa caché para garantizar datos actualizados en cada request.

### Propiedades de las Entidades

**Entidad Empresa:**
- Propiedad para nombre: `razon` (no `NombreEmpresa`)
- Ejemplo: `empresa.razon` retorna "SIAM ARCON"

**Entidad Dependencia:**
- Propiedad para nombre: `dependencia` (no `NombreDependencia`)
- Ejemplo: `dependencia.dependencia` retorna "NEUQUEN"

**Servicio Backend (`EmpresaDependenciaService`):**
```csharp
// Obtiene el nombre de la empresa
public string? ObtenerNombreEmpresaActual()
{
    var usuario = ObtenerUsuarioCompleto();
    return usuario?.Empresa?.razon;  // ✅ Usa 'razon'
}

// Obtiene el nombre de la dependencia
public string? ObtenerNombreDependenciaActual()
{
    var usuario = ObtenerUsuarioCompleto();
    return usuario?.Dependencia?.dependencia;  // ✅ Usa 'dependencia'
}
```

---

## 🚀 Ejemplo Completo: Orden de Reparación

### Frontend (OrdenReparacionForm.razor)

```razor
@inject DependenciaService dependenciaService

<h3>Orden de Reparación</h3>

@if (esNeuquen)
{
    <div class="alert alert-info">
        <i class="fas fa-info-circle"></i> Procesamiento especial para NEUQUEN
    </div>
}

<button class="btn btn-primary" @onclick="GenerarOrden">
    Generar Orden
</button>

@code {
    private bool esNeuquen = false;

    protected override async Task OnInitializedAsync()
    {
        esNeuquen = await dependenciaService.EsDependenciaAsync("NEUQUEN");
    }

    private async Task GenerarOrden()
    {
        if (esNeuquen)
        {
            // Aplicar lógica especial para NEUQUEN
            await GenerarOrdenNeuquen();
        }
        else
        {
            await GenerarOrdenEstandar();
        }
    }

    private async Task GenerarOrdenNeuquen()
    {
        // Implementación específica
    }

    private async Task GenerarOrdenEstandar()
    {
        // Implementación estándar
    }
}
```

### Backend (ReparaController.cs)

```csharp
[HttpPost("generar-orden")]
public async Task<ActionResult> GenerarOrden([FromBody] OrdenReparacionDTO orden)
{
    var esNeuquen = _empresaDependenciaService.EsDependencia("NEUQUEN");
    
    if (esNeuquen)
    {
        // Aplicar reglas especiales para NEUQUEN
        orden.PrioridadAlta = true;
        orden.TiempoEstimado = orden.TiempoEstimado * 0.8; // 20% más rápido
    }
    
    var resultado = await _reparaService.GenerarOrden(orden);
    return Ok(resultado);
}
```

---

## ✅ Resumen

- **Frontend**: Usa `DependenciaService` con `@inject`
- **Backend**: Usa `IEmpresaDependenciaService` en el constructor
- **Validación**: Siempre case-insensitive
- **Métodos principales**:
  - `EsDependenciaAsync("NEUQUEN")` / `EsDependencia("NEUQUEN")`
  - `EsEmpresaAsync("SIAM ARCON")` / `EsEmpresa("SIAM ARCON")`
  - `ObtenerDependenciaActivaAsync()` / `ObtenerNombreDependenciaActual()`
  - `ObtenerEmpresaActivaAsync()` / `ObtenerNombreEmpresaActual()`
