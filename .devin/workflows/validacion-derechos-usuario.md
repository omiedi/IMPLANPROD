---
description: Patrón completo para implementar validación de derechos de usuario en páginas Blazor usando DER_ACCE
---

# Workflow: Validación de Derechos de Usuario

## Descripción
Implementa validación de derechos de acceso granular por aplicación en páginas Blazor, con sincronización bidireccional con el sistema VB6 legacy mediante la tabla DER_ACCE.

## Arquitectura del Sistema

### Entidades Principales
- **DerAcce (DER_ACCE)**: Tabla de derechos de acceso
  - NUM_USER (short): IdFlexoft del usuario (compatibilidad VB6)
  - COD_APLI (string): Código de aplicación
  - DER_READ (short): 0=No, 1=Sí lectura
  - DER_WRITE (short): 0=No, 1=Sí escritura

- **IndiApli (INDI_APLI)**: Catálogo de aplicaciones
  - COD_APLI (string): Código único
  - DES_APLI (string): Descripción

### DTO para UI
- **DerechoUsuarioDTO**:
  - CodApli: string
  - DesApli: string
  - TieneLectura: bool (binding bidireccional)
  - TieneEscritura: bool (binding bidireccional)

## Pasos de Implementación

### 1. Configuración Inicial en la Página Blazor

```csharp
@page "/ruta-pagina"
@using IMPLANPROD.Client.Services
@inject DerechosUsuarioClientService derechosService
@inject NavigationManager NavigationManager

@code {
    // Variables de control de permisos
    private bool tieneAccesoEscritura = false;
    private bool verificandoPermisos = true;
    private const string COD_APLI = "CODIGO_APLICACION";
    private const string DES_APLI = "Descripción de la aplicación";
}
```

### 2. Verificación en OnInitializedAsync

```csharp
protected override async Task OnInitializedAsync()
{
    // Verificar derechos de acceso del usuario
    verificandoPermisos = true;
    
    // Verificar si el usuario tiene acceso de lectura mínimo
    var nivelAcceso = await derechosService.VerificarDerechoAsync(
        COD_APLI,
        DES_APLI,
        mostrarMensaje: true // Mostrar mensaje si no tiene acceso
    );

    // Si no tiene acceso (nivel 0), redirigir al inicio
    if (nivelAcceso == 0)
    {
        Console.WriteLine($"Usuario sin acceso a {COD_APLI}. Redirigiendo...");
        await Task.Delay(2000); // Esperar 2 segundos para que vea el mensaje
        NavigationManager.NavigateTo("/");
        return;
    }

    // Determinar si tiene permisos de escritura (nivel 2)
    tieneAccesoEscritura = nivelAcceso >= 2;
    
    Console.WriteLine($"Acceso a {COD_APLI}: Nivel {nivelAcceso} - Escritura: {tieneAccesoEscritura}");
    
    verificandoPermisos = false;
    
    // Cargar datos solo si tiene acceso
    await LoadAsync();
}
```

### 3. Indicadores Visuales en la UI

```html
@if (verificandoPermisos)
{
    <div class="alert alert-info">
        <div class="spinner-border spinner-border-sm me-2" role="status">
            <span class="visually-hidden">Verificando permisos...</span>
        </div>
        <span>Verificando permisos de acceso...</span>
    </div>
}
else if (!tieneAccesoEscritura)
{
    <div class="alert alert-warning">
        <i class="fas fa-exclamation-triangle me-2"></i>
        <strong>Modo Solo Lectura:</strong> Usted tiene permisos de solo lectura. 
        No puede crear, editar o dar de baja registros.
    </div>
}
```

### 4. Botones con Permisos Condicionales

```html
<!-- Botón Nuevo Registro -->
<button class="btn btn-primary" 
        @onclick="NuevoRegistro" 
        disabled="@(!tieneAccesoEscritura)">
    @if (tieneAccesoEscritura)
    {
        <i class="fas fa-plus"></i>
        <text> Nuevo Registro</text>
    }
    else
    {
        <i class="fas fa-lock"></i>
        <text> Nuevo Registro (Sin Permisos)</text>
    }
</button>
```

### 5. Botones de Acción en Tabla

```html
<!-- Botón Editar -->
<button class="btn btn-sm btn-outline-primary me-1" 
        @onclick="() => EditarRegistro(registro.Id)" 
        disabled="@(!tieneAccesoEscritura)"
        title="@(tieneAccesoEscritura ? "Editar registro" : "Sin permisos de escritura")">
    <i class="fas @(tieneAccesoEscritura ? "fa-edit" : "fa-lock")"></i>
</button>

<!-- Botón Dar de Baja -->
<button class="btn btn-sm btn-outline-danger" 
        @onclick="() => DarDeBaja(registro.Id)" 
        disabled="@(!tieneAccesoEscritura)"
        title="@(tieneAccesoEscritura ? "Dar de baja registro" : "Sin permisos de escritura")">
    <i class="fas @(tieneAccesoEscritura ? "fa-trash" : "fa-lock")"></i>
</button>
```

## Niveles de Acceso

| Nivel | Comportamiento | Botones |
|-------|----------------|---------|
| 0 (Sin acceso) | Muestra mensaje de error, espera 2 segundos, redirige a página principal | No se muestran |
| 1 (Solo lectura) | Muestra alerta de modo solo lectura, carga datos, permite visualización | ❌ Todos deshabilitados |
| 2 (Lectura y escritura) | Sin alertas restrictivas, permite todas las operaciones | ✅ Todos habilitados |

## Configuración en Base de Datos

### Requisitos previos:
1. El campo `USUARIOS.IdFlexoft` debe contener el número de usuario del sistema legacy
2. En la tabla `DER_ACCE` debe existir un registro con:
   - `NUM_USER` = valor de `USUARIOS.IdFlexoft`
   - `COD_APLI` = código de la aplicación (ej: "CLIENTES")
   - `DER_READ` = 1 (para lectura)
   - `DER_WRITE` = 1 (para escritura)
3. El usuario debe cerrar sesión y volver a iniciar sesión después de configurar los derechos

## Endpoints API Requeridos

- `GET /api/usuarios/{id}` - Obtener datos del usuario por Id (identity)
- `GET /api/DerechosUsuario/usuario/{numUser}` - Obtener derechos del usuario
- `POST /api/DerechosUsuario/usuario/{numUser}` - Guardar derechos (lista completa)

## Debugging y Logging

### Console Logs disponibles:
```csharp
Console.WriteLine($"=== DEBUG: Verificación de Derechos ===");
Console.WriteLine($"UserInfo obtenido: true");
Console.WriteLine($"IdFlexoft: {idFlexoft}");
Console.WriteLine($"Id: {id}");
Console.WriteLine($"NombreUsuario: {nombreUsuario}");
Console.WriteLine($"Acceso a {COD_APLI}: Nivel {nivelAcceso} - Escritura: {tieneAccesoEscritura}");
```

### Verificación de Datos:
1. Inspeccionar Network tab: Payload de POST
2. Verificar tabla DER_ACCE en SQL Server
3. Confirmar NUM_USER = IdFlexoft correcto
4. Validar COD_APLI existe en INDI_APLI

## Checklist de Implementación

- [ ] Inyectar DerechosUsuarioClientService
- [ ] Declarar variables de control (tieneAccesoEscritura, verificandoPermisos)
- [ ] Definir constantes COD_APLI y DES_APLI
- [ ] Verificar derechos en OnInitializedAsync
- [ ] Manejar redirección si nivel = 0
- [ ] Agregar indicadores visuales (spinner, alertas)
- [ ] Deshabilitar botones según tieneAccesoEscritura
- [ ] Agregar iconos condicionales (candado vs acción)
- [ ] Agregar tooltips informativos
- [ ] Probar con usuarios de diferentes niveles

## Compatibilidad con VB6

### Código VB6 Original:
```vb
For U& = 1 To FIN&
    CAPLI$ = Trim$(MSF.TextMatrix(U&, 4))
    DRX1% = 0: If Left(MSF.TextMatrix(U&, 2), 1) = "S" Then DRX1% = 1
    DRX2% = 0: If Left(MSF.TextMatrix(U&, 3), 1) = "S" Then DRX2% = 1
    
    SQLX$ = "SELECT * FROM DER_ACCE WHERE NUM_USER = " & USEL%
    SQLX$ = SQLX$ + " AND COD_APLI = '" & CAPLI$ & "'"
    
    If .EOF Then .AddNew
    !NUM_USER = USEL%
    !COD_APLI = CAPLI$
    !DER_READ = DRX1%
    !DER_WRITE = DRX2%
    .Update
Next U&
```

### Equivalencia C#:
- `USEL%` → `IdFlexoft` (parámetro de ruta)
- `MSF grid` → `List<DerechoUsuarioDTO>`
- `DRX1%, DRX2%` → `TieneLectura, TieneEscritura` (bool)
- `AddNew/Update` → EF Core Add/Update con SaveChangesAsync

## Consideraciones Importantes

⚠️ **IdFlexoft vs UsuarioId**: Usar IdFlexoft para DER_ACCE.NUM_USER (compatibilidad VB6)
⚠️ **Transacciones**: Backend usa transacción para garantizar atomicidad
⚠️ **Performance**: LEFT JOIN puede ser costoso con muchas aplicaciones
⚠️ **Sincronización**: Cambios se reflejan inmediatamente en VB6
✅ **Validación**: Backend valida que COD_APLI exista en INDI_APLI
