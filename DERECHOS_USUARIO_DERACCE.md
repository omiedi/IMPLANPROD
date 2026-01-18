# Sistema de Verificación de Derechos de Usuario (DERACCE)

## Descripción General

Este sistema replica la funcionalidad `DERACCE` del código VB6 original, permitiendo verificar los derechos de acceso de los usuarios sobre diferentes módulos/aplicaciones del sistema.

## Componentes Implementados

### Backend

1. **DerechosUsuarioService.cs** - Servicio que maneja la lógica de verificación
2. **DerechosUsuarioController.cs** - Controlador API con endpoints REST
3. **IDerechosUsuarioService.cs** - Interfaz del servicio

### Frontend

1. **DerechosUsuarioClientService.cs** - Servicio cliente para uso en componentes Blazor

## Niveles de Acceso

El sistema retorna los siguientes valores:

- **0** = Sin acceso (el usuario no tiene permisos)
- **1** = Solo lectura (el usuario puede ver pero no modificar)
- **2** = Lectura y escritura (el usuario tiene permisos completos)

## Funcionalidad Automática

### Creación Automática de Registros en INDI_APLI

Cuando se verifica un derecho que no existe en la tabla `INDI_APLI`, el sistema automáticamente:

1. Crea el registro en `INDI_APLI` con el código y descripción proporcionados
2. Retorna `0` (sin acceso) ya que el usuario no tiene permisos asignados aún
3. Permite que posteriormente un administrador asigne los permisos correspondientes

Esto replica exactamente la lógica del VB6:

```vb
' Código VB6 original
If CantRegistros&("INDI_APLI", CONDI$) > 0 Then GoTo 120
SQLX$ = "SELECT * FROM INDI_APLI WHERE COD_APLI = '" & MDY32$ & "'"
RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
With RS
  If .EOF Then
      .AddNew
  End If
  !COD_APLI = MDY32$  
  !DES_APLI = Left$(APM$, 128)
  .Update
End With
```

## Uso en el Frontend

### Ejemplo 1: Verificación Básica con Mensajes

```csharp
@inject DerechosUsuarioClientService derechosService

private async Task AprobarPedido()
{
    // Verificar derechos antes de ejecutar la acción
    var nivelAcceso = await derechosService.VerificarDerechoAsync(
        "APROPEDI", 
        "APROBAR PEDIDOS PENDIENTES"
    );
    
    if (nivelAcceso < 2)
    {
        // El servicio ya mostró el mensaje de error apropiado
        return;
    }
    
    // Usuario tiene permisos, continuar con la aprobación
    await ProcesarAprobacion();
}
```

### Ejemplo 2: Verificación Sin Mostrar Mensajes

```csharp
@inject DerechosUsuarioClientService derechosService

private async Task VerificarAccesoSilencioso()
{
    // Verificar sin mostrar mensajes (equivalente a "NOMESS" en VB6)
    var nivelAcceso = await derechosService.VerificarDerechoAsync(
        "MODPEDIDO", 
        "MODIFICAR PEDIDOS",
        mostrarMensaje: false  // No muestra alertas
    );
    
    if (nivelAcceso < 2)
    {
        // Manejar el error de forma personalizada
        Console.WriteLine("Usuario sin permisos");
        return;
    }
    
    // Continuar...
}
```

### Ejemplo 3: Métodos Simplificados

```csharp
@inject DerechosUsuarioClientService derechosService

// Verificar si tiene acceso de escritura
private async Task<bool> PuedeModificar()
{
    return await derechosService.TieneAccesoEscrituraAsync(
        "MODPEDIDO",
        "MODIFICAR PEDIDOS"
    );
}

// Verificar si tiene al menos acceso de lectura
private async Task<bool> PuedeVer()
{
    return await derechosService.TieneAccesoLecturaAsync(
        "VERPEDIDO",
        "VER PEDIDOS"
    );
}

// Uso en el código
protected override async Task OnInitializedAsync()
{
    if (!await PuedeVer())
    {
        // Redirigir o mostrar mensaje
        navigationManager.NavigateTo("/");
        return;
    }
    
    // Cargar datos...
}
```

### Ejemplo 4: Deshabilitar Botones Según Permisos

```razor
@inject DerechosUsuarioClientService derechosService

<button class="btn btn-primary" 
        disabled="@(!puedeModificar)" 
        @onclick="GuardarCambios">
    Guardar
</button>

@code {
    private bool puedeModificar = false;
    
    protected override async Task OnInitializedAsync()
    {
        // Verificar permisos sin mostrar mensaje
        var nivel = await derechosService.VerificarDerechoAsync(
            "MODPEDIDO",
            "MODIFICAR PEDIDOS",
            mostrarMensaje: false
        );
        
        puedeModificar = nivel >= 2;
    }
}
```

### Ejemplo 5: Verificación en Múltiples Puntos

```csharp
@inject DerechosUsuarioClientService derechosService

private async Task EliminarRegistro()
{
    // Verificar permisos específicos para eliminación
    var nivel = await derechosService.VerificarDerechoAsync(
        "ELIMREG",
        "ELIMINAR REGISTROS"
    );
    
    if (nivel < 2)
    {
        return; // Mensaje ya mostrado por el servicio
    }
    
    // Confirmar eliminación
    var confirmado = await sweetAlertService.FireAsync(new SweetAlertOptions
    {
        Title = "¿Confirmar eliminación?",
        Icon = SweetAlertIcon.Warning,
        ShowCancelButton = true
    });
    
    if (confirmado.IsConfirmed)
    {
        await ProcesarEliminacion();
    }
}
```

## Uso en el Backend (API)

### Endpoint Disponible

```
GET /api/DerechosUsuario/verificar/{numUser}/{codApli}/{desApli}
```

**Parámetros:**
- `numUser`: Número de usuario (short)
- `codApli`: Código de aplicación (string, máx 32 caracteres)
- `desApli`: Descripción de aplicación (string, máx 128 caracteres)

**Respuesta:**
- `0`: Sin acceso
- `1`: Solo lectura
- `2`: Lectura y escritura

### Ejemplo de Uso Directo

```csharp
// Desde otro servicio o controlador
public class MiController : ControllerBase
{
    private readonly IDerechosUsuarioService _derechosService;
    
    [HttpPost("aprobar")]
    public async Task<IActionResult> AprobarPedido(int pedidoId)
    {
        // Obtener número de usuario del contexto
        var numUser = ObtenerUsuarioActual();
        
        // Verificar derechos
        var nivel = await _derechosService.VerificarDerechoAsync(
            numUser,
            "APROPEDI",
            "APROBAR PEDIDOS PENDIENTES"
        );
        
        if (nivel < 2)
        {
            return Forbid("No tiene permisos para aprobar pedidos");
        }
        
        // Procesar aprobación...
        return Ok();
    }
}
```

## Mensajes Mostrados al Usuario

### Sin Acceso (Nivel 0)

```
🔒 USUARIO NO HABILITADO

Usuario no Habilitado para Ejecutar esta Opción.

Módulo: APROPEDI
Aplicación: APROBAR PEDIDOS PENDIENTES

Solicite la Autorización Correspondiente al Supervisor del Sistema.
```

### Solo Lectura (Nivel 1)

```
⚠️ ACCESO DE SOLO LECTURA

Usted solo tiene permisos de LECTURA para esta opción.
No puede realizar modificaciones.

Módulo: MODPEDIDO
Aplicación: MODIFICAR PEDIDOS

Si necesita permisos de escritura, solicite la autorización al Supervisor del Sistema.
```

## Flujo de Trabajo

1. **Usuario intenta ejecutar una acción**
2. **Sistema verifica derechos** llamando a `VerificarDerechoAsync`
3. **Si no existe en INDI_APLI**: Se crea automáticamente y retorna 0
4. **Si existe**: Se consulta DER_ACCE para el usuario
5. **Retorna nivel de acceso**: 0, 1 o 2
6. **Muestra mensaje apropiado** (si `mostrarMensaje = true`)
7. **Aplicación decide** si continuar o abortar la acción

## Administración de Permisos

Los permisos se administran a través de la pantalla de **Derechos de Usuario** donde:

1. Se listan todas las aplicaciones de `INDI_APLI`
2. Se pueden asignar permisos de lectura y/o escritura por usuario
3. Los cambios se guardan en la tabla `DER_ACCE`

## Tablas Involucradas

### INDI_APLI
- **COD_APLI**: Código de aplicación (PK, nvarchar(32))
- **DES_APLI**: Descripción de aplicación (nvarchar(128))

### DER_ACCE
- **Id**: Identificador único (PK, int identity)
- **NUM_USER**: Número de usuario (FK)
- **COD_APLI**: Código de aplicación (FK)
- **DER_READ**: Derecho de lectura (0 o 1)
- **DER_WRITE**: Derecho de escritura (0 o 1)

## Logging

El sistema registra todas las verificaciones en los logs:

```
[INFO] Verificando derecho para usuario 1, aplicación 'APROPEDI'
[INFO] Aplicación 'APROPEDI' no existe en INDI_APLI, creando registro...
[INFO] Registro creado en INDI_APLI: 'APROPEDI' - 'APROBAR PEDIDOS PENDIENTES'
[INFO] Usuario 1 no tiene derechos asignados para 'APROPEDI' - Retornando 0 (sin acceso)
```

## Mejores Prácticas

1. **Siempre verificar permisos** antes de operaciones críticas (guardar, eliminar, aprobar)
2. **Usar códigos descriptivos** para COD_APLI (ej: "APROPEDI", "MODPEDIDO")
3. **Descripciones claras** para DES_APLI que expliquen la acción
4. **Verificar en OnInitializedAsync** para deshabilitar UI según permisos
5. **Usar `mostrarMensaje: false`** cuando se verifica solo para UI (botones, etc.)
6. **Usar `mostrarMensaje: true`** cuando el usuario intenta ejecutar una acción

## Compatibilidad con VB6

Este sistema es 100% compatible con el código VB6 original:

- ✅ Misma lógica de verificación
- ✅ Mismos valores de retorno (0, 1, 2)
- ✅ Creación automática en INDI_APLI
- ✅ Truncado de strings a límites de BD
- ✅ Mensajes similares al usuario

## Ejemplo Completo de Implementación

```razor
@page "/pedidos/aprobar"
@inject DerechosUsuarioClientService derechosService
@inject IRepository repository
@inject SweetAlertService sweetAlert

<h3>Aprobar Pedidos Pendientes</h3>

@if (cargando)
{
    <p>Cargando...</p>
}
else if (!tieneAcceso)
{
    <div class="alert alert-warning">
        No tiene permisos para acceder a esta página.
    </div>
}
else
{
    <div class="table-responsive">
        <table class="table">
            <thead>
                <tr>
                    <th>Nro. Pedido</th>
                    <th>Cliente</th>
                    <th>Total</th>
                    <th>Acciones</th>
                </tr>
            </thead>
            <tbody>
                @foreach (var pedido in pedidos)
                {
                    <tr>
                        <td>@pedido.NroPed</td>
                        <td>@pedido.NombreCliente</td>
                        <td>@pedido.Total.ToString("C2")</td>
                        <td>
                            <button class="btn btn-success btn-sm" 
                                    disabled="@(!puedeAprobar)"
                                    @onclick="() => AprobarPedido(pedido.NroPed)">
                                Aprobar
                            </button>
                        </td>
                    </tr>
                }
            </tbody>
        </table>
    </div>
}

@code {
    private bool cargando = true;
    private bool tieneAcceso = false;
    private bool puedeAprobar = false;
    private List<PedidoDTO> pedidos = new();

    protected override async Task OnInitializedAsync()
    {
        // Verificar acceso a la página (con mensaje)
        var nivelAcceso = await derechosService.VerificarDerechoAsync(
            "APROPEDI",
            "APROBAR PEDIDOS PENDIENTES"
        );

        tieneAcceso = nivelAcceso >= 1; // Al menos lectura
        puedeAprobar = nivelAcceso >= 2; // Escritura para aprobar

        if (tieneAcceso)
        {
            await CargarPedidos();
        }

        cargando = false;
    }

    private async Task CargarPedidos()
    {
        var response = await repository.Get<List<PedidoDTO>>("api/pedidos/pendientes");
        if (!response.Error)
        {
            pedidos = response.Response ?? new();
        }
    }

    private async Task AprobarPedido(int nroPed)
    {
        // Verificar permisos nuevamente antes de la acción crítica
        var nivel = await derechosService.VerificarDerechoAsync(
            "APROPEDI",
            "APROBAR PEDIDOS PENDIENTES"
        );

        if (nivel < 2)
        {
            return; // Mensaje ya mostrado
        }

        // Confirmar acción
        var result = await sweetAlert.FireAsync(new SweetAlertOptions
        {
            Title = "¿Aprobar pedido?",
            Text = $"¿Está seguro de aprobar el pedido {nroPed}?",
            Icon = SweetAlertIcon.Question,
            ShowCancelButton = true,
            ConfirmButtonText = "Sí, aprobar",
            CancelButtonText = "Cancelar"
        });

        if (result.IsConfirmed)
        {
            var response = await repository.Put($"api/pedidos/{nroPed}/aprobar", null);
            
            if (!response.Error)
            {
                await sweetAlert.FireAsync("Aprobado", "Pedido aprobado exitosamente", SweetAlertIcon.Success);
                await CargarPedidos(); // Recargar lista
            }
            else
            {
                await sweetAlert.FireAsync("Error", "No se pudo aprobar el pedido", SweetAlertIcon.Error);
            }
        }
    }
}
```

## Soporte

Para dudas o problemas con el sistema de derechos:

1. Verificar logs del servidor para detalles de verificación
2. Confirmar que el usuario existe en la tabla de usuarios
3. Verificar que la aplicación existe en INDI_APLI
4. Confirmar permisos en DER_ACCE para el usuario específico
