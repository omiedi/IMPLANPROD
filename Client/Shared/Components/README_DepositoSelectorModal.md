# DepositoSelectorModal - Componente Reutilizable

## Descripción
Componente modal reutilizable para seleccionar depósitos de la tabla TADEPOSI. Permite mostrar una lista filtrada de depósitos activos (Status >= 0) y retorna el código y descripción del depósito seleccionado.

## Ubicación
```
Client/Shared/Components/DepositoSelectorModal.razor
```

## Características
- ✅ Modal Bootstrap responsivo
- ✅ Carga automática de depósitos desde la API
- ✅ Filtrado por Status >= 0 (configurable)
- ✅ Validación de selección
- ✅ Manejo de errores y estados de carga
- ✅ Interfaz intuitiva con iconos Font Awesome
- ✅ Totalmente reutilizable

## Parámetros del Componente

| Parámetro | Tipo | Descripción | Requerido |
|-----------|------|-------------|-----------|
| `IsVisible` | `bool` | Controla la visibilidad del modal | ✅ |
| `OnDepositoSeleccionado` | `EventCallback<DepositoSeleccionadoDTO>` | Evento que se dispara al seleccionar un depósito | ✅ |
| `OnModalClosed` | `EventCallback` | Evento que se dispara al cerrar el modal | ✅ |
| `StatusFilter` | `int` | Filtro para el status de depósitos (por defecto >= 0) | ❌ |
| `DepositoPorDefecto` | `int` | Código del depósito que aparece preseleccionado | ❌ |

## Clase de Respuesta

```csharp
public class DepositoSeleccionado
{
    public int CodigoDepo { get; set; }
    public string Descripcion { get; set; } = string.Empty;
    public string DescripcionAbreviada { get; set; } = string.Empty;
    public int? CodPlanta { get; set; }
}
```

## Uso Básico

### 1. Importar el componente
```razor
@using IMPLANPROD.Client.Shared.Components
```

### 2. Agregar el componente al markup
```razor
<DepositoSelectorModal IsVisible="mostrarModalDeposito"
                       OnDepositoSeleccionado="OnDepositoSeleccionado"
                       OnModalClosed="OnModalCerrado" />
```

### 3. Implementar el código C#
```csharp
@code {
    private bool mostrarModalDeposito = false;
    private DepositoSeleccionado? depositoSeleccionado;

    // Método para abrir el modal
    private async Task AbrirModalDeposito()
    {
        mostrarModalDeposito = true;
        StateHasChanged();
    }

    // Callback cuando se selecciona un depósito
    private async Task OnDepositoSeleccionado(DepositoSeleccionado deposito)
    {
        depositoSeleccionado = deposito;
        mostrarModalDeposito = false;
        StateHasChanged();
        
        // Tu lógica personalizada aquí
        Console.WriteLine($"Depósito: {deposito.CodigoDepo} - {deposito.Descripcion}");
    }

    // Callback cuando se cierra el modal
    private async Task OnModalCerrado()
    {
        mostrarModalDeposito = false;
        StateHasChanged();
    }
}
```

## Ejemplo Completo

```razor
@page "/mi-pagina"
@using IMPLANPROD.Client.Shared.Components

<button class="btn btn-primary" @onclick="AbrirModalDeposito">
    <i class="fas fa-warehouse me-2"></i>
    Seleccionar Depósito
</button>

@if (depositoSeleccionado != null)
{
    <div class="alert alert-success mt-3">
        <strong>Depósito:</strong> @depositoSeleccionado.CodigoDepo - @depositoSeleccionado.Descripcion
    </div>
}

<DepositoSelectorModal IsVisible="mostrarModalDeposito"
                       OnDepositoSeleccionado="OnDepositoSeleccionado"
                       OnModalClosed="OnModalCerrado" />

@code {
    private bool mostrarModalDeposito = false;
    private DepositoSeleccionado? depositoSeleccionado;

    private async Task AbrirModalDeposito()
    {
        mostrarModalDeposito = true;
    }

    private async Task OnDepositoSeleccionado(DepositoSeleccionado deposito)
    {
        depositoSeleccionado = deposito;
        mostrarModalDeposito = false;
        
        // Aquí puedes agregar tu lógica:
        // - Guardar en base de datos
        // - Actualizar otros campos
        // - Mostrar confirmaciones
    }

    private async Task OnModalCerrado()
    {
        mostrarModalDeposito = false;
    }
}
```

## Dependencias

### Backend
- ✅ `TadeposiController` con endpoint `/api/tadeposi/activos`
- ✅ Entidad `Tadeposi` en `Shared/Entities/`

### Frontend
- ✅ `IRepository` para llamadas a la API
- ✅ Bootstrap para estilos del modal
- ✅ Font Awesome para iconos

## API Utilizada

El componente utiliza el endpoint existente:
```
GET /api/tadeposi/activos
```

Este endpoint retorna todos los depósitos con `Status >= 0` ordenados por `CodigoDepo`.

## Estados del Componente

1. **Cargando**: Muestra spinner mientras obtiene los datos
2. **Sin datos**: Muestra mensaje cuando no hay depósitos
3. **Con datos**: Muestra el select con los depósitos disponibles
4. **Error**: Muestra mensaje de error si falla la carga

## Personalización

### Cambiar el filtro de status
```razor
<DepositoSelectorModal StatusFilter="1" ... />
```

### Establecer un depósito por defecto
```razor
<DepositoSelectorModal DepositoPorDefecto="123" ... />
```

### Estilos CSS personalizados
El componente utiliza clases Bootstrap estándar, puedes sobrescribir los estilos según necesites.

## Notas Técnicas

- Utiliza el patrón de memoria recomendado con `@inject IRepository repository`
- Maneja errores de red y muestra mensajes informativos
- Implementa loading states para mejor UX
- Compatible con validaciones de formularios
- Sigue las convenciones del proyecto IMPLANPROD

## Ejemplo de Integración en Formularios

```razor
<EditForm Model="miModelo" OnValidSubmit="GuardarFormulario">
    <div class="mb-3">
        <label class="form-label">Depósito:</label>
        <div class="input-group">
            <input class="form-control" value="@(depositoSeleccionado?.Descripcion ?? "")" readonly />
            <button type="button" class="btn btn-outline-secondary" @onclick="AbrirModalDeposito">
                <i class="fas fa-search"></i>
            </button>
        </div>
    </div>
    
    <button type="submit" class="btn btn-success">Guardar</button>
</EditForm>

<DepositoSelectorModal IsVisible="mostrarModalDeposito"
                       OnDepositoSeleccionado="OnDepositoSeleccionado"
                       OnModalClosed="OnModalCerrado" />
```

## Mantenimiento

Para modificar el componente:
1. Editar `DepositoSelectorModal.razor`
2. Mantener compatibilidad con la clase `DepositoSeleccionado`
3. Actualizar esta documentación si es necesario

## Soporte

Para dudas o mejoras, contactar al equipo de desarrollo de IMPLANPROD.
