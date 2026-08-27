# Patron de Validacion de Acceso a Paginas con PageAccessService

## Descripcion

Este patron documenta la forma estandar de validar derechos de acceso a una pagina Blazor antes de mostrar el contenido. Se basa en el ejemplo de `Client/Pages/OrdenFabricacion/OrdenFabricacionIndex.razor`.

El servicio `PageAccessService` verifica que el usuario tenga permiso para la ruta indicada. Si no tiene permiso, redirige a `/acceso-denegado` automaticamente.

## Uso Minimo Requerido

### 1. Inyectar el servicio

Agregar la directiva `@using IMPLANPROD.Client.Services` si no esta presente y el `@inject` correspondiente:

```razor
@using IMPLANPROD.Client.Services

@inject PageAccessService PageAccess
```

### 2. Agregar el flag de validacion

En la seccion `@code` declarar la variable `validando` e inicializarla en `true`:

```csharp
@code {
    private bool validando = true;

    protected override async Task OnInitializedAsync()
    {
        // Paso 1: Validar que el usuario tenga permiso para acceder a esta pagina
        // Si no tiene permiso, el servicio PageAccess redirigira a la pagina de acceso denegado
        await PageAccess.RequiereAccesoAsync("/ruta-de-la-pagina");
        validando = false;

        // ... resto de la carga de la pagina
    }
}
```

### 3. Encapsular el contenido en un condicional

Envolver todo el markup de la pagina dentro de un bloque `else` que se muestra solo cuando termina la validacion:

```razor
@if (validando)
{
    <p>Validando acceso...</p>
}
else
{
    <div class="container-fluid">
        <!-- Contenido de la pagina -->
    </div>
}
```

## Ejemplo Completo

Ejemplo extraido de `Client/Pages/OrdenFabricacion/OrdenFabricacionIndex.razor`:

```razor
@page "/ordenfabricacion"
@using IMPLANPROD.Client.Services
@inject PageAccessService PageAccess

<PageTitle>Ordenes de Fabricacion</PageTitle>

@if (validando)
{
    <p>Validando acceso...</p>
}
else
{
    <div class="container-fluid">
        <!-- Aqui sigue el contenido de la pagina -->
    </div>
}

@code {
    private bool validando = true;

    protected override async Task OnInitializedAsync()
    {
        // Paso 1: Validar que el usuario tenga permiso para acceder a esta pagina
        // Si no tiene permiso, el servicio PageAccess redirigira a la pagina de acceso denegado
        await PageAccess.RequiereAccesoAsync("/ordenfabricacion");
        validando = false;
    }
}
```

## Notas

- La ruta que se pasa a `RequiereAccesoAsync` debe coincidir con el `@page` de la pagina, o con la ruta configurada en el RBAC.
- `PageAccessService` usa prefix match: si el usuario tiene permiso para `/pedidos-venta`, tambien tendra acceso a `/pedidos-venta/crear` y `/pedidos-venta/editar/123`.
- Los roles `SUPERADMIN`, `ADMINISTRADOR` y `ADMIN` pasan la validacion sin consultar el RBAC.
- Este patron se aplica en cualquier pagina nueva que requiera control de acceso.
