# Norma de Desarrollo - Formularios Blazor (.razor)

## Objetivo

A partir de la fecha, todos los nuevos formularios desarrollados para el ERP deberán seguir esta estructura con el objetivo de mejorar la mantenibilidad, facilitar el trabajo colaborativo y reducir el tamaño de los archivos `.razor`.

Esta norma aplica únicamente a los desarrollos nuevos. Los formularios existentes podrán adecuarse gradualmente cuando se realicen modificaciones importantes.

---

# 1. Responsabilidad del archivo .razor

El archivo `.razor` debe contener principalmente:

- Marcado HTML.
- Componentes visuales.
- Enlaces (Binding) entre controles y propiedades.
- Llamadas a eventos.

Debe evitarse incluir lógica de negocio extensa dentro del archivo `.razor`.

---

# 2. Código C# separado

Cada formulario deberá poseer un archivo Code Behind:

```
ClienteEdit.razor
ClienteEdit.razor.cs
```

El archivo `.razor.cs` contendrá:

- OnInitializedAsync()
- OnParametersSetAsync()
- Eventos principales
- Coordinación entre componentes
- Llamadas a servicios

---

# 3. Utilizar clases parciales

Cuando un formulario crezca, dividir el código utilizando clases parciales.

Ejemplo:

```
ClienteEdit.razor
ClienteEdit.razor.cs
ClienteEdit.Eventos.cs
ClienteEdit.Validaciones.cs
ClienteEdit.Modales.cs
ClienteEdit.Calculos.cs
```

Cada archivo debe contener únicamente responsabilidades relacionadas.

---

# 4. Modales

Cada ventana modal deberá implementarse como un componente independiente.

Incorrecto:

- Cinco modales definidos dentro del mismo formulario.

Correcto:

```
Components/

ModalBuscarArticulo.razor
ModalBuscarProveedor.razor
ModalHistorial.razor
ModalImpresion.razor
```

El formulario únicamente deberá invocar dichos componentes.

---

# 5. Pestañas (Tabs)

Cada pestaña importante deberá convertirse en un componente.

Ejemplo:

```
ArticuloGeneral.razor
ArticuloStock.razor
ArticuloCompras.razor
ArticuloVentas.razor
```

---

# 6. Componentes reutilizables

Toda sección reutilizable deberá implementarse como componente.

Ejemplos:

- Datos del cliente
- Datos del proveedor
- Dirección
- Domicilio de entrega
- Datos fiscales

Evitar duplicar código.

---

# 7. Servicios

Toda lógica de negocio deberá implementarse en servicios.

Ejemplos:

```
ArticuloService
StockService
PrecioService
RecepcionService
OrdenFabricacionService
```

No realizar consultas SQL ni lógica de negocio dentro del formulario.

---

# 8. Validaciones

Las validaciones extensas deberán ubicarse en archivos específicos o clases auxiliares.

Ejemplo:

```
ArticuloValidator
RecepcionValidator
```

---

# 9. ViewModels

Evitar declarar gran cantidad de variables individuales.

Preferir:

```
ArticuloViewModel
RecepcionViewModel
ClienteViewModel
```

---

# 10. Organización recomendada

```
Pages/

ArticuloEdit.razor
ArticuloEdit.razor.cs
ArticuloEdit.Eventos.cs
ArticuloEdit.Validaciones.cs
ArticuloEdit.Modales.cs

Components/

ArticuloGeneral.razor
ArticuloStock.razor
ModalArticulo.razor
ModalProveedor.razor

Services/

ArticuloService.cs
StockService.cs

ViewModels/

ArticuloViewModel.cs

Helpers/

ArticuloValidator.cs
```

---

# 11. Responsabilidad única

Cada archivo debe tener una única responsabilidad.

Evitar archivos de miles de líneas mezclando:

- HTML
- SQL
- Validaciones
- Eventos
- Cálculos
- Modales
- Exportaciones

---

# 12. Tamaño recomendado

Como referencia:

- `.razor`: hasta 300 líneas.
- `.razor.cs`: preferentemente menos de 500 líneas.
- Si una clase parcial supera aproximadamente las 500 líneas, evaluar dividirla nuevamente.

Estos valores son orientativos y pueden superarse cuando exista una justificación técnica.

---

# 13. Compatibilidad

Estas normas no requieren modificar formularios existentes.

Se aplicarán obligatoriamente a los nuevos desarrollos y a las refactorizaciones importantes.

---

# 14. Norma para asistentes de IA

Toda IA que genere código para este proyecto deberá respetar esta estructura.

Al crear un nuevo formulario:

- No concentrar toda la lógica en el archivo `.razor`.
- Utilizar Code Behind (`.razor.cs`).
- Crear componentes reutilizables cuando corresponda.
- Separar los modales en componentes independientes.
- Utilizar clases parciales para organizar el código.
- Delegar la lógica de negocio a servicios.
- Mantener una clara separación entre presentación, lógica de negocio y acceso a datos.

El incumplimiento de esta estructura deberá considerarse una excepción y deberá estar debidamente justificado.