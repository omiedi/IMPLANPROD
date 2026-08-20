markdown# Guía del Componente: InputNumeroLimpio.razor

Este documento sirve como memoria persistente y definición de reglas para que Devin sepa cómo utilizar el componente reutilizable personalizado `InputNumeroLimpio.razor` dentro de este proyecto Blazor / .NET.

## Descripción General
`InputNumeroLimpio` es un componente de entrada numérico genérico y reutilizable diseñado para reemplazar al `InputNumber` estándar de Blazor. Su objetivo principal es eliminar por completo las flechas nativas de incremento/decremento (spin buttons) del navegador mediante estilos CSS, manteniendo intactas todas las capacidades nativas de binding de datos de Blazor.

---

## Especificaciones Técnicas y Parámetros

El componente soporta tipos genéricos (`@typeparam TValue`) y acepta los siguientes parámetros:

| Parámetro | Tipo | Descripción |
| :--- | :--- | :--- |
| `Value` | `TValue?` | El valor actual vinculado al campo. |
| `ValueChanged` | `EventCallback<TValue?>` | Evento que se dispara cuando el valor del input cambia. |
| `ValueExpression` | `Expression<Func<TValue?>>?` | Metadatos de validación para los formularios EditContext de Blazor. |
| `Class` | `string?` | Clases CSS personalizadas aplicadas directamente al elemento input interno. |
| `Placeholder` | `string?` | Texto fantasma que se muestra cuando el input está vacío. |
| `AdditionalAttributes` | `Dictionary<string, object>?` | Captura atributos no declarados (`CaptureUnmatchedValues`) y los reenvía al input interno. Permite usar `style`, `disabled`, `min`, `max`, `step`, `title`, etc. |

---

## Instrucciones de Uso

### 1. Vinculación de Datos Estándar (Data Binding)
Cada vez que necesites un input numérico **sin flechas de incremento** dentro o fuera de un formulario, usa este componente en lugar de `<InputNumber>`. Soporta la sintaxis estándar `@bind-Value`.

```razor
<InputNumeroLimpio @bind-Value="miModelo.Edad" 
                   Class="form-control" 
                   Placeholder="Ingresa la edad" />
```

### 2. Tipos de Datos Soportados
Al ser un componente genérico, Devin debe utilizarlo para cualquiera de los siguientes tipos de datos:
* `int` / `int?`
* `double` / `double?`
* `decimal` / `decimal?`
* `float` / `float?`
* `long` / `long?`

---

## Dependencias de la Arquitectura

Devin debe asegurarse de que estas dos piezas permanezcan intactas en el código base:

1. **El Archivo del Componente (`InputNumeroLimpio.razor`):**
   ```razor
   @using System.Linq.Expressions
   @typeparam TValue

   <div class="sin-flechas">
       <InputNumber TValue="TValue" 
                    Value="@Value" 
                    ValueChanged="@ValueChanged" 
                    ValueExpression="@ValueExpression" 
                    class="@Class" 
                    placeholder="@Placeholder"
                    @attributes="AdditionalAttributes" />
   </div>

   @code {
       [Parameter] public TValue? Value { get; set; }
       [Parameter] public EventCallback<TValue?> ValueChanged { get; set; }
       [Parameter] public Expression<Func<TValue?>>? ValueExpression { get; set; }
       [Parameter] public string? Class { get; set; }
       [Parameter] public string? Placeholder { get; set; }

       [Parameter(CaptureUnmatchedValues = true)]
       public Dictionary<string, object>? AdditionalAttributes { get; set; }
   }
   ```

   > El `@attributes="AdditionalAttributes"` en el `InputNumber` interno es obligatorio: sin él, pasar `style` o `disabled` al componente provoca un error de compilación de Razor.

2. **Reglas CSS Globales (Deben existir en app.css o site.css):**
   ```css
   .sin-flechas input[type=number]::-webkit-inner-spin-button,
   .sin-flechas input[type=number]::-webkit-outer-spin-button {
       -webkit-appearance: none;
       margin: 0;
   }
   .sin-flechas input[type=number] {
       -moz-appearance: textfield;
       appearance: textfield;
   }
   ```

## Reglas Obligatorias para Devin
* **NO utilices** la etiqueta nativa HTML `<input type="number">` ni el `<InputNumber>` estándar si el diseño de la interfaz requiere ocultar las flechas.
* **NO dupliques** el código CSS en otros archivos; depende siempre de la clase utilitaria `.sin-flechas` ya existente.
* Cuando realices refactorizaciones en formularios antiguos, verifica si existen inputs numéricos genéricos que puedan migrarse a `<InputNumeroLimpio>`.
