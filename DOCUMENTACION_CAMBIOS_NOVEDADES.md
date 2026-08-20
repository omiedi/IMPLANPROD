# 📋 Documentación de Cambios: Sistema de Registro de Novedades de Precios

**Fecha:** 15 de Abril de 2026  
**Módulo:** Órdenes de Compra - Costos  
**Tablas Afectadas:** NOVEDADES, OCOMDETA_COSTO  

---

## 📌 Resumen Ejecutivo

Se implementó y corrigió el sistema de registro automático de cambios de precios en la tabla `NOVEDADES` cuando se editan costos de órdenes de compra. El sistema detecta cambios significativos de precio (> 0.5%) y los registra para auditoría y seguimiento histórico.

---

## 🎯 Objetivo del Sistema

Registrar automáticamente en la tabla `NOVEDADES` los cambios de precios de productos cuando:
1. Se edita el costo de un producto en una orden de compra
2. La diferencia de precio es mayor al 0.5%
3. Incluso si el precio en moneda original no cambió pero la cotización sí

---

## 🔧 Cambios Realizados

### 1. **Entidad Novedades** (`Shared/Entities/Novedades.cs`)

#### Cambio de Tipo de Dato: `Codint`
- **Antes:** `short?` (smallint en SQL)
- **Después:** `int?` (int en SQL)
- **Razón:** Soportar códigos internos más grandes

#### Campos Clave Documentados:

```csharp
/// <summary>
/// Precio anterior (SIEMPRE EN PESOS)
/// Almacena el precio previo antes de la modificación convertido a pesos
/// Si es la primera vez, se guarda igual a PrecNew
/// Ejemplo: Si era USD 100 con cotiz 1050 = $105,000 pesos
/// </summary>
public decimal? PrecAnt { get; set; }

/// <summary>
/// Precio nuevo (SIEMPRE EN PESOS)
/// Almacena el nuevo precio después de la modificación convertido a pesos
/// Se calcula multiplicando el precio en moneda original por la cotización
/// Ejemplo: USD 110 * cotiz 1050 = $115,500 pesos
/// </summary>
public decimal? PrecNew { get; set; }

/// <summary>
/// Cotización de la moneda al momento del cambio
/// IMPORTANTE: Permite detectar cambios de precio por variación de cotización
/// Ejemplo: Si el precio en USD no cambió pero la cotización sí,
///          el precio en pesos será diferente y debe registrarse
/// Valores típicos: 1 (pesos), 1050 (USD), 1150 (EUR)
/// </summary>
public decimal? CotizaDolar { get; set; }
```

#### Script SQL de Migración:

```sql
-- Modificar columna Codint de smallint a int
ALTER TABLE [dbo].[NOVEDADES]
ALTER COLUMN [Codint] [int] NULL
```

---

### 2. **Servicio OCOMDETACostoService** (`Server/Services/OCOMDETACostoService.cs`)

#### Método Principal: `RegistrarNovedadAsync`

**Lógica de Negocio Implementada:**

1. **Validación de Código Interno**
   - Verifica que `CodInte > 0`
   - Si es inválido, no registra

2. **Obtención de Cotización Actual**
   - Consulta tabla `TAMONED` por código de moneda
   - Si es moneda 1 (pesos), retorna 1
   - Si no encuentra, retorna 1 por defecto

3. **Búsqueda de Precio Anterior**
   - Busca en `NOVEDADES` el último registro del producto **por código interno**
   - Obtiene `PrecNew` (ya en pesos) como precio anterior
   - Obtiene `CotizaDolar` anterior para detectar cambios de cotización

4. **Conversión a Pesos**
   - Precio nuevo: `PrUnNeto * CotizacionActual`
   - Precio anterior: Ya está en pesos (de NOVEDADES)

5. **Cálculo de Diferencia**
   - Fórmula: `|PrecioNuevo - PrecioAnterior| / PrecioAnterior * 100`
   - Umbral: **0.5%**

6. **Decisión de Registro**
   - **Primera vez** (sin precio anterior): Registra con `PrecAnt = PrecNew`
   - **Diferencia > 0.5%**: Registra el cambio
   - **Diferencia <= 0.5%**: NO registra (cambio insignificante)

#### Código Documentado:

```csharp
/// <summary>
/// Registra un cambio de precio en la tabla NOVEDADES
/// 
/// LÓGICA DE NEGOCIO:
/// 1. Busca el precio anterior en NOVEDADES (último PrecNew del producto por código)
/// 2. Obtiene la cotización anterior para detectar cambios de cotización
/// 3. Convierte ambos precios a PESOS para comparar (precio * cotización)
/// 4. Calcula diferencia porcentual entre precio anterior y nuevo (ambos en pesos)
/// 5. Solo registra si la diferencia es > 0.5%
/// 
/// CASOS ESPECIALES:
/// - Primera vez (sin precio anterior): Registra con PrecAnt = PrecNew
/// - Cambio de cotización: Aunque el precio en USD sea igual, si la cotización cambió
///   y el resultado en pesos es diferente > 0.5%, se registra
/// - Diferencia <= 0.5%: NO se registra (cambio insignificante)
/// 
/// IMPORTANTE:
/// - PrecAnt y PrecNew SIEMPRE se guardan en PESOS
/// - Se guarda la cotización actual en CotizaDolar
/// - Se guarda el número de orden de compra que originó el cambio
/// </summary>
private async Task RegistrarNovedadAsync(OCOMDETACostoDto costo, string observacion)
{
    // ... código con comentarios detallados en cada paso
}
```

#### Método Auxiliar: `ObtenerPrecioYCotizacionAnteriorAsync`

```csharp
/// <summary>
/// Obtiene el precio anterior (en pesos) y la cotización anterior de NOVEDADES
/// 
/// PROPÓSITO:
/// - Busca el último registro del producto en NOVEDADES por código interno
/// - Retorna el PrecNew (que ya está en pesos) como precio anterior
/// - Retorna también la CotizaDolar para detectar cambios de cotización
/// 
/// POR QUÉ SE NECESITA LA COTIZACIÓN ANTERIOR:
/// Si el precio en USD no cambió pero la cotización sí, el precio en pesos
/// será diferente y debe registrarse como novedad.
/// Ejemplo: USD 100 con cotiz 1050 = $105,000
///          USD 100 con cotiz 1100 = $110,000 (diferencia 4.76%)
/// </summary>
private async Task<(decimal precioAnteriorPesos, decimal cotizacionAnterior)> 
    ObtenerPrecioYCotizacionAnteriorAsync(int codInte)
{
    // Busca SOLO por código interno, no por orden ni item
    var ultimaNovedad = await _context.Novedades
        .Where(n => n.Codint == codInte)
        .OrderByDescending(n => n.AddRecord)
        .FirstOrDefaultAsync();
    
    // Retorna (PrecNew en pesos, CotizaDolar)
    // Si no existe, retorna (0, 0)
}
```

---

## 📊 Ejemplos de Funcionamiento

### Ejemplo 1: Primera Vez (Sin Registro Previo)

**Datos:**
- Producto: ABC-123
- Precio Nuevo: USD 100
- Cotización Actual: 1050
- Precio Anterior en NOVEDADES: No existe

**Cálculo:**
```
precioAnteriorPesos = 0 (no hay registro)
precioNuevoPesos = 100 * 1050 = 105,000
esPrimeraVez = true
→ precioAnteriorPesos = precioNuevoPesos = 105,000
```

**Registro en NOVEDADES:**
```
✅ SE REGISTRA
PrecAnt = 105,000 (mismo que PrecNew)
PrecNew = 105,000
CotizaDolar = 1050
```

---

### Ejemplo 2: Cambio de Precio en USD

**Datos:**
- Producto: ABC-123
- Precio Anterior en NOVEDADES: 105,000 pesos (USD 100 * 1050)
- Precio Nuevo: USD 110
- Cotización Actual: 1050

**Cálculo:**
```
precioAnteriorPesos = 105,000
cotizacionAnterior = 1050
precioNuevoPesos = 110 * 1050 = 115,500
diferencia = |115,500 - 105,000| / 105,000 * 100 = 10%
```

**Resultado:**
```
✅ SE REGISTRA (10% > 0.5%)
PrecAnt = 105,000
PrecNew = 115,500
CotizaDolar = 1050
```

---

### Ejemplo 3: Sin Cambio de Precio

**Datos:**
- Producto: ABC-123
- Precio Anterior en NOVEDADES: 105,000 pesos
- Precio Nuevo: USD 100 (mismo)
- Cotización Actual: 1050 (misma)

**Cálculo:**
```
precioAnteriorPesos = 105,000
precioNuevoPesos = 100 * 1050 = 105,000
diferencia = 0%
```

**Resultado:**
```
❌ NO SE REGISTRA (0% <= 0.5%)
```

---

### Ejemplo 4: Cambio de Cotización (Precio USD Igual) ⭐

**Datos:**
- Producto: ABC-123
- Precio Anterior en NOVEDADES: 105,000 pesos (USD 100 * 1050)
- Cotización Anterior: 1050
- Precio Nuevo: USD 100 (mismo precio en USD)
- Cotización Actual: 1100 (cotización subió)

**Cálculo:**
```
precioAnteriorPesos = 105,000
cotizacionAnterior = 1050
precioNuevoPesos = 100 * 1100 = 110,000
diferencia = |110,000 - 105,000| / 105,000 * 100 = 4.76%
```

**Resultado:**
```
✅ SE REGISTRA (4.76% > 0.5%)
PrecAnt = 105,000
PrecNew = 110,000
CotizaDolar = 1100

Aunque el precio en USD no cambió, el precio en pesos SÍ cambió por la cotización.
```

---

### Ejemplo 5: Cambio Mínimo (< 0.5%)

**Datos:**
- Producto: ABC-123
- Precio Anterior en NOVEDADES: 100,000 pesos
- Precio Nuevo: $100,400
- Cotización: 1 (pesos)

**Cálculo:**
```
precioAnteriorPesos = 100,000
precioNuevoPesos = 100,400
diferencia = |100,400 - 100,000| / 100,000 * 100 = 0.4%
```

**Resultado:**
```
❌ NO SE REGISTRA (0.4% <= 0.5%)
```

---

## 🔐 Control de Acceso

### Derecho Requerido para Ver Botón de Costos

**Código:** `VerCostosOC`  
**Descripción:** Ver y Editar Costos de Órdenes de Compra  
**Nivel Requerido:** `DER_READ >= 2` Y `DER_WRITE >= 2`

**Ubicación en Código:**
```csharp
// Client/Pages/OrdenesCompra/OrdenesCompraIndex.razor
private const string COD_APLI_COSTOS_OC = "VerCostosOC";
private const string DES_APLI_COSTOS_OC = "Ver y Editar Costos de Órdenes de Compra";

// Verificación
var nivelAccesoCostosOC = await derechosService.VerificarDerechoAsync(
    COD_APLI_COSTOS_OC,
    DES_APLI_COSTOS_OC,
    mostrarMensaje: false
);
puedeEditarCostos = nivelAccesoCostosOC >= 2;
```

**Script SQL para Otorgar Derecho:**
```sql
-- Archivo: Server/Data/Migrations/AddDerechoVerCostosOC.sql
INSERT INTO DER_ACCE (NUM_USER, COD_APLI, DER_READ, DER_WRITE)
VALUES ([ID_USUARIO], 'VerCostosOC', 2, 2)
```

---

## 📁 Archivos Modificados

| Archivo | Cambios Realizados |
|---------|-------------------|
| `Shared/Entities/Novedades.cs` | - Cambio `Codint` de `short?` a `int?`<br>- Comentarios mejorados en `PrecAnt`, `PrecNew`, `CotizaDolar`<br>- Script SQL ALTER TABLE agregado |
| `Server/Services/OCOMDETACostoService.cs` | - Método `RegistrarNovedadAsync` completamente documentado<br>- Nuevo método `ObtenerPrecioYCotizacionAnteriorAsync`<br>- Lógica corregida para comparación de precios<br>- Detección de cambios de cotización |
| `Server/Data/Migrations/AddDerechoVerCostosOC.sql` | - Script para crear derecho `VerCostosOC`<br>- Asignación automática a ADMINISTRADORES |

---

## ✅ Validaciones Implementadas

| Validación | Descripción | Acción |
|------------|-------------|--------|
| **CodInte inválido** | CodInte <= 0 | ❌ No registra |
| **Primera vez** | No hay registro previo | ✅ Registra con PrecAnt = PrecNew |
| **Diferencia <= 0.5%** | Cambio muy pequeño | ❌ No registra |
| **Diferencia > 0.5%** | Cambio significativo | ✅ Registra |
| **Cambio de cotización** | Precio USD igual, cotización diferente | ✅ Registra si diferencia en pesos > 0.5% |

---

## 🔍 Logs de Consola

El sistema genera logs detallados para debugging:

```
📊 Precio anterior encontrado para CodInte 123: 105,000.0000 pesos (cotiz=1050.0000)
✅ Novedad registrada: CodInte=123, OrdenCompra=456
   PrecAnt=105,000.0000, PrecNew=115,500.0000, Diferencia=10.00%
   Cotización Anterior=1050.0000, Cotización Actual=1050.0000
```

```
⚠️ No se registra novedad para CodInte 123: diferencia 0.40% <= 0.5%
   PrecAnterior=100,000.0000 (cotiz=1.0000), PrecNuevo=100,400.0000 (cotiz=1.0000)
```

```
ℹ️ No hay precio anterior para CodInte 789 (primera vez)
ℹ️ Primera novedad para CodInte 789: PrecAnt = PrecNew = 50,000.0000
```

---

## 🎓 Conceptos Clave para Entender el Sistema

### 1. **¿Por qué siempre en pesos?**
Para poder comparar precios de manera consistente, independientemente de la moneda original. Un producto puede comprarse en USD, EUR o pesos, pero el histórico debe ser comparable.

### 2. **¿Por qué guardar la cotización?**
Para detectar cuando el precio en moneda original no cambió, pero el precio en pesos sí (por cambio de cotización). Esto es importante para el análisis de costos.

### 3. **¿Por qué el umbral del 0.5%?**
Para evitar registrar cambios insignificantes causados por redondeos o ajustes mínimos. Solo se registran cambios relevantes.

### 4. **¿Por qué PrecAnt = PrecNew la primera vez?**
Para establecer un precio base. Si no hay precio anterior, el "anterior" es el mismo que el nuevo, indicando que es el primer registro.

---

## 🚀 Próximos Pasos (Opcional)

1. **Dashboard de Novedades:** Pantalla para visualizar el histórico de cambios
2. **Alertas:** Notificar cuando hay cambios de precio significativos (>10%)
3. **Reportes:** Análisis de evolución de precios por producto/proveedor
4. **Exportación:** Exportar histórico de novedades a Excel

---

## 📞 Contacto

Para dudas o consultas sobre este sistema:
- **Desarrollador:** [Nombre]
- **Fecha de Implementación:** 15 de Abril de 2026
- **Versión:** 1.0

---

**Fin de la Documentación**
