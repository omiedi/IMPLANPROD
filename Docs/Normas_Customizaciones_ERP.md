# Arquitectura de Customizaciones - IMPLANPROD ERP

## Objetivo

A partir de este momento queda prohibido implementar personalizaciones mediante validaciones directas del tipo:

```csharp
if (Empresa == "PEPE")
```

o

```csharp
if (CodiEmpr == 101)
```

o cualquier otra comparación que identifique directamente una empresa.

Toda personalización deberá realizarse utilizando el sistema de funcionalidades (Features), configuraciones y estrategias (Strategy Pattern).

---

# Arquitectura

Toda diferencia entre empresas deberá clasificarse en una de estas categorías.

## 1) Configuración

Cuando únicamente cambia un valor.

Ejemplos:

- Días de tolerancia
- Cantidad de decimales
- Color
- Depósito por defecto
- Permitir stock negativo
- Usa lotes
- Imprimir automáticamente
- etc.

Estas diferencias deberán almacenarse en la tabla:

EmpresaConfiguracion

Nunca deberán codificarse dentro del sistema.

---

## 2) Features

Cuando una funcionalidad existe para algunas empresas y para otras no.

Ejemplos

Recepción parcial

WhatsApp

Múltiples depósitos

Control de calidad

Planificación avanzada

etc.

Estas funcionalidades deberán almacenarse mediante Feature Flags.

El código consultará únicamente:

FeatureManager.Tiene("RecepcionParcial")

Nunca:

if (Empresa == ...)

---

## 3) Strategy Pattern

Cuando cambia completamente una lógica de negocio.

Ejemplos

Cálculo de stock

Generación de costos

Aprobación de recepción

Facturación

Producción

Para estos casos deberán implementarse interfaces.

Ejemplo

IStockService

StockServiceDefault

StockServiceEmpresaXXX

El sistema elegirá automáticamente la implementación correspondiente.

---

## 4) Eventos

Cuando una empresa necesita ejecutar acciones adicionales.

Ejemplo

Recepción aprobada

↓

Enviar WhatsApp

↓

Crear PDF

↓

Actualizar otro sistema

↓

Enviar Mail

No deberán agregarse IF dentro del proceso principal.

Se implementarán eventos.

---

# Orden de prioridad

Antes de programar cualquier personalización deberá evaluarse:

1 Configuración

2 Feature

3 Strategy

4 Evento

Solo si ninguna alternativa es suficiente podrá desarrollarse una solución específica.

---

# Queda prohibido

if (Empresa == ...)

if (CodiEmpr == ...)

switch(Empresa)

switch(CodiEmpr)

código duplicado para clientes

---

# Objetivo

El ERP deberá poder soportar cientos de empresas sin que el núcleo del sistema deba modificarse por cada cliente.