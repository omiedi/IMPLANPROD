# Normas de Arquitectura de Customizaciones
## IMPLANPROD ERP

# Objetivo

Toda personalización deberá desarrollarse utilizando un sistema de
configuración, funcionalidades (Feature Flags), estrategias (Strategy Pattern)
y eventos.

Queda prohibido identificar clientes mediante:

if (Empresa == ...)
if (CodiEmpr == ...)
switch(Empresa)

El código deberá ser independiente de la empresa.

---

# Arquitectura obligatoria

Cada implementación deberá incluir los siguientes componentes.

## 1. Base de Datos

Crear las siguientes tablas.

### EmpresaConfiguracion

Finalidad:

Guardar parámetros configurables por empresa.

Campos mínimos

Id
CodiEmpr
Modulo
Clave
Valor
TipoDato
Descripcion
Activo
AddRecord
LastUpdate
NumUsuar

Generar automáticamente:

- CREATE TABLE
- Primary Key
- Índices
- Default Values
- Restricciones

---

### EmpresaFeature

Finalidad

Determinar qué funcionalidades tiene habilitada cada empresa.

Campos

Id
CodiEmpr
CodigoFeature
Habilitado
Observaciones
AddRecord
LastUpdate
NumUsuar

Generar SQL Server completo.

---

### EmpresaStrategy

Finalidad

Determinar qué implementación utilizará una empresa.

Ejemplo

Stock

Producción

Recepción

Facturación

Campos

Id
CodiEmpr
Modulo
ClaseImplementacion
Activo

Generar SQL correspondiente.

---

### EmpresaEventos

Finalidad

Permitir registrar procesos adicionales.

Campos

Id

CodiEmpr

Evento

Clase

Orden

Activo

Generar SQL Server.

---

# Entidades

Crear automáticamente

EmpresaConfiguracion.cs

EmpresaFeature.cs

EmpresaStrategy.cs

EmpresaEvento.cs

Agregar DbSet correspondientes.

---

# Interfaces

Crear

IFeatureManager

IConfiguracionEmpresaService

IStrategyResolver

IEventoEmpresaService

---

# Servicios

Implementar

FeatureManager

ConfiguracionEmpresaService

StrategyResolver

EventoEmpresaService

Todos utilizando Dependency Injection.

---

# Controladores

Crear automáticamente

ConfiguracionEmpresaController

EmpresaFeatureController

EmpresaStrategyController

EmpresaEventoController

Con endpoints

GET

GET ID

POST

PUT

DELETE lógico

---

# Blazor

Crear las siguientes páginas.

## Configuración Empresa

Listado

Filtros

Alta

Edición

Baja lógica

Búsqueda

Paginación

---

## Features

Grilla

Empresa

Feature

Habilitado

Filtro

Botón activar

Botón desactivar

---

## Strategies

Empresa

Módulo

Implementación

Editar

---

## Eventos

Empresa

Evento

Clase

Orden

Activo

---

# FeatureManager

Todo el sistema consultará únicamente

FeatureManager.Tiene()

Nunca accederá directamente a la tabla.

---

# Configuración

Toda lectura de parámetros deberá realizarse mediante

ConfiguracionEmpresaService

Nunca consultar la tabla directamente.

---

# Strategy

Cuando una lógica cambie completamente deberá crearse

Una interfaz

Una implementación Default

Una implementación específica

Nunca utilizar IF.

---

# Eventos

Los procesos importantes deberán disparar eventos.

Ejemplo

Recepción aprobada

↓

Evento

↓

Implementaciones registradas

El proceso principal nunca conocerá qué empresa ejecuta cada acción.

---

# Procedimiento para nuevas customizaciones

Antes de escribir código deberá responderse:

1. ¿Se resuelve con un parámetro?
   → EmpresaConfiguracion

2. ¿Se resuelve habilitando una funcionalidad?
   → EmpresaFeature

3. ¿La lógica cambia completamente?
   → Strategy

4. ¿Debe ejecutarse una acción adicional?
   → Evento

Solo si ninguna respuesta es válida podrá proponerse una nueva arquitectura.

---

# Objetivo

El ERP deberá poder incorporar nuevas empresas sin modificar el núcleo del sistema.