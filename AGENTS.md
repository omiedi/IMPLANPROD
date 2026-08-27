# AGENTS.md — Convenciones del Proyecto IMPLANPROD

Este documento define las convenciones de código y arquitectura que todo colaborador del proyecto IMPLANPROD debe seguir.

---

## 1. Arquitectura General

- Solución Blazor WASM + .NET API REST.
- Backend en `IMPLANPROD.Server`.
- Cliente en `IMPLANPROD.Client`.
- Modelos compartidos en `IMPLANPROD.Shared`.

---

## 2. Entidades y Auditoría

- Toda entidad nueva debe implementar `ITimestampEntity` e `IAuditableEntity`.
- Campos de auditoría obligatorios:
  - `AddRecord` (fecha de alta)
  - `LastUpdate` (fecha de última modificación)
  - `CodiEmprNet` (int, siempre 0)
  - `NumUsuar` (int, usuario activo, mapea a `usuarios.idflexoft`)
- `CodiEmpr` siempre se maneja como `int` con valor `0` por defecto.
- Al final de cada entidad comentar el SQL original y los cambios aplicados.

---

## 3. Conversiones de Tipos

| Origen     | Destino  | Condición                    |
|------------|----------|------------------------------|
| `float`    | `decimal`| SIEMPRE                      |
| `short`    | `int`    | Solo si el usuario lo pide   |
| `CodiEmpr` | `int`    | SIEMPRE, default 0           |
| `NumUsuar` | `int`    | SIEMPRE, usuario activo      |

---

## 4. Estructura de ABM/CRUD en Blazor

Cuando se pide crear un ABM de una tabla en IMPLANPROD, usar la estructura mínima de páginas:

1. **`TablaIndex.razor`**: listado paginado con filtros e integración de impresión/exportación.
2. **`TablaForm.razor`**: componente/formulario reutilizable para alta y edición (no guarda directamente).
3. **`TablaCreate.razor`**: página de alta que usa `TablaForm` y contiene la lógica de guardado.
4. **`TablaEdit.razor`**: página de edición que carga el registro y usa `TablaForm`.
5. **`TablaBaja.razor`**: cuando aplica, realiza baja lógica cambiando `Status = -1` (no borrado físico).

### Reglas

- No asumir obligatoriamente `Interface`, `Service`, `MappingProfile` o `AutoMapper`; si el ABM se resuelve básico con consultas directas en el controller, es suficiente.
- Usar `GET` con `AsNoTracking` para evitar bloqueos.

---

## 5. Impresión y Exportación a Excel

### Cliente

- Botón: `@onclick="MostrarImpresion"` deshabilitado si no hay datos.
- Usar componente `ImpresionLista` con `UrlExportarExcel`.
- `MostrarImpresion()` obtiene todos los registros filtrados (sin paginación) y abre modal.

### Servidor

- Crear endpoint `[HttpGet("todos")]` por controlador para listados sin paginación.
- En `ExcelController`, crear endpoint `POST` específico para cada pantalla.
- Usar `_excelService.GenerarExcel(...)`.

---

## 6. Instructivos en Pantallas Blazor

### Instructivo para el Usuario

- Card Bootstrap con `bg-info text-white p-2`.
- Botón `btn btn-link` que teclea `mostrarInstructivoUsuario` con ícono chevron.
- Secciones: Descripción, Acciones, Campos importantes, Consejos.

### Instructivo para el Desarrollador

- Usar componente `Client/Shared/Components/InstructivoDesarrollador.razor`.
- Requiere `@using IMPLANPROD.Client.Shared.Components`.
- Incluir: endpoints, servicios involucrados, diagrama de flujo y validaciones.

---

## 7. Componentes de UI

- Inputs numéricos sin flechas: usar `InputNumeroLimpio` y clase `.sin-flechas`.
- Instructivos: componente `InstructivoDesarrollador`.

---

## 8. Seguridad y Derechos

- Los derechos funcionales internos por página se mantienen en `FLEXCRL` + `usuarios.idflexoft`.
- `IdFlexoft` se usa para el match con `DER_ACCE.NUM_USER` (compatibilidad VB6).
- Sin SuperAdmin activo, el login bloquea y fuerza el flujo de bootstrap.

---

## 9. Procesos de Prueba

- No iniciar `dotnet run` en segundo plano para validaciones.
- Preferir `dotnet build` y, si se requiere levantar servidor, pedir confirmación.

---

## 11. Documentación de Reglas de Negocio

- Toda modificación o nueva funcionalidad debe incluir comentarios claros que expliquen las reglas de negocio, el circuito y las decisiones del código.
- El objetivo es que un desarrollador junior (y una IA) pueda seguir el circuito sin tener que adivinar las reglas de negocio.
- Usar `<summary>` en métodos, comentarios en bloques críticos del markup y notas en puntos donde se permite comportamientos no obvios (por ejemplo, transferir sin stock).

---

## 10. Workflows Disponibles

Los siguientes procedimientos están documentados en `.devin/workflows/`:

- `crud-basico.md`
- `crud-avanzado.md`
- `master-index.md`
- `validacion-derechos-usuario.md`
- `lista-patrones.md`
