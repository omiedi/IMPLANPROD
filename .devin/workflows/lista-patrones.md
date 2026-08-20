---
description: Referencia rápida de todos los patrones de desarrollo disponibles
---

# 📋 Patrones de Desarrollo Disponibles

Referencia completa de patrones automatizados para acelerar el desarrollo en IMPLANPROD.

---

## 🔧 Patrones CRUD

### 1. PATRON BASICO - CRUD

**Comando de activación:**
```
Crea un CRUD de la tabla [NombreTabla] usando el patrón BASICO
```

**Ejemplo:**
```
Crea un CRUD de la tabla PRODUCTOS usando el patrón BASICO
```

**¿Qué genera?**
- ✅ Entity con `ITimestampEntity`, `IAuditableEntity`
- ✅ DTO (sin campos de auditoría)
- ✅ Interface del Service (`I[NombreTabla]Service`)
- ✅ Service con AutoMapper
- ✅ Controller API REST (5 endpoints)
- ✅ MappingProfile (AutoMapper)
- ✅ Actualiza `DbContext` y `Program.cs`
- ✅ UI Blazor con MudDataGrid (opcional)
- ✅ Entrada en `NavMenu`

**Conversiones de tipos:**
| Tipo Original | Tipo Final | Modo |
|---------------|------------|------|
| `float` | `decimal` | Automático |
| `short` | `int` | On demand |
| `Codiempr` | `int` (=0) | Automático |

**Campos agregados automáticamente:**
- `AddRecord` (DateTime?) - Fecha de creación
- `LastUpdate` (DateTime?) - Última actualización
- `CodiEmprNet` (int?) - Código de empresa
- `NumUsuar` (int?) - Usuario activo
- `Id` (int, identity) - Si no existe

**Documentación completa:**
- Workflow: `.windsurf/workflows/crud-basico.md`

---

### 2. PATRON AVANZADO - CRUD

**Comando de activación:**
```
Crea un CRUD de la tabla [NombreTabla] usando el patrón AVANZADO
```

**¿Qué incluye?**

Todo del **PATRON BASICO** más:

**Validaciones:**
- ✅ FluentValidation en DTOs
- ✅ Validaciones de negocio en Service
- ✅ Manejo de errores con try-catch
- ✅ Mensajes descriptivos de error

**Logging:**
- ✅ ILogger inyectado en Service
- ✅ LogInformation para operaciones exitosas
- ✅ LogWarning para situaciones anormales
- ✅ LogError para excepciones

**Permisos (on demand):**
- ✅ Atributo `[Authorize]` en Controller
- ✅ Validación de `DER_ACCE`
- ✅ Verificación `NumUsuar` vs `usuarios.idflexoft`

**Códigos HTTP:**
- 200 OK, 201 Created, 204 NoContent
- 400 BadRequest, 404 NotFound, 500 InternalServerError

**Documentación completa:**
- Workflow: `.windsurf/workflows/crud-avanzado.md`

---

## 🔄 Patrones de Sincronización

### 3. PATRON SINCRONIZACION - Central/Sucursal

**Comando de activación:**
```
Implementa sincronización para la tabla [NombreTabla]
```

**¿Qué genera?**
- ✅ Sistema de sincronización bidireccional
- ✅ Manejo de conflictos (última modificación gana)
- ✅ Tracking de cambios
- ✅ Endpoints de sincronización
- ✅ Campos: `SyncStatus`, `LastSyncDate`, `OriginBranch`
- ✅ Implementa `ISyncableEntity`

**Casos de uso:**
- Sincronizar datos entre Central y Sucursales
- Replicación de maestros (productos, clientes, etc.)
- Sincronización de fórmulas y operaciones

**Documentación completa:**
- Workflow: `.windsurf/workflows/master-index.md`

---

## 📊 Patrones Futuros (Próximamente)

### 4. PATRON REPORTES
Generación de reportes con filtros, exportación Excel/PDF

### 5. PATRON IMPORTACION/EXPORTACION
Importar/exportar datos desde Excel, CSV

### 6. PATRON MAESTRO-DETALLE
Tablas relacionadas con operaciones en cascada

### 7. PATRON AUDITORIA COMPLETA
Tabla de auditoría separada con historial de cambios

---

## 🎯 Comandos Útiles

### Listar Patrones
```
Lista los patrones
Muestra los patrones disponibles
Qué patrones tengo
```

### Ver Workflow Completo
```
Abre el workflow crud-basico
Muestra el workflow master-index
```

### Crear Nuevo Patrón
```
Crea un patrón para [funcionalidad]
```

### Listar Workflows
```
Lista los workflows disponibles
```

---

## 📚 Estructura de Archivos Generados

### Patrón BASICO/AVANZADO

```
IMPLANPROD/
├── Shared/
│   ├── Entities/
│   │   └── [NombreTabla].cs          ← Entity con auditoría
│   ├── DTOs/
│   │   └── [NombreTabla]Dto.cs       ← DTO para UI
│   ├── Interfaces/
│   │   └── I[NombreTabla]Service.cs  ← Interface del servicio
│   └── Validators/                    ← Solo AVANZADO
│       └── [NombreTabla]DtoValidator.cs
├── Server/
│   ├── Services/
│   │   └── [NombreTabla]Service.cs   ← Lógica de negocio
│   ├── Controllers/
│   │   └── [NombreTabla]Controller.cs ← API REST
│   ├── Mappings/
│   │   └── [NombreTabla]MappingProfile.cs ← AutoMapper
│   ├── Data/
│   │   └── ApplicationDbContext.cs   ← DbSet agregado
│   └── Program.cs                     ← Servicio registrado
└── Client/
    ├── Pages/
    │   └── [NombreTabla]/
    │       └── [NombreTabla]List.razor ← UI Blazor
    └── Shared/
        └── NavMenu.razor              ← Entrada de menú
```

---

## 🔍 Cómo Funciona el Sistema de Patrones

1. **Tú escribes el comando** (ej: "Crea un CRUD de la tabla PRODUCTOS usando el patrón BASICO")
2. **Yo solicito información paso a paso:**
   - Estructura SQL original
   - Campos para el DTO
   - Campo Status (si/no)
   - Conversiones de tipo
3. **Genero todos los archivos automáticamente**
4. **Actualizo configuraciones** (DbContext, Program.cs, NavMenu)
5. **Documentación incluida** (comentarios detallados, SQL original y modificaciones)

---

## 💡 Tips y Buenas Prácticas

### Cuándo usar PATRON BASICO
- ✅ Tablas simples de maestros
- ✅ Catálogos básicos
- ✅ Desarrollo rápido sin validaciones complejas

### Cuándo usar PATRON AVANZADO
- ✅ Tablas con reglas de negocio complejas
- ✅ Operaciones críticas que requieren logging
- ✅ Módulos con control de permisos
- ✅ Validaciones de datos estrictas

### Cuándo usar PATRON SINCRONIZACION
- ✅ Datos que deben replicarse entre Central y Sucursales
- ✅ Maestros compartidos (productos, clientes, fórmulas)
- ✅ Operaciones que requieren sincronización bidireccional

---

## 📞 Soporte

Si necesitas:
- **Crear un nuevo patrón personalizado**
- **Modificar un patrón existente**
- **Ayuda con un patrón específico**

Solo pregunta y te guiaré paso a paso.

---

## 🚀 Inicio Rápido

**Primer CRUD:**
```
Crea un CRUD de la tabla CLIENTES usando el patrón BASICO
```

**Ver todos los patrones:**
```
Lista los patrones
```

**Ver workflow detallado:**
```
Abre el workflow crud-basico
```
