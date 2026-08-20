---
description: Patrón BASICO para crear CRUDs completos con validaciones y estructura estándar
---

# PATRON BASICO - CRUD Completo

Sistema de generación automática de CRUDs siguiendo las convenciones del proyecto IMPLANPROD.

---

## > Instructivo para el Usuario

<details>
<summary><strong>1. Cómo Solicitar un CRUD</strong></summary>

### Comando Básico
```
Crea un CRUD de la tabla [NombreTabla] usando el patrón BASICO
```

### Ejemplo
```
Crea un CRUD de la tabla PRODUCTOS usando el patrón BASICO
```

### Información que Debes Proporcionar

El sistema te solicitará paso a paso:

1. **Estructura SQL original de la tabla**
   - Copia el CREATE TABLE completo del ERP original

2. **Campos para el DTO**
   - Lista de campos que necesitas en la interfaz de usuario
   - Ejemplo: "Id, Nombre, Precio, Stock"

3. **Campo Status**
   - ¿La tabla tiene campo Status para activo/inactivo?
   - Si no, ¿deseas agregarlo?

4. **Conversiones de tipo**
   - ¿Hay campos tipo `short` que deban convertirse a `int`?

</details>

<details>
<summary><strong>2. Qué se Genera Automáticamente</strong></summary>

### Archivos Creados

1. **Entity** (Modelo de datos)
   - `IMPLANPROD.Shared/Entities/[NombreTabla].cs`

2. **DTO** (Objeto de transferencia)
   - `IMPLANPROD.Shared/DTOs/[NombreTabla]Dto.cs`

3. **Interface del Servicio**
   - `IMPLANPROD.Shared/Interfaces/I[NombreTabla]Service.cs`

4. **Servicio**
   - `IMPLANPROD.Server/Services/[NombreTabla]Service.cs`

5. **Mapping Profile**
   - `IMPLANPROD.Server/Mappings/[NombreTabla]MappingProfile.cs`

6. **Controller API**
   - `IMPLANPROD.Server/Controllers/[NombreTabla]Controller.cs`

7. **Componente Blazor** (opcional)
   - `IMPLANPROD.Client/Pages/[NombreTabla]/[NombreTabla]List.razor`

### Modificaciones Automáticas

1. **DbContext actualizado**
   - Se agrega el DbSet en `ApplicationDbContext.cs`

2. **Servicios registrados**
   - Se registra en `Program.cs`

3. **Menú actualizado**
   - Se agrega entrada en `NavMenu.razor`

</details>

<details>
<summary><strong>3. Conversiones Automáticas</strong></summary>

### Conversiones de Tipos

El sistema aplica las siguientes conversiones:

| Tipo Original | Tipo Final | Modo | Razón |
|---------------|------------|------|-------|
| `float` | `decimal` | **AUTOMÁTICO** | Mayor precisión para datos financieros |
| `short` | `int` | **ON DEMAND** | Solo si lo solicitas explícitamente |
| `Codiempr` | `int` (=0) | **AUTOMÁTICO** | Estándar del proyecto |

### Campos Agregados Automáticamente

**Auditoría** (si no existen):
- `AddRecord` (DateTime?) - Fecha de creación
- `LastUpdate` (DateTime?) - Fecha de última actualización
- `CodiEmprNet` (int?) - Código de empresa
- `NumUsuar` (int?) - Usuario que realiza la transacción

**Identificador** (si no existe):
- `Id` (int, identity) - Clave primaria

**Status** (si lo solicitas):
- `Status` (int) - Activo (>=0) o Inactivo (-1)

</details>

<details>
<summary><strong>4. Endpoints API Generados</strong></summary>

### Rutas Disponibles

Todos los endpoints se crean en `/api/[NombreTabla]`

| Método | Ruta | Descripción |
|--------|------|-------------|
| GET | `/api/[NombreTabla]` | Obtener todos los registros |
| GET | `/api/[NombreTabla]/{id}` | Obtener un registro por ID |
| POST | `/api/[NombreTabla]` | Crear nuevo registro |
| PUT | `/api/[NombreTabla]/{id}` | Actualizar registro existente |
| DELETE | `/api/[NombreTabla]/{id}` | Eliminar registro |

### Ejemplo de Uso

```javascript
// Obtener todos
fetch('/api/productos')

// Obtener uno
fetch('/api/productos/5')

// Crear
fetch('/api/productos', {
  method: 'POST',
  body: JSON.stringify({ nombre: 'Producto 1', precio: 100 })
})
```

</details>

<details>
<summary><strong>5. Verificación Post-Generación</strong></summary>

### Checklist de Verificación

✅ **Archivos creados**
- [ ] Entity en `/Shared/Entities`
- [ ] DTO en `/Shared/DTOs`
- [ ] Interface en `/Shared/Interfaces`
- [ ] Service en `/Server/Services`
- [ ] Controller en `/Server/Controllers`
- [ ] MappingProfile en `/Server/Mappings`

✅ **Modificaciones realizadas**
- [ ] DbSet agregado en `ApplicationDbContext`
- [ ] Servicio registrado en `Program.cs`
- [ ] Entrada en `NavMenu.razor`

✅ **Compilación**
- [ ] El proyecto compila sin errores
- [ ] No hay warnings de tipos incompatibles

✅ **Pruebas básicas**
- [ ] GET /api/[tabla] retorna 200
- [ ] POST crea registros correctamente
- [ ] PUT actualiza registros
- [ ] DELETE elimina registros

</details>

---

## > Instructivo para el Desarrollador

<details>
<summary><strong>1. Arquitectura del Patrón</strong></summary>

### Capas del Sistema

```
┌─────────────────────────────────────┐
│   BLAZOR CLIENT (UI)                │
│   - Pages/[Tabla]List.razor         │
│   - Consume API REST                │
└──────────────┬──────────────────────┘
               │ HTTP
┌──────────────▼──────────────────────┐
│   API CONTROLLER                    │
│   - Controllers/[Tabla]Controller   │
│   - Endpoints REST                  │
└──────────────┬──────────────────────┘
               │ Interface
┌──────────────▼──────────────────────┐
│   SERVICE LAYER                     │
│   - Services/[Tabla]Service         │
│   - Lógica de negocio               │
└──────────────┬──────────────────────┘
               │ DbContext + AutoMapper
┌──────────────▼──────────────────────┐
│   DATA LAYER                        │
│   - Entities/[Tabla]                │
│   - ApplicationDbContext            │
│   - SQL Server                      │
└─────────────────────────────────────┘
```

### Componentes Principales

1. **Entity**: Modelo de datos (mapea tabla SQL)
2. **DTO**: Objeto de transferencia (UI ↔ API)
3. **Service**: Lógica de negocio
4. **Controller**: Endpoints REST
5. **AutoMapper**: Conversión Entity ↔ DTO

</details>

<details>
<summary><strong>2. Flujo de Datos Completo</strong></summary>

### Flujo CREATE (POST)

```
1. Usuario completa formulario en Blazor
   ↓
2. Blazor serializa datos a [Tabla]Dto
   ↓
3. HTTP POST a /api/[tabla]
   ↓
4. Controller recibe [Tabla]Dto
   ↓
5. Controller llama Service.CreateAsync(dto)
   ↓
6. Service usa AutoMapper: dto → Entity
   ↓
7. Service agrega campos de auditoría:
   - AddRecord = DateTime.Now
   - NumUsuar = usuario activo
   - CodiEmprNet = empresa activa
   ↓
8. Service guarda en DbContext
   ↓
9. DbContext ejecuta INSERT en SQL Server
   ↓
10. Service usa AutoMapper: Entity → dto
   ↓
11. Controller retorna 201 Created + dto
   ↓
12. Blazor actualiza UI con nuevo registro
```

### Flujo READ (GET)

```
1. Blazor solicita datos
   ↓
2. HTTP GET a /api/[tabla]
   ↓
3. Controller llama Service.GetAllAsync()
   ↓
4. Service consulta DbContext
   ↓
5. DbContext ejecuta SELECT en SQL Server
   ↓
6. Service usa AutoMapper: List<Entity> → List<Dto>
   ↓
7. Controller retorna 200 OK + List<Dto>
   ↓
8. Blazor renderiza datos en tabla/grid
```

</details>

<details>
<summary><strong>3. Estructura de Entity</strong></summary>

### Template de Entity

```csharp
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using IMPLANPROD.Shared.Interfaces;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa la tabla [NOMBRE_TABLA]
    /// Mapea desde: ERP_ORIGINAL → IMPLANPROD
    /// </summary>
    [Table("[NOMBRE_TABLA]")]
    public class [NombreTabla] : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA
        // ===================================================================
        
        /// <summary>
        /// Identificador único del registro
        /// </summary>
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        // ===================================================================
        // CAMPOS DE NEGOCIO
        // ===================================================================
        
        /// <summary>
        /// [Descripción del campo]
        /// Origen: [CAMPO_ORIGINAL] (tipo_original)
        /// </summary>
        [Column("[NombreCampo]")]
        public decimal Campo1 { get; set; }  // float → decimal
        
        /// <summary>
        /// Código de empresa
        /// Valor por defecto: 0
        /// </summary>
        [Column("Codiempr")]
        public int Codiempr { get; set; } = 0;

        // ===================================================================
        // CAMPOS DE AUDITORÍA - ITimestampEntity, IAuditableEntity
        // ===================================================================
        
        /// <summary>
        /// Fecha de creación del registro
        /// Se asigna automáticamente en el Service
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización
        /// Se actualiza automáticamente en el Service
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa para auditoría
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Usuario que realiza la transacción
        /// Relacionado con usuarios.idflexoft para validación de derechos
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
===================================================================
SQL ORIGINAL (ERP a migrar)
===================================================================

CREATE TABLE [NOMBRE_TABLA] (
    [Campo1] float,
    [Campo2] varchar(50),
    [Codiempr] smallint
)

===================================================================
MODIFICACIONES REALIZADAS
===================================================================

-- Agregar campo ID identity
ALTER TABLE [NOMBRE_TABLA] ADD Id INT IDENTITY(1,1) PRIMARY KEY

-- Cambiar tipo de dato float a decimal
ALTER TABLE [NOMBRE_TABLA] ALTER COLUMN Campo1 DECIMAL(18,2)

-- Cambiar Codiempr de smallint a int
ALTER TABLE [NOMBRE_TABLA] ALTER COLUMN Codiempr INT

-- Agregar campos de auditoría
ALTER TABLE [NOMBRE_TABLA] ADD AddRecord DATETIME NULL
ALTER TABLE [NOMBRE_TABLA] ADD LastUpdate DATETIME NULL
ALTER TABLE [NOMBRE_TABLA] ADD CodiEmprNet INT NULL
ALTER TABLE [NOMBRE_TABLA] ADD NumUsuar INT NULL

-- Establecer valor por defecto para Codiempr
ALTER TABLE [NOMBRE_TABLA] ADD CONSTRAINT DF_[NOMBRE_TABLA]_Codiempr DEFAULT 0 FOR Codiempr

*/
```

### Interfaces Implementadas

**ITimestampEntity**: Proporciona campos de timestamp automáticos
- `AddRecord`: Fecha de creación
- `LastUpdate`: Fecha de última modificación

**IAuditableEntity**: Proporciona campos de auditoría
- `CodiEmprNet`: Empresa que realiza la operación
- `NumUsuar`: Usuario que realiza la operación

</details>

<details>
<summary><strong>4. Estructura de DTO</strong></summary>

### Template de DTO

```csharp
namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para transferencia de datos de [NombreTabla]
    /// Usado en: Blazor UI ↔ API Controller ↔ Service
    /// 
    /// FLUJO:
    /// 1. UI → API: Usuario envía datos del formulario
    /// 2. API → Service: Controller pasa DTO al servicio
    /// 3. Service → DB: AutoMapper convierte DTO → Entity
    /// 4. DB → Service: Entity se recupera de la base de datos
    /// 5. Service → API: AutoMapper convierte Entity → DTO
    /// 6. API → UI: Controller retorna DTO al cliente
    /// </summary>
    public class [NombreTabla]Dto
    {
        /// <summary>
        /// Identificador único
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// [Descripción del campo]
        /// Usado en: [Formulario/Grid/Etc]
        /// </summary>
        public decimal Campo1 { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int Codiempr { get; set; }

        // NOTA: Los campos de auditoría NO se incluyen en el DTO
        // Se asignan automáticamente en el Service
        // - AddRecord
        // - LastUpdate
        // - CodiEmprNet
        // - NumUsuar
    }
}
```

</details>

<details>
<summary><strong>5. Estructura de Service</strong></summary>

### Template de Service

```csharp
using AutoMapper;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.Interfaces;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para gestión de [NombreTabla]
    /// 
    /// RESPONSABILIDADES:
    /// - Lógica de negocio
    /// - Conversión Entity ↔ DTO (AutoMapper)
    /// - Asignación de campos de auditoría
    /// - Interacción con DbContext
    /// </summary>
    public class [NombreTabla]Service : BaseService, I[NombreTabla]Service
    {
        private readonly ApplicationDbContext _context;
        private readonly IMapper _mapper;

        public [NombreTabla]Service(
            ApplicationDbContext context,
            IMapper mapper)
        {
            _context = context;
            _mapper = mapper;
        }

        /// <summary>
        /// Obtiene todos los registros
        /// 
        /// FLUJO:
        /// 1. DbContext consulta tabla en SQL Server
        /// 2. Retorna List<Entity>
        /// 3. AutoMapper convierte List<Entity> → List<Dto>
        /// 4. Retorna List<Dto> al Controller
        /// </summary>
        public async Task<List<[NombreTabla]Dto>> GetAllAsync()
        {
            var entities = await _context.[NombreTabla].ToListAsync();
            return _mapper.Map<List<[NombreTabla]Dto>>(entities);
        }

        /// <summary>
        /// Obtiene un registro por ID
        /// 
        /// FLUJO:
        /// 1. DbContext busca por clave primaria
        /// 2. Si no existe, retorna null
        /// 3. AutoMapper convierte Entity → Dto
        /// 4. Retorna Dto al Controller
        /// </summary>
        public async Task<[NombreTabla]Dto?> GetByIdAsync(int id)
        {
            var entity = await _context.[NombreTabla].FindAsync(id);
            return entity == null ? null : _mapper.Map<[NombreTabla]Dto>(entity);
        }

        /// <summary>
        /// Crea un nuevo registro
        /// 
        /// FLUJO:
        /// 1. AutoMapper convierte Dto → Entity
        /// 2. Asigna campos de auditoría:
        ///    - AddRecord = DateTime.Now
        ///    - NumUsuar = usuario activo
        ///    - CodiEmprNet = empresa activa
        /// 3. DbContext agrega Entity
        /// 4. SaveChanges ejecuta INSERT en SQL Server
        /// 5. AutoMapper convierte Entity → Dto
        /// 6. Retorna Dto con Id generado
        /// </summary>
        public async Task<[NombreTabla]Dto> CreateAsync([NombreTabla]Dto dto)
        {
            var entity = _mapper.Map<[NombreTabla]>(dto);
            
            // Asignar campos de auditoría
            entity.AddRecord = DateTime.Now;
            entity.NumUsuar = GetCurrentUserId(); // Usuario activo
            entity.CodiEmprNet = GetCurrentCompanyId(); // Empresa activa
            
            _context.[NombreTabla].Add(entity);
            await _context.SaveChangesAsync();
            
            return _mapper.Map<[NombreTabla]Dto>(entity);
        }

        /// <summary>
        /// Actualiza un registro existente
        /// 
        /// FLUJO:
        /// 1. DbContext busca Entity por Id
        /// 2. Si no existe, lanza excepción
        /// 3. AutoMapper actualiza propiedades de Entity con datos de Dto
        /// 4. Actualiza campos de auditoría:
        ///    - LastUpdate = DateTime.Now
        ///    - NumUsuar = usuario activo
        /// 5. SaveChanges ejecuta UPDATE en SQL Server
        /// 6. AutoMapper convierte Entity → Dto
        /// 7. Retorna Dto actualizado
        /// </summary>
        public async Task<[NombreTabla]Dto> UpdateAsync([NombreTabla]Dto dto)
        {
            var entity = await _context.[NombreTabla].FindAsync(dto.Id);
            if (entity == null)
                throw new KeyNotFoundException($"[NombreTabla] con Id {dto.Id} no encontrado");
            
            _mapper.Map(dto, entity);
            
            // Actualizar campos de auditoría
            entity.LastUpdate = DateTime.Now;
            entity.NumUsuar = GetCurrentUserId();
            
            await _context.SaveChangesAsync();
            
            return _mapper.Map<[NombreTabla]Dto>(entity);
        }

        /// <summary>
        /// Elimina un registro
        /// 
        /// FLUJO:
        /// 1. DbContext busca Entity por Id
        /// 2. Si no existe, retorna false
        /// 3. DbContext marca Entity para eliminación
        /// 4. SaveChanges ejecuta DELETE en SQL Server
        /// 5. Retorna true si exitoso
        /// </summary>
        public async Task<bool> DeleteAsync(int id)
        {
            var entity = await _context.[NombreTabla].FindAsync(id);
            if (entity == null)
                return false;
            
            _context.[NombreTabla].Remove(entity);
            await _context.SaveChangesAsync();
            
            return true;
        }
    }
}
```

</details>

<details>
<summary><strong>6. Estructura de Controller</strong></summary>

### Template de Controller

```csharp
using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Interfaces;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// API Controller para [NombreTabla]
    /// 
    /// RESPONSABILIDADES:
    /// - Exponer endpoints REST
    /// - Validar requests HTTP
    /// - Retornar códigos de estado apropiados
    /// - Delegar lógica de negocio al Service
    /// 
    /// RUTA BASE: /api/[nombrecontroller]
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class [NombreTabla]Controller : ControllerBase
    {
        private readonly I[NombreTabla]Service _service;

        public [NombreTabla]Controller(I[NombreTabla]Service service)
        {
            _service = service;
        }

        /// <summary>
        /// GET: api/[nombrecontroller]
        /// Obtiene todos los registros
        /// 
        /// FLUJO:
        /// 1. Cliente HTTP GET → /api/[nombrecontroller]
        /// 2. Controller llama Service.GetAllAsync()
        /// 3. Service retorna List<Dto>
        /// 4. Controller retorna 200 OK + List<Dto>
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<List<[NombreTabla]Dto>>> GetAll()
        {
            var result = await _service.GetAllAsync();
            return Ok(result);
        }

        /// <summary>
        /// GET: api/[nombrecontroller]/5
        /// Obtiene un registro por ID
        /// 
        /// FLUJO:
        /// 1. Cliente HTTP GET → /api/[nombrecontroller]/5
        /// 2. Controller llama Service.GetByIdAsync(5)
        /// 3. Si existe: retorna 200 OK + Dto
        /// 4. Si no existe: retorna 404 NotFound
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<[NombreTabla]Dto>> GetById(int id)
        {
            var result = await _service.GetByIdAsync(id);
            if (result == null)
                return NotFound($"[NombreTabla] con Id {id} no encontrado");
            
            return Ok(result);
        }

        /// <summary>
        /// POST: api/[nombrecontroller]
        /// Crea un nuevo registro
        /// 
        /// FLUJO:
        /// 1. Cliente HTTP POST → /api/[nombrecontroller]
        ///    Body: { "campo1": "valor", "campo2": 123 }
        /// 2. ASP.NET deserializa JSON → Dto
        /// 3. Controller llama Service.CreateAsync(dto)
        /// 4. Service crea registro y retorna Dto con Id
        /// 5. Controller retorna 201 Created + Dto
        ///    Header Location: /api/[nombrecontroller]/[id]
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<[NombreTabla]Dto>> Create([FromBody] [NombreTabla]Dto dto)
        {
            var result = await _service.CreateAsync(dto);
            return CreatedAtAction(nameof(GetById), new { id = result.Id }, result);
        }

        /// <summary>
        /// PUT: api/[nombrecontroller]/5
        /// Actualiza un registro existente
        /// 
        /// FLUJO:
        /// 1. Cliente HTTP PUT → /api/[nombrecontroller]/5
        ///    Body: { "id": 5, "campo1": "nuevo valor" }
        /// 2. ASP.NET deserializa JSON → Dto
        /// 3. Valida que id de ruta == id de body
        /// 4. Controller llama Service.UpdateAsync(dto)
        /// 5. Si existe: retorna 200 OK + Dto actualizado
        /// 6. Si no existe: retorna 404 NotFound
        /// </summary>
        [HttpPut("{id}")]
        public async Task<ActionResult<[NombreTabla]Dto>> Update(int id, [FromBody] [NombreTabla]Dto dto)
        {
            if (id != dto.Id)
                return BadRequest("El Id de la ruta no coincide con el Id del objeto");
            
            try
            {
                var result = await _service.UpdateAsync(dto);
                return Ok(result);
            }
            catch (KeyNotFoundException ex)
            {
                return NotFound(ex.Message);
            }
        }

        /// <summary>
        /// DELETE: api/[nombrecontroller]/5
        /// Elimina un registro
        /// 
        /// FLUJO:
        /// 1. Cliente HTTP DELETE → /api/[nombrecontroller]/5
        /// 2. Controller llama Service.DeleteAsync(5)
        /// 3. Si existe: retorna 204 NoContent
        /// 4. Si no existe: retorna 404 NotFound
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(int id)
        {
            var result = await _service.DeleteAsync(id);
            if (!result)
                return NotFound($"[NombreTabla] con Id {id} no encontrado");
            
            return NoContent();
        }
    }
}
```

</details>

<details>
<summary><strong>7. AutoMapper Configuration</strong></summary>

### Template de MappingProfile

```csharp
using AutoMapper;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Server.Mappings
{
    /// <summary>
    /// Configuración de AutoMapper para [NombreTabla]
    /// 
    /// PROPÓSITO:
    /// Define cómo convertir entre Entity y DTO automáticamente
    /// 
    /// CONVERSIONES:
    /// - Entity → DTO: Para retornar datos al cliente
    /// - DTO → Entity: Para guardar datos en la base de datos
    /// 
    /// CAMPOS EXCLUIDOS DEL DTO:
    /// - AddRecord (se asigna en Service)
    /// - LastUpdate (se asigna en Service)
    /// - CodiEmprNet (se asigna en Service)
    /// - NumUsuar (se asigna en Service)
    /// </summary>
    public class [NombreTabla]MappingProfile : Profile
    {
        public [NombreTabla]MappingProfile()
        {
            // Entity → DTO (para lectura)
            // DTO → Entity (para escritura)
            CreateMap<[NombreTabla], [NombreTabla]Dto>()
                .ReverseMap();
            
            // Si necesitas mapeos personalizados:
            // CreateMap<[NombreTabla], [NombreTabla]Dto>()
            //     .ForMember(dest => dest.CampoDto, 
            //                opt => opt.MapFrom(src => src.CampoEntity));
        }
    }
}
```

### Registro en Program.cs

```csharp
// Agregar AutoMapper
builder.Services.AddAutoMapper(typeof(Program));

// Registrar servicio
builder.Services.AddScoped<I[NombreTabla]Service, [NombreTabla]Service>();
```

</details>

<details>
<summary><strong>8. Integración con DbContext</strong></summary>

### Agregar DbSet

```csharp
// ApplicationDbContext.cs

public class ApplicationDbContext : DbContext
{
    // ... otros DbSets ...
    
    /// <summary>
    /// Tabla [NombreTabla]
    /// </summary>
    public DbSet<[NombreTabla]> [NombreTabla] { get; set; }
}
```

### Configuración Adicional (si es necesaria)

```csharp
protected override void OnModelCreating(ModelBuilder modelBuilder)
{
    base.OnModelCreating(modelBuilder);
    
    // Configuración de [NombreTabla]
    modelBuilder.Entity<[NombreTabla]>(entity =>
    {
        // Índices
        entity.HasIndex(e => e.Codiempr);
        
        // Valores por defecto
        entity.Property(e => e.Codiempr)
              .HasDefaultValue(0);
        
        // Precisión de decimales
        entity.Property(e => e.Campo1)
              .HasPrecision(18, 2);
    });
}
```

</details>

<details>
<summary><strong>9. Testing y Debugging</strong></summary>

### Probar con Swagger

1. Ejecutar la aplicación
2. Navegar a `/swagger`
3. Probar cada endpoint:

```
GET /api/[nombrecontroller]
  → Debe retornar lista vacía o con datos

POST /api/[nombrecontroller]
  Body: { "campo1": "test", "campo2": 123 }
  → Debe retornar 201 Created + objeto con Id

GET /api/[nombrecontroller]/{id}
  → Debe retornar 200 OK + objeto

PUT /api/[nombrecontroller]/{id}
  Body: { "id": 1, "campo1": "actualizado" }
  → Debe retornar 200 OK + objeto actualizado

DELETE /api/[nombrecontroller]/{id}
  → Debe retornar 204 NoContent
```

### Logs Importantes

```
// En Service
_logger.LogInformation("GetAllAsync: Se obtuvieron {Count} registros", result.Count);
_logger.LogInformation("CreateAsync: Registro creado con Id {Id}", entity.Id);
_logger.LogWarning("GetByIdAsync: Registro {Id} no encontrado", id);
_logger.LogError(ex, "Error al actualizar registro {Id}", dto.Id);
```

### Debugging en Visual Studio

1. **Breakpoints clave:**
   - Controller: Inicio de cada método
   - Service: Antes de SaveChangesAsync()
   - AutoMapper: Conversiones

2. **Watch variables:**
   - `dto` (datos del cliente)
   - `entity` (datos de la base de datos)
   - `_context.ChangeTracker.Entries()` (cambios pendientes)

3. **SQL Profiler:**
   - Ver queries generadas por EF Core
   - Verificar INSERTs, UPDATEs, DELETEs

</details>

---

## 📊 Resumen Técnico

| Componente | Ubicación | Propósito |
|------------|-----------|-----------|
| Entity | `/Shared/Entities/[Tabla].cs` | Modelo de datos (mapea tabla SQL) |
| DTO | `/Shared/DTOs/[Tabla]Dto.cs` | Transferencia UI ↔ API |
| Interface | `/Shared/Interfaces/I[Tabla]Service.cs` | Contrato del servicio |
| Service | `/Server/Services/[Tabla]Service.cs` | Lógica de negocio |
| Controller | `/Server/Controllers/[Tabla]Controller.cs` | Endpoints REST |
| Mapping | `/Server/Mappings/[Tabla]MappingProfile.cs` | AutoMapper config |
| DbContext | `/Server/Data/ApplicationDbContext.cs` | Acceso a datos |
| UI | `/Client/Pages/[Tabla]/[Tabla]List.razor` | Interfaz Blazor |

**Patrón de arquitectura:** Clean Architecture con separación de capas (UI → API → Service → Data)

**Tecnologías:** ASP.NET Core, Entity Framework Core, AutoMapper, Blazor WebAssembly

**Convenciones:**
- Codiempr siempre = 0
- NumUsuar = usuario activo (relacionado con usuarios.idflexoft)
- Campos de auditoría automáticos (ITimestampEntity, IAuditableEntity)
- Conversiones: float → decimal, short → int (según indicación)
