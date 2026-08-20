---
description: Sistema de Sincronización Central-Sucursal para Master, Formulas y SecOper
---

# Sistema de Sincronización Central-Sucursal

Sistema automático de sincronización de datos entre instalación Central y Sucursales para las tablas **MASTER** (productos), **FORMULAS** (componentes/BOM) y **SECOPER** (operaciones de producción).

---

## > Instructivo para el Usuario

<details>
<summary><strong>1. Configuración Inicial del Sistema</strong></summary>

### Requisitos Previos
- Carpeta compartida en red accesible desde Central y todas las Sucursales
- Permisos de lectura/escritura en la carpeta compartida
- Base de datos con tablas `SyncConfig` y `SyncProcesados` creadas

### Configurar en Central
1. Ir a **Configuración → Sincronización**
2. Marcar **"Es Central"**: ✅
3. Marcar **"Generar Scripts de Sincronización"**: ✅
4. **Ruta de Intercambio**: `C:\SyncCentral` (carpeta local que se comparte en red)
5. Guardar configuración

### Configurar en cada Sucursal
1. Ir a **Configuración → Sincronización**
2. Marcar **"Es Central"**: ❌ (desmarcar)
3. Marcar **"Generar Scripts de Sincronización"**: ✅
4. **Ruta de Intercambio**: `\\SERVIDOR\SyncCentral` (ruta UNC a la carpeta compartida)
5. Guardar configuración

</details>

<details>
<summary><strong>2. Funcionamiento Automático</strong></summary>

### En Central (Generación Automática)
Cuando se crea o modifica:
- **Producto (Master)** → Se genera `MASTER_{Codint}_{timestamp}.json`
- **Componentes (Formulas)** → Se genera `FORMULAS_{CodIConj}_{timestamp}.json`
- **Operaciones (SecOper)** → Se genera `SECOPER_{CodIConj}_{timestamp}.json`

Los archivos se guardan automáticamente en la carpeta de intercambio.

### En Sucursales (Procesamiento Automático)
- El sistema procesa archivos **cada 4 minutos** automáticamente
- No requiere intervención manual
- Actualiza las tablas Master, Formulas y SecOper
- Registra todo en auditoría

</details>

<details>
<summary><strong>3. Sincronización Manual (Opcional)</strong></summary>

Si necesitas forzar una sincronización inmediata en una Sucursal:

1. Ir a **Configuración → Sincronización**
2. Hacer clic en **"Procesar Archivos Ahora"**
3. El sistema mostrará:
   - Cantidad de archivos procesados
   - Cantidad de registros sincronizados
   - Errores si los hubiera

</details>

<details>
<summary><strong>4. Verificar Estado de Sincronización</strong></summary>

### Ver Archivos Generados
- En Central: Revisar carpeta `C:\SyncCentral`
- Buscar archivos `MASTER_*.json`, `FORMULAS_*.json`, `SECOPER_*.json`
- Verificar fecha/hora de generación

### Ver Historial de Procesamiento
1. Consultar tabla `SyncProcesados` en la base de datos
2. Campos importantes:
   - **Archivo**: Nombre del archivo procesado
   - **Tabla**: MASTER, FORMULAS o SECOPER
   - **FechaProcesamiento**: Cuándo se procesó
   - **Estado**: EXITOSO o ERROR
   - **Error**: Descripción del error si falló

### Ver Logs del Sistema
- Revisar logs de la aplicación
- Buscar entradas de `SyncBackgroundService`
- Verificar mensajes de sincronización

</details>

<details>
<summary><strong>5. Solución de Problemas Comunes</strong></summary>

### Problema: No se generan archivos en Central
**Solución:**
- Verificar que "Es Central" esté marcado
- Verificar que "Generar Scripts" esté habilitado
- Verificar permisos de escritura en la carpeta
- Verificar que la ruta de intercambio exista

### Problema: Sucursal no procesa archivos
**Solución:**
- Verificar que "Es Central" NO esté marcado
- Verificar que "Generar Scripts" esté habilitado
- Verificar conectividad de red a la carpeta compartida
- Verificar permisos de lectura en la carpeta compartida
- Revisar logs del BackgroundService

### Problema: Errores al procesar archivos
**Solución:**
- Consultar tabla `SyncProcesados` para ver el error específico
- Verificar que el archivo JSON no esté corrupto
- Verificar que la estructura de base de datos sea correcta
- Revisar logs para más detalles

### Problema: Sincronización muy lenta
**Solución:**
- Verificar velocidad de red
- Verificar que no haya muchos archivos pendientes
- Considerar limpiar archivos antiguos ya procesados

</details>

---

## > Instructivo para el Desarrollador

<details>
<summary><strong>1. Arquitectura del Sistema</strong></summary>

### Componentes Principales

**DTOs (Data Transfer Objects):**
- `SyncMasterDTO` - Sincronización de productos
- `SyncFormulasDTO` + `FormulaItemDTO` - Sincronización de componentes
- `SyncSecOperDTO` + `SecOperItemDTO` - Sincronización de operaciones
- `SyncConfigDTO` - Configuración del sistema
- `SyncResultDTO` - Resultado del procesamiento

**Servicios:**
- `ISyncService` / `SyncService` - Lógica principal de sincronización
- `SyncBackgroundService` - Procesamiento automático en segundo plano

**Controladores:**
- `MasterController` - Genera archivos al crear/actualizar productos
- `FormulasController` - Genera archivos al guardar componentes
- `SecoperController` - Genera archivos al guardar operaciones
- `SyncController` - Endpoints de configuración y procesamiento manual

**Entidades:**
- `SyncConfig` - Configuración persistente en BD
- `SyncProcesados` - Auditoría de archivos procesados

</details>

<details>
<summary><strong>2. Flujo de Sincronización</strong></summary>

### Generación en Central

```
Usuario modifica Master/Formulas/SecOper
    ↓
Controller llama a ISyncService.GenerarArchivo...Async(codint/codiConj)
    ↓
SyncService verifica: EsCentral && GeneraScriptSync
    ↓
Consulta datos de la BD
    ↓
Serializa a JSON con formato indentado
    ↓
Guarda archivo: {TABLA}_{ID}_{timestamp}.json en RutaIntercambio
    ↓
Log: "Archivo de sincronización generado: {nombre}"
```

### Procesamiento en Sucursal

```
SyncBackgroundService ejecuta cada 4 minutos
    ↓
Verifica: !EsCentral && GeneraScriptSync
    ↓
Busca archivos MASTER_*.json, FORMULAS_*.json, SECOPER_*.json
    ↓
Ordena por fecha de creación
    ↓
Para cada archivo:
    ↓
    Verifica si ya fue procesado (tabla SyncProcesados)
    ↓
    Deserializa JSON
    ↓
    Ejecuta ProcesarMasterUpsertAsync / ProcesarFormulasUpsertAsync / ProcesarSecOperUpsertAsync
    ↓
    Registra en SyncProcesados (EXITOSO o ERROR)
    ↓
Log: resultado del procesamiento
```

</details>

<details>
<summary><strong>3. Métodos Principales del SyncService</strong></summary>

### Generación de Archivos

```csharp
// Genera archivo para un producto Master
Task<bool> GenerarArchivoSyncMasterAsync(int codint)

// Genera archivo para componentes de una fórmula
Task<bool> GenerarArchivoSyncFormulasAsync(int codiConj)

// Genera archivo para operaciones de producción
Task<bool> GenerarArchivoSyncSecOperAsync(int codiConj)
```

### Procesamiento de Archivos

```csharp
// Procesa todos los archivos pendientes en la carpeta
Task<SyncResultDTO> ProcesarArchivosSyncAsync(string rutaIntercambio)

// Métodos privados de procesamiento
private Task ProcesarMasterUpsertAsync(Master master)
private Task ProcesarFormulasUpsertAsync(SyncFormulasDTO syncDto)
private Task ProcesarSecOperUpsertAsync(SyncSecOperDTO syncDto)
```

### Configuración y Auditoría

```csharp
// Obtiene configuración desde BD
SyncConfigDTO ObtenerConfiguracion()

// Guarda configuración en BD
Task<bool> GuardarConfiguracionAsync(SyncConfigDTO config)

// Registra procesamiento en auditoría
Task RegistrarProcesamientoAsync(string archivo, string tabla, 
    string operacion, string estado, string? error, int? codintAfectado)
```

</details>

<details>
<summary><strong>4. Estrategias de UPSERT</strong></summary>

### Master (MERGE SQL)
Usa `MERGE` de SQL Server para insertar o actualizar basado en `Codint`:
- **MATCHED**: Actualiza todos los campos
- **NOT MATCHED**: Inserta nuevo registro
- Maneja ~150 campos del Master

### Formulas y SecOper (DELETE + INSERT)
Usa estrategia de reemplazo completo:
1. Inicia transacción
2. `DELETE FROM tabla WHERE CodIConj = @CodIConj`
3. `INSERT` de cada componente/operación
4. Commit de transacción
5. Rollback en caso de error

**Razón:** Simplifica la lógica al reemplazar todo el conjunto de componentes/operaciones.

</details>

<details>
<summary><strong>5. BackgroundService - SyncBackgroundService</strong></summary>

### Características

```csharp
public class SyncBackgroundService : BackgroundService
{
    private readonly TimeSpan _intervalo = TimeSpan.FromMinutes(4);
    
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Espera 30s al iniciar
        await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);
        
        while (!stoppingToken.IsCancellationRequested)
        {
            // Verifica: !EsCentral && GeneraScriptSync
            // Procesa archivos
            // Espera 4 minutos
        }
    }
}
```

### Registro en Program.cs

```csharp
builder.Services.AddHostedService<SyncBackgroundService>();
```

### Inyección de Dependencias
Usa `IServiceProvider` para crear scopes y obtener `ISyncService` en cada ejecución.

</details>

<details>
<summary><strong>6. Integración en Controladores</strong></summary>

### Patrón de Integración

```csharp
public class MasterController : ControllerBase
{
    private readonly ISyncService _syncService;
    
    [HttpPost]
    public async Task<ActionResult> PostAsync(Master master)
    {
        // Guardar en BD
        await _context.SaveChangesAsync();
        await transaction.CommitAsync();
        
        // Generar archivo de sincronización
        try
        {
            await _syncService.GenerarArchivoSyncMasterAsync(master.Codint);
        }
        catch (Exception syncEx)
        {
            _logger.LogError(syncEx, "Error al generar archivo de sincronización");
            // No falla la operación principal
        }
        
        return Ok(master);
    }
}
```

**Importante:** La sincronización NO debe hacer fallar la operación principal.

</details>

<details>
<summary><strong>7. Formato de Archivos JSON</strong></summary>

### MASTER_{Codint}_{timestamp}.json

```json
{
  "Tabla": "MASTER",
  "Operacion": "UPSERT",
  "Fecha": "2026-01-23T19:30:00",
  "Datos": {
    "Codint": 12345,
    "Codempr": 0,
    "Codext": "PROD-001",
    "Descrip": "Producto de ejemplo",
    // ... ~150 campos más
  }
}
```

### FORMULAS_{CodIConj}_{timestamp}.json

```json
{
  "Tabla": "FORMULAS",
  "Operacion": "UPSERT",
  "FechaGeneracion": "2026-01-23T19:30:00",
  "CodIConj": 12345,
  "Componentes": [
    {
      "CodIConj": 12345,
      "CodiEmpr": 0,
      "CodiElem": 5678,
      "NuorIt": 10,
      "UsoBruto": 2.5,
      // ... más campos
    }
  ]
}
```

### SECOPER_{CodIConj}_{timestamp}.json

```json
{
  "Tabla": "SECOPER",
  "Operacion": "UPSERT",
  "FechaGeneracion": "2026-01-23T19:30:00",
  "CodIConj": 12345,
  "Operaciones": [
    {
      "CodIConj": 12345,
      "NumeOper": 10,
      "DescOper": "Corte",
      "MinuStan": 15.5,
      // ... más campos
    }
  ]
}
```

</details>

<details>
<summary><strong>8. Tabla de Auditoría - SyncProcesados</strong></summary>

### Estructura

```sql
CREATE TABLE SyncProcesados (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Archivo NVARCHAR(255) NOT NULL,
    Tabla NVARCHAR(50) NOT NULL,
    Operacion NVARCHAR(50) NOT NULL,
    FechaProcesamiento DATETIME NOT NULL,
    Estado NVARCHAR(50) NOT NULL,
    Error NVARCHAR(MAX) NULL,
    CodintAfectado INT NULL
)
```

### Uso

Cada archivo procesado se registra con:
- **Estado**: "EXITOSO" o "ERROR"
- **Error**: Mensaje de error si falló
- **CodintAfectado**: ID del registro afectado (Master.Codint o CodIConj)

Permite:
- Evitar reprocesar archivos ya exitosos
- Auditar qué se sincronizó y cuándo
- Diagnosticar errores

</details>

<details>
<summary><strong>9. Agregar Nueva Tabla a Sincronización</strong></summary>

### Pasos para agregar una nueva tabla

1. **Crear DTO**
```csharp
// En Shared/DTOs/
public class SyncNuevaTablaDTO
{
    public string Tabla { get; set; } = "NUEVATABLA";
    public string Operacion { get; set; } = "UPSERT";
    public DateTime FechaGeneracion { get; set; }
    public int ClavePrincipal { get; set; }
    public List<NuevaTablaItemDTO> Items { get; set; } = new();
}
```

2. **Agregar método a ISyncService**
```csharp
Task<bool> GenerarArchivoSyncNuevaTablaAsync(int clavePrincipal);
```

3. **Implementar en SyncService**
```csharp
public async Task<bool> GenerarArchivoSyncNuevaTablaAsync(int clavePrincipal)
{
    // Consultar datos
    // Crear DTO
    // Serializar a JSON
    // Guardar archivo NUEVATABLA_{clave}_{timestamp}.json
}

private async Task ProcesarNuevaTablaUpsertAsync(SyncNuevaTablaDTO syncDto)
{
    // DELETE + INSERT o MERGE según estrategia
}
```

4. **Actualizar ProcesarArchivosSyncAsync**
```csharp
var archivosNuevaTabla = Directory.GetFiles(rutaIntercambio, "NUEVATABLA_*.json");
var archivos = archivosMaster.Concat(...).Concat(archivosNuevaTabla)...

// En el foreach:
else if (nombreArchivo.StartsWith("NUEVATABLA_"))
{
    var syncDto = JsonSerializer.Deserialize<SyncNuevaTablaDTO>(json);
    await ProcesarNuevaTablaUpsertAsync(syncDto);
    await RegistrarProcesamientoAsync(...);
}
```

5. **Integrar en Controller**
```csharp
await _syncService.GenerarArchivoSyncNuevaTablaAsync(id);
```

</details>

<details>
<summary><strong>10. Testing y Debugging</strong></summary>

### Probar Generación de Archivos

1. Configurar Central con ruta local
2. Crear/modificar un Master
3. Verificar que se genere el archivo JSON
4. Abrir el JSON y verificar estructura
5. Revisar logs

### Probar Procesamiento

1. Copiar archivo JSON a carpeta de Sucursal
2. Ejecutar procesamiento manual desde UI
3. Verificar que se actualice la BD
4. Consultar tabla `SyncProcesados`
5. Revisar logs

### Probar BackgroundService

1. Configurar Sucursal
2. Copiar archivos JSON a carpeta compartida
3. Esperar 4 minutos (o reiniciar app)
4. Verificar logs de `SyncBackgroundService`
5. Verificar BD actualizada

### Logs Importantes

```
SyncBackgroundService iniciado
Iniciando procesamiento automático de sincronización...
Archivo procesado exitosamente: MASTER_12345_20260123_193000.json
Sincronización completada: 3 archivos, 15 registros sincronizados
```

</details>

---

## 📊 Resumen Técnico

| Componente | Ubicación | Propósito |
|------------|-----------|-----------|
| `SyncService.cs` | Server/Services | Lógica de generación y procesamiento |
| `SyncBackgroundService.cs` | Server/Services | Procesamiento automático cada 4 min |
| `SyncController.cs` | Server/Controllers | Endpoints de configuración |
| `SyncConfig` | Server/Data/Entities | Configuración en BD |
| `SyncProcesados` | Server/Data/Entities | Auditoría de procesamiento |
| DTOs | Shared/DTOs | Transferencia de datos JSON |

**Tablas sincronizadas:** MASTER, FORMULAS, SECOPER  
**Frecuencia automática:** 4 minutos (solo Sucursales)  
**Estrategia:** Archivos JSON en carpeta compartida  
**Auditoría:** Completa en tabla `SyncProcesados` y logs
