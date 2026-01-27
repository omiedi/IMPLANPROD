# Sistema de Sincronización Central-Sucursal

## Descripción General

Sistema de sincronización de datos entre la Central y Sucursales para la tabla MASTER (maestro de artículos/productos). Reemplaza el sistema legacy de VB6 que utilizaba scripts SQL en archivos de texto.

## Arquitectura

### Componentes Principales

1. **DTOs de Sincronización** (`Shared/DTOs/`)
   - `SyncMasterDTO`: Estructura del archivo JSON de sincronización
   - `SyncConfigDTO`: Configuración del sistema
   - `SyncResultDTO`: Resultado del procesamiento

2. **Entidad de Auditoría** (`Shared/Entities/SyncProcesado.cs`)
   - Tabla `SyncProcesados` para trazabilidad completa

3. **Servicio de Sincronización** (`Server/Services/`)
   - `ISyncService` / `SyncService`: Lógica de generación y procesamiento

4. **Controlador API** (`Server/Controllers/SyncController.cs`)
   - Endpoints para configuración y procesamiento

5. **Interfaz de Usuario** (`Client/`)
   - `SyncConfigModal.razor`: Modal de configuración
   - Integración en `MasterIndex.razor`

## Flujo de Trabajo

### Central (Genera archivos)

```
Alta/Modificación Producto
    ↓
MasterController.PostAsync() / PutAsync()
    ↓
SaveChanges() + Commit Transaction
    ↓
SyncService.GenerarArchivoSyncMasterAsync()
    ↓
Archivo JSON en carpeta Intercambio
    Formato: MASTER_{Codint}_{timestamp}.json
```

### Sucursal (Procesa archivos)

```
Usuario presiona "Sincronizar desde Central"
    ↓
SyncService.ProcesarArchivosSyncAsync()
    ↓
Lee archivos MASTER_*.json de carpeta Intercambio
    ↓
Para cada archivo:
    - Valida si ya fue procesado (tabla SyncProcesados)
    - Ejecuta MERGE SQL (UPSERT)
    - Registra en auditoría
    ↓
Retorna resultado con estadísticas
```

## Configuración

### 1. Configurar appsettings.json

```json
{
  "SyncConfig": {
    "GeneraScriptSync": true,
    "RutaIntercambio": "\\\\servidor\\intercambio",
    "EsCentral": true
  }
}
```

**Parámetros:**
- `GeneraScriptSync`: `true` para activar sincronización
- `RutaIntercambio`: Ruta de carpeta compartida (UNC o local)
- `EsCentral`: `true` = Central (genera), `false` = Sucursal (procesa)

### 2. Crear Carpeta Compartida

**En el servidor Central:**
```
1. Crear carpeta: C:\Intercambio
2. Compartir como: \\SERVIDOR\Intercambio
3. Permisos: Lectura/Escritura para usuarios del sistema
4. Mapear en Sucursales como unidad de red o usar ruta UNC
```

### 3. Ejecutar Migración de Base de Datos

**En ambas bases de datos (Central y Sucursal):**
```sql
-- Ejecutar script:
Server/Data/Migrations/CreateSyncProcesadosTable.sql
```

Esto crea:
- Tabla `SyncProcesados`
- Índices optimizados

## Uso en la Aplicación

### Configuración Inicial

1. Ir a **Ficha de Productos** (`/master-index`)
2. Clic en **"Configuración Sincronización"**
3. Activar **"Generar archivos de sincronización"**
4. Seleccionar si es **Central** o **Sucursal**
5. Ingresar **ruta de carpeta compartida**
6. Clic en **"Validar"** para verificar acceso
7. **Guardar Configuración**

### Operación en Central

**Automático:** Cada vez que se crea o modifica un producto, se genera automáticamente un archivo JSON en la carpeta de intercambio.

**Ejemplo de archivo generado:**
```
MASTER_1956_20260123_143025.json
```

**Contenido:**
```json
{
  "Tabla": "MASTER",
  "Operacion": "UPSERT",
  "Fecha": "2026-01-23T14:30:25",
  "Datos": {
    "Codint": 1956,
    "Codigo": "BOL-40X60",
    "Descripcion": "Bolsa 40x60",
    "Umedida": "UN",
    ... (todos los campos de Master)
  }
}
```

### Operación en Sucursal

1. Ir a **Ficha de Productos**
2. Aparece botón **"Sincronizar desde Central"**
3. Clic para procesar archivos pendientes
4. El sistema:
   - Lee todos los archivos `MASTER_*.json`
   - Omite archivos ya procesados
   - Ejecuta MERGE (INSERT o UPDATE según corresponda)
   - Registra en tabla de auditoría
   - Muestra resultado con estadísticas

## Ventajas sobre el Sistema Legacy (VB6)

| Aspecto | VB6 (Legacy) | .NET (Nuevo) |
|---------|--------------|--------------|
| Formato | SQL Scripts | JSON |
| Seguridad | SQL Injection vulnerable | Parametrizado seguro |
| Trazabilidad | Ninguna | Tabla SyncProcesados |
| Control de errores | Básico | Completo con logs |
| Reejecutabilidad | Problemática | MERGE idempotente |
| Validación | Ninguna | Validación de ruta y datos |
| Auditoría | No | Sí (fecha, estado, errores) |

## Tabla de Auditoría: SyncProcesados

```sql
SELECT TOP 100 
    Archivo,
    Tabla,
    Operacion,
    FechaProcesamiento,
    Estado,
    Error,
    CodintAfectado
FROM SyncProcesados
ORDER BY FechaProcesamiento DESC
```

**Estados posibles:**
- `EXITOSO`: Procesado correctamente
- `ERROR`: Error durante el procesamiento
- `PENDIENTE`: No utilizado actualmente

## Monitoreo y Troubleshooting

### Verificar Archivos Generados (Central)

```powershell
# Ver archivos generados hoy
Get-ChildItem "\\servidor\intercambio\MASTER_*.json" | 
    Where-Object {$_.CreationTime -gt (Get-Date).Date} |
    Select-Object Name, CreationTime, Length
```

### Verificar Procesamiento (Sucursal)

```sql
-- Archivos procesados hoy
SELECT 
    Archivo,
    FechaProcesamiento,
    Estado,
    CodintAfectado,
    Error
FROM SyncProcesados
WHERE CAST(FechaProcesamiento AS DATE) = CAST(GETDATE() AS DATE)
ORDER BY FechaProcesamiento DESC

-- Archivos con error
SELECT * FROM SyncProcesados WHERE Estado = 'ERROR'
```

### Logs del Sistema

Los logs se encuentran en:
```
Server/Logs/log-{fecha}.txt
```

Buscar entradas relacionadas con sincronización:
```
[INF] Archivo de sincronización generado: MASTER_1956_20260123_143025.json
[WRN] Error al generar archivo de sincronización (no afecta la operación principal)
[ERR] Error al procesar archivo: MASTER_1956_20260123_143025.json
```

## Reprocesar Archivo con Error

Si un archivo falló, se puede reprocesar:

1. **Eliminar registro de auditoría:**
```sql
DELETE FROM SyncProcesados 
WHERE Archivo = 'MASTER_1956_20260123_143025.json'
```

2. **Volver a sincronizar** desde la interfaz

## Extensión Futura: Tabla MOVISTO

El sistema está diseñado para extenderse fácilmente a otras tablas como MOVISTO (remitos).

**Pasos para agregar MOVISTO:**

1. Crear `SyncMovistoDTO` similar a `SyncMasterDTO`
2. Agregar método `GenerarArchivoSyncMovistoAsync()` en `SyncService`
3. Agregar método `ProcesarMovistoUpsertAsync()` con lógica específica:
   - Convertir `CantSalid` → `CantIngre`
   - Convertir `CodiMovi` RT → BR
4. Integrar en controladores de remitos

## Consideraciones de Seguridad

1. **Permisos de carpeta compartida**: Solo usuarios autorizados
2. **Validación de archivos**: El sistema valida estructura JSON
3. **Transacciones**: Uso de transacciones SQL para atomicidad
4. **Auditoría completa**: Todos los cambios quedan registrados
5. **No expone credenciales**: Configuración en servidor

## Soporte

Para consultas o problemas, revisar:
1. Logs del sistema
2. Tabla `SyncProcesados`
3. Permisos de carpeta compartida
4. Conectividad de red entre Central y Sucursales

---

**Versión:** 1.0  
**Fecha:** Enero 2026  
**Autor:** Sistema IMPLANPROD
