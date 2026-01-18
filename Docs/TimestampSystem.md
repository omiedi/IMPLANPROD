# Sistema de Timestamps Automáticos

Este documento explica cómo implementar y usar el sistema de timestamps automáticos en IMPLANPROD, basado en el patrón del proyecto MantenimientoImp.

## Componentes del Sistema

### 1. ITimestampEntity (Interfaz)
Ubicación: `Shared/Entities/ITimestampEntity.cs`

Define los campos requeridos para el sistema de timestamps:
```csharp
public interface ITimestampEntity
{
    DateTime AddRecord { get; set; }    // Fecha de creación del registro
    DateTime LastUpdate { get; set; }   // Fecha de última actualización
}
```

### 2. IAuditable (Interfaz del Servidor)
Ubicación: `Server/Data/Interfaces/IAuditable.cs`

Interfaz equivalente para uso en el servidor (opcional, para compatibilidad):
```csharp
public interface IAuditable
{
    DateTime AddRecord { get; set; }
    DateTime LastUpdate { get; set; }
}
```

### 3. TimestampInterceptor
Ubicación: `Server/Data/Interceptors/TimestampInterceptor.cs`

Interceptor que actualiza automáticamente los campos de timestamp antes de guardar en la base de datos.

**Funcionalidad:**
- **INSERT**: Establece `AddRecord` y `LastUpdate` con la fecha/hora actual
- **UPDATE**: Actualiza solo `LastUpdate` con la fecha/hora actual

## Cómo Implementar en una Entidad

### Paso 1: Agregar los campos a la entidad
```csharp
public class MiEntidad : ITimestampEntity
{
    // ... otros campos de la entidad ...
    
    /// <summary>
    /// Fecha de creación del registro
    /// </summary>
    public DateTime AddRecord { get; set; }
    
    /// <summary>
    /// Fecha de última actualización del registro
    /// </summary>
    public DateTime LastUpdate { get; set; }
}
```

### Paso 2: Agregar los campos a la tabla en la base de datos
```sql
ALTER TABLE [dbo].[MI_TABLA] 
ADD [AddRecord] datetime2 NOT NULL DEFAULT GETDATE(),
    [LastUpdate] datetime2 NOT NULL DEFAULT GETDATE();
```

### Paso 3: ¡Listo!
No se requiere código adicional. El interceptor se encargará automáticamente de:
- Establecer `AddRecord` cuando se crea un nuevo registro
- Actualizar `LastUpdate` cada vez que se modifica el registro

## Ejemplo Completo

```csharp
using IMPLANPROD.Shared.Entities;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

[Table("EJEMPLO")]
public class Ejemplo : ITimestampEntity
{
    [Key]
    public int Id { get; set; }
    
    [StringLength(100)]
    public string Nombre { get; set; } = string.Empty;
    
    [StringLength(500)]
    public string Descripcion { get; set; } = string.Empty;
    
    // Campos de timestamp - se actualizan automáticamente
    public DateTime AddRecord { get; set; }
    public DateTime LastUpdate { get; set; }
}
```

## Configuración del Sistema

El sistema ya está configurado en `Program.cs`:
```csharp
builder.Services.AddDbContext<DataContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"))
           .AddInterceptors(new TimestampInterceptor()));
```

## Ventajas del Sistema

1. **Automático**: No requiere código manual para actualizar fechas
2. **Consistente**: Todas las entidades que implementen la interfaz tendrán el mismo comportamiento
3. **Transparente**: Funciona sin modificar controladores o servicios existentes
4. **Auditoría**: Proporciona trazabilidad automática de cuándo se crearon y modificaron los registros

## Casos de Uso

- **Auditoría básica**: Saber cuándo se creó y modificó por última vez un registro
- **Debugging**: Identificar cuándo ocurrieron cambios problemáticos
- **Reportes**: Mostrar información de creación y actualización en interfaces de usuario
- **Sincronización**: Usar las fechas para sincronización entre sistemas

## Notas Importantes

1. **Zona horaria**: El sistema usa `DateTime.Now` (hora local del servidor)
2. **Precisión**: Los campos son `datetime2` en SQL Server para mayor precisión
3. **Performance**: El interceptor es muy eficiente y no afecta significativamente el rendimiento
4. **Compatibilidad**: Compatible con Entity Framework Core y triggers de auditoría existentes

## Entidades Recomendadas para Implementar

Considera implementar `ITimestampEntity` en:
- Tablas maestras (Master, Deudor12, etc.)
- Tablas de configuración (Tamoned, Familia, etc.)
- Tablas transaccionales importantes (Pedenca, Peddeta, etc.)
- Nuevas entidades que requieran auditoría básica
