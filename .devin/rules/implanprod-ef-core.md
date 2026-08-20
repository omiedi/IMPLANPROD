# Reglas de EF Core y Acceso a Datos en IMPLANPROD

## Consultas de lectura

- Verificar siempre que no se generen bloqueos en la base de datos.
- Usar `AsNoTracking()` en consultas de lectura (`GET`) cuando aplique.
- Evitar trackear entidades en operaciones que solo leen datos.

## Campos de auditoría en nuevas entidades

Toda tabla creada como entidad nueva debe implementar `ITimestampEntity` e `IAuditableEntity` con los siguientes campos:

```csharp
[Column("AddRecord")]
public DateTime? AddRecord { get; set; }

[Column("LastUpdate")]
public DateTime? LastUpdate { get; set; }

[Column("CodiEmprNet")]
public int? CodiEmprNet { get; set; }

[Column("NumUsuar")]
public int? NumUsuar { get; set; }
```

- `CodiEmprNet` siempre debe asignarse el valor `0` por defecto.
- `NumUsuar` representa el usuario activo. Para validar derechos en `DER_ACCE` se busca el match con `usuarios.idflexoft`.

## Conversiones de tipos

- `float` → `decimal`: SIEMPRE (automático).
- `short`/`smallint` → `int`: SOLO cuando el usuario lo solicite explícitamente.
- `CodiEmpr` → `int` con valor `0`.

## Documentación de tablas

- Al final de cada entidad nueva comentar el SQL `CREATE` original y los `ALTER` aplicados.
- Documentar el mapeo front → back → front.
