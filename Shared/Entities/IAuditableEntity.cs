using System;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Interface para entidades que requieren auditoría completa:
    /// - Campos de timestamp (AddRecord y LastUpdate)
    /// - Información de usuario y empresa
    /// </summary>
    public interface IAuditableEntity : ITimestampEntity
    {
        /// <summary>
        /// Código de empresa del usuario que creó/modificó el registro
        /// </summary>
        int? CodiEmprNet { get; set; }
        
        /// <summary>
        /// ID del usuario que creó/modificó el registro
        /// </summary>
        int? NumUsuar { get; set; }
    }
}
// ***  INSTRUCCION PARA CARGAR LOS CAMPOS CODIEMPRNET,NUMUSUAR,ADDRECORD Y LASTUPDATE CON VALORES POR DEFECTO   ****
/*DECLARE @sql NVARCHAR(MAX);

DECLARE cur CURSOR FOR
SELECT 
    'ALTER TABLE dbo.' + QUOTENAME(t.name) +
    ' ADD ' + cr.Campo + ' ' + cr.Definicion
FROM sys.tables t
CROSS JOIN (
    SELECT 'AddRecord' AS Campo, 'DATETIME NOT NULL DEFAULT (GETDATE())' AS Definicion UNION ALL
    SELECT 'LastUpdate', 'DATETIME NOT NULL DEFAULT (GETDATE())' UNION ALL
    SELECT 'CodiEmprNet', 'INT NOT NULL DEFAULT (101)' UNION ALL
    SELECT 'NumUsuar', 'INT NOT NULL DEFAULT (10)'
) cr
WHERE NOT EXISTS (
    SELECT 1
    FROM sys.columns c
    WHERE c.object_id = t.object_id
      AND c.name = cr.Campo
);

OPEN cur;
FETCH NEXT FROM cur INTO @sql;

WHILE @@FETCH_STATUS = 0
BEGIN
    PRINT @sql;
EXEC sp_executesql @sql;
FETCH NEXT FROM cur INTO @sql;
END

CLOSE cur;
DEALLOCATE cur;
*/
