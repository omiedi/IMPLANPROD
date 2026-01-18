-- Script para exportar DESCREP con saltos de línea preservados
-- Ejecutar en SQL Server Management Studio

-- Opción 1: Usar BCP (Bulk Copy Program) desde línea de comandos
-- Ejecutar esto en CMD o PowerShell:
/*
bcp "SELECT CodigoReparacion, Especificacion FROM IMPLANPROD.dbo.DESCREP WHERE CodigoReparacion = '974000710'" queryout "C:\temp\descrep_export.txt" -c -T -S localhost
*/

-- Opción 2: Ver el texto con saltos de línea visibles
SELECT 
    CodigoReparacion,
    -- Reemplazar CRLF por un marcador visible
    REPLACE(REPLACE(Especificacion, CHAR(13) + CHAR(10), '↵' + CHAR(13) + CHAR(10)), CHAR(13) + CHAR(10), CHAR(13) + CHAR(10)) as EspecificacionConMarcas,
    -- Original
    Especificacion
FROM DESCREP
WHERE CodigoReparacion = '974000710';

-- Opción 3: Exportar a XML (preserva saltos de línea)
SELECT 
    CodigoReparacion,
    Especificacion
FROM DESCREP
WHERE CodigoReparacion = '974000710'
FOR XML PATH('Registro'), ROOT('Descrep');

-- Opción 4: Mostrar cada línea por separado
DECLARE @Texto NVARCHAR(MAX);
DECLARE @CodigoRep VARCHAR(24) = '974000710';

SELECT @Texto = Especificacion 
FROM DESCREP 
WHERE CodigoReparacion = @CodigoRep;

-- Dividir por líneas y mostrar
DECLARE @Lineas TABLE (Numero INT IDENTITY(1,1), Linea NVARCHAR(MAX));

DECLARE @Pos INT = 1;
DECLARE @NextPos INT;
DECLARE @Linea NVARCHAR(MAX);

WHILE @Pos <= LEN(@Texto)
BEGIN
    SET @NextPos = CHARINDEX(CHAR(13) + CHAR(10), @Texto, @Pos);
    
    IF @NextPos = 0
    BEGIN
        -- Última línea
        SET @Linea = SUBSTRING(@Texto, @Pos, LEN(@Texto) - @Pos + 1);
        INSERT INTO @Lineas (Linea) VALUES (@Linea);
        BREAK;
    END
    ELSE
    BEGIN
        SET @Linea = SUBSTRING(@Texto, @Pos, @NextPos - @Pos);
        INSERT INTO @Lineas (Linea) VALUES (@Linea);
        SET @Pos = @NextPos + 2; -- Saltar CRLF
    END
END

-- Mostrar líneas numeradas
SELECT 
    Numero as NumLinea,
    Linea,
    LEN(Linea) as Longitud
FROM @Lineas
ORDER BY Numero;
