-- =============================================
-- Script de Migración: Ajuste de escala en BOLRECEP.NumeBoRe
-- Descripción: Cambia NumeBoRe de DECIMAL(18,4) a DECIMAL(18,0) para evitar guardar .0000
-- Fecha: 2026-03-02
-- =============================================

-- Verificar que exista la tabla y la columna
IF EXISTS (
    SELECT 1
    FROM sys.columns c
    INNER JOIN sys.objects o ON o.object_id = c.object_id
    WHERE o.type = 'U'
      AND o.name = 'BOLRECEP'
      AND c.name = 'NumeBoRe'
)
BEGIN
    -- 1) Normalizar datos: si por algún motivo hay valores con decimales, los truncamos.
    --    (En la práctica deberían ser siempre NNNNNN.0000)
    IF EXISTS (SELECT 1 FROM dbo.BOLRECEP WHERE NumeBoRe IS NOT NULL AND NumeBoRe <> FLOOR(NumeBoRe))
    BEGIN
        UPDATE dbo.BOLRECEP
        SET NumeBoRe = FLOOR(NumeBoRe)
        WHERE NumeBoRe IS NOT NULL
          AND NumeBoRe <> FLOOR(NumeBoRe);

        PRINT 'Se normalizaron registros con decimales en BOLRECEP.NumeBoRe (FLOOR)';
    END

    -- 2) Recrear índice si existe (por seguridad al cambiar el tipo)
    IF EXISTS (
        SELECT 1
        FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.BOLRECEP')
          AND name = N'IX_BOLRECEP_NumeBoRe'
    )
    BEGIN
        DROP INDEX [IX_BOLRECEP_NumeBoRe] ON [dbo].[BOLRECEP];
        PRINT 'Índice IX_BOLRECEP_NumeBoRe eliminado';
    END

    -- 3) Alterar la columna para eliminar escala
    ALTER TABLE dbo.BOLRECEP
    ALTER COLUMN NumeBoRe DECIMAL(18,0) NULL;

    PRINT 'Columna BOLRECEP.NumeBoRe alterada a DECIMAL(18,0)';

    -- 4) Volver a crear índice
    CREATE INDEX IX_BOLRECEP_NumeBoRe ON [dbo].[BOLRECEP] ([NumeBoRe] ASC);
    PRINT 'Índice IX_BOLRECEP_NumeBoRe recreado';
END
ELSE
BEGIN
    PRINT 'No se encontró dbo.BOLRECEP.NumeBoRe. No se aplicaron cambios.';
END
GO
