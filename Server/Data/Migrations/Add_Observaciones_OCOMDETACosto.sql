-- Agregar campo Observaciones a OCOMDETACosto
-- Tipo: VARCHAR(1024) NULL

IF COL_LENGTH('dbo.OCOMDETACosto', 'Observaciones') IS NULL
BEGIN
    ALTER TABLE dbo.OCOMDETACosto
    ADD Observaciones VARCHAR(1024) NULL;

    PRINT 'Columna Observaciones agregada en dbo.OCOMDETACosto.';
END
ELSE
BEGIN
    PRINT 'La columna Observaciones ya existe en dbo.OCOMDETACosto.';
END
GO

-- Validación adicional explícita de longitud (defensiva)
IF NOT EXISTS (
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_OCOMDETACosto_Observaciones_MaxLen'
      AND parent_object_id = OBJECT_ID('dbo.OCOMDETACosto')
)
BEGIN
    ALTER TABLE dbo.OCOMDETACosto
    ADD CONSTRAINT CK_OCOMDETACosto_Observaciones_MaxLen
    CHECK (Observaciones IS NULL OR LEN(Observaciones) <= 1024);

    PRINT 'Constraint CK_OCOMDETACosto_Observaciones_MaxLen creada.';
END
ELSE
BEGIN
    PRINT 'La constraint CK_OCOMDETACosto_Observaciones_MaxLen ya existe.';
END
GO
