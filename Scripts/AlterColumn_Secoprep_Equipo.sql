USE [IMPLANPROD]
GO

-- Si la columna Equipo ya existe, cambiar su tipo
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[SECOPREP]') AND name = 'Equipo')
BEGIN
    ALTER TABLE [dbo].[SECOPREP] ALTER COLUMN [Equipo] [nvarchar](6) NULL;
END
ELSE
BEGIN
    -- Si no existe, agregarla como nvarchar(6)
    ALTER TABLE [dbo].[SECOPREP] ADD [Equipo] [nvarchar](6) NULL;
END
GO
