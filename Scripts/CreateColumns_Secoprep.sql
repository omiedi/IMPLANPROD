USE [IMPLANPROD]
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[SECOPREP]') AND name = 'Minutos')
BEGIN
    ALTER TABLE [dbo].[SECOPREP] ADD [Minutos] [float] NULL;
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[SECOPREP]') AND name = 'Equipo')
BEGIN
    ALTER TABLE [dbo].[SECOPREP] ADD [Equipo] [int] NULL;
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[SECOPREP]') AND name = 'CantidadOperarios')
BEGIN
    ALTER TABLE [dbo].[SECOPREP] ADD [CantidadOperarios] [int] NULL;
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[SECOPREP]') AND name = 'Categoria')
BEGIN
    ALTER TABLE [dbo].[SECOPREP] ADD [Categoria] [int] NULL;
END
GO
