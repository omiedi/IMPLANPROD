-- Script para crear la tabla DESCREP
-- Descripción de Reparaciones (legacy VB6)

CREATE TABLE [dbo].[DESCREP] (
    [CodigoReparacion] [varchar](24) NOT NULL,
    [Especificacion] [varchar](max) NULL,
    [Status] [int] NOT NULL DEFAULT 0,
    [AddRecord] [datetime2](7) NOT NULL,
    [LastUpdate] [datetime2](7) NOT NULL,
    [CodiEmprNet] [varchar](4) NULL,
    [NumUsuar] [int] NOT NULL DEFAULT 0,
    
    CONSTRAINT [PK_DESCREP] PRIMARY KEY CLUSTERED ([CodigoReparacion] ASC)
);
GO

-- Índice para búsquedas por código
CREATE NONCLUSTERED INDEX [IX_DESCREP_CodigoReparacion] 
ON [dbo].[DESCREP] ([CodigoReparacion] ASC);
GO
