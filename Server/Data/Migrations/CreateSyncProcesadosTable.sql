-- =============================================
-- Script de Migración: Tabla SyncProcesados
-- Descripción: Crea la tabla de auditoría para sincronización Central-Sucursal
-- Fecha: 2026-01-23
-- Autor: Sistema de Sincronización
-- =============================================

-- Verificar si la tabla ya existe
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SyncProcesados]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[SyncProcesados](
        [Id] [int] IDENTITY(1,1) NOT NULL,
        [Archivo] [varchar](200) NOT NULL,
        [Tabla] [varchar](50) NOT NULL,
        [Operacion] [varchar](20) NULL,
        [FechaProcesamiento] [datetime] NOT NULL,
        [Estado] [varchar](20) NOT NULL,
        [Error] [varchar](max) NULL,
        [CodintAfectado] [int] NULL,
        [Codiempr] [smallint] NOT NULL DEFAULT 0,
        [AddRecord] [datetime] NULL,
        [LastUpdate] [datetime] NULL,
        [CodiEmprNet] [int] NULL,
        [NumUsuar] [int] NULL,
        CONSTRAINT [PK_SyncProcesados] PRIMARY KEY CLUSTERED ([Id] ASC)
    ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

    PRINT 'Tabla SyncProcesados creada exitosamente'
END
ELSE
BEGIN
    PRINT 'La tabla SyncProcesados ya existe'
END
GO

-- Crear índices para mejorar el rendimiento
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_SyncProcesados_Archivo' AND object_id = OBJECT_ID('SyncProcesados'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_SyncProcesados_Archivo] ON [dbo].[SyncProcesados]
    (
        [Archivo] ASC
    )
    INCLUDE ([Estado], [FechaProcesamiento])
    PRINT 'Índice IX_SyncProcesados_Archivo creado'
END
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_SyncProcesados_Tabla_Estado' AND object_id = OBJECT_ID('SyncProcesados'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_SyncProcesados_Tabla_Estado] ON [dbo].[SyncProcesados]
    (
        [Tabla] ASC,
        [Estado] ASC
    )
    INCLUDE ([FechaProcesamiento], [CodintAfectado])
    PRINT 'Índice IX_SyncProcesados_Tabla_Estado creado'
END
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_SyncProcesados_FechaProcesamiento' AND object_id = OBJECT_ID('SyncProcesados'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_SyncProcesados_FechaProcesamiento] ON [dbo].[SyncProcesados]
    (
        [FechaProcesamiento] DESC
    )
    INCLUDE ([Tabla], [Estado], [Archivo])
    PRINT 'Índice IX_SyncProcesados_FechaProcesamiento creado'
END
GO

PRINT 'Migración completada exitosamente'
GO
