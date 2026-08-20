-- =============================================
-- Creación de tabla TRAZAREC
-- Trazabilidad de Recepciones
-- =============================================

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TRAZAREC]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[TRAZAREC](
        [ID] [int] IDENTITY(1,1) NOT NULL,
        [CodiEmpr] [int] NOT NULL DEFAULT 0,
        [NROREMI] [int] NULL,
        [Fech_Recep] [smalldatetime] NULL,
        [Nume_Ocom] [int] NULL,
        [Cant_Remi] [decimal](18, 4) NULL,
        [Cod_Inte] [int] NULL,
        [NroPed] [int] NULL,
        [AddRecord] [datetime] NULL,
        [LastUpdate] [datetime] NULL,
        [CodiEmprNet] [int] NULL,
        [NumUsuar] [int] NULL,
     CONSTRAINT [PK_TRAZAREC] PRIMARY KEY CLUSTERED 
    (
        [ID] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
    ) ON [PRIMARY]
END
GO

-- =============================================
-- Índices para TRAZAREC
-- =============================================

-- Índice en NROREMI (Número de Remito)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NROREMI' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NROREMI] ON [dbo].[TRAZAREC]
    (
        [NROREMI] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Cod_Inte (Código Interno del Producto)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Cod_Inte' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Cod_Inte] ON [dbo].[TRAZAREC]
    (
        [Cod_Inte] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Nume_Ocom (Número de Orden de Compra)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Nume_Ocom' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Nume_Ocom] ON [dbo].[TRAZAREC]
    (
        [Nume_Ocom] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en NroPed (Número de Pedido)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NroPed' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NroPed] ON [dbo].[TRAZAREC]
    (
        [NroPed] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice compuesto en NROREMI + Cod_Inte (para búsquedas específicas de producto en un remito)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NROREMI_Cod_Inte' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NROREMI_Cod_Inte] ON [dbo].[TRAZAREC]
    (
        [NROREMI] ASC,
        [Cod_Inte] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Fech_Recep (Fecha de Recepción)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Fech_Recep' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Fech_Recep] ON [dbo].[TRAZAREC]
    (
        [Fech_Recep] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en CodiEmpr (Código de Empresa)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_CodiEmpr' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_CodiEmpr] ON [dbo].[TRAZAREC]
    (
        [CodiEmpr] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO
