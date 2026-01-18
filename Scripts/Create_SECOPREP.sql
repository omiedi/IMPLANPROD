CREATE TABLE [dbo].[SECOPREP](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [CodigoReparacion] [nvarchar](16) NOT NULL,
    [NumeroOperacion] [smallint] NOT NULL,
    [Descripcion] [nvarchar](48) NULL,
    [Status] [smallint] NULL,
    [AddRecord] [datetime2](7) NULL,
    [LastUpdate] [datetime2](7) NULL,
    [CodiEmprNet] [int] NULL,
    [NumUsuar] [int] NULL,
 CONSTRAINT [PK_SECOPREP] PRIMARY KEY CLUSTERED 
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_SECOPREP_CodigoReparacion_NumeroOperacion] ON [dbo].[SECOPREP]
(
    [CodigoReparacion] ASC,
    [NumeroOperacion] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
