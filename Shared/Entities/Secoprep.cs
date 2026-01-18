using DocumentFormat.OpenXml.Office2010.Excel;
using IMPLANPROD.Shared.Entities;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad SECOPREP - Secuencia de Operaciones por Reparación
    /// Almacena las operaciones necesarias para cada reparación
    /// </summary>
    [Table("SECOPREP")]
    public class Secoprep : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Código de reparación (vincula con INDIREP)
        /// </summary>
        [Display(Name = "Código Reparación")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        [MaxLength(16, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string CodigoReparacion { get; set; } = null!;

        /// <summary>
        /// Número de operación en la secuencia
        /// </summary>
        [Display(Name = "Número Operación")]
        [Required(ErrorMessage = "El campo {0} es obligatorio.")]
        public short NumeroOperacion { get; set; }

        /// <summary>
        /// Descripción de la operación
        /// </summary>
        [Display(Name = "Descripción")]
        [MaxLength(48, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Observación de la operación (migrado desde OBSOPREP)
        /// </summary>
        [Display(Name = "Observación")]
        [MaxLength(512, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Observacion { get; set; }

        /// <summary>
        /// Estado del registro
        /// </summary>
        public short? Status { get; set; }

        /// <summary>
        /// Horas de la operación (Float MBF)
        /// </summary>
        [Display(Name = "Horas")]
        public double? Horas { get; set; }

        /// <summary>
        /// Equipo (Máquina) - Texto 6 chars
        /// </summary>
        [Display(Name = "Equipo")]
        [MaxLength(6, ErrorMessage = "El campo {0} no puede tener más de {1} caracteres")]
        public string? Equipo { get; set; }

        /// <summary>
        /// Cantidad de Operarios
        /// </summary>
        [Display(Name = "Operarios")]
        public int? CantidadOperarios { get; set; }

        /// <summary>
        /// Categoría
        /// </summary>
        [Display(Name = "Categoría")]
        public int? Categoria { get; set; }
        
        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}
/*CREATE TABLE[dbo].[SECOPREP] (
    [Id][int] IDENTITY(1, 1) NOT NULL,
    [CodigoReparacion] [nvarchar] (16) NOT NULL,
    [NumeroOperacion] [smallint] NOT NULL,
    [Descripcion] [nvarchar] (48) NULL,
    [Status] [smallint] NULL,
    [AddRecord][datetime2](7) NULL,
    [LastUpdate] [datetime2] (7) NULL,
    [CodiEmprNet] [int] NULL,
    [NumUsuar][int] NULL,
 CONSTRAINT[PK_SECOPREP] PRIMARY KEY CLUSTERED 
(
    [Id] ASC
)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON[PRIMARY]
) ON[PRIMARY]
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_SECOPREP_CodigoReparacion_NumeroOperacion] ON[dbo].[SECOPREP]
(
    [CodigoReparacion] ASC,
    [NumeroOperacion] ASC
)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON[PRIMARY]
GO
//AGREGO EL CAMPO OBSERVACION 
ALTER TABLE [dbo].[SECOPREP]
ADD [Observacion] [nvarchar](512) NULL;
GO
--SE AGREGARON MAS CAMPOS
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[SECOPREP]') AND name = 'Horas')
BEGIN
    ALTER TABLE [dbo].[SECOPREP] ADD [Horas] [float] NULL;
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
--  *** SE MODIFICO EL CAMPO DE INT A TEXTO ****
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

*/
