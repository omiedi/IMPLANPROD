using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Representa una operación de fabricación en el sistema
    /// </summary>
    [Table("OPERFAB")]
    public class Operfab : ITimestampEntity, IAuditableEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }                   // Identity Primary Key
        /// <summary>
        /// Identificador de la sub-orden
        /// </summary>
        [MaxLength(12)]
        public string IdSubOr { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public short? CodiEmpr { get; set; }

        /// <summary>
        /// Código interno
        /// </summary>
        public int? Cod_Inte { get; set; }

        /// <summary>
        /// Código externo
        /// </summary>
        [MaxLength(24)]
        public string Cod_Exte { get; set; }

        /// <summary>
        /// Descripción
        /// </summary>
        [MaxLength(64)]
        public string Descrip { get; set; }

        /// <summary>
        /// Número de operación
        /// </summary>
        public short? NumeOper { get; set; }

        /// <summary>
        /// Descripción de la operación
        /// </summary>
        [MaxLength(48)]
        public string DescOper { get; set; }

        /// <summary>
        /// Código de máquina
        /// </summary>
        [MaxLength(6)]
        public string Maquina { get; set; }

        /// <summary>
        /// Descripción de la máquina
        /// </summary>
        [MaxLength(64)]
        public string DescMaq { get; set; }

        /// <summary>
        /// Equipo auxiliar
        /// </summary>
        [MaxLength(18)]
        public string EquiAux { get; set; }

        /// <summary>
        /// Cantidad de operaciones
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantiOps { get; set; }

        /// <summary>
        /// Cantidad programada
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantProg { get; set; }

        /// <summary>
        /// Cantidad real
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantReal { get; set; }

        /// <summary>
        /// Cantidad pendiente
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantPend { get; set; }

        /// <summary>
        /// Tiempo de preparación
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemPrep { get; set; }

        /// <summary>
        /// Tiempo fijo
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemFijo { get; set; }

        /// <summary>
        /// Tiempo variable
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemVari { get; set; }

        /// <summary>
        /// Tiempo de demora
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemDemo { get; set; }

        /// <summary>
        /// Tiempo de máquina/herramienta
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemMHer { get; set; }

        /// <summary>
        /// Tiempo de limpieza
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemLimp { get; set; }

        /// <summary>
        /// Tiempo de otros procesos
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemOtro { get; set; }

        /// <summary>
        /// Tiempo total
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemTota { get; set; }

        /// <summary>
        /// Tiempo de ejecución
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemEjec { get; set; }

        /// <summary>
        /// Tiempo pendiente
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TiemPend { get; set; }

        /// <summary>
        /// Estado de la operación
        /// </summary>
        public short? StatOper { get; set; }

        /// <summary>
        /// Fecha de comienzo del proceso
        /// </summary>
        public DateTime? ComiePro { get; set; }

        /// <summary>
        /// Fecha de término del proceso
        /// </summary>
        public DateTime? TermiPro { get; set; }

        /// <summary>
        /// Fecha de término temporal
        /// </summary>
        public DateTime? TermiTemp { get; set; }

        /// <summary>
        /// Observaciones
        /// </summary>
        [MaxLength(48)]
        public string Observa { get; set; }

        /// <summary>
        /// Historial de batch
        /// </summary>
        public string HoBatchRec { get; set; }

        /// <summary>
        /// Sección del proceso
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? SecciPro { get; set; }

        /// <summary>
        /// Identificador único global
        /// </summary>
        public Guid rowguid { get; set; }

        /// <summary>
        /// Número de orden de operación
        /// </summary>
        public short? NuOrdOpe { get; set; }

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
//Crear el indice id
//ALTER TABLE OPERFAB
//ADD Id INT IDENTITY(1,1) PRIMARY KEY;

//agregar campo numusuario
//ALTER TABLE OPERFAB
//ADD NumUsuar  INT NULL;

/*
 * INDICES
 * CREATE UNIQUE NONCLUSTERED INDEX UX_Ordefabrs_IdSubOr
ON dbo.Ordefabr (IdSubOr);

CREATE UNIQUE NONCLUSTERED INDEX UX_Ordefabr_NumeOrFa
ON dbo.Ordefabr (NumeOrFa);

 * -- =====================================================
-- Script para actualizar tipos de datos en tabla OPERFAB
-- Cambiar campos FLOAT a DECIMAL(18,4) según entidad C#
-- =====================================================

USE [TU_BASE_DE_DATOS]
GO

-- IMPORTANTE: Hacer backup de la tabla antes de ejecutar
-- SELECT * INTO OPERFAB_BACKUP FROM OPERFAB
-- GO

BEGIN TRANSACTION;

-- Modificar campos de cantidad
ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [CantiOps] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [CantProg] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [CantReal] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [CantPend] DECIMAL(18,4) NULL;

-- Modificar campos de tiempo
ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemPrep] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemFijo] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemVari] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemDemo] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemMHer] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemLimp] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemOtro] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemTota] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemEjec] DECIMAL(18,4) NULL;

ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [TiemPend] DECIMAL(18,4) NULL;

-- Modificar campo de sección
ALTER TABLE [dbo].[OPERFAB] 
ALTER COLUMN [SecciPro] DECIMAL(18,4) NULL;

-- Si todo está correcto, confirmar la transacción
COMMIT TRANSACTION;

-- Verificar los cambios
SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    NUMERIC_PRECISION,
    NUMERIC_SCALE,
    IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'OPERFAB'
    AND COLUMN_NAME IN (
        'CantiOps', 'CantProg', 'CantReal', 'CantPend',
        'TiemPrep', 'TiemFijo', 'TiemVari', 'TiemDemo',
        'TiemMHer', 'TiemLimp', 'TiemOtro', 'TiemTota',
        'TiemEjec', 'TiemPend', 'SecciPro'
    )
ORDER BY COLUMN_NAME;

PRINT 'Actualización completada exitosamente';
GO
*/