using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("Ordefabr")]


    public class OrdeFabr : ITimestampEntity, IAuditableEntity
    {
        public decimal? CanIng;
        public DateTime? FechMov;

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]

        public int Id { get; set; }

        public int? NumeOrFa { get; set; }
        public short? CodiEmpr { get; set; }
        public string? Referen { get; set; }
        public string? IdSubOr { get; set; }
        public DateTime? FechGen { get; set; }
        public int? Cod_Inte { get; set; }
        public string? Cod_Exte { get; set; }
        public string? Descrip { get; set; }
        public string? Unidad { get; set; }
        public decimal? CanProy { get; set; } // decimal(18,4)
        public string? Destino { get; set; }
        public DateTime? FechEnSol { get; set; }
        public short? Priorid { get; set; }
        public short? StatuOrf { get; set; }
        public string? NuOrProd { get; set; }
        public DateTime? FechLan { get; set; }
        public decimal? SecciPro { get; set; } // decimal(18,4)
        public DateTime? FechCierr { get; set; }
        public decimal? CanApro { get; set; } // decimal(18,4)
        public decimal? CanRech { get; set; } // decimal(18,4)
        public string? IdenRetra { get; set; }
        public DateTime? FechProyEntre { get; set; }
        public DateTime? FechTempEntre { get; set; }
        public DateTime? FechRecal { get; set; }
        public string? Observa { get; set; }
        public string? Origen { get; set; }
        public string? AnotaOrf { get; set; }
        public string? IdenLote { get; set; }
        public string? NumRem_Prov { get; set; }
        public string? Nume_Ocom { get; set; }
        public int? Nume_Cli { get; set; }
        public string? Id_Orfa { get; set; }
        public int? NroPed { get; set; }
        public int? Armador { get; set; }
        //public Guid RowGuid { get; set; }
        public decimal? LoteIndivis { get; set; } // decimal(18,4)
        public DateTime? Inicio_Informado { get; set; }
        public decimal? Canti_Fin_Proceso { get; set; } // decimal(18,4)
        public short? Status_Proceso { get; set; }
        public decimal? Canti_Iniciada { get; set; } // decimal(18,4)
        public DateTime? FinDeProceso { get; set; }
        public string? NumeroChasis { get; set; }
        public string? TipoChasis { get; set; }
        public string? LineaInterno { get; set; }
        public DateTime? FechEntMat { get; set; }
        public DateTime? DespachoFabrica { get; set; }
        public short? NuOrNroPed { get; set; }
        public int? ProdEstandar { get; set; }
        public string? Anexos { get; set; }
        public short? CantHilos { get; set; }
        public decimal? PesoBolsa { get; set; } // decimal(18,4)

        [NotMapped] // Esta propiedad no se mapeará a la base de datos
        public decimal? LoteMaximo { get; set; }

        [NotMapped] // Componentes editados por el usuario en el alta, enviados desde el cliente
        public List<IMPLANPROD.Shared.DTOs.FormulaComponenteDTO>? Componentes { get; set; }

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

//ALTER TABLE[dbo].[ORDEFABR] ALTER COLUMN[CanProy] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORDEFABR] ALTER COLUMN[SecciPro] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORDEFABR] ALTER COLUMN[CanApro] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORDEFABR] ALTER COLUMN[CanRech] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORDEFABR] ALTER COLUMN[LoteIndivis] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORDEFABR] ALTER COLUMN[Canti_fin_proceso] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORDEFABR] ALTER COLUMN[Canti_Iniciada] decimal(18,4) NULL;
//ALTER TABLE[dbo].[ORDEFABR] ALTER COLUMN[PesoBolsa] decimal(18,4) NULL;
//
//ALTER TABLE[dbo].[ORDEFABR]
//ADD Id INT IDENTITY(1,1) NOT NULL;
//ALTER TABLE[dbo].[ORDEFABR]
//ADD CONSTRAINT PK_ORDEFABR_Id PRIMARY KEY (Id);
//
//ELIMINAR ROWGUID
//ALTER TABLE[dbo].[ORDEFABR]
//DROP COLUMN rowguid;

//
//ALTER TABLE Ordefabr
//ALTER COLUMN Cod_inte INT;
//
//AGREGAR EL CAMPO NumUsuar
//ALTER TABLE ordefabr
//ADD NumUsuar INT NULL;



