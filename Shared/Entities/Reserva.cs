using DocumentFormat.OpenXml.Wordprocessing;

using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.Entities
{
    [Table("Reserva")]
    public class Reserva : ITimestampEntity, IAuditableEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Reserva_Id { get; set; }  // No acepta null, clave primaria

        public int Cod_Inte { get; set; }
        public string? Cod_Exte { get; set; }
        public string? Descrip { get; set; }
        public string? IdSubOr { get; set; }
        public short? NuOrItem { get; set; }
        public short? Status { get; set; }
        public DateTime? FechRes { get; set; }
        public decimal? CantRes { get; set; }
        public decimal? CantCons { get; set; }
        public decimal? Saldo_Entre { get; set; }
        public string? IdenLote { get; set; }
        public short? Deposito { get; set; }
        public string? UbicaLot { get; set; }
        public decimal? UsoUnit { get; set; }
        public string? Observa { get; set; }
        public string? PNL_Des { get; set; }
        public string? PLC_Has { get; set; }
        public short? CodiEmpr { get; set; }
        public int? Nume_Solint { get; set; }
        public int? Nume_Ocom { get; set; }
        public short? Marborra { get; set; }
        public string? NUMEROSERIE { get; set; }

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
// NOTA: SE MANTUVO COMO ID EL CAMPO RESERVA_ID

//ALTER TABLE[dbo].[RESERVA] ALTER COLUMN[CantRes] decimal(18,4) NULL;
//ALTER TABLE[dbo].[RESERVA] ALTER COLUMN[CantCons] decimal(18,4) NULL;
//ALTER TABLE[dbo].[RESERVA] ALTER COLUMN[Saldo_Entre] decimal(18,4) NULL;
//ALTER TABLE[dbo].[RESERVA] ALTER COLUMN[UsoUnit] decimal(18,4) NULL;

//ALTER TABLE[dbo].[RESERVA] ALTER COLUMN[Cod_Inte] int NULL;

//ALTER TABLE[dbo].[RESERVA] ALTER COLUMN[Cod_Exte] nvarchar(24) NULL;

//ALTER TABLE RESERVA
//DROP COLUMN [rowguid] ;