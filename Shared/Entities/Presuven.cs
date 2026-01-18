using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa un presupuesto de venta en el sistema.
    /// Mapea a la tabla PRESUVEN en la base de datos.
    /// </summary>
    [Table("PRESUVEN")]
    public class Presuven
    {
        /// <summary>
        /// Identificador único del presupuesto de venta (campo IDENTITY)
        /// </summary>
        [Key]
        [Column("PRESUVEN_ID")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int PresuvenId { get; set; }

        /// <summary>
        /// Número de presupuesto
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? NroPres { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Código de rubro
        /// </summary>
        [StringLength(8)]
        public string? Cod_Rubro { get; set; }

        /// <summary>
        /// Número de código de rubro
        /// </summary>
        public int? NumeCRubro { get; set; }

        /// <summary>
        /// Código de año
        /// </summary>
        [StringLength(4)]
        public string? Cod_Anio { get; set; }

        /// <summary>
        /// Fecha de emisión
        /// </summary>
        public DateTime? FechEmis { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        public int? NroClie { get; set; }

        /// <summary>
        /// Código de cliente
        /// </summary>
        [StringLength(12)]
        public string? Codi_Cli { get; set; }

        /// <summary>
        /// Razón social del cliente
        /// </summary>
        [StringLength(64)]
        public string? RasoClie { get; set; }

        /// <summary>
        /// Referencia
        /// </summary>
        [StringLength(64)]
        public string? Referen { get; set; }

        /// <summary>
        /// Atención señor
        /// </summary>
        [StringLength(48)]
        public string? AteSen { get; set; }

        /// <summary>
        /// Lista de precios
        /// </summary>
        public short? LisPre { get; set; }

        /// <summary>
        /// Condición de pago
        /// </summary>
        public short? ConPag { get; set; }

        /// <summary>
        /// Vendedor
        /// </summary>
        public short? Vended { get; set; }

        /// <summary>
        /// Moneda de emisión
        /// </summary>
        public short? MonedEmi { get; set; }

        /// <summary>
        /// Número de orden item
        /// </summary>
        public short? NuOrdIt { get; set; }

        /// <summary>
        /// Código interno
        /// </summary>
        public int? CodiInt { get; set; }

        /// <summary>
        /// Código externo
        /// </summary>
        [StringLength(24)]
        public string? CodiExt { get; set; }

        /// <summary>
        /// Descripción del producto
        /// </summary>
        [StringLength(64)]
        public string? Descrip { get; set; }

        /// <summary>
        /// Especificaciones
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? Especif { get; set; }

        /// <summary>
        /// Precio de lista
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PreList { get; set; }

        /// <summary>
        /// Tipo de descuento
        /// </summary>
        [StringLength(32)]
        public string? TDescuen { get; set; }

        /// <summary>
        /// Porcentaje de descuento
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PorcDesc { get; set; }

        /// <summary>
        /// Importe de descuento
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? ImpoDesc { get; set; }

        /// <summary>
        /// Precio unitario
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PreUnit { get; set; }

        /// <summary>
        /// Cantidad
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? Cantit { get; set; }

        /// <summary>
        /// Importe total neto
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? ImToNe { get; set; }

        /// <summary>
        /// Fecha de vencimiento
        /// </summary>
        public DateTime? FechVent { get; set; }

        /// <summary>
        /// Número de pedido
        /// </summary>
        public int? NroPedi { get; set; }
        

        /// <summary>
        /// Número de factura corto
        /// </summary>
        [StringLength(12)]
        public string? NroFac_Cor { get; set; }

        /// <summary>
        /// Número de factura largo
        /// </summary>
        [StringLength(24)]
        public string? NroFac_Lar { get; set; }

        /// <summary>
        /// Marca de no vigente
        /// </summary>
        [StringLength(1)]
        public string? MarNoVig { get; set; }

        /// <summary>
        /// Código base
        /// </summary>
        [StringLength(12)]
        public string? CBase { get; set; }

        /// <summary>
        /// Tipo de talle
        /// </summary>
        [StringLength(12)]
        public string? TTalle { get; set; }

        /// <summary>
        /// Tipo de color
        /// </summary>
        [StringLength(16)]
        public string? TColor { get; set; }

        /// <summary>
        /// Número de cotización
        /// </summary>
        [StringLength(16)]
        public string? NroCoti { get; set; }

        /// <summary>
        /// Reventa
        /// </summary>
        [StringLength(48)]
        public string? Reventa { get; set; }

        /// <summary>
        /// Antecedentes
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? Antecedentes { get; set; }

        /// <summary>
        /// Observaciones
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? Observa { get; set; }

        /// <summary>
        /// Imagen
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? Imagen { get; set; }

        /// <summary>
        /// Estado del presupuesto
        /// </summary>
        public short? Status { get; set; }

        /// <summary>
        /// Grupo de cliente
        /// </summary>
        public short? Grco_Cli { get; set; }

        /// <summary>
        /// Cantidad de bultos
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? CantBultos { get; set; }

        /// <summary>
        /// Peso neto
        /// </summary>
        [StringLength(128)]
        public string? PesoNeto { get; set; }

        /// <summary>
        /// Peso total
        /// </summary>
        [StringLength(128)]
        public string? PesoTotal { get; set; }

        /// <summary>
        /// Nomenclatura Común del Mercosur
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? NCM { get; set; }

        /// <summary>
        /// Condiciones de venta
        /// </summary>
        [StringLength(128)]
        public string? CondVta { get; set; }

        /// <summary>
        /// Identificador único global de la fila
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();

        /// <summary>
        /// Versión del presupuesto
        /// </summary>
        public short? Vers_Presu { get; set; }

        /// <summary>
        /// Número de proveedor
        /// </summary>
        public int? Num_Prov { get; set; }

        /// <summary>
        /// Costo unitario
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? COSTO_UNIT { get; set; }

        /// <summary>
        /// Vía de entrega
        /// </summary>
        public short? Via_Entrega { get; set; }

        /// <summary>
        /// Marca de borrado
        /// </summary>
        public short? MARBORRA { get; set; }

        /// <summary>
        /// Cotización de moneda
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? COTIZAMONEDA { get; set; }

        /// <summary>
        /// Lotes reservados
        /// </summary>
        [StringLength(512)]
        public string? LOTES_Reservados { get; set; }

        /// <summary>
        /// Cantidad de lotes
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? Canti_Lotes { get; set; }

        /// <summary>
        /// Total IVA 21%
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TTIVA21 { get; set; }

        /// <summary>
        /// Total IVA 10.5%
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TTIVA105 { get; set; }

        /// <summary>
        /// Total percepciones
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TOTPER { get; set; }

        /// <summary>
        /// Email del cliente
        /// </summary>
        [StringLength(64)]
        public string? Mail_Cli { get; set; }

        /// <summary>
        /// Costo FOB
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? COSTOFOB { get; set; }

        /// <summary>
        /// Presupuesto comercial real
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PresComReal { get; set; }
    }
}

/*
-- Instrucciones SQL para modificar la tabla PRESUVEN según los cambios realizados en la entidad:

-- 1. Cambiar CodiEmpr de smallint a int
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [CodiEmpr] int;

-- 2. Cambiar CodiInt de smallint a int
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [CodiInt] int;

-- 3. Cambiar todos los campos float a decimal(18,4)
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [NroPres] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [PreList] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [PorcDesc] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [ImpoDesc] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [PreUnit] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [Cantit] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [ImToNe] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [NroPedi] int;
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [COSTO_UNIT] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [COTIZAMONEDA] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [Canti_Lotes] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [TTIVA21] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [TTIVA105] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [TOTPER] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [COSTOFOB] decimal(18,4);
ALTER TABLE [dbo].[PRESUVEN] ALTER COLUMN [PresComReal] decimal(18,4);

-- 4. Verificar y mantener campo rowguid
-- NOTA: El campo rowguid debe mantenerse en la tabla
IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'PRESUVEN' AND COLUMN_NAME = 'rowguid')
BEGIN
    ALTER TABLE [dbo].[PRESUVEN] ADD [rowguid] UNIQUEIDENTIFIER ROWGUIDCOL NOT NULL DEFAULT NEWID();
END

-- NOTA: El campo PRESUVEN_ID ya es IDENTITY en la tabla original, por lo que no requiere modificación.
*/
