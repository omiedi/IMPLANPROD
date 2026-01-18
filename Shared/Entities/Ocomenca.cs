using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa el encabezado de órdenes de compra
    /// </summary>
    [Table("OCOMENCA")]
    public class Ocomenca : IAuditableEntity
    {
        /// <summary>
        /// ID único de la orden de compra (campo identity)
        /// </summary>
        [Key]
        [Column("OCOMENCA_ID")]
        public int OCOMENCA_ID { get; set; }

        /// <summary>
        /// Atributo de serie de la orden de compra
        /// </summary>
        [Column("AtSr_Ocom")]
        [MaxLength(64)]
        public string? AtSrOcom { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Observaciones de la orden de compra
        /// </summary>
        [Column("Obse_Ocom")]
        public string? ObseOcom { get; set; }

        /// <summary>
        /// Lugar de entrega
        /// </summary>
        [Column("Lentrega")]
        [MaxLength(128)]
        public string? Lentrega { get; set; }

        /// <summary>
        /// Anexo de la orden de compra
        /// </summary>
        [Column("Anex_Ocom")]
        [MaxLength(128)]
        public string? AnexOcom { get; set; }

        /// <summary>
        /// Revisión de la orden de compra
        /// </summary>
        [Column("OcomRevis")]
        public short? OcomRevis { get; set; }

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        [Column("Nume_Ocom")]
        public int? NumeOcom { get; set; }

        /// <summary>
        /// Número de pedido relacionado
        /// </summary>
        [Column("Nume_Pedi")]
        public int? NumePedi { get; set; }

        /// <summary>
        /// Referencia de la orden de compra
        /// </summary>
        [Column("Refe_Ocom")]
        [MaxLength(64)]
        public string? RefeOcom { get; set; }

        /// <summary>
        /// Fecha de emisión de la orden de compra
        /// </summary>
        [Column("Femi_Ocom")]
        public DateTime? FemiOcom { get; set; }

        /// <summary>
        /// Número de proveedor
        /// </summary>
        [Column("Npro_Ocom")]
        public int? NproOcom { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        [Column("Nume_Cli")]
        public int? NumeCli { get; set; }

        /// <summary>
        /// Razón social del proveedor
        /// </summary>
        [Column("RaSo_Prov")]
        [MaxLength(48)]
        public string? RaSoProv { get; set; }

        /// <summary>
        /// Total sumado de la orden de compra
        /// </summary>
        [Column("TSum_Ocom")]
        public short? TSumOcom { get; set; }

        /// <summary>
        /// Estado de la orden de compra
        /// </summary>
        [Column("Stat_Ocom")]
        public short? StatOcom { get; set; }

        /// <summary>
        /// Fecha de generación
        /// </summary>
        [Column("FGen_Ocom")]
        public DateTime? FGenOcom { get; set; }

        /// <summary>
        /// Usuario que generó la orden
        /// </summary>
        [Column("Usua_Gen")]
        public short? UsuaGen { get; set; }

        /// <summary>
        /// Fecha de aprobación
        /// </summary>
        [Column("FApro_Ocom")]
        public DateTime? FAproOcom { get; set; }

        /// <summary>
        /// Usuario que aprobó la orden
        /// </summary>
        [Column("Usua_Apro")]
        public int? UsuaApro { get; set; }

        /// <summary>
        /// Fecha de modificación
        /// </summary>
        [Column("FMod_Ocom")]
        public DateTime? FModOcom { get; set; }

        /// <summary>
        /// Usuario que modificó la orden
        /// </summary>
        [Column("Usua_Mod")]
        public int? UsuaMod { get; set; }

        /// <summary>
        /// Fecha de anulación
        /// </summary>
        [Column("FAnu_Ocom")]
        public DateTime? FAnuOcom { get; set; }

        /// <summary>
        /// Usuario que anuló la orden
        /// </summary>
        [Column("Usua_Anul")]
        public int? UsuaAnul { get; set; }

        /// <summary>
        /// Motivo de anulación
        /// </summary>
        [Column("Moti_Anul")]
        [MaxLength(16)]
        public string? MotiAnul { get; set; }

        /// <summary>
        /// Moneda de costo
        /// </summary>
        [Column("Mone_Cos")]
        public short? MoneCos { get; set; }

        /// <summary>
        /// Moneda de emisión
        /// </summary>
        [Column("Mone_Emis")]
        public short? MoneEmis { get; set; }

        /// <summary>
        /// Tipo de cambio
        /// </summary>
        [Column("Tcam_Ocom", TypeName = "decimal(18,4)")]
        public decimal? TcamOcom { get; set; }

        /// <summary>
        /// Total bruto de la orden de compra
        /// </summary>
        [Column("ToBru_Ocom", TypeName = "decimal(18,4)")]
        public decimal? ToBruOcom { get; set; }

        /// <summary>
        /// Porcentaje de descuento
        /// </summary>
        [Column("Pdes_Ocom")]
        [MaxLength(24)]
        public string? PdesOcom { get; set; }

        /// <summary>
        /// Importe de descuento
        /// </summary>
        [Column("Ides_Ocom", TypeName = "decimal(18,4)")]
        public decimal? IdesOcom { get; set; }

        /// <summary>
        /// Importe de IVA
        /// </summary>
        [Column("Iiva_Ocom", TypeName = "decimal(18,4)")]
        public decimal? IivaOcom { get; set; }

        /// <summary>
        /// Importe otros 1
        /// </summary>
        [Column("Iotr1_Ocom", TypeName = "decimal(18,4)")]
        public decimal? Iotr1Ocom { get; set; }

        /// <summary>
        /// Importe otros 2
        /// </summary>
        [Column("Iotr2_Ocom", TypeName = "decimal(18,4)")]
        public decimal? Iotr2Ocom { get; set; }

        /// <summary>
        /// Total de la orden de compra
        /// </summary>
        [Column("Total_Ocom", TypeName = "decimal(18,4)")]
        public decimal? TotalOcom { get; set; }

        /// <summary>
        /// Condición de pago
        /// </summary>
        [Column("Cpag_Ocom")]
        public short? CpagOcom { get; set; }

        /// <summary>
        /// Embalaje de la orden de compra
        /// </summary>
        [Column("Embala_Ocom")]
        public short? EmbalaOcom { get; set; }

        /// <summary>
        /// Porcentaje de IVA
        /// </summary>
        [Column("PorcIva")]
        [MaxLength(16)]
        public string? PorcIva { get; set; }

        /// <summary>
        /// ID del comprobante
        /// </summary>
        [Column("Id_Comprob")]
        [MaxLength(16)]
        public string? IdComprob { get; set; }

        /// <summary>
        /// Código de carpeta de impresión
        /// </summary>
        [Column("CodCarpImp")]
        [MaxLength(12)]
        public string? CodCarpImp { get; set; }

        /// <summary>
        /// GUID único del registro
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();

        /// <summary>
        /// Responsable de la orden de compra
        /// </summary>
        [Column("Respon_Ocompra")]
        [MaxLength(128)]
        public string? ResponOcompra { get; set; }

        /// <summary>
        /// Vía de entrega
        /// </summary>
        [Column("Via_Entrega")]
        public short? ViaEntrega { get; set; }

        /// <summary>
        /// Cotización del dólar
        /// </summary>
        [Column("COTIZA_DOLAR", TypeName = "decimal(18,4)")]
        public decimal? CotizaDolar { get; set; }

        /// <summary>
        /// Orden de compra del cliente
        /// </summary>
        [Column("OComCliente")]
        [MaxLength(16)]
        public string? OComCliente { get; set; }

        /// <summary>
        /// ID de comprobante externo
        /// </summary>
        [Column("IdCompExt")]
        public int? IdCompExt { get; set; }

        /// <summary>
        /// Código de empresa .NET
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

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
    }
}

/*
-- =====================================================
-- INSTRUCCIONES SQL PARA MODIFICAR LA TABLA OCOMENCA
-- =====================================================

-- 1. CAMBIAR CAMPOS DE SMALLINT A INT
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [CodiEmpr] INT NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Npro_Ocom] INT NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Usua_Apro] INT NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Usua_Mod] INT NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Usua_Anul] INT NULL;

-- 2. AGREGAR NUEVOS CAMPOS
ALTER TABLE [dbo].[OCOMENCA] ADD [CodiEmprNet] INT NULL;
ALTER TABLE [dbo].[OCOMENCA] ADD [NumUsuar] INT NULL;
ALTER TABLE [dbo].[OCOMENCA] ADD [AddRecord] DATETIME2 NOT NULL DEFAULT GETDATE();
ALTER TABLE [dbo].[OCOMENCA] ADD [LastUpdate] DATETIME2 NOT NULL DEFAULT GETDATE();

-- 3. CAMBIAR CAMPOS DE FLOAT A DECIMAL(18,4)
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Tcam_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [ToBru_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Ides_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Iiva_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Iotr1_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Iotr2_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [Total_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMENCA] ALTER COLUMN [COTIZA_DOLAR] DECIMAL(18,4) NULL;

-- 4. VERIFICAR Y ASEGURAR QUE EXISTE EL CAMPO FGEN_OCOM
-- (Este campo debería existir en la tabla original, pero verificamos por seguridad)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('[dbo].[OCOMENCA]') AND name = 'FGen_Ocom')
BEGIN
    ALTER TABLE [dbo].[OCOMENCA] ADD [FGen_Ocom] DATETIME2 NULL;
    PRINT 'Campo FGen_Ocom agregado a la tabla OCOMENCA';
END
ELSE
BEGIN
    PRINT 'Campo FGen_Ocom ya existe en la tabla OCOMENCA';
END

-- NOTA: Ejecutar estos comandos en SQL Server Management Studio
-- para aplicar los cambios a la estructura de la tabla existente.
*/
//CREATE INDEX IX_OCOMENCA_NUME_OCOM ON OCOMENCA (NUME_OCOM);

//--Ejemplo: si se consulta frecuentemente por NUME_OCOM y STAT_OCOM
//CREATE INDEX IX_OCOMENCA_NUME_OCOM_STAT ON OCOMENCA (NUME_OCOM, STAT_OCOM);