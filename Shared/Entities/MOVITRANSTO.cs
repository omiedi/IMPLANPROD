using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa la tabla MOVITRANSTO para movimientos de transferencia de stock
    /// </summary>
    [Table("MOVITRANSTO")]
    public class MOVITRANSTO : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único del registro (clave primaria identity)
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("Id")]
        public int Id { get; set; }

        /// <summary>
        /// Fecha de registro del movimiento
        /// </summary>
        [Column("FeReMovi")]
        public DateTime? FeReMovi { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public short? CodiEmpr { get; set; }

        /// <summary>
        /// Identificador del movimiento
        /// </summary>
        [Column("IdenMovi")]
        [StringLength(32)]
        public string? IdenMovi { get; set; }

        /// <summary>
        /// Descripción larga del remitente
        /// </summary>
        [Column("DeLaRemi")]
        public string? DeLaRemi { get; set; }

        /// <summary>
        /// Descripción larga del movimiento
        /// </summary>
        [Column("Descri_Lar")]
        public string? Descri_Lar { get; set; }

        /// <summary>
        /// Fecha del movimiento
        /// </summary>
        [Column("FechMov")]
        public DateTime? FechMov { get; set; }

        /// <summary>
        /// Código interno del producto (cambiado de smallint a int)
        /// </summary>
        [Column("Cod_Inte")]
        public int? Cod_Inte { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        [Column("Cod_Exte")]
        [StringLength(24)]
        public string? Cod_Exte { get; set; }

        /// <summary>
        /// Descripción del producto
        /// </summary>
        [Column("Descrip")]
        [StringLength(64)]
        public string? Descrip { get; set; }

        /// <summary>
        /// Identificador del lote
        /// </summary>
        [Column("IdenLote")]
        [StringLength(16)]
        public string? IdenLote { get; set; }

        /// <summary>
        /// Lote de consumo
        /// </summary>
        [Column("LotConsum")]
        public string? LotConsum { get; set; }

        /// <summary>
        /// Posición desde
        /// </summary>
        [Column("PNL_Des")]
        [StringLength(32)]
        public string? PNL_Des { get; set; }

        /// <summary>
        /// Posición hasta
        /// </summary>
        [Column("PLC_Has")]
        [StringLength(32)]
        public string? PLC_Has { get; set; }

        /// <summary>
        /// Código de movimiento
        /// </summary>
        [Column("CodiMovi")]
        [StringLength(2)]
        public string? CodiMovi { get; set; }

        /// <summary>
        /// Documento primario corto
        /// </summary>
        [Column("DoPriCor")]
        [StringLength(12)]
        public string? DoPriCor { get; set; }

        /// <summary>
        /// Documento primario largo
        /// </summary>
        [Column("DoPriLar")]
        [StringLength(24)]
        public string? DoPriLar { get; set; }

        /// <summary>
        /// Documento secundario corto
        /// </summary>
        [Column("DoSecCor")]
        [StringLength(12)]
        public string? DoSecCor { get; set; }

        /// <summary>
        /// Documento secundario largo
        /// </summary>
        [Column("DoSecLar")]
        [StringLength(24)]
        public string? DoSecLar { get; set; }

        /// <summary>
        /// Documento original
        /// </summary>
        [Column("DocuOrig")]
        [StringLength(24)]
        public string? DocuOrig { get; set; }

        /// <summary>
        /// Número de remito del cliente
        /// </summary>
        [Column("NumRem_Cli")]
        [StringLength(16)]
        public string? NumRem_Cli { get; set; }

        /// <summary>
        /// Número de orden ítem
        /// </summary>
        [Column("NuorItem")]
        public short? NuorItem { get; set; }

        /// <summary>
        /// Número de pedido del cliente
        /// </summary>
        [Column("NuPedCli")]
        public int? NuPedCli { get; set; }

        /// <summary>
        /// Número de orden de pedido
        /// </summary>
        [Column("NumeOrdPed")]
        public short? NumeOrdPed { get; set; }

        /// <summary>
        /// Referencia corta
        /// </summary>
        [Column("ReferCor")]
        [StringLength(64)]
        public string? ReferCor { get; set; }

        /// <summary>
        /// Costo de reposición (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("CostoRep", TypeName = "decimal(18,4)")]
        public decimal? CostoRep { get; set; }

        /// <summary>
        /// Valor unitario (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("ValoUnit", TypeName = "decimal(18,4)")]
        public decimal? ValoUnit { get; set; }

        /// <summary>
        /// Valor unitario con IVA (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("ValoUnitIva", TypeName = "decimal(18,4)")]
        public decimal? ValoUnitIva { get; set; }

        /// <summary>
        /// Valor total del ítem (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("ValTotIt", TypeName = "decimal(18,4)")]
        public decimal? ValTotIt { get; set; }

        /// <summary>
        /// Valor total del ítem con IVA (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("ValTotItIva", TypeName = "decimal(18,4)")]
        public decimal? ValTotItIva { get; set; }

        /// <summary>
        /// Moneda del valor
        /// </summary>
        [Column("MoneVal")]
        public short? MoneVal { get; set; }

        /// <summary>
        /// Tipo de cambio (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("TipoCam", TypeName = "decimal(18,4)")]
        public decimal? TipoCam { get; set; }

        /// <summary>
        /// Depósito del movimiento (cambiado de smallint a int)
        /// </summary>
        [Column("DepoMovi")]
        public int? DepoMovi { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        [Column("Nume_Clie")]
        public int? Nume_Clie { get; set; }

        /// <summary>
        /// Sucursal del cliente
        /// </summary>
        [Column("Sucu_Clie")]
        public short? Sucu_Clie { get; set; }

        /// <summary>
        /// Número de proveedor (cambiado de smallint a int)
        /// </summary>
        [Column("Nume_Prov")]
        public int? Nume_Prov { get; set; }

        /// <summary>
        /// Cantidad de ingreso (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("CantIngre", TypeName = "decimal(18,4)")]
        public decimal? CantIngre { get; set; }

        /// <summary>
        /// Cantidad de salida (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("CantSalid", TypeName = "decimal(18,4)")]
        public decimal? CantSalid { get; set; }

        /// <summary>
        /// Cantidad de saldo (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("CantSaldo", TypeName = "decimal(18,4)")]
        public decimal? CantSaldo { get; set; }

        /// <summary>
        /// Estado de verificación
        /// </summary>
        [Column("Verificado")]
        public short? Verificado { get; set; }

        /// <summary>
        /// Causa del movimiento
        /// </summary>
        [Column("Causa")]
        public short? Causa { get; set; }

        /// <summary>
        /// Usuario de la transacción
        /// </summary>
        [Column("User_Tran")]
        [StringLength(32)]
        public string? User_Tran { get; set; }

        /// <summary>
        /// Código de carpeta de impresión
        /// </summary>
        [Column("CodCarpImp")]
        [StringLength(12)]
        public string? CodCarpImp { get; set; }

        /// <summary>
        /// Verificado por
        /// </summary>
        [Column("VerificadoPor")]
        public short? VerificadoPor { get; set; }

        /// <summary>
        /// Fecha y hora de verificación
        /// </summary>
        [Column("FHor_Verif")]
        public DateTime? FHor_Verif { get; set; }

        /// <summary>
        /// Tipo de transacción
        /// </summary>
        [Column("TipoTran")]
        public short? TipoTran { get; set; }

        /// <summary>
        /// Orden de compra retorno abierto
        /// </summary>
        [Column("OCOMPRA_RTOABIERTO")]
        [StringLength(64)]
        public string? OCOMPRA_RTOABIERTO { get; set; }

        /// <summary>
        /// Número de envío
        /// </summary>
        [Column("NroEnvio")]
        public int? NroEnvio { get; set; }

        /// <summary>
        /// Número de punto de venta
        /// </summary>
        [Column("NUMPVTA")]
        public short? NUMPVTA { get; set; }

        /// <summary>
        /// Documento primario número
        /// </summary>
        [Column("DOPRINUM")]
        public int? DOPRINUM { get; set; }

        /// <summary>
        /// ID de calidad
        /// </summary>
        [Column("IdCalidad")]
        public int? IdCalidad { get; set; }

        /// <summary>
        /// Orden de registro del movimiento
        /// </summary>
        [Column("OrdenRegistroMovimiento")]
        public short? OrdenRegistroMovimiento { get; set; }

        /// <summary>
        /// Cantidad rechazada (cambiado de float a decimal con 4 decimales)
        /// </summary>
        [Column("CantidadRechazada", TypeName = "decimal(18,4)")]
        public decimal? CantidadRechazada { get; set; }

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

/*
 * INSTRUCCIONES SQL PARA MODIFICAR LOS CAMPOS EN LA BASE DE DATOS:
 * 
 * -- Agregar campo Id identity como clave primaria
 * ALTER TABLE [dbo].[MOVITRANSTO] ADD [Id] INT IDENTITY(1,1) NOT NULL;
 * ALTER TABLE [dbo].[MOVITRANSTO] ADD CONSTRAINT PK_MOVITRANSTO PRIMARY KEY ([Id]);
 * 
 * -- Cambiar NumUsuar de smallint a int
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [NumUsuar] INT NULL;
 * 
 * -- Cambiar Cod_Inte de smallint a int
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [Cod_Inte] INT NULL;
 * 
 * -- Cambiar DepoMovi de smallint a int
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [DepoMovi] INT NULL;
 * 
 * -- Cambiar Nume_Prov de smallint a int
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [Nume_Prov] INT NULL;
 * 
 * -- Cambiar CostoRep de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [CostoRep] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar ValoUnit de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [ValoUnit] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar ValoUnitIva de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [ValoUnitIva] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar ValTotIt de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [ValTotIt] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar ValTotItIva de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [ValTotItIva] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar TipoCam de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [TipoCam] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar CantIngre de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [CantIngre] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar CantSalid de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [CantSalid] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar CantSaldo de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [CantSaldo] DECIMAL(18,4) NULL;
 * 
 * -- Cambiar CantidadRechazada de float a decimal(18,4)
 * ALTER TABLE [dbo].[MOVITRANSTO] ALTER COLUMN [CantidadRechazada] DECIMAL(18,4) NULL;
 */
