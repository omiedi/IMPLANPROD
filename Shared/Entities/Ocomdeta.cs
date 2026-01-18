using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa el detalle de órdenes de compra
    /// </summary>
    [Table("OCOMDETA")]
    public class Ocomdeta
    {
        /// <summary>
        /// ID único del detalle de orden de compra (campo identity)
        /// </summary>
        [Key]
        [Column("OCOMDETA_ID")]
        public int OCOMDETA_ID { get; set; }

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        [Column("Nume_Ocom")]
        public int? NumeOcom { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Número de pedido relacionado
        /// </summary>
        [Column("Nume_Pedi")]
        public int? NumePedi { get; set; }

        /// <summary>
        /// Número de orden de item
        /// </summary>
        [Column("NuOrd_Item")]
        public short? NuOrdItem { get; set; }

        /// <summary>
        /// Fecha de emisión de la orden de compra
        /// </summary>
        [Column("Femi_Ocom")]
        public DateTime? FemiOcom { get; set; }

        /// <summary>
        /// Número de proveedor
        /// </summary>
        [Column("Npro_Ocom")]
        public short? NproOcom { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        [Column("Nume_Cli")]
        public int? NumeCli { get; set; }

        /// <summary>
        /// Razón social del proveedor
        /// </summary>
        [Column("RaSo_Prov")]
        [MaxLength(64)]
        public string? RaSoProv { get; set; }

        /// <summary>
        /// Código interno del producto
        /// </summary>
        [Column("Cod_Inte")]
        public short? CodInte { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        [Column("Cod_Exte")]
        [MaxLength(24)]
        public string? CodExte { get; set; }

        /// <summary>
        /// Descripción del item
        /// </summary>
        [Column("Descritem")]
        [MaxLength(64)]
        public string? Descritem { get; set; }

        /// <summary>
        /// Descripción libre del item
        /// </summary>
        [Column("Delibre")]
        public string? Delibre { get; set; }

        /// <summary>
        /// Indicador contable
        /// </summary>
        [Column("IContab")]
        public short? IContab { get; set; }

        /// <summary>
        /// Sector de la orden de compra
        /// </summary>
        [Column("Sect_Ocom")]
        [MaxLength(12)]
        public string? SectOcom { get; set; }

        /// <summary>
        /// Unidad de medida
        /// </summary>
        [Column("Uni_Mas")]
        [MaxLength(4)]
        public string? UniMas { get; set; }

        /// <summary>
        /// Color de la orden de compra
        /// </summary>
        [Column("Colo_Ocom")]
        [MaxLength(16)]
        public string? ColoOcom { get; set; }

        /// <summary>
        /// Medida de la orden de compra
        /// </summary>
        [Column("Medi_Ocom")]
        [MaxLength(20)]
        public string? MediOcom { get; set; }

        /// <summary>
        /// Valor unitario
        /// </summary>
        [Column("Valo_Unid")]
        [MaxLength(12)]
        public string? ValoUnid { get; set; }

        /// <summary>
        /// Unidades de la orden de compra
        /// </summary>
        [Column("Unids_Ocom", TypeName = "decimal(18,4)")]
        public decimal? UnidsOcom { get; set; }

        /// <summary>
        /// Cantidad de la orden de compra
        /// </summary>
        [Column("Cant_Ocom", TypeName = "decimal(18,4)")]
        public decimal? CantOcom { get; set; }

        /// <summary>
        /// Unidad de compra
        /// </summary>
        [Column("Uni_Com")]
        [MaxLength(4)]
        public string? UniCom { get; set; }

        /// <summary>
        /// Coeficiente de unidades
        /// </summary>
        [Column("Coef_Unids", TypeName = "decimal(18,4)")]
        public decimal? CoefUnids { get; set; }

        /// <summary>
        /// Precio unitario
        /// </summary>
        [Column("Pre_Unit", TypeName = "decimal(18,4)")]
        public decimal? PreUnit { get; set; }

        /// <summary>
        /// Descuento del item
        /// </summary>
        [Column("Desc_Item", TypeName = "decimal(18,4)")]
        public decimal? DescItem { get; set; }

        /// <summary>
        /// Precio unitario neto
        /// </summary>
        [Column("PrUn_Neto", TypeName = "decimal(18,4)")]
        public decimal? PrUnNeto { get; set; }

        /// <summary>
        /// Importe neto del item
        /// </summary>
        [Column("INet_Item", TypeName = "decimal(18,4)")]
        public decimal? INetItem { get; set; }

        /// <summary>
        /// Fecha de entrega solicitada
        /// </summary>
        [Column("Fentre_Sol")]
        public DateTime? FentreSol { get; set; }

        /// <summary>
        /// Consideración de material
        /// </summary>
        [Column("Consimat")]
        [MaxLength(2)]
        public string? Consimat { get; set; }

        /// <summary>
        /// Cantidad recibida
        /// </summary>
        [Column("Cant_Rec", TypeName = "decimal(18,4)")]
        public decimal? CantRec { get; set; }

        /// <summary>
        /// Código de operación
        /// </summary>
        [Column("Cod_Oper")]
        public short? CodOper { get; set; }

        /// <summary>
        /// Documento original
        /// </summary>
        [Column("Doc_Orig")]
        [MaxLength(12)]
        public string? DocOrig { get; set; }

        /// <summary>
        /// Número de orden de servicio
        /// </summary>
        [Column("N_OrServ", TypeName = "decimal(18,4)")]
        public decimal? NOrServ { get; set; }

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
        /// Número de orden de trabajo
        /// </summary>
        [Column("Nro_Otrab")]
        [MaxLength(16)]
        public string? NroOtrab { get; set; }

        /// <summary>
        /// Número de grupo
        /// </summary>
        [Column("Nro_Grup")]
        [MaxLength(10)]
        public string? NroGrup { get; set; }

        /// <summary>
        /// Identificador de requerimiento
        /// </summary>
        [Column("IdenReque")]
        [MaxLength(24)]
        public string? IdenReque { get; set; }

        /// <summary>
        /// Precio unitario neto en pesos
        /// </summary>
        [Column("PrUn_Neto_Pes", TypeName = "decimal(18,4)")]
        public decimal? PrUnNetoPes { get; set; }

        // ===== PROPIEDADES AUXILIARES PARA EL FRONTEND =====
        // Estas propiedades no se mapean a la base de datos, solo se usan en el frontend
        // para manejar los valores de Sector yDestino de forma más amigable

        /// <summary>
        /// Propiedad auxiliar para el Sector en el frontend
        /// Se mapea al campo StatOcom al guardar
        /// Campo de valor libre 1
        /// </summary>
        [Column("VALOBADA_1")]
        [MaxLength(256)]
        public string? Valobada1 { get; set; }

        /// <summary>
        /// Campo de valor libre 2
        /// </summary>
        [Column("VALOBADA_2")]
        [MaxLength(256)]
        public string? Valobada2 { get; set; }

        /// <summary>
        /// Campo de valor libre 3
        /// </summary>
        [Column("VALOBADA_3")]
        [MaxLength(256)]
        public string? Valobada3 { get; set; }

        /// <summary>
        /// Número de usuario
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

        /// <summary>
        /// Código de empresa .NET
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

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
        /// GUID único del registro
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();

        // ===== PROPIEDADES AUXILIARES PARA EL FRONTEND =====
        // Estas propiedades no se mapean a la base de datos, solo se usan en el frontend
        // para manejar los valores de Sector y Destino de forma más amigable

        /// <summary>
        /// Propiedad auxiliar para el Sector en el frontend
        /// Se mapea al campo StatOcom al guardar (máximo 12 caracteres)
        /// </summary>
        [NotMapped]
        public string SectorItem { get; set; } = string.Empty;

        /// <summary>
        /// Propiedad auxiliar para el Destino en el frontend
        /// Se mapea al campo IContab al guardar
        /// </summary>
        [NotMapped]
        public int? DestinoItem { get; set; }
    }
}

/*
-- =====================================================
-- INSTRUCCIONES SQL PARA MODIFICAR LA TABLA OCOMDETA
-- =====================================================

-- 1. AGREGAR CAMPO ID IDENTITY COMO CLAVE PRIMARIA
ALTER TABLE [dbo].[OCOMDETA] ADD [OCOMDETA_ID] INT IDENTITY(1,1) NOT NULL;
ALTER TABLE [dbo].[OCOMDETA] ADD CONSTRAINT PK_OCOMDETA PRIMARY KEY ([OCOMDETA_ID]);

-- 2. CAMBIAR CAMPO CodiEmpr DE SMALLINT A INT
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [CodiEmpr] INT NULL;

-- 3. AGREGAR NUEVOS CAMPOS
ALTER TABLE [dbo].[OCOMDETA] ADD [NumUsuar] INT NULL;
ALTER TABLE [dbo].[OCOMDETA] ADD [CodiEmprNet] INT NULL;
ALTER TABLE [dbo].[OCOMDETA] ADD [AddRecord] DATETIME2 NULL;
ALTER TABLE [dbo].[OCOMDETA] ADD [LastUpdate] DATETIME2 NULL;

-- 4. CAMBIAR CAMPOS DE FLOAT A DECIMAL(18,4) PARA MAYOR PRECISIÓN
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [Unids_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [Cant_Ocom] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [Coef_Unids] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [Pre_Unit] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [Desc_Item] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [PrUn_Neto] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [INet_Item] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [Cant_Rec] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [N_OrServ] DECIMAL(18,4) NULL;
ALTER TABLE [dbo].[OCOMDETA] ALTER COLUMN [PrUn_Neto_Pes] DECIMAL(18,4) NULL;

-- NOTA: Ejecutar estos comandos en SQL Server Management Studio
-- para aplicar los cambios a la estructura de la tabla existente.
-- 
-- SECCIÓN 1: Agrega campo ID identity como clave primaria
-- SECCIÓN 2: Cambia CodiEmpr de SMALLINT a INT
-- SECCIÓN 3: Agrega nuevos campos para control y auditoría
-- SECCIÓN 4: Convierte campos FLOAT a DECIMAL(18,4) para mayor precisión
--           en cálculos monetarios y cantidades
*/
