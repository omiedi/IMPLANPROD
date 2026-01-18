using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad MACONSIG - Tabla de consignaciones
    /// </summary>
    [Table("MACONSIG")]
    public class Maconsig : IAuditableEntity
    {
        /// <summary>
        /// Identificador interno autoincremental (IDENTITY) y clave primaria.
        /// Se agrega para usar una PK numérica por rendimiento y mantener los campos originales como datos de negocio.
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Fecha de emisión de la orden de compra
        /// </summary>
        [Column("Femi_Ocom")]
        public DateTime? FemiOcom { get; set; }

        /// <summary>
        /// Número de proveedor de la orden de compra
        /// </summary>
        [Column("Npro_Ocom")]
        public int? NproOcom { get; set; }

        /// <summary>
        /// Razón social del proveedor
        /// </summary>
        [Column("RaSo_Prov")]
        [MaxLength(64)]
        public string? RaSoProv { get; set; }

        /// <summary>
        /// Sección ejecutiva
        /// </summary>
        [Column("Secc_Ejec", TypeName = "decimal(18,4)")]
        public decimal? SeccEjec { get; set; }

        /// <summary>
        /// Denominación de la sección
        /// </summary>
        [Column("Deno_Secc")]
        [MaxLength(64)]
        public string? DenoSecc { get; set; }

        /// <summary>
        /// Número de orden de fabricación
        /// </summary>
        [Column("Nro_Orfa")]
        [MaxLength(12)]
        public string? NroOrfa { get; set; }

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        [Column("Nro_Ocom")]
        [MaxLength(14)]
        public string? NroOcom { get; set; }

        /// <summary>
        /// Código interno del producto
        /// </summary>
        [Column("Cod_Inte")]
        public int? CodInte { get; set; }

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
        /// Lote de la orden de compra
        /// </summary>
        [Column("Lote_Ocom")]
        [MaxLength(16)]
        public string? LoteOcom { get; set; }

        /// <summary>
        /// Unidad de medida
        /// </summary>
        [Column("Unidad")]
        [MaxLength(4)]
        public string? Unidad { get; set; }

        /// <summary>
        /// Cantidad asignada
        /// </summary>
        [Column("Cant_Asig", TypeName = "decimal(18,4)")]
        public decimal? CantAsig { get; set; }

        /// <summary>
        /// Cantidad entregada
        /// </summary>
        [Column("Cant_Entreg", TypeName = "decimal(18,4)")]
        public decimal? CantEntreg { get; set; }

        /// <summary>
        /// Fecha de entrega
        /// </summary>
        [Column("Fecha_Entreg")]
        public DateTime? FechaEntreg { get; set; }

        /// <summary>
        /// Número de remito de consignación
        /// </summary>
        [Column("NRemi_Consig")]
        [MaxLength(16)]
        public string? NRemiConsig { get; set; }

        /// <summary>
        /// Valor unitario
        /// </summary>
        [Column("Valo_Unit", TypeName = "decimal(18,4)")]
        public decimal? ValoUnit { get; set; }

        /// <summary>
        /// Código de moneda
        /// </summary>
        [Column("Moneda")]
        public short? Moneda { get; set; }

        /// <summary>
        /// Valor total
        /// </summary>
        [Column("Valo_Tot", TypeName = "decimal(18,4)")]
        public decimal? ValoTot { get; set; }

        /// <summary>
        /// Cantidad rendida
        /// </summary>
        [Column("Cant_Rendid", TypeName = "decimal(18,4)")]
        public decimal? CantRendid { get; set; }

        /// <summary>
        /// Saldo pendiente
        /// </summary>
        [Column("Saldo_Pend", TypeName = "decimal(18,4)")]
        public decimal? SaldoPend { get; set; }

        /// <summary>
        /// Observaciones
        /// </summary>
        [Column("Observa")]
        public string? Observa { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Devolución/Ajuste
        /// </summary>
        [Column("Devo_Ajus", TypeName = "decimal(18,4)")]
        public decimal? DevoAjus { get; set; }

        /// <summary>
        /// Fecha de entrega (campo adicional)
        /// </summary>
        [Column("Fech_Entreg")]
        public DateTime? FechEntreg { get; set; }

        /// <summary>
        /// Cantidad devuelta
        /// </summary>
        [Column("Cant_Devu", TypeName = "decimal(18,4)")]
        public decimal? CantDevu { get; set; }

        /// <summary>
        /// Marca de borrado
        /// </summary>
        [Column("MarBorra")]
        public short? MarBorra { get; set; }

        /// <summary>
        /// GUID único del registro
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();

        /// <summary>
        /// Número de orden de item
        /// </summary>
        [Column("NuOrdItem")]
        public short? NuOrdItem { get; set; }

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
INSTRUCCIONES SQL PARA MODIFICAR LA TABLA EXISTENTE:
Ejecutar en SQL Server Management Studio para actualizar los tipos de datos:

-- SECCIÓN 0: Agregar campo ID IDENTITY como clave primaria
ALTER TABLE MACONSIG ADD Id INT IDENTITY(1,1) NOT NULL;
ALTER TABLE MACONSIG ADD CONSTRAINT PK_MACONSIG_Id PRIMARY KEY (Id);

- AGREGAR NUEVOS CAMPOS
ALTER TABLE [dbo].[MACONSIG] ADD [CodiEmprNet] INT NULL;
ALTER TABLE [dbo].[MACONSIG] ADD [NumUsuar] INT NULL;
ALTER TABLE [dbo].[MACONSIG] ADD [AddRecord] DATETIME2 NOT NULL DEFAULT GETDATE();
ALTER TABLE [dbo].[MACONSIG] ADD [LastUpdate] DATETIME2 NOT NULL DEFAULT GETDATE();


-- SECCIÓN 1: Cambiar campos SMALLINT a INT
ALTER TABLE MACONSIG ALTER COLUMN Npro_Ocom INT;
ALTER TABLE MACONSIG ALTER COLUMN Cod_Inte INT;
ALTER TABLE MACONSIG ALTER COLUMN CodiEmpr INT;

-- SECCIÓN 2: Cambiar campos FLOAT a DECIMAL(18,4) para mayor precisión
ALTER TABLE MACONSIG ALTER COLUMN Secc_Ejec DECIMAL(18,4);
ALTER TABLE MACONSIG ALTER COLUMN Cant_Asig DECIMAL(18,4);
ALTER TABLE MACONSIG ALTER COLUMN Cant_Entreg DECIMAL(18,4);
ALTER TABLE MACONSIG ALTER COLUMN Valo_Unit DECIMAL(18,4);
ALTER TABLE MACONSIG ALTER COLUMN Valo_Tot DECIMAL(18,4);
ALTER TABLE MACONSIG ALTER COLUMN Cant_Rendid DECIMAL(18,4);
ALTER TABLE MACONSIG ALTER COLUMN Saldo_Pend DECIMAL(18,4);
ALTER TABLE MACONSIG ALTER COLUMN Devo_Ajus DECIMAL(18,4);
ALTER TABLE MACONSIG ALTER COLUMN Cant_Devu DECIMAL(18,4);

-- NOTAS IMPORTANTES:
-- 1. Ejecutar las secciones en orden (0, 1, 2)
-- 2. Hacer backup de la tabla antes de ejecutar
-- 3. Verificar que no haya valores que excedan los límites de INT antes de ejecutar la Sección 1
-- 4. El campo Id se agregará automáticamente con valores secuenciales para registros existentes
*/
