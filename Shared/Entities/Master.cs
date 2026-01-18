using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace IMPLANPROD.Shared.Entities

{
    /// <summary>
    /// Entidad Master. Representa la tabla [Master] con todos los campos especificados.
    /// Todos los campos float se mapean como decimal para mayor precisión.
    /// </summary>
    public class Master : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Clave primaria, identidad. Equivale a Codint en la base de datos.
        /// </summary>
        [Key]
        public int Id { get; set; }

        public int Codint { get; set; }

        // Fecha de revisión
        public DateTime? Ferevis { get; set; }

        // Código de empresa
        public short? CodiEmpr { get; set; }

        // Observaciones (texto largo)
        public string? Observa { get; set; }

        // Especificación (texto largo)
        public string? Especificacion { get; set; }

        // Descomposición automática
        public short? DesCompAut { get; set; }

        // Descripción de composición
        [MaxLength(64)]
        public string? DescriComp { get; set; }

        // Depósito de entrada
        public short? DepoEntra { get; set; }

        // Depósito de consumo
        public short? DepoCon { get; set; }

        // Código
        [MaxLength(24)]
        public string? Codigo { get; set; }

        // Descripción
        [MaxLength(64)]
        public string? Descripcion { get; set; }

        // Descripción dos
        [MaxLength(128)]
        public string? Descridos { get; set; }

        // Estado
        public short? Status { get; set; }

        // Alias
        [MaxLength(24)]
        public string? Alias { get; set; }

        // Unidad de medida
        [MaxLength(4)]
        public string? Umedida { get; set; }

        // Tiempo de mate
        public short? Timate { get; set; }

        // Crítico de reposición
        public short? Critrep { get; set; }

        // Moneda de costo
        public short? Monecos { get; set; }

        // Fecha de costo
        public DateTime? Fecosto { get; set; }

        // Porcentaje de descuento de costo
        public decimal? PorDescCos { get; set; }

        // Costo unitario
        public decimal? Cosuni { get; set; }

        // Porcentaje nacional
        public decimal? Pornacio { get; set; }

        // Porcentaje comisión de venta
        public decimal? PorComVen { get; set; }

        // Proveedor habilitado
        public short? Provhabi { get; set; }

        // Stock mínimo
        public decimal? Stomin { get; set; }

        // Stock máximo
        public decimal? Stomax { get; set; }

        // Lote de reposición
        public decimal? Lorepos { get; set; }

        // Lote indivisible
        public decimal? Lindivis { get; set; }

        // Peso unitario
        public decimal? Pesuni { get; set; }

        // Peso por metro
        public decimal? PesMetro { get; set; }

        // Color
        [MaxLength(32)]
        public string? Color { get; set; }

        // Talle
        [MaxLength(32)]
        public string? Talle { get; set; }

        // Descomposición
        public short? Descompo { get; set; }

        // Grupo de IVA
        public short? Grupiva { get; set; }

        // Grupo de comisión de venta
        public short? Grucoven { get; set; }

        // Código de línea
        [MaxLength(8)]
        public string? CLinea { get; set; }

        // Código de familia
        [MaxLength(8)]
        public string? Cfamilia { get; set; }

        // Número de plano
        [MaxLength(24)]
        public string? NPlano { get; set; }

        // Revisión
        [MaxLength(16)]
        public string? Revision { get; set; }

        // Indicador de trazabilidad
        public short? Trazable { get; set; }

        // Revisión de fórmula
        public DateTime? Revi_Formu { get; set; }

        // Fecha de aprobación de fórmula
        public DateTime? FechApFor { get; set; }

        // Usuario aprobación fórmula
        [MaxLength(24)]
        public string? UsuApFor { get; set; }

        // Código interna
        [MaxLength(64)]
        public string? Cod_Interna { get; set; }

        // Bien de capital
        public short? Bien_Cap { get; set; }

        // Código NCM
        [MaxLength(32)]
        public string? Codigo_Ncm { get; set; }

        // Código secundario
        [MaxLength(32)]
        public string? Codigo_Sec { get; set; }

        // Proveedor unidad de medida
        [MaxLength(32)]
        public string? Pro_Umed { get; set; }

        // Id de IVA
        [MaxLength(8)]
        public string? Iva_Id { get; set; }

        // Archivo PDF (texto largo)
        public string? ArchiPdf { get; set; }

        // Familia
        [MaxLength(4)]
        public string? Familia { get; set; }

        // Límite de venta indivisible
        public decimal? Lindiven { get; set; }

        // Cantidad decimal
        [MaxLength(2)]
        public string? CanDeci { get; set; }

        // Consumo promedio
        public decimal? ConsuProm { get; set; }

        // Costo calculado
        public decimal? CosCal { get; set; }

        // Fecha de costo calculado
        public DateTime? FeCosCal { get; set; }

        // Estado de movimiento
        public short? Stat_Movi { get; set; }

        // Clasificación ATC
        [MaxLength(64)]
        public string? Clas_ATC { get; set; }

        // Versión BC
        [MaxLength(32)]
        public string? Vers_BC { get; set; }

        // Fecha de vigencia BC
        public DateTime? FeVig_BC { get; set; }

        // Fecha de revisión BC
        public DateTime? FeRev_BC { get; set; }

        // Cantidad aprobada
        public short? CantApro { get; set; }

        // Letra tipo de composición
        [MaxLength(16)]
        public string? LeTiCom { get; set; }

        // Piezas por caja
        public int? PiezasXCaja { get; set; }

        // Metros cúbicos por caja
        public decimal? M3Caja { get; set; }

        // Kilo neto por caja
        public decimal? KiloNetoCaja { get; set; }

        // Kilo bruto por caja
        public decimal? KiloBrutoCaja { get; set; }

        // Autonomía
        [MaxLength(10)]
        public string? Autonomia { get; set; }

        // Valor obada 1
        [MaxLength(512)]
        public string? VALOBADA_1 { get; set; }

        // Valor obada 2
        [MaxLength(512)]
        public string? VALOBADA_2 { get; set; }

        // Valor obada 3
        [MaxLength(512)]
        public string? VALOBADA_3 { get; set; }

        // Autonomía proyectada
        [MaxLength(10)]
        public string? Autono_Proy { get; set; }

        // Usuario precio peso
        [MaxLength(24)]
        public string? UsuPrePes { get; set; }

        // Usuario precio dólar
        [MaxLength(24)]
        public string? UsuPreDol { get; set; }

        // Mano de obra real
        public decimal? MANO_OBRA_REAL { get; set; }

        // Tamaño de etiqueta
        public short? Tamano_Etiqueta { get; set; }

        // Cuenta contable I
        public short? CueContaI { get; set; }

        // Molde
        public short? Molde { get; set; }

        // Consumo por corte
        public short? Consumo_X_Corte { get; set; }

        // Gramaje
        public decimal? GRAMAJE { get; set; }

        // Ancho
        public decimal? ANCHO { get; set; }

        // Grado actual
        public short? GR_ACTU { get; set; }

        // Código ARBA
        [MaxLength(16)]
        public string? Cod_Arba { get; set; }

        // Micrones
        public decimal? MICRONES { get; set; }

        // Valor obada 4
        [MaxLength(512)]
        public string? VALOBADA_4 { get; set; }

        // Valor obada 5
        [MaxLength(512)]
        public string? VALOBADA_5 { get; set; }

        // Valor obada 6
        [MaxLength(512)]
        public string? VALOBADA_6 { get; set; }

        [Column("rowguid")]
        public Guid RowGuid { get; set; }

        [NotMapped] // Esta propiedad no se mapeará a la base de datos
        public decimal? valorPesos { get; set; } // Propiedad virtual que no se almacena en la tabla

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
//CREATE INDEX IX_Master_Codigo ON Master(Codigo) INCLUDE (Descripcion, Status);
//CREATE INDEX IX_Master_Descripcion ON Master(Descripcion) INCLUDE (Codigo, Status);
