using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa a un Deudor en el sistema.
    /// Mapea a la tabla Deudor12 en la base de datos.
    /// </summary>
    public class Deudor12 : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único autoincremental para Deudor12 (nuevo campo IDENTITY).
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de cliente (se mantiene como int, pero ya no es identity).
        /// </summary>
        [Column("Nume_Cli")]
        public int Nume_Cli { get; set; }

        /// <summary>
        /// Línea de entrega 0 del cliente
        /// </summary>
        [StringLength(255)]
        public string? Lentre0_Cli { get; set; }

        /// <summary>
        /// Línea de entrega 1 del cliente
        /// </summary>
        [StringLength(255)]
        public string? Lentre1_Cli { get; set; }

        /// <summary>
        /// Número de transacción del cliente
        /// </summary>
        public decimal? NTrans_Cli { get; set; }

        /// <summary>
        /// Línea de entrega 2 del cliente
        /// </summary>
        [StringLength(255)]
        public string? Lentre2_Cli { get; set; }

        /// <summary>
        /// Hora de entrega
        /// </summary>
        [StringLength(255)]
        public string? Hora_Ent { get; set; }

        /// <summary>
        /// Zona del cliente
        /// </summary>
        [StringLength(255)]
        public string? Zona_Cli { get; set; }

        /// <summary>
        /// Zona del representante
        /// </summary>
        [StringLength(255)]
        public string? Zona_Rep { get; set; }

        /// <summary>
        /// Grupo de división
        /// </summary>
        public decimal? Gru_Div { get; set; }

        /// <summary>
        /// Código del cliente
        /// </summary>
        [StringLength(255)]
        public string? Codi_Cli { get; set; }

        /// <summary>
        /// Razón social del cliente
        /// </summary>
        [StringLength(255)]
        [Required(ErrorMessage = "La Razón Social es obligatoria")]
        public string? Raso_Cli { get; set; }

        /// <summary>
        /// Nombre de fantasía del cliente
        /// </summary>
        [StringLength(255)]
        public string? Fant_Cli { get; set; }

        /// <summary>
        /// Estado del cliente
        /// </summary>
        public short? Stat_Cli { get; set; }

        /// <summary>
        /// Domicilio del cliente
        /// </summary>
        [StringLength(255)]
        public string? Domi_Cli { get; set; }

        /// <summary>
        /// Código postal del cliente
        /// </summary>
        [StringLength(255)]
        public string? Cpos_Cli { get; set; }

        /// <summary>
        /// Localidad del cliente
        /// </summary>
        [StringLength(255)]
        public string? Loca_Cli { get; set; }

        /// <summary>
        /// País del cliente
        /// </summary>
        [StringLength(255)]
        public string? Pais_Cli { get; set; }

        /// <summary>
        /// Provincia del cliente
        /// </summary>
        [StringLength(255)]
        public string? Prov_Cli { get; set; }

        /// <summary>
        /// Teléfono del cliente
        /// </summary>
        [StringLength(255)]
        public string? Tele_Cli { get; set; }

        /// <summary>
        /// Fax del cliente
        /// </summary>
        [StringLength(255)]
        public string? Faxi_Cli { get; set; }

        /// <summary>
        /// Email del cliente
        /// </summary>
        [StringLength(255)]
        public string? Mail_Cli { get; set; }

        /// <summary>
        /// Página web del cliente
        /// </summary>
        [StringLength(255)]
        public string? PagWeb { get; set; }

        /// <summary>
        /// Posición IVA del cliente
        /// </summary>
        public short? Piva_Cli { get; set; }

        /// <summary>
        /// CUIT del cliente
        /// </summary>
        [StringLength(255)]
        public string? Cuit_Cli { get; set; }

        /// <summary>
        /// Número de IIBB del cliente
        /// </summary>
        [StringLength(255)]
        public string? Nibr_Cli { get; set; }

        /// <summary>
        /// Proveedor del cliente
        /// </summary>
        [StringLength(255)]
        public string? Prvd_Cli { get; set; }

        /// <summary>
        /// Cuenta contable del cliente
        /// </summary>
        [StringLength(255)]
        public string? Ctac_Cli { get; set; }

        /// <summary>
        /// Horario del cliente
        /// </summary>
        [StringLength(255)]
        public string? Hora_Cli { get; set; }

        /// <summary>
        /// Condición de pago del cliente
        /// </summary>
        public decimal? Cpag_Cli { get; set; }

        /// <summary>
        /// Lista de precios del cliente
        /// </summary>
        public decimal? Lpre_Cli { get; set; }

        /// <summary>
        /// Vendedor asignado al cliente
        /// </summary>
        public decimal? Vend_Cli { get; set; }

        /// <summary>
        /// Tipo de comprobante preferido del cliente
        /// </summary>
        public short? Tcpr_Cli { get; set; }

        /// <summary>
        /// Categoría de IIBB del cliente
        /// </summary>
        [StringLength(255)]
        public string? CaIb_Cli { get; set; }

        /// <summary>
        /// Estado de despacho del cliente
        /// </summary>
        [StringLength(255)]
        public string? Dsta_Cli { get; set; }

        /// <summary>
        /// Límite de crédito del cliente
        /// </summary>
        public decimal? Lcre_Cli { get; set; }

        /// <summary>
        /// Saldo del cliente
        /// </summary>
        public decimal? Sald_Cli { get; set; }

        /// <summary>
        /// Cuota máxima del cliente
        /// </summary>
        public decimal? CuMa_Cli { get; set; }

        /// <summary>
        /// Grupo de compra del cliente
        /// </summary>
        public decimal? GrCo_Cli { get; set; }

        /// <summary>
        /// Notas del cliente
        /// </summary>
        public string? Nota_Cli { get; set; }

        /// <summary>
        /// Observaciones de entrega
        /// </summary>
        public string? Obser_Entre { get; set; }

        /// <summary>
        /// Recargo por embalaje
        /// </summary>
        public decimal? Recar_Embal { get; set; }

        /// <summary>
        /// Marca ya facturado
        /// </summary>
        public decimal? Mar_Ya_Fac { get; set; }

        /// <summary>
        /// Posición IVA
        /// </summary>
        [StringLength(255)]
        public string? DePosIva { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public short? CodiEmpr { get; set; }

        /// <summary>
        /// Límite de crédito en cuenta corriente
        /// </summary>
        public decimal? LCred_CCte { get; set; }

        /// <summary>
        /// Envío de mail de deuda
        /// </summary>
        public decimal? EnvMaiDeuda { get; set; }

        /// <summary>
        /// Cuenta contable de cobranza
        /// </summary>
        [StringLength(48)]
        public string? Ctact_Cob { get; set; }

        /// <summary>
        /// Teléfono de cobranza
        /// </summary>
        [StringLength(48)]
        public string? Tele_Cob { get; set; }

        /// <summary>
        /// Email de cobranza
        /// </summary>
        [StringLength(128)]
        public string? Mail_Cob { get; set; }

        /// <summary>
        /// Días de recepción para cobranza
        /// </summary>
        [StringLength(64)]
        public string? DiasRec_Cob { get; set; }

        /// <summary>
        /// Días de pago para cobranza
        /// </summary>
        [StringLength(64)]
        public string? DiasPago_Cob { get; set; }

        /// <summary>
        /// Seguimiento de cobranza
        /// </summary>
        public string? Segui_Cob { get; set; }

        /// <summary>
        /// Email para facturación
        /// </summary>
        [StringLength(128)]
        public string? MailFactura_Cob { get; set; }

        /// <summary>
        /// Celular del cliente
        /// </summary>
        [StringLength(128)]
        public string? CELU_CLI { get; set; }

        /// <summary>
        /// Línea de entrega 3 del cliente
        /// </summary>
        [StringLength(10)]
        public string? Lentre3_Cli { get; set; }

        /// <summary>
        /// Jurisdicción de IIBB del cliente
        /// </summary>
        [StringLength(2)]
        public string? JurisdiccioIIBB_CLI { get; set; }

        /// <summary>
        /// Régimen de factura de crédito
        /// </summary>
        public short? RegimenFacturaCredito { get; set; }

        /// <summary>
        /// Fecha de alta del cliente
        /// </summary>
        public DateTime? FechAlta { get; set; }

        /// <summary>
        /// Provincia de entrega del cliente
        /// </summary>
        [StringLength(2)]
        public string? ProvEntre_Cli { get; set; }

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
