using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("MASUMCLI")]
    public class Masumcli : IAuditableEntity
    {
        public int? Nume_Cli { get; set; }

        public int? NumeIng { get; set; }

        [MaxLength(64)]
        public string? Descripcion { get; set; }

        [MaxLength(64)]
        public string? Yacimiento { get; set; }

        [MaxLength(64)]
        public string? Referencia { get; set; }

        [Column(TypeName = "ntext")]
        public string? TarAcordadas { get; set; }

        [Column(TypeName = "ntext")]
        public string? Observaciones { get; set; }

        [Column(TypeName = "ntext")]
        public string? DepInterviniente { get; set; }

        [Column(TypeName = "smalldatetime")]
        public DateTime? FechRecep { get; set; }

        [MaxLength(64)]
        public string? UsuApFor_1 { get; set; }

        [MaxLength(64)]
        public string? UsuApFor_2 { get; set; }

        [MaxLength(64)]
        public string? UsuApFor_3 { get; set; }

        [MaxLength(64)]
        public string? UsuApFor_4 { get; set; }

        [Column(TypeName = "ntext")]
        public string? EstadoRec { get; set; }

        [MaxLength(64)]
        public string? Ubicacion { get; set; }

        [Column(TypeName = "smalldatetime")]
        public DateTime? FechVerif { get; set; }

        [Column(TypeName = "ntext")]
        public string? Trab_Efect { get; set; }

        [Column(TypeName = "smalldatetime")]
        public DateTime? FechTrabEfect { get; set; }

        [Column(TypeName = "ntext")]
        public string? EstEntrega { get; set; }

        public short? Status { get; set; }

        [Column(TypeName = "smalldatetime")]
        public DateTime? FechEnt { get; set; }

        [MaxLength(64)]
        public string? Lentre_Cli { get; set; }

        [MaxLength(16)]
        public string? Patente { get; set; }

        [MaxLength(48)]
        public string? Conductor { get; set; }

        [MaxLength(48)]
        public string? Piva_Cli { get; set; }

        [MaxLength(24)]
        public string? Cuit_Cli { get; set; }

        [MaxLength(48)]
        public string? Nom_Tra { get; set; }

        [Column(TypeName = "smalldatetime")]
        public DateTime? Fech_Remito { get; set; }

        [MaxLength(18)]
        public string? NumRemito { get; set; }

        public int? CodiEmpr { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }
    }
}
