using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("SyncConfig")]
    public class SyncConfig
    {
        [Key]
        public int Id { get; set; }

        [Column("GeneraScriptSync")]
        public bool GeneraScriptSync { get; set; }

        [Column("RutaIntercambio")]
        [MaxLength(500)]
        public string RutaIntercambio { get; set; } = string.Empty;

        [Column("EsCentral")]
        public bool EsCentral { get; set; }

        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }
    }
}
