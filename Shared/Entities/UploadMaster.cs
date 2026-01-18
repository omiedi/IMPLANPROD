using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.Entities
{
    public class UploadMaster : ITimestampEntity, IAuditableEntity
    {   [Key]
        public int Id { get; set; }
        public int Codint { get; set; }
        public string? Nombre { get; set; }
        public string? StoredFileName { get; set; }
        public string? ContentType { get; set; }
        public DateTime? FechaSubida { get; set; } 

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

//CREATE TABLE UploadMaster (
//    Id INT PRIMARY KEY IDENTITY(1,1),
//    codint int NOT NULL,
//    Nombre NVARCHAR(MAX) NULL,
//    StoredFileName NVARCHAR(MAX) NULL,
//    ContentType NVARCHAR(MAX) NULL
//);

//ALTER TABLE UploadMaster
//ADD FechaSubida DATETIME DEFAULT GETDATE();
