using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;

namespace IMPLANPROD.Shared.Entities
{
    public class PedidoAuditoria : ITimestampEntity, IAuditableEntity
    {
        [Key]
        [Column("IdAuditoria")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int IdAuditoria { get; set; }

        public string? Tabla { get; set; }
        public int NroPed { get; set; }
        public string? Accion { get; set; }
        public string? Campo { get; set; }
        public string? ValorAntiguo { get; set; }
        public string? ValorNuevo { get; set; }
        public DateTime FechaCambio { get; set; }
        public string? Usuario { get; set; }
        public int? NumUsuario { get; set; }  // Nullable int

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
//---------------------------------------------------
//--1.Tabla de Auditoría
//---------------------------------------------------
//select * from PEDIDOS_AUDITORIA
//CREATE TABLE PEDIDOS_AUDITORIA (
//    IdAuditoria INT IDENTITY(1,1) PRIMARY KEY,       -- ID autoincremental
//    Tabla VARCHAR(20) NOT NULL,                      -- PEDENCA o PEDDETA
//    NroPed INT NOT NULL,                             -- Número de pedido
//    Accion VARCHAR(10) NOT NULL,                     -- INSERT / UPDATE / DELETE
//    Campo VARCHAR(50) NULL,                          -- Nombre del campo o descripción del ítem
//    ValorAntiguo NVARCHAR(4000) NULL,               -- Valor anterior
//    ValorNuevo NVARCHAR(4000) NULL,                 -- Valor nuevo
//    FechaCambio DATETIME NOT NULL DEFAULT GETDATE(), -- Fecha y hora del cambio
//    Usuario NVARCHAR(50) NULL,                      -- Usuario (opcional)
//    NumUsuario INT NULL                              -- Número de usuario (opcional)
//);
//GO


//---------------------------------------------------
//--Trigger optimizado para PEDENCA
//---------------------------------------------------
//IF OBJECT_ID('trg_Auditoria_PEDENCA', 'TR') IS NOT NULL
//    DROP TRIGGER trg_Auditoria_PEDENCA;
//GO

//CREATE TRIGGER trg_Auditoria_PEDENCA
//ON PEDENCA
//AFTER INSERT, UPDATE, DELETE
//AS
//BEGIN
//    SET NOCOUNT ON;

//DECLARE @FechaCambio DATETIME = GETDATE();

//-- =======================
//--UPDATE reales
//-- =======================
//INSERT INTO PEDIDOS_AUDITORIA(Tabla, NroPed, Accion, Campo, ValorAntiguo, ValorNuevo, FechaCambio)
//    SELECT
//        'PEDENCA',
//        i.NroPed,
//        'UPDATE',
//        'DomiEnt/LocaEnt/NTranspo/Vended/NumUsuar/TDescApli',
//        CONCAT(
//            'DomiEnt:', d.DomiEnt,
//            ', LocaEnt:', d.LocaEnt,
//            ', NTranspo:', d.NTranspo,
//            ', Vended:', d.Vended,
//            ', NumUsuar:', d.NumUsuar,
//            ', TDescApli:', d.TDescApli
//        ),
//        CONCAT(
//            'DomiEnt:', i.DomiEnt,
//            ', LocaEnt:', i.LocaEnt,
//            ', NTranspo:', i.NTranspo,
//            ', Vended:', i.Vended,
//            ', NumUsuar:', i.NumUsuar,
//            ', TDescApli:', i.TDescApli
//        ),
//        @FechaCambio
//    FROM INSERTED i
//    JOIN DELETED d ON i.NroPed = d.NroPed
//    WHERE
//        ISNULL(i.DomiEnt,'')   <> ISNULL(d.DomiEnt,'')
//        OR ISNULL(i.LocaEnt,'')<> ISNULL(d.LocaEnt,'')
//        OR ISNULL(i.NTranspo,0)<> ISNULL(d.NTranspo,0)
//        OR ISNULL(i.Vended,0)  <> ISNULL(d.Vended,0)
//        OR ISNULL(i.NumUsuar,0)<> ISNULL(d.NumUsuar,0)
//        OR ISNULL(i.TDescApli,'')<> ISNULL(d.TDescApli,'');

//-- =======================
//--INSERT nuevos
//-- =======================
//INSERT INTO PEDIDOS_AUDITORIA(Tabla, NroPed, Accion, Campo, ValorNuevo, FechaCambio)
//    SELECT
//        'PEDENCA',
//        i.NroPed,
//        'INSERT',
//        'Registro Nuevo',
//        CONCAT(
//            'DomiEnt:',i.DomiEnt,
//            ', LocaEnt:',i.LocaEnt,
//            ', NTranspo:',i.NTranspo,
//            ', Vended:',i.Vended,
//            ', NumUsuar:',i.NumUsuar,
//            ', TDescApli:',i.TDescApli
//        ),
//        @FechaCambio
//    FROM INSERTED i
//    LEFT JOIN DELETED d ON i.NroPed = d.NroPed
//    WHERE d.NroPed IS NULL;

//-- =======================
//--DELETE
//-- =======================
//INSERT INTO PEDIDOS_AUDITORIA(Tabla, NroPed, Accion, Campo, ValorAntiguo, FechaCambio)
//    SELECT
//        'PEDENCA',
//        d.NroPed,
//        'DELETE',
//        'Registro Eliminado',
//        CONCAT(
//            'DomiEnt:',d.DomiEnt,
//            ', LocaEnt:',d.LocaEnt,
//            ', NTranspo:',d.NTranspo,
//            ', Vended:',d.Vended,
//            ', NumUsuar:',d.NumUsuar,
//            ', TDescApli:',d.TDescApli
//        ),
//        @FechaCambio
//    FROM DELETED d
//    LEFT JOIN INSERTED i ON i.NroPed = d.NroPed
//    WHERE i.NroPed IS NULL;
//END;
//GO

//-------------------------------------------------- -
//--Trigger optimizado para PEDDETA
//---------------------------------------------------
//IF OBJECT_ID('trg_Auditoria_PEDDETA', 'TR') IS NOT NULL
//    DROP TRIGGER trg_Auditoria_PEDDETA;
//GO

//CREATE TRIGGER trg_Auditoria_PEDDETA
//ON PEDDETA
//AFTER INSERT, DELETE
//AS
//BEGIN
//    SET NOCOUNT ON;

//DECLARE @FechaCambio DATETIME = GETDATE();

//-- =======================
//--UPDATE reales: solo ítems que cambiaron
//    -- =======================
//    INSERT INTO PEDIDOS_AUDITORIA (Tabla, NroPed, Accion, Campo, ValorAntiguo, ValorNuevo, FechaCambio)
//    SELECT 
//        'PEDDETA',
//    i.NroPed,
//        'UPDATE',
//        'Item',
//    CONCAT(
//            'CanPed:', CONVERT(varchar(20), CAST(d.CanPed AS decimal(18, 2))),
//        ', PreUnid:', CONVERT(varchar(20), CAST(d.PreUnid AS decimal(18, 2))),
//        ', CantEnt:', CONVERT(varchar(20), CAST(d.CantEnt AS decimal(18, 2))),
//        ', NroOcom:', d.NroOcom,
//        ', Vended:', d.Vended,
//        ', NuOrItem:', d.NuOrItem,
//        ', CodiInt:', d.CodiInt,
//        ', CodiExt:', d.CodiExt,
//        ', Descrip:', d.Descrip,
//        ', MoneEmis:', CONVERT(varchar(20), CAST(d.MoneEmis AS decimal(18, 2))),
//        ', TipoCam:', CONVERT(varchar(20), CAST(d.TipoCam AS decimal(18, 2)))
//        ),
//    CONCAT(
//            'CanPed:', CONVERT(varchar(20), CAST(i.CanPed AS decimal(18, 2))),
//        ', PreUnid:', CONVERT(varchar(20), CAST(i.PreUnid AS decimal(18, 2))),
//        ', CantEnt:', CONVERT(varchar(20), CAST(i.CantEnt AS decimal(18, 2))),
//        ', NroOcom:', i.NroOcom,
//        ', Vended:', i.Vended,
//        ', NuOrItem:', i.NuOrItem,
//        ', CodiInt:', i.CodiInt,
//        ', CodiExt:', i.CodiExt,
//        ', Descrip:', i.Descrip,
//        ', MoneEmis:', CONVERT(varchar(20), CAST(i.MoneEmis AS decimal(18, 2))),
//        ', TipoCam:', CONVERT(varchar(20), CAST(i.TipoCam AS decimal(18, 2)))
//        ),
//    @FechaCambio
//    FROM INSERTED i
//    JOIN DELETED d ON i.NroPed = d.NroPed AND i.NuOrItem = d.NuOrItem
//    WHERE
//        ISNULL(i.CanPed,0)     <> ISNULL(d.CanPed,0)
//        OR ISNULL(i.PreUnid,0) <> ISNULL(d.PreUnid,0)
//        OR ISNULL(i.CantEnt,0) <> ISNULL(d.CantEnt,0)
//        OR ISNULL(i.NroOcom,'')<> ISNULL(d.NroOcom,'')
//        OR ISNULL(i.Vended,0)  <> ISNULL(d.Vended,0)
//        OR ISNULL(i.CodiInt,0) <> ISNULL(d.CodiInt,0)
//        OR ISNULL(i.CodiExt,'')<> ISNULL(d.CodiExt,'')
//        OR ISNULL(i.Descrip,'')<> ISNULL(d.Descrip,'')
//        OR ISNULL(i.MoneEmis,0)<> ISNULL(d.MoneEmis,0)
//        OR ISNULL(i.TipoCam,0) <> ISNULL(d.TipoCam,0);

//-- =======================
//--INSERT nuevos
//-- =======================
//INSERT INTO PEDIDOS_AUDITORIA(Tabla, NroPed, Accion, Campo, ValorNuevo, FechaCambio)
//    SELECT
//        'PEDDETA',
//        i.NroPed,
//        'INSERT',
//        'Item Nuevo',
//        CONCAT(
//            'CanPed:',   CONVERT(varchar(20), CAST(i.CanPed AS decimal(18,2))),
//            ', PreUnid:',CONVERT(varchar(20), CAST(i.PreUnid AS decimal(18,2))),
//            ', CantEnt:',CONVERT(varchar(20), CAST(i.CantEnt AS decimal(18,2))),
//            ', NroOcom:', i.NroOcom,
//            ', Vended:', i.Vended,
//            ', NuOrItem:', i.NuOrItem,
//            ', CodiInt:', i.CodiInt,
//            ', CodiExt:', i.CodiExt,
//            ', Descrip:', i.Descrip,
//            ', MoneEmis:',CONVERT(varchar(20), CAST(i.MoneEmis AS decimal(18,2))),
//            ', TipoCam:', CONVERT(varchar(20), CAST(i.TipoCam AS decimal(18,2)))
//        ),
//        @FechaCambio
//    FROM INSERTED i
//    LEFT JOIN DELETED d ON i.NroPed = d.NroPed AND i.NuOrItem = d.NuOrItem
//    WHERE d.NuOrItem IS NULL
//       OR ISNULL(i.CanPed,0)     <> ISNULL(d.CanPed,0)
//       OR ISNULL(i.PreUnid,0)   <> ISNULL(d.PreUnid,0)
//       OR ISNULL(i.CantEnt,0)   <> ISNULL(d.CantEnt,0)
//       OR ISNULL(i.NroOcom,'')  <> ISNULL(d.NroOcom,'')
//       OR ISNULL(i.Vended,0)    <> ISNULL(d.Vended,0)
//       OR ISNULL(i.CodiInt,0)   <> ISNULL(d.CodiInt,0)
//       OR ISNULL(i.CodiExt,'')  <> ISNULL(d.CodiExt,'')
//       OR ISNULL(i.Descrip,'')  <> ISNULL(d.Descrip,'')
//       OR ISNULL(i.MoneEmis,0)  <> ISNULL(d.MoneEmis,0)
//       OR ISNULL(i.TipoCam,0)   <> ISNULL(d.TipoCam,0);

//-- =======================
//--DELETE reales
//-- =======================
//INSERT INTO PEDIDOS_AUDITORIA(Tabla, NroPed, Accion, Campo, ValorAntiguo, FechaCambio)
//    SELECT
//        'PEDDETA',
//        d.NroPed,
//        'DELETE',
//        'Item Eliminado',
//        CONCAT(
//            'CanPed:',   CONVERT(varchar(20), CAST(d.CanPed AS decimal(18,2))),
//            ', PreUnid:',CONVERT(varchar(20), CAST(d.PreUnid AS decimal(18,2))),
//            ', CantEnt:',CONVERT(varchar(20), CAST(d.CantEnt AS decimal(18,2))),
//            ', NroOcom:', d.NroOcom,
//            ', Vended:', d.Vended,
//            ', NuOrItem:', d.NuOrItem,
//            ', CodiInt:', d.CodiInt,
//            ', CodiExt:', d.CodiExt,
//            ', Descrip:', d.Descrip,
//            ', MoneEmis:',CONVERT(varchar(20), CAST(d.MoneEmis AS decimal(18,2))),
//            ', TipoCam:', CONVERT(varchar(20), CAST(d.TipoCam AS decimal(18,2)))
//        ),
//        @FechaCambio
//    FROM DELETED d
//    LEFT JOIN INSERTED i ON i.NroPed = d.NroPed AND i.NuOrItem = d.NuOrItem
//    WHERE i.NuOrItem IS NULL;
//END;
//GO
