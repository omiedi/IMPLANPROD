using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad Movisto con auditoría completa automática
    /// Implementa IAuditableEntity para timestamps y seguimiento de usuario/empresa
    /// </summary>
    [Table("movisto")]
    public class Movisto : IAuditableEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("Id")]
        public int Id { get; set; }

        [Column("FeReMovi")]
        public DateTime? FeReMovi { get; set; }

        /// <summary>
        /// Código de empresa 
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// Se actualiza automáticamente con el ID del usuario actual
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

        [Column("IdenMovi")]
        [StringLength(32)]
        public string? IdenMovi { get; set; }

        [Column("DeLaRemi")]
        public string? DeLaRemi { get; set; }

        [Column("Descri_Lar")]
        public string? Descri_Lar { get; set; }

        [Column("FechMov")]
        public DateTime? FechMov { get; set; }

        [Column("Cod_Inte")]
        public int? Cod_Inte { get; set; }

        [Column("Cod_Exte")]
        [StringLength(24)]
        public string? Cod_Exte { get; set; }

        [Column("Descrip")]
        [StringLength(64)]
        public string? Descrip { get; set; }

        [Column("IdenLote")]
        [StringLength(16)]
        public string? IdenLote { get; set; }

        [Column("LotConsum")]
        public string? LotConsum { get; set; }

        [Column("PNL_Des")]
        [StringLength(32)]
        public string? PNL_Des { get; set; }

        [Column("PLC_Has")]
        [StringLength(32)]
        public string? PLC_Has { get; set; }

        [Column("CodiMovi")]
        [StringLength(2)]
        public string? CodiMovi { get; set; }

        [Column("DoPriCor")]
        [StringLength(12)]
        public string? DoPriCor { get; set; }

        [Column("DoPriLar")]
        [StringLength(24)]
        public string? DoPriLar { get; set; }

        [Column("DoSecCor")]
        [StringLength(12)]
        public string? DoSecCor { get; set; }

        [Column("DoSecLar")]
        [StringLength(24)]
        public string? DoSecLar { get; set; }

        [Column("DocuOrig")]
        [StringLength(24)]
        public string? DocuOrig { get; set; }

        [Column("NumRem_Cli")]
        [StringLength(16)]
        public string? NumRem_Cli { get; set; }

        [Column("NuorItem")]
        public int? NuorItem { get; set; }

        [Column("NuPedCli")]
        public int? NuPedCli { get; set; }

        [Column("NumeOrdPed")]
        public short? NumeOrdPed { get; set; }

        [Column("ReferCor")]
        [StringLength(64)]
        public string? ReferCor { get; set; }

        [Column("CostoRep")]
        public decimal? CostoRep { get; set; }

        [Column("ValoUnit")]
        public decimal? ValoUnit { get; set; }

        [Column("MoneVal")]
        public short? MoneVal { get; set; }

        [Column("TipoCam")]
        public decimal? TipoCam { get; set; }

        [Column("DepoMovi")]
        public int? DepoMovi { get; set; }

        [Column("Nume_Clie")]
        public int? Nume_Clie { get; set; }

        [Column("Sucu_Clie")]
        public short? Sucu_Clie { get; set; }

        [Column("Nume_Prov")]
        public int? Nume_Prov { get; set; }

        [Column("CantIngre")]
        public decimal? CantIngre { get; set; }

        [Column("CantSalid")]
        public decimal? CantSalid { get; set; }

        [Column("CantSaldo")]
        public decimal? CantSaldo { get; set; }

        [Column("Verificado")]
        public short? Verificado { get; set; }

        [Column("Causa")]
        public short? Causa { get; set; }

        [Column("User_Tran")]
        [StringLength(32)]
        public string? User_Tran { get; set; }

        [Column("CodCarpImp")]
        [StringLength(12)]
        public string? CodCarpImp { get; set; }

        [Column("ValoUnitIva")]
        public decimal? ValoUnitIva { get; set; }

        [Column("VerificadoPor")]
        public short? VerificadoPor { get; set; }

        [Column("FHor_Verif")]
        public DateTime? FHor_Verif { get; set; }

        [Column("ValTotIt")]
        public decimal? ValTotIt { get; set; }

        [Column("ValTotItIva")]
        public decimal? ValTotItIva { get; set; }

        [Column("TipoTran")]
        public short? TipoTran { get; set; }

        [Column("VALOBADA_1")]
        [StringLength(256)]
        public string? VALOBADA_1 { get; set; }

        [Column("VALOBADA_2")]
        [StringLength(256)]
        public string? VALOBADA_2 { get; set; }

        [Column("OCOMPRA_RTOABIERTO")]
        [StringLength(64)]
        public string? OCOMPRA_RTOABIERTO { get; set; }

        [Column("NroEnvio")]
        public int? NroEnvio { get; set; }

        [Column("NUMPVTA")]
        public short? NUMPVTA { get; set; }

        [Column("DOPRINUM")]
        public int? DOPRINUM { get; set; }

        [Column("IdCalidad")]
        public int? IdCalidad { get; set; }

        [Column("ID_OPERACION")]
        public short? ID_OPERACION { get; set; }

        [Column("PESOKG")]
        public decimal? PESOKG { get; set; }

        [Column("METROS")]
        public decimal? METROS { get; set; }

        [Column("CierreId")]
        public int? CierreId { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================
        
        /// <summary>
        /// Fecha de creación del registro
        /// Se establece automáticamente cuando se crea el movimiento de stock
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente cada vez que se modifica el movimiento
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// Se actualiza automáticamente con la empresa del usuario actual
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }
    }
}

//PRIMERO PARA QUE SE CREE LA TABLA ARCON2.DBO.MOVISTO Y SE MIGREN TODOS LOS DATOS < 01/MAY/2004
//SELECT * 
//INTO arcon2.dbo.movisto --tabla destion
//FROM arcon.dbo.movisto -- tabla origen
//where fechmov < '01/MAY/2004'


//INSERT INTO arcon2.dbo.movisto --tabla destion
//SELECT  * FROM arcon.dbo.movisto -- tabla origen
//where fechmov >= '01/MAY/2004' AND FECHMOV < '01/MAY/2006'

//INSERT INTO arcon2.dbo.movisto --tabla destion
//SELECT  * FROM arcon.dbo.movisto -- tabla origen
//where fechmov >= '01/MAY/2006' AND FECHMOV < '01/MAY/2008'


//crear los indices mediante el script
//DIO ERROR QUE LA TABLA YA EXISTE PERO CREO LOS INDICES, HAY QEU PROBAR SI SE PUEDE EJECUTAR LOS INDICES SOLAMENTE
///****** Object:  Index [ICantIngre]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICantIngre] ON[dbo].[MOVISTO]
//(
//    [CantIngre] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [ICantSalid]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICantSalid] ON[dbo].[MOVISTO]
//(
//    [CantSalid] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [ICod_Inte]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICod_Inte] ON[dbo].[MOVISTO]
//(
//    [Cod_Inte] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [ICodiEmpr]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICodiEmpr] ON[dbo].[MOVISTO]
//(
//    [CodiEmpr] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [ICodiMovi]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICodiMovi] ON[dbo].[MOVISTO]
//(
//    [CodiMovi] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [IDepoMovi]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDepoMovi] ON[dbo].[MOVISTO]
//(
//    [DepoMovi] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IDoPriCor]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDoPriCor] ON[dbo].[MOVISTO]
//(
//    [DoPriCor] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IDoPriLar]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDoPriLar] ON[dbo].[MOVISTO]
//(
//    [DoPriLar] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [IDOPRINUM]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDOPRINUM] ON[dbo].[MOVISTO]
//(
//    [DOPRINUM] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IDoSecCor]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDoSecCor] ON[dbo].[MOVISTO]
//(
//    [DoSecCor] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IDoSecLar]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDoSecLar] ON[dbo].[MOVISTO]
//(
//    [DoSecLar] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [IFechMov]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IFechMov] ON[dbo].[MOVISTO]
//(
//    [FechMov] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IIdenLote]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IIdenLote] ON[dbo].[MOVISTO]
//(
//    [IdenLote] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [index_1733581214]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE UNIQUE NONCLUSTERED INDEX [index_1733581214] ON[dbo].[MOVISTO]
//(
//    [rowguid] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [INume_Clie]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [INume_Clie] ON[dbo].[MOVISTO]
//(
//    [Nume_Clie] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//GO
///****** Object:  Index [INUMPVTA]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [INUMPVTA] ON[dbo].[MOVISTO]
//(
//    [NUMPVTA] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//GO
///****** Object:  Index [INumUsuar]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [INumUsuar] ON[dbo].[MOVISTO]
//(
//    [NumUsuar] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [INuPedCli]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [INuPedCli] ON[dbo].[MOVISTO]
//(
//    [NuPedCli] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO

//--ELIMINE DESDE LA TABLA MOVISTO > INDEXES UN INDICE QUE TENIA NOMBRE index_1733581214

//Y LUEGO PUDE ELIMINAR EL CAMPO ROWGUID
//-- 3. Eliminar la columna

//ALTER TABLE movisto DROP COLUMN rowguid;


//--Cambiar columnas FLOAT a DECIMAL(18,4)
//ALTER TABLE dbo.movisto ALTER COLUMN CostoRep     DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN ValoUnit     DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN TipoCam      DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN CantIngre    DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN CantSalid    DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN CantSaldo    DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN ValoUnitIva  DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN ValTotIt     DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN ValTotItIva  DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN PESOKG       DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN METROS       DECIMAL(18,4) NULL;

//--Cambiar Cod_Inte de SMALLINT a INT
//ALTER TABLE dbo.movisto ALTER COLUMN Cod_Inte INT NULL;


//Cod_inte, CantSalid y Cantingre generaban error 
//borrar el indice DROP INDEX ICantSalid ON dbo.movisto;

//cambiar el tipo de dato a decimal : ALTER TABLE dbo.movisto ALTER COLUMN CantSalid    DECIMAL(18,4) NULL;

//volver a generar el indice: CREATE INDEX ICantSalid ON dbo.movisto (CantSalid);


//lo mismo con los otros 2

//cambiar el depomovi a int
//DROP INDEX IDepoMovi ON dbo.movisto;

//ALTER TABLE dbo.movisto
//ALTER COLUMN DepoMovi INT NULL;

//CREATE NONCLUSTERED INDEX IDepoMovi
//ON dbo.movisto (DepoMovi); PRIMERO PARA QUE SE CREE LA TABLA ARCON2.DBO.MOVISTO Y SE MIGREN TODOS LOS DATOS < 01/MAY/2004
//SELECT * 
//INTO arcon2.dbo.movisto --tabla destion
//FROM arcon.dbo.movisto -- tabla origen
//where fechmov < '01/MAY/2004'


//INSERT INTO arcon2.dbo.movisto --tabla destion
//SELECT  * FROM arcon.dbo.movisto -- tabla origen
//where fechmov >= '01/MAY/2004' AND FECHMOV < '01/MAY/2006'

//INSERT INTO arcon2.dbo.movisto --tabla destion
//SELECT  * FROM arcon.dbo.movisto -- tabla origen
//where fechmov >= '01/MAY/2006' AND FECHMOV < '01/MAY/2008'


//crear los indices mediante el script
//DIO ERROR QUE LA TABLA YA EXISTE PERO CREO LOS INDICES, HAY QEU PROBAR SI SE PUEDE EJECUTAR LOS INDICES SOLAMENTE
///****** Object:  Index [ICantIngre]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICantIngre] ON[dbo].[MOVISTO]
//(
//    [CantIngre] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [ICantSalid]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICantSalid] ON[dbo].[MOVISTO]
//(
//    [CantSalid] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [ICod_Inte]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICod_Inte] ON[dbo].[MOVISTO]
//(
//    [Cod_Inte] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [ICodiEmpr]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICodiEmpr] ON[dbo].[MOVISTO]
//(
//    [CodiEmpr] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [ICodiMovi]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [ICodiMovi] ON[dbo].[MOVISTO]
//(
//    [CodiMovi] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [IDepoMovi]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDepoMovi] ON[dbo].[MOVISTO]
//(
//    [DepoMovi] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IDoPriCor]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDoPriCor] ON[dbo].[MOVISTO]
//(
//    [DoPriCor] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IDoPriLar]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDoPriLar] ON[dbo].[MOVISTO]
//(
//    [DoPriLar] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [IDOPRINUM]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDOPRINUM] ON[dbo].[MOVISTO]
//(
//    [DOPRINUM] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IDoSecCor]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDoSecCor] ON[dbo].[MOVISTO]
//(
//    [DoSecCor] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IDoSecLar]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IDoSecLar] ON[dbo].[MOVISTO]
//(
//    [DoSecLar] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [IFechMov]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IFechMov] ON[dbo].[MOVISTO]
//(
//    [FechMov] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
//SET ANSI_PADDING ON
//GO
///****** Object:  Index [IIdenLote]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [IIdenLote] ON[dbo].[MOVISTO]
//(
//    [IdenLote] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [index_1733581214]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE UNIQUE NONCLUSTERED INDEX [index_1733581214] ON[dbo].[MOVISTO]
//(
//    [rowguid] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [INume_Clie]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [INume_Clie] ON[dbo].[MOVISTO]
//(
//    [Nume_Clie] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//GO
///****** Object:  Index [INUMPVTA]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [INUMPVTA] ON[dbo].[MOVISTO]
//(
//    [NUMPVTA] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//GO
///****** Object:  Index [INumUsuar]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [INumUsuar] ON[dbo].[MOVISTO]
//(
//    [NumUsuar] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO
///****** Object:  Index [INuPedCli]    Script Date: 24/06/2025 11:50:58 ******/
//CREATE NONCLUSTERED INDEX [INuPedCli] ON[dbo].[MOVISTO]
//(
//    [NuPedCli] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON[PRIMARY]
//GO

//--ELIMINE DESDE LA TABLA MOVISTO > INDEXES UN INDICE QUE TENIA NOMBRE index_1733581214

//Y LUEGO PUDE ELIMINAR EL CAMPO ROWGUID
//-- 3. Eliminar la columna

//ALTER TABLE movisto DROP COLUMN rowguid;


//--Cambiar columnas FLOAT a DECIMAL(18,4)
//ALTER TABLE dbo.movisto ALTER COLUMN CostoRep     DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN ValoUnit     DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN TipoCam      DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN CantIngre    DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN CantSalid    DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN CantSaldo    DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN ValoUnitIva  DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN ValTotIt     DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN ValTotItIva  DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN PESOKG       DECIMAL(18,4) NULL;
//ALTER TABLE dbo.movisto ALTER COLUMN METROS       DECIMAL(18,4) NULL;

//--Cambiar Cod_Inte de SMALLINT a INT
//ALTER TABLE dbo.movisto ALTER COLUMN Cod_Inte INT NULL;

//--Cambiar NumUsuar de SMALLINT a INT
//ALTER TABLE dbo.movisto ALTER COLUMN NumUsuar INT NULL;


//Cod_inte, CantSalid y Cantingre generaban error 
//borrar el indice DROP INDEX ICantSalid ON dbo.movisto;

//cambiar el tipo de dato a decimal : ALTER TABLE dbo.movisto ALTER COLUMN CantSalid    DECIMAL(18,4) NULL;

//volver a generar el indice: CREATE INDEX ICantSalid ON dbo.movisto (CantSalid);


//lo mismo con los otros 2

//cambiar el depomovi a int
//DROP INDEX IDepoMovi ON dbo.movisto;

//ALTER TABLE dbo.movisto
//ALTER COLUMN DepoMovi INT NULL;

//CREATE NONCLUSTERED INDEX IDepoMovi
//ON dbo.movisto (DepoMovi);

//quitar el indice DROP INDEX INumUsuar ON dbo.movisto

//modificar a int ALTER TABLE dbo.movisto ALTER COLUMN NumUsuar INT NULL;
//volver a crear el indice CREATE INDEX INumUsuar ON dbo.movisto (NumUsuar);

//ALTER TABLE movisto ADD CierreId INT NULL;

//ALTER TABLE MOVISTO ALTER COLUMN Nume_Prov INT;

/*
-- INSTRUCCIONES SQL PARA IMPLEMENTAR AUDITORÍA COMPLETA EN MOVISTO:

-- 1. Modificar CodiEmpr de short a int para compatibilidad con IAuditableEntity
ALTER TABLE movisto ALTER COLUMN CodiEmpr INT NULL;

-- 2. Agregar campos de timestamp automáticos
ALTER TABLE movisto 
ADD [AddRecord] datetime2 NOT NULL DEFAULT GETDATE(),
    [LastUpdate] datetime2 NOT NULL DEFAULT GETDATE();

-- 3. Crear índices para optimizar consultas de auditoría
CREATE INDEX IX_Movisto_AddRecord ON movisto (AddRecord);
CREATE INDEX IX_Movisto_LastUpdate ON movisto (LastUpdate);
CREATE INDEX IX_Movisto_NumUsuar ON movisto (NumUsuar);
CREATE INDEX IX_Movisto_CodiEmpr ON movisto (CodiEmpr);

-- 4. Crear índice compuesto para consultas de auditoría por usuario y fecha
CREATE INDEX IX_Movisto_Auditoria ON movisto (NumUsuar, CodiEmpr, LastUpdate);

-- 5. Comentarios en los nuevos campos
EXEC sys.sp_addextendedproperty 
    @name=N'MS_Description', 
    @value=N'Fecha de creación del registro (automática)', 
    @level0type=N'SCHEMA', @level0name=N'dbo', 
    @level1type=N'TABLE', @level1name=N'movisto', 
    @level2type=N'COLUMN', @level2name=N'AddRecord';

EXEC sys.sp_addextendedproperty 
    @name=N'MS_Description', 
    @value=N'Fecha de última actualización del registro (automática)', 
    @level0type=N'SCHEMA', @level0name=N'dbo', 
    @level1type=N'TABLE', @level1name=N'movisto', 
    @level2type=N'COLUMN', @level2name=N'LastUpdate';

EXEC sys.sp_addextendedproperty 
    @name=N'MS_Description', 
    @value=N'Código de empresa del usuario que creó/modificó el registro (automático)', 
    @level0type=N'SCHEMA', @level0name=N'dbo', 
    @level1type=N'TABLE', @level1name=N'movisto', 
    @level2type=N'COLUMN', @level2name=N'CodiEmpr';

EXEC sys.sp_addextendedproperty 
    @name=N'MS_Description', 
    @value=N'ID del usuario que creó/modificó el registro (automático)', 
    @level0type=N'SCHEMA', @level0name=N'dbo', 
    @level1type=N'TABLE', @level1name=N'movisto', 
    @level2type=N'COLUMN', @level2name=N'NumUsuar';
*/
