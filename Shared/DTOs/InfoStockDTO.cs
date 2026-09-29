namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Fila del reporte InfoStock: stock por depósito o inventario general.
    /// Fuente: MOVISTO agrupado por Cod_Inte/Cod_Exte (saldo = Σ CantIngre − CantSalid),
    /// enriquecido con MASTER (Umedida, Timate→Tadeposi.Descripcion como tipo de material).
    /// </summary>
    public class InfoStockDTO
    {
        /// <summary>Código interno del producto (MASTER.Codint)</summary>
        public int? CodigoInterno { get; set; }

        /// <summary>Código del producto (MOVISTO.Cod_Exte)</summary>
        public string? Codigo { get; set; }

        /// <summary>Descripción del producto (MOVISTO.Descrip)</summary>
        public string? Descripcion { get; set; }

        /// <summary>Unidad de medida (UnidadMedidas.DescripcionAbreviada)</summary>
        public string? UMedida { get; set; }

        /// <summary>Tipo de material (Tadeposi.Descripcion, vía MASTER.Timate)</summary>
        public string? TMaterial { get; set; }

        /// <summary>Saldo = Σ(CantIngre − CantSalid), nulos como 0</summary>
        public decimal Saldo { get; set; }
    }
}
