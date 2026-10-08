namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Fila del reporte "Órdenes de Compra Pendientes" (réplica del reporte VB6
    /// de OCOMDETA: Stat_Ocom &lt; 3, &lt;&gt; -1 y Cant_Rec &lt; Cant_Ocom - 0.001).
    /// Los campos replican el Rst del VB6:
    ///   Descripcion  = DESCRITEM, o DELIBRE cuando Cod_Inte &lt;= 0 (ítem libre)
    ///   Moneda       = "" si Mone_Emis &lt;= 1, "U$S" si = 2, sino TAMONED.Descr (3 chars)
    ///   CantPend     = Cant_Ocom - Cant_Rec (si &lt; 0.05 se muestra 0)
    ///   PreUnidPesos = PrUn_Neto * tipoCambio (TAMONED.Cotizacion; 1 si moneda local)
    ///   ValorPend    = PreUnidPesos * CantPend
    /// </summary>
    public class OcomPendienteDTO
    {
        /// <summary>Número de orden de compra (nume_ocom)</summary>
        public int? NumeOcom { get; set; }

        /// <summary>Fecha de emisión (Femi_Ocom)</summary>
        public DateTime? FemiOcom { get; set; }

        /// <summary>Número de proveedor (Npro_Ocom)</summary>
        public short? NproOcom { get; set; }

        /// <summary>Razón social del proveedor (RaSo_Prov) - para encabezado de grupo</summary>
        public string? RaSoProv { get; set; }

        /// <summary>Código externo del producto (Cod_Exte)</summary>
        public string? CodExte { get; set; }

        /// <summary>Descripción: DESCRITEM, o DELIBRE si el ítem es libre (Cod_Inte &lt;= 0)</summary>
        public string? Descripcion { get; set; }

        /// <summary>Cantidad comprada (Cant_Ocom)</summary>
        public decimal? CantOcom { get; set; }

        /// <summary>Unidad de compra (Uni_Com)</summary>
        public string? UniCom { get; set; }

        /// <summary>Moneda abreviada ("", "U$S" o descripción TAMONED de 3 chars)</summary>
        public string Moneda { get; set; } = string.Empty;

        /// <summary>Precio unitario neto en moneda de emisión (PrUn_Neto)</summary>
        public decimal? PrunNeto { get; set; }

        /// <summary>Importe = Cant_Ocom * PrUn_Neto</summary>
        public decimal Importe { get; set; }

        /// <summary>Fecha de entrega solicitada (Fentre_Sol)</summary>
        public DateTime? FentreSol { get; set; }

        /// <summary>Cantidad recibida (Cant_Rec)</summary>
        public decimal? CantRec { get; set; }

        /// <summary>Cantidad pendiente = Cant_Ocom - Cant_Rec (&lt; 0.05 → 0)</summary>
        public decimal CantPend { get; set; }

        /// <summary>Precio unitario neto en pesos = PrUn_Neto * tipoCambio</summary>
        public decimal PreUnidPesos { get; set; }

        /// <summary>Valor pendiente = PreUnidPesos * CantPend</summary>
        public decimal ValorPend { get; set; }

        /// <summary>
        /// Días de atraso respecto de Fentre_Sol (solo reporte Pendientes/Vencidas:
        /// se calcula cuando Fentre_Sol &lt; hoy; en otro caso queda null).
        /// </summary>
        public int? DiasVencido { get; set; }
    }
}
