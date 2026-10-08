namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Renglón del informe de Cumplimiento de Entregas (COCUENT09 VB6: IC-R / IC-P).
    /// Reemplaza el registro COCUENT del archivo de trabajo COCUSEL.DAT.
    /// </summary>
    public class CumplimientoEntregaDTO
    {
        public int NumeClie { get; set; }
        public string RazonSocial { get; set; } = string.Empty;

        /// <summary>Número de remito: Val(Mid(DoPriCor, 4)) como en VB6.</summary>
        public long NroRemito { get; set; }
        /// <summary>Documento largo del remito (MOVISTO.DoPriLar).</summary>
        public string RemitoLargo { get; set; } = string.Empty;
        public DateTime? FechaRemito { get; set; }

        public int NroPedido { get; set; }
        public DateTime? FechaPedido { get; set; }
        public DateTime? FechaSolicitada { get; set; }

        public int CodInte { get; set; }
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public string Lote { get; set; } = string.Empty;
        public decimal Cantidad { get; set; }

        public int PlazoSolicitado { get; set; }
        public int PlazoReal { get; set; }
        public int Atraso { get; set; }
        public decimal Puntaje { get; set; }

        public string FechaRemitoTexto => FechaRemito?.ToString("dd/MM/yy") ?? string.Empty;
        public string FechaPedidoTexto => FechaPedido?.ToString("dd/MM/yy") ?? string.Empty;
        public string FechaSolicitadaTexto => FechaSolicitada?.ToString("dd/MM/yy") ?? string.Empty;
        public string NroPedidoTexto => NroPedido > 0 ? NroPedido.ToString() : string.Empty;
        public string AtrasoTexto => Atraso > 0 ? Atraso.ToString() : string.Empty;
    }

    /// <summary>Índice de cumplimiento acumulado por cliente.</summary>
    public class CumplimientoEntregaClienteDTO
    {
        public int NumeClie { get; set; }
        public string RazonSocial { get; set; } = string.Empty;
        public int Renglones { get; set; }
        public decimal Cantidad { get; set; }
        public int TotalAtraso { get; set; }
        public decimal TotalPuntaje { get; set; }
        public decimal IndiceCumplimiento { get; set; }
    }

    /// <summary>Resultado completo del informe (detalle + totales).</summary>
    public class CumplimientoEntregaResultadoDTO
    {
        /// <summary>"REMITO" (IC-R) o "PEDIDO" (IC-P).</summary>
        public string Tipo { get; set; } = "REMITO";
        public DateTime FechaDesde { get; set; }
        public DateTime FechaHasta { get; set; }
        public int ToleranciaDias { get; set; }
        public bool ToleranciaAplicada { get; set; }
        public int DepositoExpedicion { get; set; }
        public string? Advertencia { get; set; }

        public List<CumplimientoEntregaDTO> Detalle { get; set; } = new();
        public List<CumplimientoEntregaClienteDTO> Clientes { get; set; } = new();

        public decimal TotalCantidad { get; set; }
        public int TotalAtraso { get; set; }
        public decimal TotalPuntaje { get; set; }
        public decimal IndiceCumplimiento { get; set; }
    }

    /// <summary>Tolerancia de entrega en días (FLEXCRL PROBUS='COCUENT0', CLABUS='TOLERANC').</summary>
    public class ToleranciaEntregaDTO
    {
        public int Dias { get; set; }
    }
}
