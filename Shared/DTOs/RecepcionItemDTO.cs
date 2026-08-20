using System;

namespace IMPLANPROD.Shared.DTOs
{
    public class RecepcionItemDTO
    {
        public int CodigoInterno { get; set; }
        public int? NumeBore { get; set; } = 0;
        public int? NumeroOrdenCompra { get; set; }
        public DateTime? FechaEntregaSolicitada { get; set; }
        public DateTime? FechaRecepcion { get; set; }
        public string CodigoExterno { get; set; }
        public string DescripcionItem { get; set; }
        public decimal CantidadNominal { get; set; }
        public string Unidad { get; set; }
        public decimal? CantidadRechazada { get; set; } = 0;
        public decimal? CantidadFaltante { get; set; } = 0;
        public decimal? CanToApro { get; set; } = 0;
        // Depósito para el ingreso normal
        public int CodigoDeposito { get; set; } = 0;
        public string NombreDeposito { get; set; } = "";
        // Depósito para la cantidad rechazada
        public int CodigoDepositoRechazo { get; set; } = 0;
        public string NombreDepositoRechazo { get; set; } = "";

        public decimal? CanToRech { get; set; } = 0;
        public decimal? CanToFalta { get; set; } = 0;
        public int Status { get; set; } = 0;

        // Campos para rechazo
        public int? MotivoRechazo { get; set; }
        public int? DestinoRechazo { get; set; }
        public string? ObservacionesRechazo { get; set; }

    }
}
