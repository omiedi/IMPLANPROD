namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para previsualización de registros del archivo legacy REPLACOP.DAT.
    ///
    /// Estructura del archivo REPLACOP.DAT (128 bytes por registro):
    /// - Posición  1, longitud  2: Número de cliente (Int16 - CVI en VB6)
    /// - Posición  3, longitud  2: Código interno    (Int16 - CVI en VB6)
    /// - Posición  5, longitud 28: Código de cliente (string, Windows-1252)
    /// - Posición 33, longitud 96: Observaciones     (string, Windows-1252)
    ///
    /// CVI en VB6 = BitConverter.ToInt16 en C# (little-endian, 2 bytes).
    /// </summary>
    public class ReplacopPreviewDTO
    {
        /// <summary>Número de cliente (CVI, pos 1-2)</summary>
        public short NumeroCliente { get; set; }

        /// <summary>Código interno (CVI, pos 3-4)</summary>
        public short CodigoInterno { get; set; }

        /// <summary>Código de cliente (string, pos 5-32, 28 chars)</summary>
        public string CodigoCliente { get; set; } = "";

        /// <summary>Observaciones (string, pos 33-128, 96 chars)</summary>
        public string Observaciones { get; set; } = "";
    }
}
