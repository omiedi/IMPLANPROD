using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el detalle de órdenes de compra en modo edición
    /// Incluye datos calculados y referencias a otras tablas según lógica VB6
    /// </summary>
    public class OcomdetaEditDTO
    {
        // ===== DATOS BÁSICOS DEL ITEM =====
        public int OCOMDETA_ID { get; set; }
        public int? NumeOcom { get; set; }
        public short? NuOrdItem { get; set; }
        public short? CodInte { get; set; }
        public string? CodExte { get; set; }
        public string? Descritem { get; set; }
        public string? Delibre { get; set; }

        // ===== CANTIDADES Y PRECIOS =====
        public decimal? CantOcom { get; set; }           // Cantidad original solicitada
        public decimal? CantRec { get; set; }            // Cantidad recibida
        public decimal? CantidadPendiente { get; set; }  // Calculado: CantOcom - CantRec
        public string? UniCom { get; set; }              // Unidad de compra
        public string? UniMas { get; set; }              // Unidad de medida del master
        public decimal? PreUnit { get; set; }            // Precio unitario
        public decimal? DescItem { get; set; }           // Descuento del item
        public decimal? PrUnNeto { get; set; }           // Precio unitario neto
        public decimal? ImporteTotal { get; set; }       // Calculado: CantOcom * PrUnNeto

        // ===== FECHAS =====
        public DateTime? FemiOcom { get; set; }          // Fecha de emisión
        public DateTime? FentreSol { get; set; }         // Fecha de entrega solicitada

        // ===== SECTOR Y DESTINO =====
        public string? SectOcom { get; set; }            // Sector
        public short? IContab { get; set; }              // Código de cuenta contable
        public string? DescripcionCuenta { get; set; }   // Descripción de la cuenta (desde PLANCUEN)

        // ===== OTROS CAMPOS =====
        public string? MediOcom { get; set; }            // Pedido interno/Medida
        public string? Consimat { get; set; }            // Consideración de material
        public short? StatOcom { get; set; }             // Estado del item
        public short? MoneEmis { get; set; }             // Moneda de emisión

        // ===== CAMPOS CALCULADOS SEGÚN VB6 =====
        public decimal? PrecioConTipoCambio { get; set; } // PreUnit * TipoCambio
        public bool EsMaterialTextoLibre { get; set; }    // CodInte <= 0
        public string? UnidadMedidaFinal { get; set; }    // UniMas o desde Master si está vacío

        // ===== PROPIEDADES DE SOLO LECTURA PARA LA UI =====
        public bool TieneCantidadPendiente => CantidadPendiente > 0.05m;
        public bool EsItemCompleto => CantidadPendiente <= 0.05m;
        public string EstadoDescripcion => StatOcom switch
        {
            < 0 => "Pendiente",
            0 => "Activo",
            >= 9 => "Anulado",
            _ => "Desconocido"
        };

        /// <summary>
        /// Convierte una entidad Ocomdeta a OcomdetaEditDTO
        /// </summary>
        public static OcomdetaEditDTO FromEntity(Ocomdeta entity, string? descripcionCuenta = null, string? unidadMedidaMaster = null, decimal tipoCambio = 1m)
        {
            var dto = new OcomdetaEditDTO
            {
                OCOMDETA_ID = entity.OCOMDETA_ID,
                NumeOcom = entity.NumeOcom,
                NuOrdItem = entity.NuOrdItem,
                CodInte = entity.CodInte,
                CodExte = entity.CodExte,
                Descritem = entity.Descritem,
                Delibre = entity.Delibre,
                CantOcom = entity.CantOcom,
                CantRec = entity.CantRec,
                UniCom = entity.UniCom,
                UniMas = entity.UniMas,
                PreUnit = entity.PreUnit,
                DescItem = entity.DescItem,
                PrUnNeto = entity.PrUnNeto,
                FemiOcom = entity.FemiOcom,
                FentreSol = entity.FentreSol,
                SectOcom = entity.SectOcom,
                IContab = entity.IContab,
                MediOcom = entity.MediOcom,
                Consimat = entity.Consimat,
                StatOcom = entity.StatOcom,
                MoneEmis = entity.MoneEmis,
                DescripcionCuenta = descripcionCuenta
            };

            // Cálculos según lógica VB6
            dto.CantidadPendiente = (dto.CantOcom ?? 0) - (dto.CantRec ?? 0);
            if (dto.CantidadPendiente < 0.05m) dto.CantidadPendiente = 0;

            dto.PrecioConTipoCambio = (dto.PrUnNeto ?? 0) * tipoCambio;
            dto.ImporteTotal = (dto.CantOcom ?? 0) * (dto.PrUnNeto ?? 0);
            dto.EsMaterialTextoLibre = (dto.CodInte ?? 0) <= 0;

            // Unidad de medida final: usar UniMas o la del master si está vacía
            dto.UnidadMedidaFinal = !string.IsNullOrEmpty(dto.UniMas) ? dto.UniMas : unidadMedidaMaster;

            return dto;
        }
    }
}
