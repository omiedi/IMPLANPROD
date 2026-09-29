namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para gestionar costos de items de órdenes de compra
    /// Se usa en la UI para mostrar y editar costos con control de acceso
    /// </summary>
    public class OCOMDETACostoDto
    {
        /// <summary>
        /// ID único del registro de costo
        /// </summary>
        public int OCOMDETA_ID { get; set; }

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        public int? NumeOcom { get; set; }

        /// <summary>
        /// Número de item dentro de la orden
        /// </summary>
        public short? NuOrdItem { get; set; }

        /// <summary>
        /// Código interno del producto
        /// </summary>
        public short? CodInte { get; set; }

        /// <summary>
        /// Código externo del producto (mostrar en UI)
        /// </summary>
        public string? CodExte { get; set; }

        /// <summary>
        /// Descripción del item (mostrar en UI)
        /// </summary>
        public string? Descritem { get; set; }

        /// <summary>
        /// Unidad de medida de compra (mostrar en UI)
        /// </summary>
        public string? UniCom { get; set; }

        /// <summary>
        /// Cantidad solicitada (mostrar en UI)
        /// </summary>
        public decimal? CantOcom { get; set; }

        /// <summary>
        /// Precio unitario bruto - costo de lista editable por ítem
        /// </summary>
        public decimal? PreBruto { get; set; }

        /// <summary>
        /// Porcentaje de descuento global (0-100) aplicado sobre el bruto.
        /// Se ingresa una sola vez en el modal y se persiste por fila.
        /// </summary>
        public decimal? PorcDesc { get; set; }

        /// <summary>
        /// Precio unitario neto - COSTO (persistido).
        /// Valor final = PreBruto * (1 - PorcDesc/100). Si no hay bruto
        /// cargado (registros legacy), se conserva el valor ya guardado.
        /// </summary>
        public decimal? PrUnNeto { get; set; }

        /// <summary>
        /// Observaciones del costo del item (máximo 1024 caracteres)
        /// </summary>
        [System.ComponentModel.DataAnnotations.MaxLength(1024)]
        public string? Observaciones { get; set; }

        /// <summary>
        /// Fecha de entrega solicitada (mostrar en UI)
        /// </summary>
        public DateTime? FentreSol { get; set; }

        /// <summary>
        /// Moneda de emisión (para buscar descripción en TAMONED)
        /// </summary>
        public short? MoneEmis { get; set; }

        /// <summary>
        /// Descripción de la moneda (calculado desde TAMONED)
        /// </summary>
        public string? DescripcionMoneda { get; set; }

        /// <summary>
        /// Estado del item (solo se muestran < 9)
        /// </summary>
        public short? StatOcom { get; set; }

        /// <summary>
        /// Precio unitario neto calculado = PreBruto * (1 - PorcDesc/100).
        /// Si PreBruto no está cargado, devuelve el neto ya persistido.
        /// Se redondea a 4 decimales (compatibilidad con decimal(18,4)).
        /// </summary>
        public decimal NetoCalculado =>
            Math.Round((PreBruto ?? PrUnNeto ?? 0m) * (1m - (PorcDesc ?? 0m) / 100m), 4);

        /// <summary>
        /// Importe calculado (Cantidad * Neto con descuento aplicado)
        /// Se calcula en el cliente
        /// </summary>
        public decimal? Importe => (CantOcom ?? 0) * NetoCalculado;
    }
}
