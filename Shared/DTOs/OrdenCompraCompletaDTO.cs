using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para crear/actualizar una orden de compra completa con distribución de IVA
    /// Incluye encabezado, detalle y distribución de IVA
    /// </summary>
    public class OrdenCompraCompletaDTO
    {
        /// <summary>
        /// Datos del encabezado de la orden de compra
        /// </summary>
        public required Ocomenca Encabezado { get; set; }

        /// <summary>
        /// Lista de items del detalle de la orden
        /// </summary>
        public List<Ocomdeta> Detalle { get; set; } = new List<Ocomdeta>();

        /// <summary>
        /// Distribución de IVA de la orden
        /// Equivalente a los arrays IMPOGRA e IMPOIVA del VB6
        /// </summary>
        public DistribucionIvaDTO? DistribucionIva { get; set; }

        /// <summary>
        /// Código de empresa para los registros DETAIVA
        /// Equivalente a CodiEmp% del VB6
        /// </summary>
        public int? CodigoEmpresa { get; set; }

        /// <summary>
        /// Indica si se debe actualizar la distribución de IVA
        /// </summary>
        public bool ActualizarIva { get; set; } = true;

        /// <summary>
        /// Valida que la orden completa sea correcta
        /// </summary>
        /// <returns>Lista de errores de validación</returns>
        public List<string> Validar()
        {
            var errores = new List<string>();

            // Validar encabezado
            if (Encabezado == null)
            {
                errores.Add("El encabezado de la orden es obligatorio");
                return errores;
            }

            // Validar número de orden
            if (!Encabezado.NumeOcom.HasValue || Encabezado.NumeOcom <= 0)
            {
                errores.Add("El número de orden debe ser mayor a cero");
            }

            // Validar proveedor
            if (!Encabezado.NproOcom.HasValue || Encabezado.NproOcom <= 0)
            {
                errores.Add("Debe seleccionar un proveedor válido");
            }

            // Validar fecha de emisión
            if (!Encabezado.FemiOcom.HasValue)
            {
                errores.Add("La fecha de emisión es obligatoria");
            }
            else if (Encabezado.FemiOcom.Value.Date > DateTime.Now.Date)
            {
                errores.Add("La fecha de emisión no puede ser mayor a la fecha actual");
            }

            // Validar distribución de IVA si está presente
            if (DistribucionIva != null && ActualizarIva)
            {
                // El TotalOcom ya viene con el descuento aplicado desde el frontend
                var totalConDescuento = Encabezado.TotalOcom ?? 0m;
                if (!DistribucionIva.EsValida(totalConDescuento))
                {
                    errores.Add($"La distribución de IVA no coincide con el total de la orden. " +
                               $"Total esperado: {totalConDescuento:C2}, " +
                               $"Total distribuido: {DistribucionIva.TotalBaseImponible:C2}");
                }
            }

            return errores;
        }
    }
}
