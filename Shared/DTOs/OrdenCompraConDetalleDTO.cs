using IMPLANPROD.Shared.Entities;
using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para manejar la orden de compra completa con su detalle
    /// Incluye validaciones para asegurar integridad de datos
    /// </summary>
    public class OrdenCompraConDetalleDTO
    {
        /// <summary>
        /// Encabezado de la orden de compra
        /// </summary>
        [Required(ErrorMessage = "El encabezado de la orden es requerido")]
        public Ocomenca Encabezado { get; set; } = new();

        /// <summary>
        /// Detalle de items de la orden de compra
        /// </summary>
        [Required(ErrorMessage = "El detalle de la orden es requerido")]
        [MinLength(1, ErrorMessage = "Debe incluir al menos un item en la orden")]
        public List<Ocomdeta> Detalle { get; set; } = new();

        /// <summary>
        /// Distribución de IVA por tasas aplicadas
        /// </summary>
        public List<DetaivaDTO> DistribucionIva { get; set; } = new();

        /// <summary>
        /// Indica si es una edición (true) o una nueva orden (false)
        /// </summary>
        public bool EsEdicion { get; set; } = false;

        /// <summary>
        /// Valida la consistencia de los datos de la orden completa
        /// </summary>
        /// <returns>Lista de errores encontrados</returns>
        public List<string> Validar()
        {
            var errores = new List<string>();

            // Validar encabezado
            if (Encabezado == null)
            {
                errores.Add("El encabezado de la orden es requerido");
                return errores;
            }

            // Validar proveedor
            if (!Encabezado.NproOcom.HasValue || Encabezado.NproOcom <= 0)
            {
                errores.Add("Debe seleccionar un proveedor");
            }

            // Validar detalle
            if (Detalle == null || !Detalle.Any())
            {
                errores.Add("Debe incluir al menos un item en la orden");
                return errores;
            }

            // Validar items del detalle
            for (int i = 0; i < Detalle.Count; i++)
            {
                var item = Detalle[i];
                var prefijo = $"Item {i + 1}:";

                if (!item.CodInte.HasValue || item.CodInte <= 0)
                {
                    errores.Add($"{prefijo} Debe seleccionar un producto");
                }

                if (!item.CantOcom.HasValue || item.CantOcom <= 0)
                {
                    errores.Add($"{prefijo} La cantidad debe ser mayor a cero");
                }

                if (!item.PreUnit.HasValue || item.PreUnit <= 0)
                {
                    errores.Add($"{prefijo} El precio unitario debe ser mayor a cero");
                }

                if (string.IsNullOrWhiteSpace(item.Descritem))
                {
                    errores.Add($"{prefijo} La descripción del producto es requerida");
                }
            }

            // Validar consistencia de números de orden
            if (Encabezado.NumeOcom.HasValue)
            {
                foreach (var item in Detalle)
                {
                    if (item.NumeOcom.HasValue && item.NumeOcom != Encabezado.NumeOcom)
                    {
                        errores.Add("Inconsistencia en números de orden entre encabezado y detalle");
                        break;
                    }
                }
            }

            return errores;
        }

        /// <summary>
        /// Calcula el total neto de la orden (suma de todos los items)
        /// </summary>
        public decimal CalcularTotalNeto()
        {
            return Detalle?.Sum(d => (d.CantOcom ?? 0) * (d.PreUnit ?? 0)) ?? 0;
        }

        /// <summary>
        /// Calcula el total de IVA de la orden
        /// </summary>
        public decimal CalcularTotalIva()
        {
            return DistribucionIva?.Sum(d => d.ImpoIva ?? 0) ?? 0;
        }

        /// <summary>
        /// Calcula el total final de la orden (neto + IVA)
        /// </summary>
        public decimal CalcularTotalFinal()
        {
            return CalcularTotalNeto() + CalcularTotalIva();
        }
    }

    /// <summary>
    /// DTO para el detalle de IVA
    /// </summary>
    public class DetaivaDTO
    {
        /// <summary>
        /// Número de orden de IVA (0-5)
        /// </summary>
        public int? Norit { get; set; }

        /// <summary>
        /// Tasa de IVA aplicada (ej: 21.00 para 21%)
        /// </summary>
        public decimal? TasaIva { get; set; }

        /// <summary>
        /// Base imponible para esta tasa
        /// </summary>
        public decimal? BaseImpo { get; set; }

        /// <summary>
        /// Importe de IVA calculado
        /// </summary>
        public decimal? ImpoIva { get; set; }

        /// <summary>
        /// Calcula el importe de IVA basado en la base imponible y la tasa
        /// </summary>
        public void CalcularImporteIva()
        {
            if (BaseImpo.HasValue && TasaIva.HasValue)
            {
                ImpoIva = BaseImpo.Value * (TasaIva.Value / 100);
            }
            else
            {
                ImpoIva = 0;
            }
        }
    }
}
