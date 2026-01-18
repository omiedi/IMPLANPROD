using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el cierre avanzado de órdenes de fabricación
    /// Contiene todos los datos necesarios para procesar el cierre con persistencia completa
    /// </summary>
    public class DatosCierreAvanzadoDTO
    {
        /// <summary>
        /// ID de la orden de fabricación a cerrar
        /// </summary>
        [Required]
        public int IdOrden { get; set; }

        /// <summary>
        /// ID de la suborden (IdSubOr) para las anotaciones
        /// </summary>
        [Required]
        public string IdSubOr { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad aprobada en este cierre
        /// </summary>
        [Required]
        [Range(0, double.MaxValue, ErrorMessage = "La cantidad aprobada debe ser mayor o igual a 0")]
        public decimal CantidadAprobada { get; set; }

        /// <summary>
        /// Cantidad rechazada/scrap en este cierre
        /// </summary>
        [Range(0, double.MaxValue, ErrorMessage = "La cantidad rechazada debe ser mayor o igual a 0")]
        public decimal CantidadRechazada { get; set; }

        /// <summary>
        /// Depósito de entrada para el producto terminado
        /// </summary>
        public int? DepositoEntrada { get; set; }

        /// <summary>
        /// Depósito de salida/consumo para los componentes
        /// </summary>
        public int? DepositoSalida { get; set; }

        /// <summary>
        /// Depósito para el material scrap/rechazado
        /// </summary>
        public int? DepositoScrap { get; set; }

        /// <summary>
        /// Fecha del cierre
        /// </summary>
        [Required]
        public DateTime FechaCierre { get; set; }

        /// <summary>
        /// Indica si la orden se cierra completamente (1) o se mantiene abierta (0)
        /// Basado en la validación de reservas y saldo pendiente
        /// </summary>
        [Required]
        public int CierraCompleto { get; set; }

        /// <summary>
        /// Lista de componentes con las cantidades editadas por el usuario
        /// </summary>
        public List<ComponenteCierreDTO> Componentes { get; set; } = new List<ComponenteCierreDTO>();

        /// <summary>
        /// Modo de cierre de calidad (equivalente a ModoCierreCalidad% del VB6)
        /// 1 = Modo calidad activo, 0 = Modo normal
        /// </summary>
        public int ModoCierreCalidad { get; set; } = 0;
    }

    /// <summary>
    /// DTO para los componentes en el cierre con cantidades editables
    /// </summary>
    public class ComponenteCierreDTO
    {
        /// <summary>
        /// Código interno del componente
        /// </summary>
        public int CodIElem { get; set; }

        /// <summary>
        /// Código del elemento/componente
        /// </summary>
        public string CodXElem { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del componente
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Uso bruto unitario del componente
        /// </summary>
        public decimal UsoBruto { get; set; }

        /// <summary>
        /// Cantidad a utilizar (editable por el usuario)
        /// </summary>
        public decimal CantidadAUtilizar { get; set; }

        /// <summary>
        /// Stock actual del componente
        /// </summary>
        public decimal Stock { get; set; }
    }
}
