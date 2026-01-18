using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para transportar información del último destino y sector utilizados por un proveedor
    /// Se utiliza para cargar automáticamente estos valores al seleccionar un proveedor en órdenes de compra
    /// </summary>
    public class UltimoDestinoSectorDTO
    {
        /// <summary>
        /// Último destino utilizado para el proveedor (código de cuenta del plan de cuentas)
        /// Corresponde al campo CodiEmpr de OCOMENCA
        /// </summary>
        public int? UltimoDestino { get; set; }

        /// <summary>
        /// Último sector utilizado para el proveedor (máximo 12 caracteres)
        /// Corresponde al campo SectOcom de OCOMDETA
        /// </summary>
        [MaxLength(12)]
        public string UltimoSector { get; set; } = string.Empty;

        /// <summary>
        /// Fecha de emisión de la última orden de compra encontrada
        /// Se utiliza para información y logging
        /// </summary>
        public DateTime? FechaUltimaOrden { get; set; }

        /// <summary>
        /// Número de la última orden de compra encontrada
        /// Se utiliza para información y logging
        /// </summary>
        public int? NumeroUltimaOrden { get; set; }

        /// <summary>
        /// Constructor por defecto que inicializa valores vacíos
        /// </summary>
        public UltimoDestinoSectorDTO()
        {
        }

        /// <summary>
        /// Constructor con parámetros para inicializar todos los valores
        /// </summary>
        /// <param name="ultimoDestino">Código del último destino utilizado</param>
        /// <param name="ultimoSector">Último sector utilizado</param>
        /// <param name="fechaUltimaOrden">Fecha de la última orden</param>
        /// <param name="numeroUltimaOrden">Número de la última orden</param>
        public UltimoDestinoSectorDTO(int? ultimoDestino, string ultimoSector, DateTime? fechaUltimaOrden, int? numeroUltimaOrden)
        {
            UltimoDestino = ultimoDestino;
            UltimoSector = ultimoSector ?? string.Empty;
            FechaUltimaOrden = fechaUltimaOrden;
            NumeroUltimaOrden = numeroUltimaOrden;
        }

        /// <summary>
        /// Indica si el DTO contiene datos válidos (al menos destino o sector)
        /// </summary>
        public bool TieneDatos => UltimoDestino.HasValue || !string.IsNullOrEmpty(UltimoSector);

        /// <summary>
        /// Representación en string para debugging
        /// </summary>
        /// <returns>String con información del DTO</returns>
        public override string ToString()
        {
            return $"UltimoDestinoSectorDTO: Destino={UltimoDestino}, Sector='{UltimoSector}', Fecha={FechaUltimaOrden:yyyy-MM-dd}, Orden={NumeroUltimaOrden}";
        }
    }
}
