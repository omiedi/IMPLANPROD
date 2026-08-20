using System;
using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar datos en la vista de índice de órdenes de fabricación
    /// Contiene solo los campos necesarios para la tabla de OrdenFabricacionIndex.razor
    /// </summary>
    public class OrdefabrIndexDTO
    {
        /// <summary>
        /// Identificador único de la orden de fabricación
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Fecha de generación de la orden
        /// </summary>
        public DateTime? FechGen { get; set; }

        /// <summary>
        /// Identificador de la suborden
        /// </summary>
        public string? IdSubOr { get; set; }

        /// <summary>
        /// Cantidad proyectada
        /// </summary>
        public decimal? CanProy { get; set; }

        /// <summary>
        /// Cantidad aprobada
        /// </summary>
        public decimal? CanApro { get; set; }

        /// <summary>
        /// Identificador de la orden de fabricación (código alfanumérico)
        /// </summary>
        public string? Id_Orfa { get; set; }

        /// <summary>
        /// Estado de la orden de fabricación
        /// </summary>
        public short? StatuOrf { get; set; }

        /// <summary>
        /// Código externo del producto a fabricar
        /// </summary>
        public string? Cod_Exte { get; set; }

        public string? Unidad { get; set; }

        public DateTime? FechEnSol { get; set; }

        public DateTime? FechLan { get; set; }

        public DateTime? FechCierr { get; set; }

        /// <summary>
        /// Descripción del producto a fabricar
        /// </summary>
        public string? Descrip { get; set; }

        /// <summary>
        /// Código interno del producto a fabricar (necesario para consultar fórmula)
        /// </summary>
        public int? Cod_Inte { get; set; }

        /// <summary>
        /// Cantidad de cierres realizados (antes era Cant.Aprob.)
        /// </summary>
        public decimal? Cierres { get; set; }

        /// <summary>
        /// Suma de cantidades salidas desde ORFAPEN (nueva columna Cant.Aprob.)
        /// </summary>
        public decimal? CantAprobOrfapen { get; set; }

        /// <summary>
        /// Cantidad pendiente de aprobar (Cierres - CantAprobOrfapen)
        /// </summary>
        public decimal? PendienteAprobar { get; set; }

        /// <summary>
        /// Cantidad cerrada desde ORFAPEN (CantIngre con DOPRICOR='la orden de fabricacion activa' y CODIMOVI='OF')
        /// </summary>
        public decimal? CantidadCerrada { get; set; }

        /// <summary>
        /// Cantidad aprobada desde ORFAPEN (CantSalid con DOPRICOR='la orden de fabricacion activa' y CODIMOVI='OF')
        /// </summary>
        public decimal? CantidadAprobadaOrfapen { get; set; }

        /// <summary>
        /// Cantidad pendiente de aprobación desde ORFAPEN (CantIngre - CantSalid)
        /// </summary>
        public decimal? PendienteAprobacion { get; set; }

        public int? DepoEntrada { get; set; }
        public int? DepoSalida { get; set; }
        public int? DepoScrap { get; set; }
        public decimal? CantAprobada { get; set; }
        public decimal? CantRechazada { get; set; }

        /// <summary>
        /// Anotaciones de la orden de fabricación
        /// </summary>
        public string? AnotaOrf { get; set; }

    }
}
