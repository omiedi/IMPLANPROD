using System;
using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO con todos los datos necesarios para imprimir una Orden de Fabricación en formato A4.
    /// Se utiliza después de crear una orden en /orden-fabricacion-create.
    /// El controlador OrdeFabrController lo arma a partir de Ordefabr + Master + Usuarios + Secoper.
    /// </summary>
    public class OrdenFabricacionImpresionDTO
    {
        // --- Datos propios de la Orden de Fabricación ---

        /// <summary>Número interno de la orden (NumeOrFa).</summary>
        public int? NumeOrFa { get; set; }

        /// <summary>Número visible/alfanumérico de la orden (IdSubOr).</summary>
        public string? IdSubOr { get; set; }

        /// <summary>Fecha de generación de la orden (FechGen).</summary>
        public DateTime? FechGen { get; set; }

        /// <summary>Fecha de generación formateada para la impresión (dd/MM/yyyy).</summary>
        public string? FechGenFormateada { get; set; }

        /// <summary>Código externo del producto (Cod_Exte).</summary>
        public string? Cod_Exte { get; set; }

        /// <summary>Descripción del producto (Descrip).</summary>
        public string? Descrip { get; set; }

        /// <summary>Cantidad proyectada/lote (CanProy).</summary>
        public decimal? CanProy { get; set; }

        /// <summary>Cantidad proyectada/lote formateada para mostrar en el remito.</summary>
        public string? CanProyFormateada { get; set; }

        /// <summary>Código interno del producto (Cod_Inte).</summary>
        public int? Cod_Inte { get; set; }

        /// <summary>Unidad de medida del producto (Unidad).</summary>
        public string? Unidad { get; set; }

        // --- Datos del usuario que lanzó la orden ---

        /// <summary>IdFlexoft del usuario que lanzó la orden (NumUsuar).</summary>
        public int? NumUsuar { get; set; }

        /// <summary>Nombre del usuario obtenido de la tabla Usuarios (según IdFlexoft).</summary>
        public string? NombreUsuario { get; set; }

        // --- Datos del master / producto ---

        /// <summary>Fecha de revisión del master (Master.Ferevis) para el Codint de la orden.</summary>
        public DateTime? Ferevis { get; set; }

        /// <summary>Fecha de revisión formateada para la impresión (dd/MM/yyyy).</summary>
        public string? FerevisFormateada { get; set; }

        /// <summary>
        /// Revisión del producto tomada de MASTER.Revision (MASTER.Codint = ORDEFABR.Cod_Inte).
        /// Es el valor que se imprime en el casillero "Revisión" del formulario A4.
        /// </summary>
        public string? Revision { get; set; }

        // --- Detalle de operaciones (Secoper) ---

        /// <summary>Lista de operaciones de fabricación asociadas al producto.</summary>
        public List<OrdenFabricacionOperacionDTO> Operaciones { get; set; } = new();
    }

    /// <summary>
    /// Línea de operación incluida en la impresión de la Orden de Fabricación.
    /// Mapea la tabla SECOPER filtrada por CodIConj y StatOper = 1, ordenada por NumeOper.
    /// </summary>
    public class OrdenFabricacionOperacionDTO
    {
        /// <summary>Número de operación (NumeOper).</summary>
        public short? NumeOper { get; set; }

        /// <summary>Descripción de la operación (DescOper).</summary>
        public string? DescOper { get; set; }

        /// <summary>Observaciones de la operación (ObseOper). Se imprimen debajo de la descripción.</summary>
        public string? ObseOper { get; set; }
    }

    /// <summary>
    /// DTO con los datos del Vale de Materiales para una Orden de Fabricación.
    /// Mantiene el mismo encabezado que la Orden de Fabricación e incluye la lista de materiales reservados.
    /// </summary>
    public class OrdenFabricacionConsumoImpresionDTO
    {
        /// <summary>Número interno de la orden (NumeOrFa).</summary>
        public int? NumeOrFa { get; set; }

        /// <summary>Número visible/alfanumérico de la orden (IdSubOr).</summary>
        public string? IdSubOr { get; set; }

        /// <summary>Fecha de generación de la orden (FechGen).</summary>
        public DateTime? FechGen { get; set; }

        /// <summary>Fecha de generación formateada para la impresión (dd/MM/yyyy).</summary>
        public string? FechGenFormateada { get; set; }

        /// <summary>Código externo del producto (Cod_Exte).</summary>
        public string? Cod_Exte { get; set; }

        /// <summary>Descripción del producto (Descrip).</summary>
        public string? Descrip { get; set; }

        /// <summary>Cantidad proyectada/lote (CanProy).</summary>
        public decimal? CanProy { get; set; }

        /// <summary>Cantidad proyectada/lote formateada para mostrar en el remito.</summary>
        public string? CanProyFormateada { get; set; }

        /// <summary>Código interno del producto (Cod_Inte).</summary>
        public int? Cod_Inte { get; set; }

        /// <summary>Unidad de medida del producto (Unidad).</summary>
        public string? Unidad { get; set; }

        /// <summary>IdFlexoft del usuario que lanzó la orden (NumUsuar).</summary>
        public int? NumUsuar { get; set; }

        /// <summary>Nombre del usuario obtenido de la tabla Usuarios (según IdFlexoft).</summary>
        public string? NombreUsuario { get; set; }

        /// <summary>Fecha de revisión del master (Master.Ferevis) para el Codint de la orden.</summary>
        public DateTime? Ferevis { get; set; }

        /// <summary>Fecha de revisión formateada para la impresión (dd/MM/yyyy).</summary>
        public string? FerevisFormateada { get; set; }

        /// <summary>
        /// Revisión del producto tomada de MASTER.Revision (MASTER.Codint = ORDEFABR.Cod_Inte).
        /// Es el valor que se imprime en el casillero "Revisión" del Vale de Materiales.
        /// </summary>
        public string? Revision { get; set; }

        /// <summary>Lista de materiales reservados para esta OF (Reserva con CantRes > 0).</summary>
        public List<OrdenFabricacionMaterialConsumoDTO> Materiales { get; set; } = new();
    }

    /// <summary>
    /// Línea de material incluida en el Vale de Materiales.
    /// </summary>
    public class OrdenFabricacionMaterialConsumoDTO
    {
        /// <summary>Código externo del material (Reserva.Cod_Exte).</summary>
        public string? Cod_Exte { get; set; }

        /// <summary>Descripción del material (Reserva.Descrip).</summary>
        public string? Descrip { get; set; }

        /// <summary>Unidad de medida del material (Master.Umedida).</summary>
        public string? Umedida { get; set; }

        /// <summary>Uso por unidad del producto (Reserva.UsoUnit).</summary>
        public decimal? UsoUnit { get; set; }

        /// <summary>Uso por unidad formateado para la impresión.</summary>
        public string? UsoUnitFormateada { get; set; }

        /// <summary>Cantidad reservada / consumo total (Reserva.CantRes).</summary>
        public decimal? CantRes { get; set; }

        /// <summary>Cantidad reservada formateada para la impresión.</summary>
        public string? CantResFormateada { get; set; }
    }
}
