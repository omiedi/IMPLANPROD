using System;
using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO con todos los datos necesarios para imprimir el "Plan de Inspección y Ensayo" en formato A4.
    /// Se utiliza desde el modal PlanInspeccionModal, accesible en /ingenieria/documentos.
    /// El controlador CodiprocController lo arma a partir de Master + Codiproc + Plainsp.
    /// </summary>
    public class PlanInspeccionImpresionDTO
    {
        // --- Datos del producto (Master) ---

        /// <summary>Código interno del producto (Master.Codint).</summary>
        public int? CodInte { get; set; }

        /// <summary>Código externo del producto (Master.Codigo).</summary>
        public string? Codigo { get; set; }

        /// <summary>Denominación / descripción del producto (Master.Descripcion).</summary>
        public string? Denominacion { get; set; }

        /// <summary>Revisión del producto (Master.Revision).</summary>
        public string? Revision { get; set; }

        /// <summary>
        /// Origen del plan de inspección (por ejemplo, número de Orden de Fabricación).
        /// Impreso desde /ingenieria/documentos el plan no está atado a una OF específica,
        /// por lo que este campo queda vacío; se deja preparado para reutilizar el mismo
        /// formulario desde otros contextos que sí conozcan el origen.
        /// </summary>
        public string? Origen { get; set; }

        /// <summary>Fecha de impresión (hoy), formateada dd/MM/yyyy.</summary>
        public string? FechaImpresionFormateada { get; set; }

        // --- Datos del encabezado del plan (Codiproc) ---

        /// <summary>Nivel de Calidad Aceptable (Codiproc.AQL).</summary>
        public int? AQL { get; set; }

        /// <summary>Nivel porcentual de inspección (Codiproc.NivelPct).</summary>
        public int? NivelPct { get; set; }

        /// <summary>Frecuencia de Inspección (Codiproc.FrecInspeccion).</summary>
        public string? FrecInspeccion { get; set; }

        /// <summary>Puntos de Detención Obligatoria (Codiproc.PuntoDetencion).</summary>
        public string? PuntoDetencion { get; set; }

        /// <summary>Nombre de quien preparó el plan (Codiproc.NombrePreparo).</summary>
        public string? NombrePreparo { get; set; }

        /// <summary>Fecha de preparación formateada dd/MM/yyyy (Codiproc.FechaPreparo).</summary>
        public string? FechaPreparoFormateada { get; set; }

        /// <summary>Nombre de quien aprobó el plan (Codiproc.NombreAprobo).</summary>
        public string? NombreAprobo { get; set; }

        /// <summary>Fecha de aprobación formateada dd/MM/yyyy (Codiproc.FechaAprobo).</summary>
        public string? FechaAproboFormateada { get; set; }

        /// <summary>
        /// Observaciones generales del plan (Codiproc.Observaciones).
        /// Se imprimen al pie de cada hoja y se repiten en todas las páginas generadas.
        /// </summary>
        public string? Observaciones { get; set; }

        // --- Detalle (Plainsp) ---

        /// <summary>
        /// Ítems del plan de inspección, en el orden definido por el usuario (Plainsp.NroOrden).
        /// Se imprimen tal cual están, incluyendo filas totalmente vacías que el usuario
        /// haya dejado como separador visual entre bloques (no se filtran ni se colapsan).
        /// </summary>
        public List<PlanInspeccionItemImpresionDTO> Items { get; set; } = new();
    }

    /// <summary>
    /// Línea de detalle del Plan de Inspección, mapea un registro de PLAINSP.
    /// </summary>
    public class PlanInspeccionItemImpresionDTO
    {
        /// <summary>Dimensión según plano (Plainsp.DimSegunPlano).</summary>
        public string? DimSegunPlano { get; set; }

        /// <summary>Instrumento de medición (Plainsp.Instrumento).</summary>
        public string? Instrumento { get; set; }

        /// <summary>Resultado / observaciones (Plainsp.Observaciones).</summary>
        public string? Observaciones { get; set; }
    }
}
