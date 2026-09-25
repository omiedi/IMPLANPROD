namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO que representa un ítem del informe de implosión.
    /// Muestra en qué productos (conjuntos) es componente un ítem específico.
    /// </summary>
    public class ImplosionItemDTO
    {
        /// <summary>
        /// Código interno del conjunto (producto padre). Equivale a Master.Codint.
        /// </summary>
        public int CodIConj { get; set; }

        /// <summary>
        /// Id (clave primaria) del conjunto en la tabla MASTER.
        /// Necesario para navegar a /master/edit/{Id}, ya que la ruta usa la PK
        /// y no el código interno (Codint).
        /// </summary>
        public int IdMaster { get; set; }

        /// <summary>
        /// Código externo del conjunto (producto padre)
        /// </summary>
        public string Codigo { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del conjunto (producto padre)
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Unidad de medida del conjunto
        /// </summary>
        public string Umedida { get; set; } = string.Empty;

        /// <summary>
        /// Uso bruto del componente en la fórmula del conjunto
        /// </summary>
        public decimal UsoBruto { get; set; }

        /// <summary>
        /// Indica si el uso bruto fue ajustado por el modo por mil (PORMILES)
        /// </summary>
        public bool AjustadoPorMil { get; set; }
    }
}
