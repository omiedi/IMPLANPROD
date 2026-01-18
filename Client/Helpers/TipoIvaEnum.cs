namespace IMPLANPROD.Client.Helpers
{
    /// <summary>
    /// Enumeración de tipos de IVA según AFIP.
    /// </summary>
    public enum TipoIvaEnum
    {
        /// <summary>
        /// Consumidor Final
        /// </summary>
        ConsumidorFinal = 0,
        /// <summary>
        /// IVA Responsable Inscripto
        /// </summary>
        IvaResponsableInscripto = 1,
        /// <summary>
        /// IVA Responsable No Inscripto
        /// </summary>
        IvaResponsableNoInscripto = 2,
        /// <summary>
        /// IVA No Responsable
        /// </summary>
        IvaNoResponsable = 3,
        /// <summary>
        /// IVA Exento
        /// </summary>
        IvaExento = 4,
        /// <summary>
        /// Operaciones de Comercio Exterior
        /// </summary>
        ComercioExterior = 5,
        /// <summary>
        /// Contribuyente Monotributo
        /// </summary>
        Monotributo = 6,
        /// <summary>
        /// Exento Zona Promovida
        /// </summary>
        ExentoZonaPromovida = 7
    }
}
