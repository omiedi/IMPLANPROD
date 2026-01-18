namespace IMPLANPROD.Client.Helpers
{
    /// <summary>
    /// Clase helper para manejar tipos de IVA y sus descripciones
    /// Centraliza la lógica de conversión entre códigos numéricos y descripciones
    /// </summary>
    public static class TipoIvaHelper
    {
        /// <summary>
        /// Obtiene la descripción legible del enum TipoIvaEnum
        /// </summary>
        /// <param name="tipo">Tipo de IVA enum</param>
        /// <returns>Descripción legible del tipo de IVA</returns>
        public static string ObtenerDescripcionIva(TipoIvaEnum tipo)
        {
            return tipo switch
            {
                TipoIvaEnum.ConsumidorFinal => "A CONSUMIDOR FINAL",
                TipoIvaEnum.IvaResponsableInscripto => "IVA RESPONSABLE INSCRIPTO",
                TipoIvaEnum.IvaResponsableNoInscripto => "IVA RESPONSABLE NO INSCRIPTO",
                TipoIvaEnum.IvaNoResponsable => "IVA NO RESPONSABLE",
                TipoIvaEnum.IvaExento => "IVA EXENTO",
                TipoIvaEnum.ComercioExterior => "OPs.COMERCIO EXTERIOR",
                TipoIvaEnum.Monotributo => "CONTRIBUYENTE MONOTRIBUTO",
                TipoIvaEnum.ExentoZonaPromovida => "EXENTO - ZONA PROMOVIDA",
                _ => tipo.ToString()
            };
        }

        /// <summary>
        /// Obtiene la descripción legible del tipo de IVA a partir del código numérico
        /// Usado para campos como Piva_Cli en PROVED12 y DEUDOR12
        /// </summary>
        /// <param name="codigoIva">Código numérico del tipo de IVA (0-7)</param>
        /// <returns>Descripción legible del tipo de IVA</returns>
        public static string ObtenerDescripcionIvaPorCodigo(int? codigoIva)
        {
            if (!codigoIva.HasValue)
                return "SIN ESPECIFICAR";

            if (Enum.IsDefined(typeof(TipoIvaEnum), codigoIva.Value))
            {
                var tipoIva = (TipoIvaEnum)codigoIva.Value;
                return ObtenerDescripcionIva(tipoIva);
            }

            return $"CÓDIGO DESCONOCIDO ({codigoIva})";
        }

        /// <summary>
        /// Obtiene la descripción completa con código y descripción
        /// Formato: "1 - IVA RESPONSABLE INSCRIPTO"
        /// </summary>
        /// <param name="codigoIva">Código numérico del tipo de IVA</param>
        /// <returns>Descripción completa con código y texto</returns>
        public static string ObtenerDescripcionCompletaPorCodigo(int? codigoIva)
        {
            if (!codigoIva.HasValue)
                return "Sin especificar";

            var descripcion = ObtenerDescripcionIvaPorCodigo(codigoIva);
            return $"{codigoIva} - {descripcion}";
        }

        /// <summary>
        /// Obtiene todos los tipos de IVA disponibles para usar en combos/selects
        /// </summary>
        /// <returns>Lista de tuplas con código y descripción</returns>
        public static List<(int Codigo, string Descripcion)> ObtenerTodosLosTiposIva()
        {
            var tipos = new List<(int Codigo, string Descripcion)>();
            
            foreach (TipoIvaEnum tipo in Enum.GetValues(typeof(TipoIvaEnum)))
            {
                tipos.Add(((int)tipo, ObtenerDescripcionIva(tipo)));
            }
            
            return tipos.OrderBy(t => t.Codigo).ToList();
        }

        /// <summary>
        /// Verifica si un código de IVA es válido
        /// </summary>
        /// <param name="codigoIva">Código a verificar</param>
        /// <returns>True si el código es válido</returns>
        public static bool EsCodigoIvaValido(int? codigoIva)
        {
            return codigoIva.HasValue && Enum.IsDefined(typeof(TipoIvaEnum), codigoIva.Value);
        }

        /// <summary>
        /// Obtiene el enum TipoIvaEnum a partir del código numérico
        /// </summary>
        /// <param name="codigoIva">Código numérico</param>
        /// <returns>Enum TipoIvaEnum o null si no es válido</returns>
        public static TipoIvaEnum? ObtenerTipoIvaPorCodigo(int? codigoIva)
        {
            if (EsCodigoIvaValido(codigoIva))
                return (TipoIvaEnum)codigoIva!.Value;
            
            return null;
        }
    }
}
