namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar un registro del archivo Ocompra_Sucursal.txt
    /// Este archivo se genera en la sucursal para sincronizar pedidos de reposición con la central
    /// 
    /// Formato del archivo:
    /// - Nombre: Ocompra_Sucursal AAAA_MM_DD HHMMSS.TXT
    /// - Delimitador: punto y coma (;)
    /// - Primera línea: encabezado con nombres de columnas
    /// - Líneas siguientes: datos de cada registro
    /// - Última línea opcional: mensaje con formato ||MENSAJE||;texto
    /// 
    /// Campos según código VB6 original:
    /// 1. SUCURSAL (8 caracteres) - NCLIENTE ARCON O COMODORO
    /// 2. CODIGO (8 caracteres) - Código interno del producto
    /// 3. NRO.PED. (8 caracteres) - Número de pedido
    /// 4. CTE.PED. (8 caracteres) - Cliente del pedido
    /// 5. USUARIO (8 caracteres) - Usuario que generó el pedido
    /// 6. OCOMPRA (8 caracteres) - Número de orden de compra
    /// 7. CANTIDAD (10 caracteres) - Cantidad solicitada
    /// 8. SOLENTRE (10 caracteres) - Fecha solicitud de entrega (formato dd/MM/yyyy)
    /// 9. VERS-COMP (8 caracteres) - Versión
    /// 10. FEMISION (8 caracteres) - Fecha de emisión (formato dd/MM/yy)
    /// 11. O.COMPRA CLIENTE (16 caracteres) - Orden de compra del cliente
    /// 12. R.SOCIAL CTE.PEDIDO (32 caracteres) - Razón social del cliente del pedido
    /// </summary>
    public class OcompraSucursalArchivoDTO
    {
        /// <summary>
        /// Número de sucursal (NCLIENTE ARCON O COMODORO)
        /// Corresponde al campo NC_CONFIG1 en VB6
        /// </summary>
        public string Sucursal { get; set; } = "";

        /// <summary>
        /// Código interno del producto (CI1% en VB6)
        /// Es el Codint de la tabla masters
        /// </summary>
        public string Codigo { get; set; } = "";

        /// <summary>
        /// Número de pedido (NPX& en VB6)
        /// Es el número de pedido generado
        /// </summary>
        public string NroPedido { get; set; } = "";

        /// <summary>
        /// Cliente del pedido (NROCLIENTE en VB6)
        /// Es el cliente asociado al pedido
        /// </summary>
        public string CtePedido { get; set; } = "";

        /// <summary>
        /// Usuario que generó el pedido (USER1% en VB6)
        /// En VB6 se calcula como USNBR% Mod 100
        /// </summary>
        public string Usuario { get; set; } = "";

        /// <summary>
        /// Número de orden de compra (OCOMPRA& en VB6)
        /// Es el número de la orden de compra de sucursal
        /// </summary>
        public string OCompra { get; set; } = "";

        /// <summary>
        /// Cantidad solicitada (CANSTOK en VB6)
        /// Es la cantidad de stock solicitada
        /// </summary>
        public string Cantidad { get; set; } = "";

        /// <summary>
        /// Fecha solicitud de entrega (FENTRESOL% en VB6)
        /// Formato: dd/MM/yyyy
        /// Si la fecha es menor a hoy, se usa la fecha de hoy
        /// </summary>
        public string SolEntre { get; set; } = "";

        /// <summary>
        /// Versión (VERSIONX% en VB6)
        /// Es el número de versión
        /// </summary>
        public string Version { get; set; } = "";

        /// <summary>
        /// Fecha de emisión (FEHOY% en VB6)
        /// Formato: dd/MM/yy
        /// Es la fecha actual cuando se genera el archivo
        /// </summary>
        public string Femision { get; set; } = "";

        /// <summary>
        /// Orden de compra del cliente (ORCOMCLIENTE$ en VB6)
        /// Longitud: 16 caracteres
        /// </summary>
        public string OCompraCliente { get; set; } = "";

        /// <summary>
        /// Razón social del cliente del pedido (RAZO_NC_PEDIDO$ en VB6)
        /// Longitud: 32 caracteres
        /// </summary>
        public string RSocialCtePedido { get; set; } = "";
    }

    /// <summary>
    /// DTO para representar el resultado de procesar un archivo Ocompra_Sucursal.txt
    /// Contiene la lista de registros y el mensaje opcional
    /// </summary>
    public class OcompraSucursalArchivoResultadoDTO
    {
        /// <summary>
        /// Lista de registros leídos del archivo
        /// </summary>
        public List<OcompraSucursalArchivoDTO> Registros { get; set; } = new();

        /// <summary>
        /// Mensaje opcional que viene al final del archivo
        /// Formato: ||MENSAJE||;texto
        /// Longitud máxima: 1000 caracteres
        /// </summary>
        public string? Mensaje { get; set; }

        /// <summary>
        /// Nombre del archivo procesado
        /// </summary>
        public string NombreArchivo { get; set; } = "";

        /// <summary>
        /// Fecha y hora de generación del archivo (extraída del nombre)
        /// Formato: AAAA_MM_DD HHMMSS
        /// </summary>
        public string FechaGeneracion { get; set; } = "";
    }
}
