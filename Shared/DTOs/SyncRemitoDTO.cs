namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para sincronización de Remitos entre Sucursal Neuquén y Central
    /// Se genera a partir de archivos TXT con delimitador punto y coma
    /// </summary>
    public class SyncRemitoDTO
    {
        /// <summary>
        /// Nombre de la tabla que se sincroniza
        /// </summary>
        public string Tabla { get; set; } = "MOVISTO";

        /// <summary>
        /// Operación a realizar: siempre INSERT para remitos
        /// </summary>
        public string Operacion { get; set; } = "INSERT";

        /// <summary>
        /// Fecha y hora de generación del archivo de sincronización
        /// </summary>
        public DateTime FechaGeneracion { get; set; }

        /// <summary>
        /// Número de remito (extraído del archivo TXT)
        /// </summary>
        public string NumeroRemito { get; set; } = string.Empty;

        /// <summary>
        /// Lista de movimientos de stock del remito (filas de detalle del TXT)
        /// Cada movimiento generará un INSERT en la tabla MOVISTO
        /// </summary>
        public List<RemitoItemDTO> Movimientos { get; set; } = new();
    }

    /// <summary>
    /// Representa un movimiento de stock individual de un remito
    /// Corresponde a una fila de detalle del archivo TXT
    /// Se mapea a la tabla MOVISTO
    /// </summary>
    public class RemitoItemDTO
    {
        // === Campos mapeados del archivo TXT ===
        
        /// <summary>
        /// Cantidad ingreso (del TXT: CANTINGRE)
        /// </summary>
        public decimal? CantIngre { get; set; }

        /// <summary>
        /// Cantidad salida (del TXT: CANTSALID)
        /// </summary>
        public decimal? CantSalid { get; set; }

        /// <summary>
        /// Código externo del producto (del TXT: CODEXTE)
        /// </summary>
        public string? CodExte { get; set; }

        /// <summary>
        /// Código interno del producto (del TXT: CODINT)
        /// </summary>
        public int? CodInte { get; set; }

        /// <summary>
        /// Descripción larga del producto (del TXT: DESCRILAR)
        /// </summary>
        public string? DescriLar { get; set; }

        /// <summary>
        /// Número de remito corto (del TXT: REMITO, transformado de RT-X-0001-00018694 → RT.18694)
        /// Se usa para DoPriCor, DoPriLar, IdenLote
        /// </summary>
        public string? DoPriCor { get; set; }

        /// <summary>
        /// Número de remito largo (del TXT: REMITO, transformado de RT-X-0001-00018694 → RT.18694)
        /// Se usa para DoPriLar
        /// </summary>
        public string? DoPriLar { get; set; }

        /// <summary>
        /// Número de remito numérico (del TXT: REMITO, extrae últimas 8 posiciones: 18694)
        /// Se usa para DOPRINUM
        /// </summary>
        public int? DOPRINUM { get; set; }

        /// <summary>
        /// Número de documento secundario corto (para órdenes de reparación)
        /// </summary>
        public string? DoSecCor { get; set; }

        /// <summary>
        /// Número de documento secundario largo (para órdenes de reparación)
        /// </summary>
        public string? DoSecLar { get; set; }

        /// <summary>
        /// Número de cliente (del TXT: N_Clie)
        /// </summary>
        public int? Nume_Clie { get; set; }

        /// <summary>
        /// Número de punto de venta (del TXT: REMITO, extrae posición 6-10: 1)
        /// Se usa para NUMPVTA
        /// </summary>
        public int? NUMPVTA { get; set; }

        /// <summary>
        /// Número de orden ítem (del TXT: NUORI)
        /// Se usa para NuorItem
        /// </summary>
        public int? NuorItem { get; set; }

        /// <summary>
        /// Razón social del cliente (del TXT: R.SOCIAL)
        /// Se usa para ReferCor
        /// </summary>
        public string? ReferCor { get; set; }

        /// <summary>
        /// Fecha de movimiento del remito (del TXT: FECHMOV)
        /// Formato: Jul 24 2026 12:00AM
        /// Se usa para FeReMovi
        /// </summary>
        public DateTime? FeReMovi { get; set; }

        // === Campos con valores por defecto ===

        /// <summary>
        /// Tipo de cambio (valor por defecto: 1)
        /// </summary>
        public int? TipoCam { get; set; } = 1;

        /// <summary>
        /// Tipo de transacción (valor por defecto: 0)
        /// </summary>
        public int? TipoTran { get; set; } = 0;

        /// <summary>
        /// Valor unitario (valor por defecto: 0)
        /// </summary>
        public decimal? ValoUnit { get; set; } = 0;

        /// <summary>
        /// Valor unitario con IVA (valor por defecto: 0)
        /// </summary>
        public decimal? ValoUnitIva { get; set; } = 0;

        /// <summary>
        /// Valor total ítem (valor por defecto: 0)
        /// </summary>
        public decimal? ValTotIt { get; set; } = 0;

        /// <summary>
        /// Valor total ítem con IVA (valor por defecto: 0)
        /// </summary>
        public decimal? ValTotItIva { get; set; } = 0;

        /// <summary>
        /// Verificado por (valor por defecto: 0)
        /// </summary>
        public int? VerificadoPor { get; set; } = 0;

        /// <summary>
        /// Cantidad saldo (valor por defecto: 0)
        /// </summary>
        public decimal? CantSaldo { get; set; } = 0;

        /// <summary>
        /// Código de empresa (valor por defecto: 0)
        /// </summary>
        public int? CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Fecha de movimiento (valor por defecto: fecha del día del INSERT)
        /// Formato: Jul 27 2026 12:00AM
        /// </summary>
        public DateTime? FechMov { get; set; }

        /// <summary>
        /// Fecha y hora de verificación (valor por defecto: fecha del día del INSERT)
        /// Formato: Jul 27 2026 12:00AM
        /// </summary>
        public DateTime? FHor_Verif { get; set; }

        /// <summary>
        /// ID de operación (valor por defecto: 0)
        /// </summary>
        public int? ID_OPERACION { get; set; } = 0;

        /// <summary>
        /// ID de calidad (valor por defecto: 0)
        /// </summary>
        public int? IdCalidad { get; set; } = 0;

        /// <summary>
        /// Moneda valor (valor por defecto: 1)
        /// </summary>
        public int? MoneVal { get; set; } = 1;

        /// <summary>
        /// Número de proveedor (valor por defecto: 0)
        /// </summary>
        public int? Nume_Prov { get; set; } = 0;

        /// <summary>
        /// Número de orden de pedido (valor por defecto: 0)
        /// </summary>
        public int? NumeOrdPed { get; set; } = 0;

        /// <summary>
        /// Número de pedido cliente (valor por defecto: 0)
        /// </summary>
        public int? NuPedCli { get; set; } = 0;

        /// <summary>
        /// Número de usuario (valor por defecto: usuario activo)
        /// </summary>
        public int? NumUsuar { get; set; }

        /// <summary>
        /// Código de movimiento (valor por defecto: 'RT')
        /// </summary>
        public string? CodiMovi { get; set; } = "RT";

        /// <summary>
        /// Depósito de movimiento (valor por defecto: 1)
        /// </summary>
        public int? DepoMovi { get; set; } = 1;
    }
}
