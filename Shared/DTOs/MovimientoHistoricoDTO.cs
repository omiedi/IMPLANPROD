namespace IMPLANPROD.Shared.DTOs
{
    public class MovimientoHistoricoDTO
    {
        public int Id { get; set; }
        public DateTime Fecha { get; set; }
        public string? TipoMovimiento { get; set; }
        public string? DocumentoNumero { get; set; }
        public string? Transaccion { get; set; }
        public int? Deposito { get; set; }
        public string? IdBulto { get; set; }
        public string? Referencia { get; set; }
        public decimal? Entrada { get; set; }
        public decimal? Salida { get; set; }
        public decimal Saldo { get; set; }
        public int CodInte { get; set; }

        /// <summary>
        /// Clave primaria real de MOVISTO (MOVISTO.Id).
        /// El campo Id hereda Cod_Inte por compatibilidad con código legado,
        /// por lo que se agrega MovistoId para identificar el registro exacto.
        /// </summary>
        public int MovistoId { get; set; }

        /// <summary>
        /// Fecha y hora real del movimiento (MOVISTO.FeReMovi).
        /// En VB6 se mostraba FechMov en la grilla (solo fecha) y FeReMovi en el MsgBox.
        /// </summary>
        public DateTime? FeReMovi { get; set; }

        /// <summary>
        /// Número de usuario que registró el movimiento (MOVISTO.NumUsuar).
        /// Mapea a usuarios.idflexoft (compatibilidad VB6).
        /// </summary>
        public int? NumUsuar { get; set; }

        /// <summary>
        /// Nombre del usuario resuelto desde la tabla Usuarios (IdFlexoft = NumUsuar).
        /// Equivalente a USERNAM$(NUMEROUSUARIO%) en VB6.
        /// </summary>
        public string? NombreUsuario { get; set; }
    }

    public class FiltroMovimientosDTO
    {
        public int CodigoRepuesto { get; set; }
        public int? Deposito { get; set; }
        public DateTime? FechaDesde { get; set; }
        public DateTime? FechaHasta { get; set; }
    }
}
