using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar un remito que puede ser anulado
    /// </summary>
    public class RemitoAnularDTO
    {
        /// <summary>
        /// Documento primario corto (número de remito)
        /// </summary>
        public string DoPriCor { get; set; } = string.Empty;
        /// <summary>
        /// Documento primario largo (número de remito largo)
        /// </summary>
        public string DoPriLar { get; set; } = string.Empty;

        /// <summary>
        /// Número de cliente
        /// </summary>
        public int NumeClie { get; set; }

        /// <summary>
        /// Razón social del cliente
        /// </summary>
        public string RazonSocial { get; set; } = string.Empty;

        /// <summary>
        /// Fecha del movimiento
        /// </summary>
        public DateTime FechMov { get; set; }

        /// <summary>
        /// Referencia corta del cliente
        /// </summary>
        public string ReferCor { get; set; } = string.Empty;
    }
}
