using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Interfaz para el servicio de movimientos de stock
    /// Replica la funcionalidad de MOVISTOK de VB6
    /// </summary>
    public interface IMoviStockService
    {
        /// <summary>
        /// Registra un movimiento de stock en la tabla MOVISTO
        /// Equivalente a la rutina MOVISTOK de VB6
        /// </summary>
        /// <param name="movimiento">Datos del movimiento de stock</param>
        /// <param name="user">Usuario que ejecuta la operación</param>
        /// <returns>Resultado de la operación</returns>
        Task<ActionResult> RegistrarMovimientoAsync(MoviStockDTO movimiento, ClaimsPrincipal user);

        /// <summary>
        /// Valida si un movimiento de stock es válido antes de registrarlo
        /// </summary>
        /// <param name="movimiento">Datos del movimiento a validar</param>
        /// <returns>True si el movimiento es válido, False en caso contrario</returns>
        Task<bool> ValidarMovimientoAsync(MoviStockDTO movimiento);

        /// <summary>
        /// Valida si existe un remito duplicado (mismo prefijo y número)
        /// </summary>
        /// <param name="prefijo">Prefijo del remito (NUMPVTA)</param>
        /// <param name="numero">Número del remito (DOPRINUM)</param>
        /// <returns>True si existe, False si no existe</returns>
        Task<bool> ExisteRemitoDuplicadoAsync(short prefijo, int numero);

        /// <summary>
        /// Obtiene los remitos de un mes específico para anulación
        /// </summary>
        /// <param name="anio">Año</param>
        /// <param name="mes">Mes (1-12)</param>
        /// <param name="numeroCliente">Número de cliente (opcional)</param>
        /// <returns>Lista de remitos del mes</returns>
        Task<List<RemitoAnularDTO>> ObtenerRemitosMesAsync(int anio, int mes, int? numeroCliente = null);

        /// <summary>
        /// Obtiene los remitos de un mes específico para consulta general
        /// Incluye también los remitos anulados
        /// </summary>
        /// <param name="anio">Año</param>
        /// <param name="mes">Mes (1-12)</param>
        /// <param name="numeroCliente">Número de cliente (opcional)</param>
        /// <returns>Lista de remitos del mes (incluyendo anulados)</returns>
        Task<List<RemitoAnularDTO>> ObtenerRemitosMesTodosAsync(int anio, int mes, int? numeroCliente = null);

        /// <summary>
        /// Obtiene el último número de remito registrado en MOVISTO
        /// Filtra por código de movimiento RT y fecha mayor o igual a la fecha de nueva numeración
        /// </summary>
        /// <param name="fechaNuevaNumeracion">Fecha desde la cual considerar los remitos (opcional, default 01/01/1980)</param>
        /// <returns>Último número de remito</returns>
        Task<int> ObtenerUltimoNumeroRemitoAsync(DateTime? fechaNuevaNumeracion = null);

        /// <summary>
        /// Anula un remito específico
        /// Replica la funcionalidad de anulación de remitos de VB6
        /// 1. Determina si el remito está en MOVISTO o MOVIREPA según CodiMovi
        /// 2. Actualiza los pedidos vinculados (reduce cantidad entregada)
        /// 3. Marca el remito como anulado (CantSalid=0, ReferCor con prefijo "(Anul)")
        /// </summary>
        /// <param name="request">Datos del remito a anular</param>
        /// <returns>Resultado de la anulación</returns>
        Task<AnularRemitoResponse> AnularRemitoAsync(AnularRemitoRequest request);
    }
}
