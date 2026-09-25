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
        /// <returns>Tupla con (IsValid, ErrorMessage) - IsValid indica si es válido, ErrorMessage contiene el mensaje de error si no lo es</returns>
        Task<(bool IsValid, string ErrorMessage)> ValidarMovimientoAsync(MoviStockDTO movimiento);

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
        /// Obtiene los remitos de traslado del mes (códigos RC y DV)
        /// </summary>
        /// <param name="anio">Año</param>
        /// <param name="mes">Mes (1-12)</param>
        /// <param name="numeroCliente">Número de cliente (opcional)</param>
        /// <returns>Lista de remitos de traslado del mes</returns>
        Task<List<RemitoAnularDTO>> ObtenerRemitosTrasladoMesAsync(int anio, int mes, int? numeroCliente = null);

        /// <summary>
        /// Obtiene el último número de remito registrado en MOVISTO
        /// Filtra por código de movimiento y fecha y hora mayor o igual a la fecha de nueva numeración
        /// </summary>
        /// <param name="fechaNuevaNumeracion">Fecha y hora desde la cual considerar los remitos (opcional, default 01/01/1980 00:00)</param>
        /// <param name="codigoMovimiento">Código de movimiento a filtrar (opcional, default "RT" para remitos de cliente, usar "RC" para traslados)</param>
        /// <returns>Último número de remito</returns>
        Task<int> ObtenerUltimoNumeroRemitoAsync(DateTime? fechaNuevaNumeracion = null, string codigoMovimiento = "RT");

        /// <summary>
        /// Obtiene el último número de remito registrado en MOVISTO para un prefijo (NUMPVTA) específico.
        /// Se usa para proponer el siguiente número libre de un punto de venta.
        /// </summary>
        /// <param name="prefijo">Prefijo del punto de venta</param>
        /// <returns>Último número de remito para ese prefijo</returns>
        Task<int> ObtenerUltimoNumeroPorPrefijoAsync(short prefijo);

        /// <summary>
        /// Anula un remito específico
        /// Replica la funcionalidad de anulación de remitos de VB6
        /// 1. Determina si el remito está en MOVISTO o MOVIREPA según CodiMovi
        /// 2. Actualiza los pedidos vinculados (reduce cantidad entregada)
        /// 3. Marca el remito como anulado (CantSalid=0, CantIngre=0, PESOKG=0, METROS=0, ReferCor con prefijo "(Anul)")
        /// </summary>
        /// <param name="request">Datos del remito a anular</param>
        /// <returns>Resultado de la anulación</returns>
        Task<AnularRemitoResponse> AnularRemitoAsync(AnularRemitoRequest request);

        /// <summary>
        /// Obtiene la cantidad de filas (items) que componen un remito.
        /// Se usa para mostrar en el mensaje de confirmación antes de anular.
        /// La búsqueda se realiza por DoPriLar + Nume_Clie + FechMov en MOVISTO o MOVIREPA.
        /// </summary>
        /// <param name="request">Datos del remito a consultar</param>
        /// <returns>Cantidad de filas del remito (0 si no se encuentra)</returns>
        Task<int> ObtenerCantidadFilasRemitoAsync(AnularRemitoRequest request);
    }
}
