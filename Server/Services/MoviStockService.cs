using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para manejar movimientos de stock
    /// Replica la funcionalidad de la rutina MOVISTOK de VB6
    /// </summary>
    public class MoviStockService : IMoviStockService
    {
        private readonly DataContext _context;
        private readonly IUsuarioFlexoftService _usuarioFlexoftService;
        private readonly ICodiempresa _codiempresa;
        private readonly ILogger<MoviStockService> _logger;
        private readonly ISyncService _syncService;

        public MoviStockService(
            DataContext context,
            IUsuarioFlexoftService usuarioFlexoftService,
            ICodiempresa codiempresa,
            ILogger<MoviStockService> logger,
            ISyncService syncService)
        {
            _context = context;
            _usuarioFlexoftService = usuarioFlexoftService;
            _codiempresa = codiempresa;
            _logger = logger;
            _syncService = syncService;
        }

        /// <summary>
        /// Registra un movimiento de stock en la tabla MOVISTO
        /// Equivalente a la rutina MOVISTOK de VB6
        /// </summary>
        public async Task<ActionResult> RegistrarMovimientoAsync(MoviStockDTO movimiento, ClaimsPrincipal user)
        {
            try
            {
                // Obtener información del usuario y empresa
                int? idFlexoft = _usuarioFlexoftService.InterObtenerIdFlexoft(user);
                short? codigoEmpresa = _codiempresa.IObtenerCodiempresa(user);

                if (!idFlexoft.HasValue)
                {
                    _logger.LogWarning("Intento de registro de movimiento sin usuario válido");
                    return new BadRequestObjectResult("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Validaciones básicas
                var (isValid, errorMessage) = await ValidarMovimientoAsync(movimiento);
                if (!isValid)
                {
                    return new BadRequestObjectResult(errorMessage);
                }

                // Determinar la tabla destino según el código de movimiento
                // Si CMOV$ empieza con "*" graba en MOVIREPA, sino en MOVISTO
                string tablaDestino = movimiento.CodigoMovimiento.StartsWith("*") ? "MOVIREPA" : "MOVISTO";
                
                _logger.LogInformation($"Registrando movimiento en tabla {tablaDestino} para producto {movimiento.CodigoInterno}");

                // Procesar números de punto de venta y documento para remitos (RT)
                short numPvta = 0;
                long doPriNum = 0;
                
                // Si vienen los campos numéricos directamente, usarlos (preferencia)
                if (movimiento.PrefijoRemito.HasValue)
                {
                    numPvta = movimiento.PrefijoRemito.Value;
                }
                if (movimiento.NumeroRemito.HasValue)
                {
                    doPriNum = movimiento.NumeroRemito.Value;
                }
                
                // Si no vienen los campos numéricos, intentar extraerlos del DocumentoPrimarioLargo (compatibilidad)
                if (numPvta == 0 && doPriNum == 0 && 
                    movimiento.CodigoMovimiento.ToUpper() == "RT" && 
                    movimiento.DocumentoPrimarioLargo?.ToUpper().StartsWith("RT") == true)
                {
                    // Extraer NUMPVTA del documento (posiciones 6-9)
                    if (movimiento.DocumentoPrimarioLargo.Length >= 10)
                    {
                        if (short.TryParse(movimiento.DocumentoPrimarioLargo.Substring(5, 4), out short pvta))
                            numPvta = pvta;
                    }
                    
                    // Extraer DOPRINUM del documento (posiciones 11-18)
                    if (movimiento.DocumentoPrimarioLargo.Length >= 19)
                    {
                        if (long.TryParse(movimiento.DocumentoPrimarioLargo.Substring(10, 8), out long priNum))
                            doPriNum = priNum;
                    }
                }
                
                _logger.LogInformation("Números de remito procesados: NUMPVTA={NumPvta}, DOPRINUM={DoPriNum}", numPvta, doPriNum);

                // Obtener información del producto
                var producto = await _context.masters
                    .Where(m => m.Codint == movimiento.CodigoInterno)
                    .Select(m => new { m.Codigo, m.Descripcion })
                    .FirstOrDefaultAsync();

                string codigoExterno = producto?.Codigo?.Substring(0, Math.Min(24, producto.Codigo.Length)) ?? "";
                string descripcion = producto?.Descripcion?.Substring(0, Math.Min(64, producto.Descripcion.Length)) ?? "";

                // Procesar documentos primario y secundario
                string doPriLar = string.IsNullOrEmpty(movimiento.DocumentoPrimarioLargo?.Trim()) 
                    ? movimiento.DocumentoPrimarioCorto ?? "" 
                    : movimiento.DocumentoPrimarioLargo.Trim();
                
                string doSecLar = string.IsNullOrEmpty(movimiento.DocumentoSecundarioLargo?.Trim()) 
                    ? movimiento.DocumentoSecundarioCorto ?? "" 
                    : movimiento.DocumentoSecundarioLargo.Trim();

                // Procesar número de pedido cliente y orden del pedido
                // Usar directamente los valores del DTO que vienen del cliente
                int nuPedCli = (int)(movimiento.NumeroPedidoCliente ?? 0);
                short numeOrdPed = movimiento.NumeroOrdenPedido ?? 0;
                
                _logger.LogInformation("Procesando movimiento: NumeroPedidoCliente={NuPedCli}, NumeroOrdenPedido={NumeOrdPed}", 
                    nuPedCli, numeOrdPed);

                // Calcular tipo de cambio (función TICAMONE en VB6)
                decimal tipoCambio = await ObtenerTipoCambioAsync(movimiento.Moneda ?? 0);

                // Crear el nuevo registro de movimiento
                var nuevoMovimiento = new Movisto
                {
                    // Campos básicos del movimiento
                    FeReMovi = DateTime.Now,
                    FechMov = movimiento.Fecha,
                    CodiEmpr = 0, // Siempre 0 para remitos
                    NumUsuar = idFlexoft,
                    
                    // Información del producto
                    Cod_Inte = movimiento.CodigoInterno,
                    Cod_Exte = codigoExterno,
                    Descrip = descripcion,
                    IdenLote = movimiento.Lote?.Substring(0, Math.Min(16, movimiento.Lote.Length)),
                    
                    // Código de movimiento y documentos
                    CodiMovi = movimiento.CodigoMovimiento.Substring(0, Math.Min(2, movimiento.CodigoMovimiento.Length)),
                    DoPriCor = movimiento.DocumentoPrimarioCorto?.Substring(0, Math.Min(12, movimiento.DocumentoPrimarioCorto.Length)),
                    DoPriLar = doPriLar.Substring(0, Math.Min(24, doPriLar.Length)),
                    DoSecCor = movimiento.DocumentoSecundarioCorto?.Substring(0, Math.Min(12, movimiento.DocumentoSecundarioCorto.Length)),
                    DoSecLar = doSecLar.Substring(0, Math.Min(24, doSecLar.Length)),
                    
                    // Números de documento para remitos
                    NUMPVTA = numPvta,
                    DOPRINUM = (int?)doPriNum,
                    
                    // Información adicional
                    NuorItem = movimiento.NumeroOrdenItem ?? 0,
                    DocuOrig = movimiento.DocumentoCompra?.Substring(0, Math.Min(24, movimiento.DocumentoCompra.Length)),
                    IdenMovi = movimiento.IdentificadorMovimiento?.Substring(0, Math.Min(32, movimiento.IdentificadorMovimiento.Length)),
                    NuPedCli = nuPedCli,
                    NumeOrdPed = numeOrdPed,
                    ReferCor = movimiento.Referencia?.Substring(0, Math.Min(64, movimiento.Referencia.Length)),
                    
                    // Descripción larga - si está vacía usa la descripción del producto
                    Descri_Lar = string.IsNullOrEmpty(movimiento.DescripcionLarga?.Trim()) 
                        ? descripcion 
                        : movimiento.DescripcionLarga,
                    
                    // Valores monetarios
                    ValoUnit = movimiento.ValorUnitario ?? 0,
                    ValTotIt = Math.Round((movimiento.ValorUnitario ?? 0) * Math.Round(Math.Abs(movimiento.Cantidad), 4), 2),
                    MoneVal = movimiento.Moneda ?? 0,
                    TipoCam = tipoCambio,
                    ValoUnitIva = movimiento.PrecioConIva ?? 0,
                    ValTotItIva = Math.Round((movimiento.PrecioConIva ?? 0) * Math.Round(Math.Abs(movimiento.Cantidad), 4), 2),
                    
                    // Depósito
                    DepoMovi = movimiento.Deposito,
                    
                    // Cliente/Proveedor
                    Nume_Clie = movimiento.NumeroCliente >= 0 ? movimiento.NumeroCliente : 0,
                    Nume_Prov = movimiento.NumeroCliente < 0 ? (short)Math.Abs(movimiento.NumeroCliente.Value) : (short)0,
                    
                    // Cantidades - Positivo=Ingreso, Negativo=Salida
                    CantIngre = movimiento.Cantidad >= 0 ? Math.Round(movimiento.Cantidad, 4) : 0,
                    CantSalid = movimiento.Cantidad < 0 ? Math.Round(Math.Abs(movimiento.Cantidad), 4) : 0,
                    CantSaldo = 0,
                    
                    // Campos de verificación
                    VerificadoPor = 0,
                    FHor_Verif = new DateTime(1900, 1, 1), // Equivalente a NUNCA$ en VB6
                    
                    // Campos opcionales
                    TipoTran = movimiento.TipoTransferencia,
                    OCOMPRA_RTOABIERTO = movimiento.OrdenCompraRetornoAbierto?.Substring(0, Math.Min(64, movimiento.OrdenCompraRetornoAbierto.Length)),
                    IdCalidad = (int?)movimiento.IdCalidad,
                    ID_OPERACION = movimiento.IdOperacion,
                    PESOKG = movimiento.PesoKg > 0 || movimiento.Metros > 0 ? movimiento.PesoKg : null,
                    METROS = movimiento.PesoKg > 0 || movimiento.Metros > 0 ? movimiento.Metros : null
                };

                // Procesar lote de consumo para rechazos y faltantes
                if (movimiento.Referencia?.StartsWith("RECHAZO") == true || 
                    movimiento.Referencia?.StartsWith("FALTANTE") == true)
                {
                    string loteConsum = (movimiento.Lote ?? "").PadRight(24).Substring(0, 24) + 
                                       Math.Abs(movimiento.Cantidad).ToString("F4");
                    nuevoMovimiento.LotConsum = loteConsum;
                }

                // Agregar el movimiento a la base de datos
                _context.Movistos.Add(nuevoMovimiento);
                await _context.SaveChangesAsync();

                _logger.LogInformation($"Movimiento de stock registrado exitosamente. ID: {nuevoMovimiento.Id}");

                // Generar archivo JSON de sincronización si es un remito (RT) y está configurado
                if (movimiento.CodigoMovimiento.ToUpper() == "RT" && idFlexoft.HasValue)
                {
                    try
                    {
                        // Buscar todos los movimientos del mismo remito en la base de datos
                        // (incluyendo el que acabamos de insertar)
                        var movimientosRemito = await _context.Movistos
                            .Where(m => m.DoPriLar == doPriLar && m.CodiMovi == "RT")
                            .OrderBy(m => m.Id)
                            .ToListAsync();

                        if (movimientosRemito.Any())
                        {
                            // Mapear todos los movimientos a RemitoItemDTO
                            var remitoItems = movimientosRemito.Select(m => new RemitoItemDTO
                            {
                                CantIngre = m.CantIngre ?? 0,
                                CantSalid = m.CantSalid ?? 0,
                                CodExte = m.Cod_Exte,
                                CodInte = m.Cod_Inte,
                                DescriLar = m.Descri_Lar,
                                DoPriCor = m.DoPriCor,
                                DoPriLar = m.DoPriLar,
                                DOPRINUM = m.DOPRINUM,
                                Nume_Clie = m.Nume_Clie,
                                NUMPVTA = m.NUMPVTA,
                                NuorItem = m.NuorItem,
                                ReferCor = m.ReferCor,
                                FeReMovi = m.FeReMovi,
                                TipoCam = m.TipoCam.HasValue ? (int?)m.TipoCam.Value : null,
                                TipoTran = m.TipoTran,
                                ValoUnit = m.ValoUnit,
                                ValoUnitIva = m.ValoUnitIva,
                                ValTotIt = m.ValTotIt,
                                ValTotItIva = m.ValTotItIva,
                                VerificadoPor = m.VerificadoPor,
                                CantSaldo = m.CantSaldo,
                                CodiEmpr = m.CodiEmpr,
                                FechMov = m.FechMov,
                                FHor_Verif = m.FHor_Verif,
                                ID_OPERACION = m.ID_OPERACION,
                                IdCalidad = m.IdCalidad,
                                MoneVal = m.MoneVal,
                                Nume_Prov = m.Nume_Prov,
                                NumeOrdPed = m.NumeOrdPed,
                                NuPedCli = m.NuPedCli,
                                NumUsuar = m.NumUsuar ?? idFlexoft.Value,
                                CodiMovi = m.CodiMovi,
                                DepoMovi = m.DepoMovi
                            }).ToList();

                            var cantidadItems = remitoItems.Count;
                            await _syncService.GenerarArchivoSyncRemitoDesdeDatosAsync(remitoItems, idFlexoft.Value);
                            _logger.LogInformation($"Archivo JSON de sincronización generado para remito: {doPriLar} con {cantidadItems} items");
                        }
                    }
                    catch (Exception syncEx)
                    {
                        _logger.LogWarning(syncEx, "No se pudo generar archivo JSON de sincronización para remito: {DoPriLar}", doPriLar);
                        // No fallar el registro del movimiento si falla la generación del JSON
                    }
                }

                // Actualizar PEDDETA si el movimiento es un remito (RT o RC) y tiene número de pedido
                if ((movimiento.CodigoMovimiento.ToUpper() == "RT" || movimiento.CodigoMovimiento.ToUpper() == "RC") && nuPedCli > 0 && movimiento.Cantidad < 0)
                {
                    _logger.LogInformation("Actualizando PEDDETA para remito {CodigoMovimiento}: NuPedCli={NuPedCli}, CodigoInterno={CodigoInterno}, NumeOrdPed={NumeOrdPed}",
                        movimiento.CodigoMovimiento.ToUpper(), nuPedCli, movimiento.CodigoInterno, numeOrdPed);
                    
                    await ActualizarCantidadEntregadaPedidoAsync(
                        nuPedCli, 
                        movimiento.CodigoInterno, 
                        movimiento.NumeroCliente ?? 0,
                        Math.Abs(movimiento.Cantidad),
                        doPriLar.Substring(0, Math.Min(12, doPriLar.Length)),
                        numeOrdPed);
                }
                
                return new OkObjectResult(new { 
                    Id = nuevoMovimiento.Id, 
                    Mensaje = "Movimiento de stock registrado exitosamente" 
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar movimiento de stock: {Message}", ex.Message);
                return new BadRequestObjectResult($"Error al registrar el movimiento: {ex.Message}");
            }
        }

        /// <summary>
        /// Valida si un movimiento de stock es válido antes de registrarlo
        /// </summary>
        public async Task<(bool IsValid, string ErrorMessage)> ValidarMovimientoAsync(MoviStockDTO movimiento)
        {
            try
            {
                // Validaciones básicas
                if (movimiento.CodigoInterno <= 0)
                {
                    return (false, $"Código interno del producto inválido: {movimiento.CodigoInterno}");
                }

                if (movimiento.Deposito <= 0)
                {
                    return (false, $"Depósito inválido: {movimiento.Deposito}");
                }

                if (string.IsNullOrEmpty(movimiento.CodigoMovimiento))
                {
                    return (false, "Código de movimiento vacío");
                }

                if (movimiento.Cantidad == 0)
                {
                    return (false, "Cantidad de movimiento es cero");
                }

                // Verificar que el producto existe en la tabla Master
                bool productoExiste = await _context.masters
                    .AnyAsync(m => m.Codint == movimiento.CodigoInterno);

                if (!productoExiste)
                {
                    return (false, $"El producto con código interno {movimiento.CodigoInterno} no existe en la tabla Master");
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al validar movimiento de stock");
                return (false, $"Error al validar movimiento de stock: {ex.Message}");
            }
        }

        /// <summary>
        /// Valida si existe un remito duplicado (mismo prefijo y número)
        /// Busca en MOVISTO por NUMPVTA y DOPRINUM
        /// </summary>
        /// <param name="prefijo">Prefijo del remito (NUMPVTA)</param>
        /// <param name="numero">Número del remito (DOPRINUM)</param>
        /// <returns>True si existe, False si no existe</returns>
        public async Task<bool> ExisteRemitoDuplicadoAsync(short prefijo, int numero)
        {
            try
            {
                _logger.LogInformation("Validando existencia de remito {Prefijo}-{Numero}", prefijo, numero);

                // Buscar en MOVISTO si existe un movimiento con ese prefijo y número
                // CodiMovi = "RT" para remitos de cliente
                var existe = await _context.Set<Movisto>()
                    .AnyAsync(m => m.NUMPVTA == prefijo && 
                                   m.DOPRINUM == numero && 
                                   m.CodiMovi == "RT");

                if (existe)
                {
                    _logger.LogWarning("Ya existe un remito con prefijo {Prefijo} y número {Numero}", prefijo, numero);
                }

                return existe;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al validar remito duplicado {Prefijo}-{Numero}", prefijo, numero);
                throw;
            }
        }

        /// <summary>
        /// Obtiene los remitos de un mes específico para anulación
        /// Replica la lógica del VB6 para listar remitos del mes
        /// </summary>
        /// <param name="anio">Año</param>
        /// <param name="mes">Mes (1-12)</param>
        /// <param name="numeroCliente">Número de cliente (opcional)</param>
        /// <returns>Lista de remitos del mes</returns>
        public async Task<List<RemitoAnularDTO>> ObtenerRemitosMesAsync(int anio, int mes, int? numeroCliente = null)
        {
            try
            {
                _logger.LogInformation("Obteniendo remitos del mes {Mes}/{Anio}, Cliente: {Cliente}", 
                    mes, anio, numeroCliente?.ToString() ?? "Todos");

                // Obtener código de movimiento configurado
                var codigoMov = await ObtenerCodigoMovimientoConfiguradoAsync();
                if (string.IsNullOrWhiteSpace(codigoMov) || codigoMov.Length != 2)
                {
                    codigoMov = "RT";
                }

                // Calcular fechas de inicio y fin del mes
                var fechaInicio = new DateTime(anio, mes, 1);
                var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);

                _logger.LogInformation("Buscando remitos entre {FechaInicio} y {FechaFin} con código {CodigoMov}", 
                    fechaInicio, fechaFin, codigoMov);

                // Consultar MOVISTO para obtener remitos distintos del mes
                var query = _context.Set<Movisto>()
                    .Where(m => m.FechMov >= fechaInicio && 
                               m.FechMov <= fechaFin &&
                               (m.CodiMovi == codigoMov || m.CodiMovi == "DV") &&
                               m.CantSalid > 0 &&
                               !m.ReferCor.StartsWith("ANUL")); // Excluir anulados

                // Aplicar filtro de cliente si se especificó
                if (numeroCliente.HasValue)
                {
                    query = query.Where(m => m.Nume_Clie == numeroCliente.Value);
                    _logger.LogInformation("Aplicando filtro de cliente: {NumeroCliente}", numeroCliente.Value);
                }

                var remitos = await query
                    .GroupBy(m => new { m.DoPriCor, m.DoPriLar, m.Nume_Clie, m.FechMov, m.ReferCor })
                    .Select(g => new
                    {
                        DoPriCor = g.Key.DoPriCor,
                        DoPriLar = g.Key.DoPriLar,
                        NumeClie = g.Key.Nume_Clie,
                        FechMov = g.Key.FechMov,
                        ReferCor = g.Key.ReferCor
                    })
                    .OrderBy(r => r.FechMov)
                    .ToListAsync();

                _logger.LogInformation("Remitos encontrados: {Count}", remitos.Count);

                // Obtener razones sociales de los clientes
                var numerosClientes = remitos.Select(r => r.NumeClie).Distinct().ToList();
                var clientes = await _context.Set<Deudor12>()
                    .Where(d => numerosClientes.Contains(d.Nume_Cli))
                    .Select(d => new { d.Nume_Cli, d.Raso_Cli })
                    .ToListAsync();

                var clientesDic = clientes.ToDictionary(c => c.Nume_Cli, c => c.Raso_Cli ?? $"CLIENTE Nro.{c.Nume_Cli}");

                // Mapear a DTO
                var resultado = remitos.Select(r => new RemitoAnularDTO
                {
                    DoPriCor = r.DoPriCor ?? string.Empty,
                    DoPriLar = r.DoPriLar ?? string.Empty,
                    NumeClie = r.NumeClie ?? 0,
                    RazonSocial = clientesDic.ContainsKey(r.NumeClie ?? 0) 
                        ? clientesDic[r.NumeClie ?? 0] 
                        : $"CLIENTE Nro.{r.NumeClie}",
                    FechMov = r.FechMov ?? DateTime.MinValue,
                    ReferCor = r.ReferCor ?? string.Empty
                }).ToList();

                return resultado;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener remitos del mes {Mes}/{Anio}", mes, anio);
                throw;
            }
        }

        public async Task<List<RemitoAnularDTO>> ObtenerRemitosMesTodosAsync(int anio, int mes, int? numeroCliente = null)
        {
            try
            {
                _logger.LogInformation("Obteniendo TODOS los remitos del mes {Mes}/{Anio}, Cliente: {Cliente}",
                    mes, anio, numeroCliente?.ToString() ?? "Todos");

                var codigoMov = await ObtenerCodigoMovimientoConfiguradoAsync();
                if (string.IsNullOrWhiteSpace(codigoMov) || codigoMov.Length != 2)
                {
                    codigoMov = "RT";
                }

                var fechaInicio = new DateTime(anio, mes, 1);
                var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);

                _logger.LogInformation("Buscando TODOS los remitos entre {FechaInicio} y {FechaFin} con código {CodigoMov}",
                    fechaInicio, fechaFin, codigoMov);

                var query = _context.Set<Movisto>()
                    .Where(m => m.FechMov >= fechaInicio &&
                               m.FechMov <= fechaFin &&
                               (m.CodiMovi == codigoMov || m.CodiMovi == "DV"));

                if (numeroCliente.HasValue)
                {
                    query = query.Where(m => m.Nume_Clie == numeroCliente.Value);
                    _logger.LogInformation("Aplicando filtro de cliente (TODOS): {NumeroCliente}", numeroCliente.Value);
                }

                var remitos = await query
                    .GroupBy(m => new { m.DoPriLar, m.Nume_Clie, m.FechMov })
                    .Select(g => new RemitoAnularDTO
                    {
                        DoPriLar = g.Key.DoPriLar ?? "",
                        NumeClie = g.Key.Nume_Clie ?? 0,
                        FechMov = g.Key.FechMov ?? DateTime.MinValue,
                        ReferCor = g.FirstOrDefault().ReferCor ?? "",
                        RazonSocial = _context.Set<Deudor12>()
                            .Where(c => c.Nume_Cli == (g.Key.Nume_Clie ?? 0))
                            .Select(c => c.Raso_Cli)
                            .FirstOrDefault() ?? ""
                    })
                    .OrderBy(r => r.FechMov)
                    .ThenBy(r => r.DoPriLar)
                    .ToListAsync();

                _logger.LogInformation("Se encontraron {Cantidad} remitos", remitos.Count);
                return remitos;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener remitos del mes {Mes}/{Anio}", mes, anio);
                throw;
            }
        }

        /// <summary>
        /// Obtiene los remitos de traslado del mes (códigos RC y DV)
        /// </summary>
        public async Task<List<RemitoAnularDTO>> ObtenerRemitosTrasladoMesAsync(int anio, int mes, int? numeroCliente = null)
        {
            try
            {
                _logger.LogInformation("Obteniendo REMITOS DE TRASLADO del mes {Mes}/{Anio}, Cliente: {Cliente}",
                    mes, anio, numeroCliente?.ToString() ?? "Todos");

                var fechaInicio = new DateTime(anio, mes, 1);
                var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);

                _logger.LogInformation("Buscando remitos de traslado entre {FechaInicio} y {FechaFin} con códigos RC y DV",
                    fechaInicio, fechaFin);

                var query = _context.Set<Movisto>()
                    .Where(m => m.FechMov >= fechaInicio &&
                               m.FechMov <= fechaFin &&
                               (m.CodiMovi == "RC" || m.CodiMovi == "DV"));

                if (numeroCliente.HasValue)
                {
                    query = query.Where(m => m.Nume_Clie == numeroCliente.Value);
                    _logger.LogInformation("Aplicando filtro de cliente (TRASLADO): {NumeroCliente}", numeroCliente.Value);
                }

                var remitos = await query
                    .GroupBy(m => new { m.DoPriLar, m.Nume_Clie, m.FechMov })
                    .Select(g => new RemitoAnularDTO
                    {
                        DoPriLar = g.Key.DoPriLar ?? "",
                        NumeClie = g.Key.Nume_Clie ?? 0,
                        FechMov = g.Key.FechMov ?? DateTime.MinValue,
                        ReferCor = g.FirstOrDefault().ReferCor ?? "",
                        RazonSocial = _context.Set<Deudor12>()
                            .Where(c => c.Nume_Cli == (g.Key.Nume_Clie ?? 0))
                            .Select(c => c.Raso_Cli)
                            .FirstOrDefault() ?? ""
                    })
                    .OrderBy(r => r.FechMov)
                    .ThenBy(r => r.DoPriLar)
                    .ToListAsync();

                _logger.LogInformation("Se encontraron {Cantidad} remitos de traslado", remitos.Count);
                return remitos;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener remitos de traslado del mes {Mes}/{Anio}", mes, anio);
                throw;
            }
        }

        /// <summary>
        /// Obtiene el código de movimiento configurado en FLEXCRL
        /// </summary>
        private async Task<string> ObtenerCodigoMovimientoConfiguradoAsync()
        {
            try
            {
                var config = await _context.Set<FLEXCRL>()
                    .AsNoTracking()
                    .FirstOrDefaultAsync(f => f.CLABUS == "" && f.PROBUS == "CODREFAC");

                return config?.VALCONTROL ?? "RT";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener código de movimiento configurado");
                return "RT";
            }
        }

        /// <summary>
        /// Obtiene el tipo de cambio para una moneda específica
        /// Equivalente a la función TICAMONE en VB6
        /// </summary>
        private async Task<decimal> ObtenerTipoCambioAsync(short moneda)
        {
            try
            {
                // Aquí se debería implementar la lógica para obtener el tipo de cambio
                // Por ahora retornamos 1 como valor por defecto
                // En VB6 esto se obtiene de una tabla de monedas
                return 1;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener tipo de cambio para moneda {Moneda}", moneda);
                return 1; // Valor por defecto
            }
        }

        /// <summary>
        /// Anula un remito específico
        /// Replica la funcionalidad de anulación de remitos de VB6
        /// 
        /// LÓGICA VB6 REPLICADA:
        /// 1. Busca CodiMovi en MOVISTO por DoPriCor, si no existe busca en MOVIREPA
        /// 2. Determina la tabla a usar (MOVISTO o MOVIREPA) según el CodiMovi encontrado
        /// 3. Actualiza pedidos vinculados: reduce cantidad entregada, aumenta saldo pendiente
        /// 4. Marca remito como anulado: CantSalid=0, PESOKG=0, METROS=0, ReferCor con prefijo "(Anul)"
        /// </summary>
        public async Task<AnularRemitoResponse> AnularRemitoAsync(AnularRemitoRequest request)
        {
            var response = new AnularRemitoResponse();

            try
            {
                _logger.LogInformation("Iniciando anulación de remito {NumeroRemito} para cliente {NumeroCliente}", 
                    request.NumeroRemito, request.NumeroCliente);

                // PASO 1: Determinar si el remito está en MOVISTO o MOVIREPA
                // Replica: COMOV$ = TRIM$(VALOBADA("MOVISTO", "CODIMOVI", "DoPriCor = '" & REMISEL$ & "'", "NOMESS"))
                var codiMoviMovisto = await _context.Set<Movisto>()
                    .Where(m => m.DoPriCor == request.NumeroRemito)
                    .Select(m => m.CodiMovi)
                    .FirstOrDefaultAsync();

                string tablaUtilizar;
                string codiMovi;

                if (!string.IsNullOrEmpty(codiMoviMovisto))
                {
                    // El remito está en MOVISTO
                    tablaUtilizar = "MOVISTO";
                    codiMovi = codiMoviMovisto;
                    _logger.LogInformation("Remito encontrado en MOVISTO con CodiMovi: {CodiMovi}", codiMovi);
                }
                else
                {
                    // Buscar en MOVIREPA
                    // Replica: If COMOV$ = "" Then COMOV$ = TRIM$(VALOBADA("MOVIREPA", "CODIMOVI", "DoPriCor = '" & REMISEL$ & "'", "NOMESS"))
                    var codiMoviMovirepa = await _context.Set<Movirepa>()
                        .Where(m => m.DoPriCor == request.NumeroRemito)
                        .Select(m => m.CodiMovi)
                        .FirstOrDefaultAsync();

                    if (!string.IsNullOrEmpty(codiMoviMovirepa))
                    {
                        tablaUtilizar = "MOVIREPA";
                        codiMovi = codiMoviMovirepa;
                        _logger.LogInformation("Remito encontrado en MOVIREPA con CodiMovi: {CodiMovi}", codiMovi);
                    }
                    else
                    {
                        response.Exitoso = false;
                        response.Mensaje = $"No se encontró el remito {request.NumeroRemito} en ninguna tabla";
                        return response;
                    }
                }

                response.TablaUtilizada = tablaUtilizar;

                // PASO 2: Obtener movimientos del remito para actualizar pedidos vinculados
                // Replica: SELECT * FROM MOVISTO WHERE DoPriCor = ... AND Nume_Clie = ... AND FechMov = ...
                List<dynamic> movimientos;
                
                if (tablaUtilizar == "MOVISTO")
                {
                    movimientos = await _context.Set<Movisto>()
                        .Where(m => m.DoPriCor == request.NumeroRemito &&
                                   m.Nume_Clie == request.NumeroCliente &&
                                   m.FechMov == request.FechaMovimiento)
                        .Select(m => new
                        {
                            m.NuPedCli,
                            m.Cod_Inte,
                            m.CantSalid,
                            m.DoPriLar
                        })
                        .Cast<dynamic>()
                        .ToListAsync();
                }
                else
                {
                    movimientos = await _context.Set<Movirepa>()
                        .Where(m => m.DoPriCor == request.NumeroRemito &&
                                   m.Nume_Clie == request.NumeroCliente &&
                                   m.FechMov == request.FechaMovimiento)
                        .Select(m => new
                        {
                            m.NuPedCli,
                            m.Cod_Inte,
                            m.CantSalid,
                            m.DoPriLar
                        })
                        .Cast<dynamic>()
                        .ToListAsync();
                }

                // PASO 3: Actualizar pedidos vinculados
                // Replica el código VB6 que actualiza PEDDETA
                int pedidosActualizados = 0;

                foreach (var movimiento in movimientos)
                {
                    int? nroPedido = movimiento.NuPedCli;
                    
                    if (nroPedido.HasValue && nroPedido.Value > 0)
                    {
                        int codigoInterno = movimiento.Cod_Inte ?? 0;
                        decimal cantidadSalida = movimiento.CantSalid ?? 0;

                        // Buscar el detalle del pedido
                        // Replica: SELECT * FROM PEDDETA WHERE NroPed = ... AND CodiInt = ... AND NroClie = ... AND CantEnt >= ...
                        var pedidoDetalle = await _context.Set<Peddeta>()
                            .Where(pd => pd.NroPed == nroPedido.Value &&
                                        pd.CodiInt == codigoInterno &&
                                        pd.NroClie == request.NumeroCliente &&
                                        pd.CantEnt >= (cantidadSalida - 0.05m) && // Tolerancia de 0.05
                                        pd.Status < 9)
                            .FirstOrDefaultAsync();

                        if (pedidoDetalle != null)
                        {
                            // Actualizar cantidad entregada y saldo
                            // Replica: ENTREGA# = XVALO(!CANTENT) - CANTSALD#
                            //          SALDO# = XVALO(!CanSaldo) + CANTSALD#
                            decimal nuevaCantidadEntregada = (pedidoDetalle.CantEnt ?? 0) - cantidadSalida;
                            decimal nuevoSaldo = (pedidoDetalle.CanSaldo ?? 0) + cantidadSalida;

                            pedidoDetalle.CantEnt = nuevaCantidadEntregada;
                            pedidoDetalle.CanSaldo = nuevoSaldo;
                            pedidoDetalle.Status = 0; // Cambiar a pendiente
                            pedidoDetalle.NroRemi = ""; // Limpiar número de remito

                            _logger.LogInformation("Actualizando pedido {NroPedido}, producto {CodigoInterno}: " +
                                "CantEnt: {CantEnt} -> {NuevaCantEnt}, CanSaldo: {CanSaldo} -> {NuevoSaldo}",
                                nroPedido.Value, codigoInterno,
                                pedidoDetalle.CantEnt + cantidadSalida, nuevaCantidadEntregada,
                                pedidoDetalle.CanSaldo - cantidadSalida, nuevoSaldo);

                            pedidosActualizados++;
                        }
                    }
                }

                response.PedidosActualizados = pedidosActualizados;

                // PASO 4: Marcar remito como anulado
                // Replica: UPDATE MOVISTO SET CANTSALID=0, REFERCOR='(Anul) ' + REFERCOR, PESOKG=0, METROS=0
                int registrosAfectados;

                if (tablaUtilizar == "MOVISTO")
                {
                    var movistosAnular = await _context.Set<Movisto>()
                        .Where(m => m.DoPriCor == request.NumeroRemito &&
                                   m.Nume_Clie == request.NumeroCliente &&
                                   m.FechMov == request.FechaMovimiento)
                        .ToListAsync();

                    foreach (var movisto in movistosAnular)
                    {
                        movisto.CantSalid = 0;
                        movisto.PESOKG = 0;
                        movisto.METROS = 0;
                        
                        // Agregar prefijo "(Anul) " si no lo tiene ya
                        if (!movisto.ReferCor?.StartsWith("(Anul)") ?? true)
                        {
                            movisto.ReferCor = "(Anul) " + (movisto.ReferCor ?? "");
                        }
                    }

                    registrosAfectados = movistosAnular.Count;
                }
                else
                {
                    var movirepasAnular = await _context.Set<Movirepa>()
                        .Where(m => m.DoPriCor == request.NumeroRemito &&
                                   m.Nume_Clie == request.NumeroCliente &&
                                   m.FechMov == request.FechaMovimiento)
                        .ToListAsync();

                    foreach (var movirepa in movirepasAnular)
                    {
                        movirepa.CantSalid = 0;
                        movirepa.PESOKG = 0;
                        movirepa.METROS = 0;
                        
                        // Agregar prefijo "(Anul) " si no lo tiene ya
                        if (!movirepa.ReferCor?.StartsWith("(Anul)") ?? true)
                        {
                            movirepa.ReferCor = "(Anul) " + (movirepa.ReferCor ?? "");
                        }
                    }

                    registrosAfectados = movirepasAnular.Count;
                }

                // Guardar todos los cambios
                await _context.SaveChangesAsync();

                response.Exitoso = true;
                response.RegistrosMovimientosAfectados = registrosAfectados;
                response.Mensaje = $"Remito {request.NumeroRemito} anulado exitosamente. " +
                    $"Registros afectados: {registrosAfectados}, Pedidos actualizados: {pedidosActualizados}";

                _logger.LogInformation("Anulación completada: {Mensaje}", response.Mensaje);

                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al anular remito {NumeroRemito}", request.NumeroRemito);
                response.Exitoso = false;
                response.Mensaje = $"Error al anular remito: {ex.Message}";
                return response;
            }
        }

        /// <summary>
        /// Actualiza la cantidad entregada (CantEnt) en PEDDETA cuando se registra un remito
        /// Suma la cantidad remitida a la cantidad ya entregada y actualiza el saldo
        /// </summary>
        private async Task ActualizarCantidadEntregadaPedidoAsync(
            int numeroPedido, 
            int codigoInterno, 
            int numeroCliente,
            decimal cantidadRemitida,
            string numeroRemito,
            short numeOrdPed = 0)
        {
            try
            {
                _logger.LogInformation("Actualizando cantidad entregada para pedido {NroPed}, producto {CodiInt}, NuOrItem {NuOrItem}, cantidad {Cantidad}",
                    numeroPedido, codigoInterno, numeOrdPed, cantidadRemitida);

                // Buscar el detalle del pedido
                // Si numeOrdPed > 0, buscar por ese campo específico, sino buscar solo por producto
                var query = _context.Set<Peddeta>()
                    .Where(pd => pd.NroPed == numeroPedido &&
                                pd.CodiInt == codigoInterno &&
                                pd.NroClie == numeroCliente &&
                                pd.Status < 9); // Solo pedidos activos

                // Si se especifica NuOrItem, filtrar también por ese campo
                if (numeOrdPed > 0)
                {
                    query = query.Where(pd => pd.NuOrItem == numeOrdPed);
                }

                var pedidoDetalle = await query.FirstOrDefaultAsync();

                if (pedidoDetalle != null)
                {
                    // Sumar la cantidad remitida a la cantidad ya entregada
                    decimal cantidadEntregadaAnterior = pedidoDetalle.CantEnt ?? 0;
                    decimal nuevaCantidadEntregada = cantidadEntregadaAnterior + cantidadRemitida;
                    
                    // Actualizar saldo (restar la cantidad remitida del saldo)
                    decimal saldoAnterior = pedidoDetalle.CanSaldo ?? 0;
                    decimal nuevoSaldo = saldoAnterior - cantidadRemitida;

                    pedidoDetalle.CantEnt = nuevaCantidadEntregada;
                    pedidoDetalle.CanSaldo = nuevoSaldo;
                    pedidoDetalle.FeUltEnt = DateTime.Now; // Fecha de última entrega
                    pedidoDetalle.NroRemi = numeroRemito; // Número de remito

                    // Si el saldo es <= 0, marcar como entregado completamente
                    if (nuevoSaldo <= 0.005m)
                    {
                        pedidoDetalle.Status = 8; // Completado
                    }

                    await _context.SaveChangesAsync();

                    _logger.LogInformation("Pedido actualizado: {NroPed}, producto {CodiInt}: " +
                        "CantEnt: {CantEntAnterior} -> {NuevaCantEnt}, CanSaldo: {SaldoAnterior} -> {NuevoSaldo}",
                        numeroPedido, codigoInterno,
                        cantidadEntregadaAnterior, nuevaCantidadEntregada,
                        saldoAnterior, nuevoSaldo);
                }
                else
                {
                    _logger.LogWarning("No se encontró detalle de pedido para actualizar: Pedido {NroPed}, Producto {CodiInt}, Cliente {NroClie}",
                        numeroPedido, codigoInterno, numeroCliente);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar cantidad entregada del pedido {NroPed}, producto {CodiInt}",
                    numeroPedido, codigoInterno);
                // No lanzar excepción para no interrumpir el registro del movimiento
            }
        }

        /// <summary>
        /// Obtiene el último número de remito registrado en MOVISTO
        /// Filtra por código de movimiento y fecha y hora mayor o igual a la fecha de nueva numeración
        /// Para remitos de traslado (RC), usa el campo DoPriLar con formato RC-XXXX o RC.XXXX
        /// </summary>
        /// <param name="fechaNuevaNumeracion">Fecha y hora desde la cual considerar los remitos (opcional, default 01/01/1980 00:00)</param>
        /// <param name="codigoMovimiento">Código de movimiento a filtrar (opcional, default "RT" para remitos de cliente, usar "RC" para traslados)</param>
        /// <returns>Último número de remito</returns>
        public async Task<int> ObtenerUltimoNumeroRemitoAsync(DateTime? fechaNuevaNumeracion = null, string codigoMovimiento = "RT")
        {
            try
            {
                // Si no se proporciona fecha, usar 01/01/1980 00:00 como default
                DateTime fechaFiltro = fechaNuevaNumeracion ?? new DateTime(1980, 1, 1, 0, 0, 0);

                _logger.LogInformation("Obteniendo último número de remito para código {CodigoMovimiento} desde fecha y hora: {FechaFiltro:dd/MM/yyyy HH:mm}", codigoMovimiento, fechaFiltro);

                int resultado = 0;

                if (codigoMovimiento == "RC")
                {
                    // Para remitos de traslado, usar DoPriLar con formato RC-XXXX o RC.XXXX
                    var remitosTraslado = await _context.Set<Movisto>()
                        .Where(m => m.CodiMovi == codigoMovimiento && m.FeReMovi >= fechaFiltro && m.DoPriLar != null)
                        .Select(m => m.DoPriLar)
                        .ToListAsync();

                    // Extraer el número de cada DoPriLar y encontrar el máximo
                    foreach (var doprilar in remitosTraslado)
                    {
                        int numero = ExtraerNumeroDeDoprilar(doprilar);
                        if (numero > resultado)
                        {
                            resultado = numero;
                        }
                    }

                    _logger.LogInformation("Último número de remito encontrado para código {CodigoMovimiento} usando DoPriLar: {UltimoNumero}", codigoMovimiento, resultado);
                }
                else
                {
                    // Para remitos de cliente (RT), usar DOPRINUM
                    var ultimoRemito = await _context.Set<Movisto>()
                        .Where(m => m.CodiMovi == codigoMovimiento && m.FeReMovi >= fechaFiltro)
                        .OrderByDescending(m => m.DOPRINUM)
                        .Select(m => m.DOPRINUM)
                        .FirstOrDefaultAsync();

                    resultado = ultimoRemito ?? 0;
                    _logger.LogInformation("Último número de remito encontrado para código {CodigoMovimiento} usando DOPRINUM: {UltimoNumero}", codigoMovimiento, resultado);
                }

                return resultado;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener último número de remito");
                return 0;
            }
        }

        /// <summary>
        /// Extrae el número de un campo DoPriLar con formato RC-XXXX o RC.XXXX
        /// </summary>
        private int ExtraerNumeroDeDoprilar(string doprilar)
        {
            if (string.IsNullOrWhiteSpace(doprilar))
            {
                return 0;
            }

            // Intentar formato RC-XXXX
            if (doprilar.StartsWith("RC-"))
            {
                string numeroStr = doprilar.Substring(3);
                if (int.TryParse(numeroStr, out int numero))
                {
                    return numero;
                }
            }

            // Intentar formato RC.XXXX
            if (doprilar.StartsWith("RC."))
            {
                string numeroStr = doprilar.Substring(3);
                if (int.TryParse(numeroStr, out int numero))
                {
                    return numero;
                }
            }

            return 0;
        }
    }
}
