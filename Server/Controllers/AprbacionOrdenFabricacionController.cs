using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Data;
using System.Linq;
using System.Security.Claims;



namespace IMPLANPROD.Server.Controllers
{
    // TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Actualización del controlador para usar el servicio centralizado
    // de contexto de usuario en lugar de inyectar múltiples servicios relacionados con el usuariO

    /// <summary>
    /// Controlador para la gestión de órdenes de fabricación
    /// </summary>
    [ApiController]
    [Route("/api/AprobacionOrdenFabricacion")]
    public class AprbacionOrdenFabricacionController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<AprbacionOrdenFabricacionController> _logger;
        private readonly IUserContextService _userContextService;
        private readonly IMoviStockService _moviStockService;

        /// <summary>
        /// Constructor del controlador de transferencias de stock
        /// </summary>
        /// <param name="context">Contexto de datos</param>
        /// <param name="logger">Logger para registro de eventos</param>
        /// <param name="userContextService">Servicio centralizado de contexto de usuario</param>
        /// <param name="moviStockService">Servicio para movimientos de stock</param>
        public AprbacionOrdenFabricacionController(DataContext context, ILogger<AprbacionOrdenFabricacionController> logger, IUserContextService userContextService, IMoviStockService moviStockService)
        {
            _context = context;
            _logger = logger;
            _userContextService = userContextService;
            _moviStockService = moviStockService;
        }
        [HttpPost("pendientesAprobacion")]
        public async Task<ActionResult<List<Orfapen>>> GetOFabricacionPendientes([FromBody] OFabricacionRequestDTO request)
        {
            try
            {
                _logger.LogInformation($"Obteniendo órdenes pendientes de aprobación para usuario {request.IdFlexoftUsuario} con depósitos: {(request.Depositos != null ? string.Join(", ", request.Depositos.Select(d => d.ToString())) : "TODOS (SUPERADMIN)")}");

                // Validar parámetros de entrada
                if (request == null)
                {
                    return BadRequest("Los parámetros de solicitud son requeridos");
                }

                // Si es SUPERADMIN (IdFlexoftUsuario null y Depositos null), no aplicar filtros
                bool esSuperAdmin = request.IdFlexoftUsuario == null && (request.Depositos == null || !request.Depositos.Any());

                if (!esSuperAdmin)
                {
                    if (request.IdFlexoftUsuario <= 0)
                    {
                        return BadRequest("El ID del usuario es requerido");
                    }

                    if (request.Depositos == null || !request.Depositos.Any())
                    {
                        return BadRequest("Al menos un depósito es requerido");
                    }
                }

                // Consultar órdenes de fabricación pendientes de aprobación
                // Buscar órdenes que tengan movimientos en ORFAPEN con cantidades pendientes de aprobar


                if (!esSuperAdmin && (request == null || request.IdFlexoftUsuario <= 0 || request.Depositos == null || !request.Depositos.Any()))
                    return BadRequest("Parámetros inválidos");

                // Construir filtro de depósitos (solo si no es SUPERADMIN)
                var queryBase = esSuperAdmin
                    ? _context.Orfapens
                    : (request.IdFlexoftUsuario == 14
                        ? _context.Orfapens.Where(o => request.Depositos.Contains(o.DepoMovi!.Value) || o.DepoMovi.Value >= 30)
                        : _context.Orfapens.Where(o => request.Depositos.Contains(o.DepoMovi!.Value)));

                // Aplicar filtro de texto ANTES del GroupBy para mejorar rendimiento
                if (!string.IsNullOrWhiteSpace(request.Filter))
                {
                    var terms = request.Filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        queryBase = queryBase.Where(x =>
                            (x.DoPriCor ?? "").ToLower().Contains(term) ||
                            (x.Cod_Exte ?? "").ToLower().Contains(term) ||
                            (x.IDTEXT ?? "").ToLower().Contains(term) ||
                            (x.Cod_Inte.HasValue && x.Cod_Inte.Value.ToString().Contains(term)) ||
                            (x.DepoMovi.HasValue && x.DepoMovi.Value.ToString().Contains(term)) ||
                            (x.FechMov.HasValue && x.FechMov.Value.ToString().ToLower().Contains(term))
                        );
                    }
                }

                var query = queryBase
                    .Where(o => o.CodiMovi == "OF")
                    .GroupBy(o => new { o.DoPriCor, o.IDTEXT })
                    .Where(g => g.Sum(x => x.CantIngre) > g.Sum(x => x.CantSalid))
                    .Select(g => new Orfapen
                    {
                        Cod_Inte = g.FirstOrDefault().Cod_Inte,
                        Cod_Exte = g.FirstOrDefault().Cod_Exte,
                        CantIngre = g.Sum(x => x.CantIngre),
                        CantSalid = g.Sum(x => x.CantSalid),
                        DoPriCor = g.Key.DoPriCor,
                        FechMov = g.FirstOrDefault().FechMov,
                        IDTEXT = g.Key.IDTEXT,
                        DepoMovi = g.FirstOrDefault().DepoMovi,
                        Ajuste = g.Sum(x => x.CantIngre ?? 0) - g.Sum(x => x.CantSalid ?? 0)
                    });

                var pagination = new PaginationDTO
                {
                    Page = request.Page <= 0 ? 1 : request.Page,
                    RecordsNumber = request.RecordsNumber <= 0 ? 10 : request.RecordsNumber,
                    Filter = request.Filter
                };

                var ordenesPendientes = await query
                    .OrderByDescending(o => o.DoPriCor)
                    .ThenByDescending(o => o.FechMov)
                    .Paginate(pagination)
                    .ToListAsync();

                _logger.LogInformation($"Se encontraron {ordenesPendientes.Count} órdenes pendientes (página {pagination.Page}, records {pagination.RecordsNumber})");

                return Ok(ordenesPendientes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener órdenes pendientes de aprobación");
                return StatusCode(500, "Error interno del servidor al obtener las órdenes pendientes");
            }
        }

        [HttpPost("totalPages")]
        public async Task<ActionResult<int>> GetPagesPendientes([FromBody] OFabricacionRequestDTO request)
        {
            try
            {
                // Si es SUPERADMIN (IdFlexoftUsuario null y Depositos null), no aplicar filtros
                bool esSuperAdmin = request.IdFlexoftUsuario == null && (request.Depositos == null || !request.Depositos.Any());

                if (!esSuperAdmin && (request == null || request.IdFlexoftUsuario <= 0 || request.Depositos == null || !request.Depositos.Any()))
                {
                    return BadRequest("Parámetros inválidos");
                }

                var queryBase = esSuperAdmin
                    ? _context.Orfapens
                    : (request.IdFlexoftUsuario == 14
                        ? _context.Orfapens.Where(o => request.Depositos.Contains(o.DepoMovi!.Value) || o.DepoMovi.Value >= 30)
                        : _context.Orfapens.Where(o => request.Depositos.Contains(o.DepoMovi!.Value)));

                // Aplicar filtro de texto ANTES del GroupBy para mejorar rendimiento
                if (!string.IsNullOrWhiteSpace(request.Filter))
                {
                    var terms = request.Filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        queryBase = queryBase.Where(x =>
                            (x.DoPriCor ?? "").ToLower().Contains(term) ||
                            (x.Cod_Exte ?? "").ToLower().Contains(term) ||
                            (x.IDTEXT ?? "").ToLower().Contains(term) ||
                            (x.Cod_Inte.HasValue && x.Cod_Inte.Value.ToString().Contains(term)) ||
                            (x.DepoMovi.HasValue && x.DepoMovi.Value.ToString().Contains(term)) ||
                            (x.FechMov.HasValue && x.FechMov.Value.ToString().ToLower().Contains(term))
                        );
                    }
                }

                var query = queryBase
                    .Where(o => o.CodiMovi == "OF")
                    .GroupBy(o => new { o.DoPriCor, o.IDTEXT })
                    .Where(g => g.Sum(x => x.CantIngre) > g.Sum(x => x.CantSalid))
                    .Select(g => new Orfapen
                    {
                        Cod_Inte = g.FirstOrDefault().Cod_Inte,
                        Cod_Exte = g.FirstOrDefault().Cod_Exte,
                        CantIngre = g.Sum(x => x.CantIngre),
                        CantSalid = g.Sum(x => x.CantSalid),
                        DoPriCor = g.Key.DoPriCor,
                        FechMov = g.FirstOrDefault().FechMov,
                        IDTEXT = g.Key.IDTEXT,
                        DepoMovi = g.FirstOrDefault().DepoMovi,
                        Ajuste = g.Sum(x => x.CantIngre ?? 0) - g.Sum(x => x.CantSalid ?? 0)
                    });

                var recordsNumber = request.RecordsNumber <= 0 ? 10 : request.RecordsNumber;
                var count = await query.CountAsync();
                var totalPages = (int)Math.Ceiling(count / (double)recordsNumber);
                if (totalPages <= 0) totalPages = 1;

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener totalPages de pendientes de aprobación");
                return StatusCode(500, "Error interno del servidor al obtener totalPages");
            }
        }

        /// <summary>
        /// Procesa la aprobación de una orden de fabricación con cierre existente
        /// Replica la lógica VB6 original para aprobación de órdenes
        /// </summary>
        /// <param name="request">Datos de la aprobación</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPost("procesarAprobacionConCierre")]
        public async Task<IActionResult> ProcesarAprobacionConCierre([FromBody] AprobacionCierreDTO request)
        {
            try
            {
                // Verificar autenticación
                if (!_userContextService.IsAuthenticated())
                {
                    return BadRequest("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                // Obtener datos del registro ORFAPEN seleccionado
                var datosOrfapen = await ObtenerDatosOrfapen(request.IdText, request.IdSubOr);
                if (datosOrfapen == null)
                {
                    return NotFound("No se encontró el registro ORFAPEN especificado.");
                }

                // Calcular saldo pendiente: ORFAPEN.CANTINGRE - ORFAPEN.CANTSALID
                decimal saldo = (datosOrfapen.CantIngre ?? 0) - (datosOrfapen.CantSalid ?? 0);
                if (saldo <= 0)
                {
                    return BadRequest("No hay saldo pendiente para aprobar.");
                }

                // Obtener datos adicionales de la orden de fabricación
                var ordenFabricacion = await _context.Ordefabrs
                    .FirstOrDefaultAsync(o => o.IdSubOr == request.IdSubOr);
                
                if (ordenFabricacion == null)
                {
                    return NotFound($"No se encontró la orden de fabricación {request.IdSubOr}.");
                }

                // Calcular cantidad a aprobar según lógica VB6
                decimal cantidadAprobar = await CalcularCantidadAprobar(ordenFabricacion, saldo);

                // Validar que la cantidad solicitada no supere el saldo
                if (request.CantidadAprobada > saldo)
                {
                    return BadRequest($"Imposible aprobar más cantidad que la pendiente. Saldo disponible: {saldo}");
                }

                // Usar la cantidad solicitada o la calculada
                decimal cantidadFinal = request.CantidadAprobada > 0 ? request.CantidadAprobada : cantidadAprobar;

                // PASO 1: Procesar movimiento de stock estándar (MoviStockService maneja su propia transacción)
                await ProcesarMovimientoStock(datosOrfapen, cantidadFinal, request.DepositoEntrada);

                // PASO 2: Procesar consumo de componentes (cada movimiento se guarda individualmente)
                _logger.LogInformation($"Iniciando PASO 2: ProcesarConsumoComponentes para orden {request.IdSubOr}, IDTEXT: {request.IdText}, cantidad: {cantidadFinal}");
                await ProcesarConsumoComponentes(request.IdSubOr, request.IdText, cantidadFinal);
                _logger.LogInformation($"PASO 2 completado: ProcesarConsumoComponentes");

                // PASO 3: Usar transacción solo para cambios en EF Core
                using var transaction = await _context.Database.BeginTransactionAsync();
                try
                {
                    // Agregar registro ORFAPEN para aprobación
                    await AgregarRegistroOrfapenAprobacion(datosOrfapen, cantidadFinal, request.DepositoEntrada, request.IdText);

                    // Actualizar orden de fabricación
                    bool ordenCerrada = await ActualizarOrdenFabricacion(ordenFabricacion, cantidadFinal, request.IdText);

                    // Eliminar reservas si la orden se cerró completamente
                    if (ordenCerrada)
                    {
                        await EliminarReservas(request.IdSubOr);
                    }

                    // Confirmar transacción para cambios en EF Core
                    await _context.SaveChangesAsync();
                    await transaction.CommitAsync();

                    return Ok(new { 
                        mensaje = "Aprobación procesada exitosamente", 
                        cantidadAprobada = cantidadFinal,
                        ordenCerrada = ordenCerrada
                    });
                }
                catch (Exception ex)
                {
                    await transaction.RollbackAsync();
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar aprobación con cierre");
                return StatusCode(500, $"Error al procesar la aprobación: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene los datos del registro ORFAPEN según IDTEXT e IdSubOr
        /// Equivalente a la consulta VB6: SELECT TOP 1 ReferCor,FechMov,COD_INTE...
        /// </summary>
        private async Task<Orfapen?> ObtenerDatosOrfapen(string idText, string idSubOr)
        {
            return await _context.Orfapens
                .Where(o => o.DoPriCor == idSubOr && 
                           o.CodiMovi == "OF" && 
                           o.IDTEXT == idText && 
                           o.CantIngre > 0)
                .OrderBy(o => o.FeReMovi)
                .FirstOrDefaultAsync();
        }

        /// <summary>
        /// Calcula la cantidad a aprobar según la lógica VB6
        /// </summary>
        private Task<decimal> CalcularCantidadAprobar(OrdeFabr orden, decimal saldo)
        {
            decimal cantidadAprobar = saldo;

            // Lógica VB6: If CANTFINPROCESO1 > 0 Then...
            if (orden.Canti_Fin_Proceso.HasValue && orden.Canti_Fin_Proceso.Value > 0)
            {
                decimal cantidadPendiente = orden.Canti_Fin_Proceso.Value - (orden.CanApro ?? 0);
                if (cantidadPendiente >= 1)
                {
                    cantidadAprobar = cantidadPendiente;
                }
            }

            return Task.FromResult(cantidadAprobar);
        }

        /// <summary>
        /// Procesa el movimiento de stock estándar
        /// Equivalente a: Call MOVISTOK(FF%, CIII%, NORFA$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC, DEPOEN%, CANTI)
        /// </summary>
        private async Task ProcesarMovimientoStock(Orfapen datosOrfapen, decimal cantidad, int depositoEntrada)
        {
            var movimientoStock = new MoviStockDTO
            {
                Fecha = datosOrfapen.FechMov ?? DateTime.Now,
                CodigoInterno = datosOrfapen.Cod_Inte ?? 0,
                Lote = datosOrfapen.IdenLote ?? "",
                CodigoMovimiento = datosOrfapen.CodiMovi ?? "OF",
                DocumentoPrimarioLargo = ProcesarCodigoLongitud(datosOrfapen.DoPriCor ?? ""),
                DocumentoPrimarioCorto = ProcesarCodigoLongitud(datosOrfapen.DoPriCor ?? ""),
                DocumentoSecundarioLargo = ProcesarCodigoLongitud(datosOrfapen.DoSecCor ?? ""),
                DocumentoSecundarioCorto = ProcesarCodigoLongitud(datosOrfapen.DoSecCor ?? ""),
                Referencia = datosOrfapen.ReferCor ?? "",
                NumeroOrdenItem = 1,
                ValorUnitario = 0,
                Moneda = 1,
                NumeroCliente = datosOrfapen.Nume_Clie ?? 0,
                Deposito = depositoEntrada,
                Cantidad = cantidad,
                TipoTransferencia = null
            };

            var resultado = await _moviStockService.RegistrarMovimientoAsync(movimientoStock, User);
            if (resultado is BadRequestObjectResult)
            {
                throw new Exception("Error al procesar el movimiento de stock en la aprobación");
            }
        }

        /// <summary>
        /// Agrega registro ORFAPEN para la aprobación
        /// Equivalente a: Call AGREGA_REGISTRO_NORFAPEN(..., "AproCierre")
        /// </summary>
        private async Task AgregarRegistroOrfapenAprobacion(Orfapen datosOrfapen, decimal cantidad, int depositoEntrada, string idText)
        {
            var idFlexoft = _userContextService.GetCurrentUserId();
            if (!idFlexoft.HasValue)
            {
                throw new Exception("No se pudo obtener el contexto del usuario");
            }

            // Generar ID de texto si está vacío
            string idTextoFinal = string.IsNullOrWhiteSpace(idText) ? GenerarIdTextoAleatorio() : idText;

            // Procesar códigos para asegurar longitud de 11 caracteres
            string dopriProcesado = ProcesarCodigoLongitud(datosOrfapen.DoPriCor ?? "");

            // Generar DOSEC cambiando la 'O' por 'I' en la orden de fabricación
            string ordenFabricacion = datosOrfapen.DoPriCor ?? "";
            string dosecProcesado = ProcesarCodigoLongitud(ordenFabricacion.Replace("OF-", "IF-"));

            // Obtener código externo del producto
            var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == datosOrfapen.Cod_Inte);
            string codigoExterno = master?.Codigo ?? "";

            // Crear el nuevo registro ORFAPEN para aprobación
            // Para AproCierre + OF: cantidad positiva en CantSalid (salida para compensar el ingreso del cierre)
            // Para AproCierre + CF: cantidad positiva en CantIngre (ingreso para compensar la salida del cierre)
            var nuevoRegistro = new Orfapen
            {
                FechMov = datosOrfapen.FechMov,
                Cod_Inte = datosOrfapen.Cod_Inte,
                Cod_Exte = codigoExterno,
                IdenLote = datosOrfapen.IdenLote ?? "",
                CodiMovi = datosOrfapen.CodiMovi ?? "OF",
                DoPriCor = dopriProcesado,
                DoPriLar = dopriProcesado,
                DoSecCor = dosecProcesado, // Cambiar PF- por IF-
                DoSecLar = dosecProcesado, // Cambiar PF- por IF-
                ReferCor = "Fabricación Propia", // Referencia fija para aprobación
                ValoUnit = 0,
                MoneVal = 1,
                Nume_Clie = datosOrfapen.Nume_Clie ?? 0,
                DepoMovi = depositoEntrada,
                CantIngre = 0, // Para aprobación OF: no hay ingreso
                CantSalid = cantidad, // Para aprobación OF: cantidad positiva en CantSalid
                FeReMovi = DateTime.Now,
                NumUsuar = idFlexoft.Value % 100,
                STATUS = 0,
                IDTEXT = idTextoFinal,
                DepoConsu = 0,
                IDSITUACION = "AproCierre"
            };

            _context.Orfapens.Add(nuevoRegistro);
        }

        /// <summary>
        /// Procesa el consumo de componentes
        /// Equivalente a: Sub CONSUMODECOMPONENTES(NORFA$, IDTX$, CantAprobada)
        /// </summary>
        private async Task ProcesarConsumoComponentes(string idSubOr, string idText, decimal cantidadAprobada)
        {
            try
            {
                _logger.LogInformation($"Iniciando procesamiento de consumo de componentes para orden {idSubOr}, IDTEXT: {idText}, cantidad: {cantidadAprobada}");

                // Obtener código del producto terminado (padre) desde la orden de fabricación
                var ordenFabricacion = await _context.Ordefabrs
                    .FirstOrDefaultAsync(o => o.IdSubOr == idSubOr);
                string codigoPadre = ordenFabricacion?.Cod_Exte ?? "";

                // Construir condición base: DOPRICOR = 'NORFA$' AND IDTEXT = 'IDTX$' AND IDSITUACION = 'GenCierre'
                var condicionBase = $"DoPriCor = '{idSubOr}' AND IDTEXT = '{idText}' AND IDSITUACION = 'GenCierre'";

                // Obtener cantidad total de producto terminado: CANTINGRE donde CODIMOVI = 'OF'
                var cantidadTotalProductoTerminado = await _context.Orfapens
                    .Where(o => o.DoPriCor == idSubOr && 
                               o.IDTEXT == idText && 
                               o.IDSITUACION == "GenCierre" && 
                               o.CodiMovi == "OF")
                    .SumAsync(o => o.CantIngre ?? 0);

                _logger.LogInformation($"Cantidad total producto terminado encontrada: {cantidadTotalProductoTerminado}");

                if (cantidadTotalProductoTerminado <= 0)
                {
                    _logger.LogWarning($"No se encontró cantidad de producto terminado para {idSubOr} con IDTEXT {idText}");
                    return;
                }

                // Obtener registros de componentes (CF) desde ORFAPEN
                var componentesOrfapen = await _context.Orfapens
                    .Where(o => o.DoPriCor == idSubOr && 
                               o.IDTEXT == idText && 
                               o.IDSITUACION == "GenCierre" && 
                               o.CodiMovi == "CF")
                    .ToListAsync();

                _logger.LogInformation($"Encontrados {componentesOrfapen.Count} componentes para procesar");

                int numeroOrdenItem = 0;
                int componentesProcesados = 0;

                foreach (var componente in componentesOrfapen)
                {
                    if (componente.Cod_Inte == null || componente.Cod_Inte <= 0) 
                    {
                        _logger.LogWarning($"Componente sin código interno válido, saltando...");
                        continue;
                    }

                    int codigoInterno = componente.Cod_Inte.Value;
                    _logger.LogInformation($"Procesando componente {componentesProcesados + 1}/{componentesOrfapen.Count} - Código: {codigoInterno}");

                    // Extraer datos del componente
                    int depositoConsumo = componente.DepoConsu ?? 0;
                    string referencia = componente.ReferCor ?? "";
                    DateTime fechaMovimiento = componente.FechMov ?? DateTime.Now;
                    string lote = componente.IdenLote?.Trim() ?? "";
                    
                    // Procesar códigos con longitud de 11 caracteres
                    string dopriProcesado = ProcesarCodigoLongitud(componente.DoPriCor ?? "");
                    string dosecProcesado = ProcesarCodigoLongitud(componente.DoSecCor ?? "");
                    
                    decimal cantidadOriginalAConsumir = componente.CantSalid ?? 0;

                    // Calcular consumo proporcional: (CantOrigAConsumir / CantTotProdTerminado) * CantAprobada
                    decimal consumoProporcional = (cantidadOriginalAConsumir / cantidadTotalProductoTerminado) * cantidadAprobada;

                    _logger.LogInformation($"Componente {codigoInterno}: Cantidad original: {cantidadOriginalAConsumir}, Consumo proporcional calculado: {consumoProporcional}");

                    // Verificar si la unidad de medida es "GK" para dividir por 1000
                    var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == codigoInterno);
                    if (master != null && master.Umedida?.Trim().ToUpper() == "GK")
                    {
                        consumoProporcional = consumoProporcional / 1000;
                        _logger.LogInformation($"Componente {codigoInterno}: Unidad GK detectada, consumo ajustado: {consumoProporcional}");
                    }

                    // Solo procesar si el consumo es significativo (>= 0.05)
                    if (Math.Abs(consumoProporcional) >= 0.05m)
                    {
                        _logger.LogInformation($"Procesando movimiento para componente {codigoInterno} con consumo: {consumoProporcional}");

                        // Obtener lote de la reserva
                        var reserva = await _context.Reservas
                            .Where(r => r.IdSubOr == idSubOr && r.Cod_Inte == codigoInterno)
                            .FirstOrDefaultAsync();
                        
                        string loteReserva = reserva?.IdenLote ?? lote;

                        // Preparar datos para movimiento
                        string codigoMovimiento = "CF";
                        string dopriMovimiento = idSubOr;
                        string dosecMovimiento = "PF-" + idSubOr.Substring(3).Trim(); // PF- para movimientos CF
                        string referenciaMovimiento = $"CONSUMO {codigoPadre}"; // CONSUMO + código padre
                        
                        // Determinar depósito: si DP2% = 99 usar depósito predeterminado del producto
                        int depositoFinal = depositoConsumo;
                        if (depositoConsumo == 99)
                        {
                            depositoFinal = master?.DepoCon ?? depositoConsumo;
                        }

                        numeroOrdenItem++;

                        try
                        {
                            // Generar movimiento de stock para el componente (cantidad negativa = consumo)
                            _logger.LogInformation($"Iniciando movimiento de stock para componente {codigoInterno}, depósito: {depositoFinal}, cantidad: {-consumoProporcional}");
                            
                            await ProcesarMovimientoStockComponente(
                                fechaMovimiento, 
                                codigoInterno, 
                                loteReserva, 
                                codigoMovimiento,
                                dopriMovimiento, 
                                dosecMovimiento, 
                                referenciaMovimiento, 
                                numeroOrdenItem,
                                depositoFinal, 
                                -consumoProporcional);

                            _logger.LogInformation($"Movimiento de stock completado para componente {codigoInterno}");

                            // Pequeña pausa para evitar conflictos de transacción
                            await Task.Delay(50);

                            // Agregar registro ORFAPEN para el consumo del componente
                            _logger.LogInformation($"Agregando registro ORFAPEN para componente {codigoInterno}");
                            
                            await AgregarRegistroOrfapenConsumo(
                                fechaMovimiento, 
                                codigoInterno, 
                                loteReserva, 
                                codigoMovimiento,
                                dopriMovimiento, 
                                dosecMovimiento, 
                                referenciaMovimiento, 
                                depositoFinal,
                                consumoProporcional, 
                                idText);

                            // Actualizar reserva del componente
                            _logger.LogInformation($"Actualizando reserva para componente {codigoInterno}");
                            await ActualizarReservaComponente(idSubOr, codigoInterno, consumoProporcional);

                            componentesProcesados++;
                            _logger.LogInformation($"Componente {codigoInterno} procesado exitosamente ({componentesProcesados}/{componentesOrfapen.Count})");
                        }
                        catch (Exception exComponente)
                        {
                            _logger.LogError(exComponente, $"Error al procesar componente {codigoInterno}: {exComponente.Message}");
                            throw;
                        }
                    }
                    else
                    {
                        _logger.LogInformation($"Componente {codigoInterno} saltado por consumo insignificante: {consumoProporcional}");
                    }
                }

                _logger.LogInformation($"Procesamiento completado: {componentesProcesados} de {componentesOrfapen.Count} componentes procesados para orden {idSubOr}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al procesar consumo de componentes para orden {idSubOr}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Procesa movimiento de stock para componente consumido
        /// Equivalente a: Call MOVISTOK(...) en VB6
        /// </summary>
        private async Task ProcesarMovimientoStockComponente(DateTime fecha, int codigoInterno, string lote, 
            string codigoMovimiento, string dopri, string dosec, string referencia, int numeroOrdenItem, 
            int deposito, decimal cantidad)
        {
            try
            {
                var movimientoStock = new MoviStockDTO
                {
                    Fecha = fecha,
                    CodigoInterno = codigoInterno,
                    Lote = lote,
                    CodigoMovimiento = codigoMovimiento,
                    DocumentoPrimarioLargo = ProcesarCodigoLongitud(dopri),
                    DocumentoPrimarioCorto = ProcesarCodigoLongitud(dopri),
                    DocumentoSecundarioLargo = ProcesarCodigoLongitud(dosec),
                    DocumentoSecundarioCorto = ProcesarCodigoLongitud(dosec),
                    Referencia = referencia,
                    NumeroOrdenItem = numeroOrdenItem,
                    ValorUnitario = 0,
                    Moneda = 1,
                    NumeroCliente = 0,
                    Deposito = deposito,
                    Cantidad = cantidad,
                    TipoTransferencia = null
                };

                var resultado = await _moviStockService.RegistrarMovimientoAsync(movimientoStock, User);
                
                if (resultado is BadRequestObjectResult badRequest)
                {
                    var errorMessage = badRequest.Value?.ToString() ?? "Error desconocido";
                    _logger.LogError($"MoviStockService retornó BadRequest para componente {codigoInterno}: {errorMessage}");
                    _logger.LogError($"Datos enviados al servicio - CodigoInterno: {codigoInterno}, Deposito: {deposito}, Cantidad: {cantidad}, CodigoMovimiento: '{codigoMovimiento}', Lote: '{lote}'");
                    throw new Exception($"Error al procesar movimiento de stock para componente {codigoInterno}: {errorMessage}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Excepción al procesar movimiento de stock para componente {codigoInterno}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Agrega registro ORFAPEN para consumo de componente
        /// Equivalente a: Call AGREGA_REGISTRO_NORFAPEN(..., "AproCierre") en VB6
        /// </summary>
        private async Task AgregarRegistroOrfapenConsumo(DateTime fecha, int codigoInterno, string lote, 
            string codigoMovimiento, string dopri, string dosec, string referencia, int deposito, 
            decimal cantidad, string idText)
        {
            var idFlexoft = _userContextService.GetCurrentUserId();
            if (!idFlexoft.HasValue)
            {
                throw new Exception("No se pudo obtener el contexto del usuario");
            }

            // Obtener código externo del producto
            var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == codigoInterno);
            string codigoExterno = master?.Codigo ?? "";

            // Crear el registro ORFAPEN para consumo de componente
            // Para AproCierre + CF: cantidad positiva en CantIngre (ingreso para compensar la salida del cierre)
            var registroConsumo = new Orfapen
            {
                FechMov = fecha,
                Cod_Inte = codigoInterno,
                Cod_Exte = codigoExterno,
                IdenLote = lote,
                CodiMovi = codigoMovimiento,
                DoPriCor = ProcesarCodigoLongitud(dopri),
                DoPriLar = ProcesarCodigoLongitud(dopri),
                DoSecCor = ProcesarCodigoLongitud(dosec),
                DoSecLar = ProcesarCodigoLongitud(dosec),
                ReferCor = referencia,
                ValoUnit = 0,
                MoneVal = 1,
                Nume_Clie = 0,
                DepoMovi = (short?)0, // Para consumo no hay depósito de movimiento
                CantIngre = Math.Abs(cantidad), // Para aprobación CF: cantidad positiva en CantIngre
                CantSalid = 0, // Para consumo no hay salida
                FeReMovi = DateTime.Now,
                NumUsuar = idFlexoft.Value % 100,
                STATUS = 0,
                IDTEXT = idText,
                DepoConsu = (short?)deposito, // Depósito desde donde se consume
                IDSITUACION = "AproCierre"
            };

            _context.Orfapens.Add(registroConsumo);
        }

        /// <summary>
        /// Actualiza reserva de componente específico
        /// Equivalente a: Call ACTUALIZARESERVA(NORFA$, CI2%, CONSUMOX) en VB6
        /// </summary>
        private async Task ActualizarReservaComponente(string idSubOr, int codigoComponente, decimal cantidadConsumida)
        {
            var reservas = await _context.Reservas
                .Where(r => r.IdSubOr == idSubOr && r.Cod_Inte == codigoComponente)
                .OrderBy(r => r.Reserva_Id)
                .ToListAsync();

            decimal cantidadRestante = cantidadConsumida;

            foreach (var reserva in reservas)
            {
                if (cantidadRestante <= 0) break;

                decimal cantidadReserva = reserva.CantRes ?? 0;
                if (cantidadReserva <= 0) continue;

                if (cantidadRestante >= cantidadReserva)
                {
                    // Consumir toda la reserva
                    cantidadRestante -= cantidadReserva;
                    _context.Reservas.Remove(reserva);
                }
                else
                {
                    // Consumir parcialmente la reserva
                    reserva.CantRes = cantidadReserva - cantidadRestante;
                    cantidadRestante = 0;
                    _context.Reservas.Update(reserva);
                }
            }

            _logger.LogInformation($"Actualizada reserva para componente {codigoComponente} en orden {idSubOr}, cantidad consumida: {cantidadConsumida}");
        }

        /// <summary>
        /// Actualiza las reservas de un componente específico (método original - mantenido para compatibilidad)
        /// </summary>
        private async Task ActualizarReservasComponente(string idSubOr, int codigoComponente, decimal cantidadConsumo)
        {
            var reservas = await _context.Reservas
                .Where(r => r.IdSubOr == idSubOr && r.Cod_Inte == codigoComponente)
                .OrderBy(r => r.Reserva_Id)
                .ToListAsync();

            decimal cantidadRestante = cantidadConsumo;

            foreach (var reserva in reservas)
            {
                if (cantidadRestante <= 0) break;

                decimal cantidadReserva = reserva.CantRes ?? 0;
                if (cantidadReserva <= 0) continue;

                if (cantidadRestante >= cantidadReserva)
                {
                    // Consumir toda la reserva
                    cantidadRestante -= cantidadReserva;
                    _context.Reservas.Remove(reserva);
                }
                else
                {
                    // Consumir parcialmente la reserva
                    reserva.CantRes = cantidadReserva - cantidadRestante;
                    cantidadRestante = 0;
                    _context.Reservas.Update(reserva);
                }
            }
        }

        /// <summary>
        /// Actualiza la orden de fabricación con las cantidades aprobadas
        /// Equivalente a la lógica VB6 de actualización de ORDEFABR
        /// </summary>
        private Task<bool> ActualizarOrdenFabricacion(OrdeFabr orden, decimal cantidadAprobada, string? idText)
        {
            decimal cantidadFinProceso = orden.Canti_Fin_Proceso ?? 0;
            decimal aprobadoTotal = (orden.CanApro ?? 0) + cantidadAprobada;
            bool ordenCerrada = false;

            // Actualizar campos según lógica VB6
            if (aprobadoTotal > cantidadFinProceso)
            {
                orden.Canti_Fin_Proceso = aprobadoTotal;
                orden.Status_Proceso = 2;
            }

            orden.CanApro = aprobadoTotal;
            // orden.CanRech se mantiene igual en aprobación

            // Verificar si se debe cerrar la orden
            if ((orden.CanApro - orden.CanProy) > 0.05m)
            {
                orden.StatuOrf = 3; // Marcar como cerrada
                ordenCerrada = true;
            }

            orden.FechCierr = DateTime.Now;

            // Agregar anotación en AnotaOrf
            string nombreUsuario = ObtenerNombreUsuario();
            string anotacionAprobacion = FormatearNuevaAnotacion($"Aprobacion de cierre id: {idText}   -- Cantidad: {cantidadAprobada:F1}", nombreUsuario);

            if (string.IsNullOrEmpty(orden.AnotaOrf))
            {
                orden.AnotaOrf = anotacionAprobacion;
            }
            else
            {
                orden.AnotaOrf += "\r\n" + anotacionAprobacion;
            }

            _context.Ordefabrs.Update(orden);
            return Task.FromResult(ordenCerrada);
        }

        /// <summary>
        /// Obtiene el nombre del usuario actual desde el ClaimsPrincipal
        /// </summary>
        /// <returns>Nombre del usuario o "Usuario" si no se puede obtener</returns>
        private string ObtenerNombreUsuario()
        {
            try
            {
                var user = _userContextService.GetCurrentUser();
                if (user == null) return "Usuario";

                var nombreClaim = user.FindFirst(ClaimTypes.Name)?.Value;
                if (!string.IsNullOrEmpty(nombreClaim))
                    return nombreClaim;

                var usernameClaim = user.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                if (!string.IsNullOrEmpty(usernameClaim))
                    return usernameClaim;

                return "Usuario";
            }
            catch
            {
                return "Usuario";
            }
        }

        /// <summary>
        /// Formatea una nueva anotación para agregarla al campo AnotaOrf
        /// </summary>
        /// <param name="textoAnotacion">Texto de la anotación</param>
        /// <param name="nombreUsuario">Nombre del usuario que registra la anotación</param>
        /// <returns>Texto formateado para el campo AnotaOrf</returns>
        private string FormatearNuevaAnotacion(string textoAnotacion, string nombreUsuario)
        {
            string ttxyz = new string(' ', 76);
            string hos = DateTime.Now.ToString("dd/MM/yy");
            ttxyz = Repla(ttxyz, hos, 1, 8);
            ttxyz = Repla(ttxyz, textoAnotacion, 11, 56);
            ttxyz = Repla(ttxyz, nombreUsuario, 69, 24);
            return ttxyz;
        }

        /// <summary>
        /// Reemplaza una porción de texto en una cadena
        /// Equivalente a la función REPLA de VB6
        /// </summary>
        private string Repla(string target, string replacement, int start, int length)
        {
            if (string.IsNullOrEmpty(target))
                target = new string(' ', Math.Max(start + length, 76));

            if (target.Length < start + length)
                target = target.PadRight(start + length);

            char[] chars = target.ToCharArray();
            for (int i = 0; i < Math.Min(replacement.Length, length); i++)
            {
                chars[start - 1 + i] = replacement[i];
            }
            return new string(chars);
        }

        /// <summary>
        /// Elimina las reservas cuando la orden se cierra completamente
        /// Equivalente a: DELETE FROM Reserva WHERE IdSubOr = NORFA$
        /// </summary>
        private async Task EliminarReservas(string idSubOr)
        {
            var reservas = await _context.Reservas
                .Where(r => r.IdSubOr == idSubOr)
                .ToListAsync();

            if (reservas.Any())
            {
                _context.Reservas.RemoveRange(reservas);
                _logger.LogInformation($"Eliminadas {reservas.Count} reservas para orden {idSubOr}");
            }
        }

        /// <summary>
        /// Procesa un código para asegurar longitud de 11 caracteres (como en VB6)
        /// While Len(DOPRICO$) < 11: DOPRICO$ = Mid$(DOPRICO$, 1, 3) & "0" & Mid$(DOPRICO$, 4): Wend
        /// </summary>
        private string ProcesarCodigoLongitud(string codigo)
        {
            if (string.IsNullOrEmpty(codigo)) return "";

            string resultado = codigo;
            while (resultado.Length < 11)
            {
                if (resultado.Length >= 3)
                {
                    resultado = resultado.Substring(0, 3) + "0" + resultado.Substring(3);
                }
                else
                {
                    resultado = resultado + "0";
                }
            }

            return resultado.Length > 11 ? resultado.Substring(0, 11) : resultado;
        }

        /// <summary>
        /// Genera un ID de texto aleatorio
        /// </summary>
        private string GenerarIdTextoAleatorio()
        {
            var random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, 10)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }


    }
}
