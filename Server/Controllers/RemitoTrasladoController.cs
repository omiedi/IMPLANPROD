using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para manejar remitos de traslado a sucursal
    /// Genera movimientos de stock y archivos de intercambio para sincronización
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
    public class RemitoTrasladoController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IMoviStockService _moviStockService;
        private readonly ILogger<RemitoTrasladoController> _logger;
        private readonly FlexcrlService _flexcrlService;
        private readonly RecepcionMaterialesService _recepcionMaterialesService;

        public RemitoTrasladoController(
            DataContext context,
            IMoviStockService moviStockService,
            ILogger<RemitoTrasladoController> logger,
            FlexcrlService flexcrlService,
            RecepcionMaterialesService recepcionMaterialesService)
        {
            _context = context;
            _moviStockService = moviStockService;
            _logger = logger;
            _flexcrlService = flexcrlService;
            _recepcionMaterialesService = recepcionMaterialesService;
        }

        /// <summary>
        /// Crea un remito de traslado a sucursal
        /// Genera movimientos de stock (RC/RS) y archivo de intercambio .txt
        /// </summary>
        /// <param name="request">Datos del remito de traslado</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPost]
        public async Task<ActionResult> CrearRemitoTraslado([FromBody] RemitoTrasladoDTO request)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                _logger.LogInformation("Iniciando creación de remito de traslado - Prefijo: {Prefijo}, Número: {Numero}", 
                    request.Prefijo, request.Numero);

                // Validaciones
                if (request.DepositoOrigen == request.DepositoDestino)
                {
                    return BadRequest("El depósito de origen y destino no pueden ser el mismo");
                }

                if (!request.Items.Any())
                {
                    return BadRequest("Debe haber al menos un item en el remito");
                }

                // Validar que no exista un remito duplicado
                bool existeDuplicado = await _moviStockService.ExisteRemitoDuplicadoAsync(request.Prefijo, request.Numero);
                if (existeDuplicado)
                {
                    // Si está en modo autonumeración, recalcular el número automáticamente
                    // Usar CLABUS vacío según VB6: GRACONTROL("", "NURETRA", MONUME$)
                    var tipoNumeracion = await _flexcrlService.ObtenerControlAsync("", "NURETRA");
                    if (tipoNumeracion != null && tipoNumeracion.VALCONTROL == "AUTONUME")
                    {
                        _logger.LogWarning("Número duplicado detectado en modo AUTONUME, recalculando automáticamente - Prefijo: {Prefijo}, Número: {Numero}",
                            request.Prefijo, request.Numero);

                        // Obtener fecha de nueva numeración
                        var fechaNumeracion = await _flexcrlService.ObtenerControlAsync("RTRASLA", "FENURET");
                        DateTime? fechaFiltro = null;
                        if (fechaNumeracion != null && !string.IsNullOrWhiteSpace(fechaNumeracion.VALCONTROL))
                        {
                            // Intentar formato con hora (dd/MM/yy HH:mm)
                            if (DateTime.TryParseExact(fechaNumeracion.VALCONTROL, "dd/MM/yy HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime fecha))
                            {
                                fechaFiltro = fecha;
                            }
                            // Por compatibilidad, intentar solo fecha
                            else if (DateTime.TryParseExact(fechaNumeracion.VALCONTROL, "dd/MM/yy", null, System.Globalization.DateTimeStyles.None, out fecha))
                            {
                                fechaFiltro = fecha;
                            }
                        }

                        // Obtener el último número y calcular el siguiente (usar "RC" para remitos de traslado)
                        int ultimoNumero = await _moviStockService.ObtenerUltimoNumeroRemitoAsync(fechaFiltro, "RC");
                        int nuevoNumero = ultimoNumero + 1;

                        _logger.LogInformation("Nuevo número calculado: {NuevoNumero} (último: {UltimoNumero})", nuevoNumero, ultimoNumero);

                        // Actualizar el request con el nuevo número
                        request.Numero = nuevoNumero;

                        // Volver a validar con el nuevo número (máximo 5 intentos para evitar loop infinito)
                        int intentos = 0;
                        while (intentos < 5)
                        {
                            existeDuplicado = await _moviStockService.ExisteRemitoDuplicadoAsync(request.Prefijo, request.Numero);
                            if (!existeDuplicado)
                            {
                                break;
                            }

                            intentos++;
                            request.Numero++;
                            _logger.LogWarning("Número {Numero} también duplicado, intentando con {NuevoNumero} (intento {Intentos})",
                                request.Numero - 1, request.Numero, intentos);
                        }

                        if (existeDuplicado)
                        {
                            return BadRequest($"No se pudo encontrar un número disponible después de 5 intentos. Último número probado: {request.Numero}");
                        }

                        _logger.LogInformation("Número disponible encontrado: {Numero}", request.Numero);
                    }
                    else
                    {
                        return BadRequest($"Ya existe un remito con prefijo {request.Prefijo} y número {request.Numero}");
                    }
                }

                // Obtener la razón social del cliente
                string razonSocialCliente = await ObtenerRazonSocialCliente(request.NumeroCliente);

                // Generar documento del remito en formato RC.1234567
                string documentoRemito = $"RC.{request.Numero}";

                // Generar movimientos de stock para cada item (RC para salida, RS para ingreso)
                short contadorItem = 0;
                foreach (var item in request.Items)
                {
                    // Si el producto es trazable, procesar por lote
                    if (item.Trazable == 1 && item.Lotes.Any())
                    {
                        foreach (var lote in item.Lotes)
                        {
                            contadorItem++;
                            // Movimiento 1 - SALIDA (RC) desde depósito de origen
                            var movimientoRC = new MoviStockDTO
                            {
                                Fecha = request.Fecha,
                                CodigoInterno = item.CodInt,
                                Lote = lote.Lote,
                                CodigoMovimiento = "RC",
                                DocumentoPrimarioLargo = documentoRemito, // doprilar: RT.0001
                                DocumentoPrimarioCorto = documentoRemito, // dopricor: RT.0001
                                DocumentoSecundarioLargo = "", // doseclar: vacío
                                DocumentoSecundarioCorto = "", // doseccor: vacío
                                Referencia = razonSocialCliente,
                                NumeroOrdenItem = contadorItem, // nuoritem: enumerador del orden de los items
                                NumeroPedidoCliente = item.NroPedido, // nupedcli: nro de pedido
                                NumeroOrdenPedido = (short?)item.NuOrItemPedido, // numeordped: orden en peddeta
                                NumeroCliente = request.NumeroCliente, // NroClie para actualizar PEDDETA
                                ValorUnitario = 0,
                                Deposito = request.DepositoOrigen,
                                Cantidad = -lote.Cantidad, // Negativo para salida
                                NumeroRemito = request.Numero,
                                PrefijoRemito = request.Prefijo
                            };

                            var resultadoRC = await _moviStockService.RegistrarMovimientoAsync(movimientoRC, User);
                            if (!(resultadoRC is OkObjectResult))
                            {
                                throw new Exception($"Error al registrar movimiento RC para lote {lote.Lote}");
                            }

                            // Movimiento 2 - INGRESO (RS) a depósito de destino
                            var movimientoRS = new MoviStockDTO
                            {
                                Fecha = request.Fecha,
                                CodigoInterno = item.CodInt,
                                Lote = lote.Lote,
                                CodigoMovimiento = "RS",
                                DocumentoPrimarioLargo = documentoRemito, // doprilar: RT.0001
                                DocumentoPrimarioCorto = documentoRemito, // dopricor: RT.0001
                                DocumentoSecundarioLargo = "", // doseclar: vacío
                                DocumentoSecundarioCorto = "", // doseccor: vacío
                                Referencia = "De Deposito Central",
                                NumeroOrdenItem = contadorItem, // nuoritem: enumerador del orden de los items
                                NumeroPedidoCliente = item.NroPedido, // nupedcli: nro de pedido
                                NumeroOrdenPedido = (short?)item.NuOrItemPedido, // numeordped: orden en peddeta
                                NumeroCliente = request.NumeroCliente, // NroClie para actualizar PEDDETA
                                ValorUnitario = 0,
                                Deposito = request.DepositoDestino,
                                Cantidad = lote.Cantidad, // Positivo para ingreso
                                NumeroRemito = request.Numero,
                                PrefijoRemito = request.Prefijo
                            };

                            var resultadoRS = await _moviStockService.RegistrarMovimientoAsync(movimientoRS, User);
                            if (!(resultadoRS is OkObjectResult))
                            {
                                throw new Exception($"Error al registrar movimiento RS para lote {lote.Lote}");
                            }

                            _logger.LogInformation("Movimientos RC/RS registrados - Producto: {Codigo}, Lote: {Lote}, Cantidad: {Cantidad}", 
                                item.Codigo, lote.Lote, lote.Cantidad);
                        }
                    }
                    else
                    {
                        contadorItem++;
                        // Producto no trazable, movimiento único sin lote
                        // Movimiento 1 - SALIDA (RC) desde depósito de origen
                        var movimientoRC = new MoviStockDTO
                        {
                            Fecha = request.Fecha,
                            CodigoInterno = item.CodInt,
                            Lote = "",
                            CodigoMovimiento = "RC",
                            DocumentoPrimarioLargo = documentoRemito, // doprilar: RT.0001
                            DocumentoPrimarioCorto = documentoRemito, // dopricor: RT.0001
                            DocumentoSecundarioLargo = "", // doseclar: vacío
                            DocumentoSecundarioCorto = "", // doseccor: vacío
                            Referencia = razonSocialCliente,
                            NumeroOrdenItem = contadorItem, // nuoritem: enumerador del orden de los items
                            NumeroPedidoCliente = item.NroPedido, // nupedcli: nro de pedido
                            NumeroOrdenPedido = (short?)item.NuOrItemPedido, // numeordped: orden en peddeta
                            NumeroCliente = request.NumeroCliente, // NroClie para actualizar PEDDETA
                            ValorUnitario = 0,
                            Deposito = request.DepositoOrigen,
                            Cantidad = -item.Cantidad, // Negativo para salida
                            NumeroRemito = request.Numero,
                            PrefijoRemito = request.Prefijo
                        };

                        var resultadoRC = await _moviStockService.RegistrarMovimientoAsync(movimientoRC, User);
                        if (!(resultadoRC is OkObjectResult))
                        {
                            throw new Exception($"Error al registrar movimiento RC para producto {item.Codigo}");
                        }

                        // Movimiento 2 - INGRESO (RS) a depósito de destino
                        var movimientoRS = new MoviStockDTO
                        {
                            Fecha = request.Fecha,
                            CodigoInterno = item.CodInt,
                            Lote = "",
                            CodigoMovimiento = "RS",
                            DocumentoPrimarioLargo = documentoRemito, // doprilar: RT.0001
                            DocumentoPrimarioCorto = documentoRemito, // dopricor: RT.0001
                            DocumentoSecundarioLargo = "", // doseclar: vacío
                            DocumentoSecundarioCorto = "", // doseccor: vacío
                            Referencia = "De Deposito Central",
                            NumeroOrdenItem = contadorItem, // nuoritem: enumerador del orden de los items
                            NumeroPedidoCliente = item.NroPedido, // nupedcli: nro de pedido
                            NumeroOrdenPedido = (short?)item.NuOrItemPedido, // numeordped: orden en peddeta
                            NumeroCliente = request.NumeroCliente, // NroClie para actualizar PEDDETA
                            ValorUnitario = 0,
                            Deposito = request.DepositoDestino,
                            Cantidad = item.Cantidad, // Positivo para ingreso
                            NumeroRemito = request.Numero,
                            PrefijoRemito = request.Prefijo
                        };

                        var resultadoRS = await _moviStockService.RegistrarMovimientoAsync(movimientoRS, User);
                        if (!(resultadoRS is OkObjectResult))
                        {
                            throw new Exception($"Error al registrar movimiento RS para producto {item.Codigo}");
                        }

                        _logger.LogInformation("Movimientos RC/RS registrados - Producto: {Codigo}, Cantidad: {Cantidad}", 
                            item.Codigo, item.Cantidad);
                    }
                }

                // Generar archivo .txt en ruta temporal
                await GenerarArchivoTrasladoAsync(request, documentoRemito, razonSocialCliente);

                await transaction.CommitAsync();

                _logger.LogInformation("Remito de traslado creado exitosamente - Prefijo: {Prefijo}, Número: {Numero}", 
                    request.Prefijo, request.Numero);

                return Ok(new { 
                    mensaje = "Remito de traslado creado exitosamente",
                    prefijo = request.Prefijo,
                    numero = request.Numero
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, "Error al crear remito de traslado");
                return BadRequest($"Error al crear remito de traslado: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene la razón social del cliente desde DEUDOR12
        /// </summary>
        private async Task<string> ObtenerRazonSocialCliente(int numeroCliente)
        {
            try
            {
                var cliente = await _context.Set<Deudor12>()
                    .Where(c => c.Nume_Cli == numeroCliente)
                    .FirstOrDefaultAsync();

                if (cliente == null)
                {
                    _logger.LogWarning("Cliente no encontrado: {NumeroCliente}", numeroCliente);
                    return "";
                }

                return cliente.Raso_Cli ?? "";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener razón social del cliente {NumeroCliente}", numeroCliente);
                return "";
            }
        }

        /// <summary>
        /// Genera el archivo de intercambio .txt para el remito de traslado
        /// Formato: ReTraslado_RT_X_0001_000187626_8_2026_113841.txt
        /// </summary>
        private async Task GenerarArchivoTrasladoAsync(RemitoTrasladoDTO request, string documentoRemito, string razonSocialCliente)
        {
            try
            {
                // Obtener la ruta temporal desde FLEXCRL
                const string CLABUS_PARAM = "RTRASLA";
                const string PROBUS_PARAM = "RUTATEMP";

                var control = await _flexcrlService.ObtenerControlAsync(CLABUS_PARAM, PROBUS_PARAM);
                
                if (control == null || string.IsNullOrWhiteSpace(control.VALCONTROL))
                {
                    throw new Exception("No está configurada la ruta temporal para remitos de traslado");
                }

                string rutaTemporal = control.VALCONTROL;

                // Verificar que el directorio existe
                if (!Directory.Exists(rutaTemporal))
                {
                    Directory.CreateDirectory(rutaTemporal);
                }

                // Generar nombre del archivo según formato requerido
                // ReTraslado_RT_X_0001_000187626_8_2026_113841.txt
                string documentoSinPunto = documentoRemito.Replace("RT.", "");
                string nombreArchivo = $"ReTraslado_{documentoSinPunto}_{request.Fecha:dd}_{request.Fecha:MM}_{request.Fecha:yyyy}_{DateTime.Now:HHmmss}.txt";
                string rutaCompleta = Path.Combine(rutaTemporal, nombreArchivo);

                // Generar contenido del archivo
                var lineas = new List<string>();

                // Calcular anchos máximos de cada campo basado en los datos
                var anchos = CalcularAnchosCampos(request, documentoRemito);

                // Agregar encabezado con nombres de campos y padding
                string encabezado = GenerarEncabezadoArchivo(anchos);
                lineas.Add(encabezado);

                // Por cada artículo, agregar línea con los datos requeridos
                foreach (var item in request.Items)
                {
                    if (item.Trazable == 1 && item.Lotes.Any())
                    {
                        // Producto trazable con lotes - una línea por lote
                        foreach (var lote in item.Lotes)
                        {
                            string linea = GenerarLineaArchivo(
                                request.Fecha,
                                documentoRemito,
                                request.NumeroCliente,
                                lote.Lote,
                                item.CodInt,
                                item.Codigo,
                                item.Descripcion,
                                lote.Cantidad,
                                item.NroPedido ?? 0,
                                item.NuOrItemPedido ?? 0,
                                anchos
                            );
                            lineas.Add(linea);
                        }
                    }
                    else
                    {
                        // Producto no trazable - una línea sin lote
                        string linea = GenerarLineaArchivo(
                            request.Fecha,
                            documentoRemito,
                            request.NumeroCliente,
                            "",
                            item.CodInt,
                            item.Codigo,
                            item.Descripcion,
                            item.Cantidad,
                            item.NroPedido ?? 0,
                            item.NuOrItemPedido ?? 0,
                            anchos
                        );
                        lineas.Add(linea);
                    }
                }

                // Escribir archivo
                await System.IO.File.WriteAllLinesAsync(rutaCompleta, lineas);

                _logger.LogInformation("Archivo de traslado generado: {Ruta}", rutaCompleta);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar archivo de traslado");
                throw new Exception($"Error al generar archivo de traslado: {ex.Message}");
            }
        }

        /// <summary>
        /// Calcula los anchos máximos de cada campo basado en los datos del request
        /// </summary>
        private Dictionary<string, int> CalcularAnchosCampos(RemitoTrasladoDTO request, string documentoRemito)
        {
            var anchos = new Dictionary<string, int>
            {
                { "FechaGeneracion", 19 }, // dd/MM/yyyy HH:mm:ss
                { "Remito", documentoRemito.Length },
                { "NumeroCliente", request.NumeroCliente.ToString().Length },
                { "Lote", 0 },
                { "CodInt", 0 },
                { "CodExt", 0 },
                { "Descripcion", 0 },
                { "Cantidad", 0 },
                { "NroPedido", 0 },
                { "MoneVal", 1 },
                { "Nume_Clie", request.NumeroCliente.ToString().Length }
            };

            // Recorrer todos los items para calcular anchos máximos
            foreach (var item in request.Items)
            {
                if (item.Trazable == 1 && item.Lotes.Any())
                {
                    foreach (var lote in item.Lotes)
                    {
                        anchos["Lote"] = Math.Max(anchos["Lote"], lote.Lote?.Length ?? 0);
                        anchos["CodInt"] = Math.Max(anchos["CodInt"], item.CodInt.ToString().Length);
                        anchos["CodExt"] = Math.Max(anchos["CodExt"], item.Codigo?.Length ?? 0);
                        anchos["Descripcion"] = Math.Max(anchos["Descripcion"], item.Descripcion?.Length ?? 0);
                        anchos["Cantidad"] = Math.Max(anchos["Cantidad"], lote.Cantidad.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture).Length);
                        anchos["NroPedido"] = Math.Max(anchos["NroPedido"], (item.NroPedido ?? 0).ToString().Length);
                    }
                }
                else
                {
                    anchos["CodInt"] = Math.Max(anchos["CodInt"], item.CodInt.ToString().Length);
                    anchos["CodExt"] = Math.Max(anchos["CodExt"], item.Codigo?.Length ?? 0);
                    anchos["Descripcion"] = Math.Max(anchos["Descripcion"], item.Descripcion?.Length ?? 0);
                    anchos["Cantidad"] = Math.Max(anchos["Cantidad"], item.Cantidad.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture).Length);
                    anchos["NroPedido"] = Math.Max(anchos["NroPedido"], (item.NroPedido ?? 0).ToString().Length);
                }
            }

            // Asegurar anchos mínimos para encabezados
            anchos["FechaGeneracion"] = Math.Max(anchos["FechaGeneracion"], "Fecha Generación".Length);
            anchos["Remito"] = Math.Max(anchos["Remito"], "Remito".Length);
            anchos["NumeroCliente"] = Math.Max(anchos["NumeroCliente"], "Número Cliente".Length);
            anchos["Lote"] = Math.Max(anchos["Lote"], "Lote".Length);
            anchos["CodInt"] = Math.Max(anchos["CodInt"], "COD_INTE".Length);
            anchos["CodExt"] = Math.Max(anchos["CodExt"], "COD_EXTE".Length);
            anchos["Descripcion"] = Math.Max(anchos["Descripcion"], "Descripción".Length);
            anchos["Cantidad"] = Math.Max(anchos["Cantidad"], "CANTSALID".Length);
            anchos["NroPedido"] = Math.Max(anchos["NroPedido"], "Número Pedido".Length);
            anchos["MoneVal"] = Math.Max(anchos["MoneVal"], "MoneVal".Length);
            anchos["Nume_Clie"] = Math.Max(anchos["Nume_Clie"], "Nume_Clie".Length);

            return anchos;
        }

        /// <summary>
        /// Genera el encabezado del archivo .txt con los nombres de los campos
        /// Campos: Fecha generación, Remito, Número cliente, Lote, COD_INTE, COD_EXTE, Descripción, CANTSALID, Número pedido, MoneVal, Nume_Clie
        /// </summary>
        private string GenerarEncabezadoArchivo(Dictionary<string, int> anchos)
        {
            // Usamos padding para alinear los encabezados con los datos
            return $"{"Fecha Generación".PadRight(anchos["FechaGeneracion"])};" +
                   $"{"Remito".PadRight(anchos["Remito"])};" +
                   $"{"Número Cliente".PadRight(anchos["NumeroCliente"])};" +
                   $"{"Lote".PadRight(anchos["Lote"])};" +
                   $"{"COD_INTE".PadRight(anchos["CodInt"])};" +
                   $"{"COD_EXTE".PadRight(anchos["CodExt"])};" +
                   $"{"Descripción".PadRight(anchos["Descripcion"])};" +
                   $"{"CANTSALID".PadRight(anchos["Cantidad"])};" +
                   $"{"Número Pedido".PadRight(anchos["NroPedido"])};" +
                   $"{"MoneVal".PadRight(anchos["MoneVal"])};" +
                   $"{"Nume_Clie".PadRight(anchos["Nume_Clie"])}";
        }

        /// <summary>
        /// Genera una línea del archivo .txt con el formato requerido
        /// Campos: Fecha generación, Remito, Número cliente, Lote, COD_INTE, COD_EXTE, Descripción, CANTSALID, Número pedido, MoneVal, Nume_Clie
        /// </summary>
        private string GenerarLineaArchivo(
            DateTime fecha,
            string remito,
            int numeroCliente,
            string lote,
            int codInt,
            string codExt,
            string descripcion,
            decimal cantidad,
            long nroPedido,
            int nuOrItem,
            Dictionary<string, int> anchos)
        {
            // Formato de fecha de generación: dd/MM/yyyy HH:mm:ss
            string fechaGeneracion = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            // Formatear cantidad con decimales según sistema
            string cantidadFormateada = cantidad.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);

            // Construir línea con separadores y padding según anchos calculados
            return $"{fechaGeneracion.PadRight(anchos["FechaGeneracion"])};" +
                   $"{remito.PadRight(anchos["Remito"])};" +
                   $"{numeroCliente.ToString().PadRight(anchos["NumeroCliente"])};" +
                   $"{(lote ?? "").PadRight(anchos["Lote"])};" +
                   $"{codInt.ToString().PadRight(anchos["CodInt"])};" +
                   $"{(codExt ?? "").PadRight(anchos["CodExt"])};" +
                   $"{(descripcion ?? "").PadRight(anchos["Descripcion"])};" +
                   $"{cantidadFormateada.PadRight(anchos["Cantidad"])};" +
                   $"{nroPedido.ToString().PadRight(anchos["NroPedido"])};" +
                   $"{"1".PadRight(anchos["MoneVal"])};" +
                   $"{numeroCliente.ToString().PadRight(anchos["Nume_Clie"])}";
        }

        /// <summary>
        /// Mueve los archivos de la ruta temporal a la ruta definitiva
        /// Esta función se usa cuando se confirma que los archivos están listos para ser procesados por la sucursal
        /// </summary>
        /// <returns>Resultado de la operación</returns>
        [HttpPost("mover-archivos")]
        public async Task<ActionResult> MoverArchivosADefinitivo()
        {
            try
            {
                _logger.LogInformation("Iniciando movimiento de archivos de traslado a ruta definitiva");

                // Obtener rutas desde FLEXCRL
                const string CLABUS_PARAM = "RTRASLA";
                const string PROBUS_PARAM_RUTATEMP = "RUTATEMP";
                const string PROBUS_PARAM_RUTADEFI = "RUTADEFI";

                var controlTemp = await _flexcrlService.ObtenerControlAsync(CLABUS_PARAM, PROBUS_PARAM_RUTATEMP);
                var controlDef = await _flexcrlService.ObtenerControlAsync(CLABUS_PARAM, PROBUS_PARAM_RUTADEFI);

                if (controlTemp == null || string.IsNullOrWhiteSpace(controlTemp.VALCONTROL))
                {
                    return BadRequest("No está configurada la ruta temporal para remitos de traslado");
                }

                if (controlDef == null || string.IsNullOrWhiteSpace(controlDef.VALCONTROL))
                {
                    return BadRequest("No está configurada la ruta definitiva para remitos de traslado");
                }

                string rutaTemporal = controlTemp.VALCONTROL;
                string rutaDefinitiva = controlDef.VALCONTROL;

                // Verificar que la ruta temporal existe
                if (!Directory.Exists(rutaTemporal))
                {
                    return BadRequest($"La ruta temporal no existe: {rutaTemporal}");
                }

                // Verificar que la ruta definitiva existe, si no crearla
                if (!Directory.Exists(rutaDefinitiva))
                {
                    Directory.CreateDirectory(rutaDefinitiva);
                    _logger.LogInformation("Creada ruta definitiva: {Ruta}", rutaDefinitiva);
                }

                // Obtener todos los archivos .txt de la ruta temporal
                var archivos = Directory.GetFiles(rutaTemporal, "*.txt");
                int archivosMovidos = 0;

                foreach (var archivo in archivos)
                {
                    try
                    {
                        string nombreArchivo = Path.GetFileName(archivo);
                        string destino = Path.Combine(rutaDefinitiva, nombreArchivo);

                        // Si ya existe un archivo con el mismo nombre en destino, agregar timestamp
                        if (System.IO.File.Exists(destino))
                        {
                            string nombreSinExtension = Path.GetFileNameWithoutExtension(nombreArchivo);
                            string extension = Path.GetExtension(nombreArchivo);
                            destino = Path.Combine(rutaDefinitiva, $"{nombreSinExtension}_{DateTime.Now:yyyyMMddHHmmss}{extension}");
                        }

                        // Mover archivo
                        System.IO.File.Move(archivo, destino);
                        archivosMovidos++;

                        _logger.LogInformation("Archivo movido: {Origen} -> {Destino}", archivo, destino);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error al mover archivo: {Archivo}", archivo);
                    }
                }

                _logger.LogInformation("Movimiento de archivos completado. Archivos movidos: {Cantidad}", archivosMovidos);

                return Ok(new { 
                    mensaje = "Archivos movidos exitosamente",
                    archivosMovidos = archivosMovidos
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al mover archivos a ruta definitiva");
                return BadRequest($"Error al mover archivos: {ex.Message}");
            }
        }

        /// <summary>
        /// Procesa los archivos de recepción de materiales de central
        /// Solo disponible cuando el sistema está configurado como PROCESADOR (no GENERADOR)
        /// </summary>
        /// <returns>Resultado del procesamiento</returns>
        [HttpPost("recepcion-materiales")]
        public async Task<ActionResult<RecepcionMaterialesResult>> RecepcionMaterialesCentral()
        {
            try
            {
                _logger.LogInformation("Iniciando recepción de materiales de central");

                // Obtener ruta definitiva desde FLEXCRL
                const string CLABUS_PARAM = "RTRASLA";
                const string PROBUS_PARAM_RUTADEFI = "RUTADEFI";
                const string PROBUS_PARAM_DEPODEST = "DEPODEST";

                var controlRuta = await _flexcrlService.ObtenerControlAsync(CLABUS_PARAM, PROBUS_PARAM_RUTADEFI);
                var controlDepo = await _flexcrlService.ObtenerControlAsync(CLABUS_PARAM, PROBUS_PARAM_DEPODEST);

                if (controlRuta == null || string.IsNullOrWhiteSpace(controlRuta.VALCONTROL))
                {
                    return BadRequest("No está configurada la ruta definitiva para recepción de materiales");
                }

                if (controlDepo == null || string.IsNullOrWhiteSpace(controlDepo.VALCONTROL))
                {
                    return BadRequest("No está configurado el depósito destino");
                }

                string rutaDefinitiva = controlRuta.VALCONTROL;
                int depositoDestino = int.TryParse(controlDepo.VALCONTROL, out int depo) ? depo : 0;

                if (depositoDestino == 0)
                {
                    return BadRequest("El depósito destino configurado no es válido");
                }

                // Procesar archivos usando el servicio
                var resultado = await _recepcionMaterialesService.ProcesarArchivosAsync(rutaDefinitiva, depositoDestino, User);

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar recepción de materiales");
                return BadRequest($"Error al procesar recepción de materiales: {ex.Message}");
            }
        }
    }
}
