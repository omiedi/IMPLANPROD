using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para procesar la recepción de materiales de central
    /// Replica la funcionalidad de RecibirRemitoTraslado de VB6
    /// </summary>
    public class RecepcionMaterialesService
    {
        private readonly DataContext _context;
        private readonly IMoviStockService _moviStockService;
        private readonly IUsuarioFlexoftService _usuarioFlexoftService;
        private readonly ILogger<RecepcionMaterialesService> _logger;

        public RecepcionMaterialesService(
            DataContext context,
            IMoviStockService moviStockService,
            IUsuarioFlexoftService usuarioFlexoftService,
            ILogger<RecepcionMaterialesService> logger)
        {
            _context = context;
            _moviStockService = moviStockService;
            _usuarioFlexoftService = usuarioFlexoftService;
            _logger = logger;
        }

        /// <summary>
        /// Procesa los archivos de recepción de materiales de la ruta configurada
        /// </summary>
        /// <param name="rutaDefinitiva">Ruta donde se encuentran los archivos</param>
        /// <param name="depositoDestino">Depósito destino configurado</param>
        /// <param name="user">Usuario autenticado que ejecuta el proceso (necesario para MOVISTO)</param>
        /// <returns>Resultado del procesamiento</returns>
        public async Task<RecepcionMaterialesResult> ProcesarArchivosAsync(string rutaDefinitiva, int depositoDestino, ClaimsPrincipal user)
        {
            var resultado = new RecepcionMaterialesResult();

            try
            {
                _logger.LogInformation("Iniciando procesamiento de archivos de recepción en ruta: {Ruta}", rutaDefinitiva);

                // Validar que la ruta exista
                if (!Directory.Exists(rutaDefinitiva))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = $"La ruta configurada no existe: {rutaDefinitiva}";
                    resultado.Errores.Add(resultado.Mensaje);
                    return resultado;
                }

                // Asegurar que la ruta termine con \
                if (!rutaDefinitiva.EndsWith("\\"))
                {
                    rutaDefinitiva += "\\";
                }

                // Buscar archivos con patrón ReTraslado_RC*.TXT
                var archivos = Directory.GetFiles(rutaDefinitiva, "ReTraslado_RC*.TXT");

                if (!archivos.Any())
                {
                    resultado.Exitoso = true;
                    resultado.Mensaje = "No se encontraron archivos de Remito de Traslado en la ruta configurada. No hay materiales pendientes de recepción.";
                    return resultado;
                }

                _logger.LogInformation("Se encontraron {Cantidad} archivos para procesar", archivos.Length);

                // Procesar cada archivo
                foreach (var archivo in archivos)
                {
                    try
                    {
                        var nombreArchivo = Path.GetFileName(archivo);
                        _logger.LogInformation("Procesando archivo: {Archivo}", nombreArchivo);

                        // Validar nombre del archivo
                        if (!ValidarNombreArchivo(nombreArchivo))
                        {
                            resultado.Errores.Add($"El archivo tiene un nombre incorrecto: {nombreArchivo}");
                            continue;
                        }

                        // Procesar el archivo
                        var items = await LeerArchivoAsync(archivo);
                        if (!items.Any())
                        {
                            resultado.Errores.Add($"El archivo no contiene items válidos: {nombreArchivo}");
                            continue;
                        }

                        // Procesar cada item
                        bool huboErroresEnArchivo = false;
                        foreach (var item in items)
                        {
                            try
                            {
                                await ProcesarItemAsync(item, depositoDestino, user);
                                resultado.RegistrosProcesados++;
                            }
                            catch (Exception ex)
                            {
                                huboErroresEnArchivo = true;
                                _logger.LogError(ex, "Error al procesar item del archivo {Archivo}: {Error}", nombreArchivo, ex.Message);
                                resultado.Errores.Add($"Error al procesar item en {nombreArchivo}: {ex.Message}");
                            }
                        }

                        // Solo eliminar el archivo si todos los items se procesaron sin errores
                        if (huboErroresEnArchivo)
                        {
                            _logger.LogWarning("El archivo {Archivo} no se elimina porque tuvo errores de procesamiento", nombreArchivo);
                            continue;
                        }

                        File.Delete(archivo);
                        resultado.ArchivosProcesados++;
                        _logger.LogInformation("Archivo procesado y eliminado: {Archivo}", nombreArchivo);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error al procesar archivo {Archivo}: {Error}", archivo, ex.Message);
                        resultado.Errores.Add($"Error al procesar archivo {Path.GetFileName(archivo)}: {ex.Message}");
                    }
                }

                resultado.Exitoso = resultado.ArchivosProcesados > 0;
                resultado.Mensaje = $"Proceso finalizado. Se han recepcionado los materiales de central exitosamente. Archivos procesados: {resultado.ArchivosProcesados}, Registros procesados: {resultado.RegistrosProcesados}";

                return resultado;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error general al procesar archivos de recepción");
                resultado.Exitoso = false;
                resultado.Mensaje = $"Error general: {ex.Message}";
                resultado.Errores.Add(resultado.Mensaje);
                return resultado;
            }
        }

        /// <summary>
        /// Valida el formato del nombre del archivo
        /// Replica VALIDARNOMBRERETRASLADO de VB6
        /// Formato esperado: ReTraslado_RC.NNNN_DD_MM_YYYY_HHMMSS.TXT
        /// </summary>
        private bool ValidarNombreArchivo(string nombreArchivo)
        {
            try
            {
                // Validar extensión .TXT
                if (!nombreArchivo.EndsWith(".TXT", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                // Quitar extensión
                var baseNombre = nombreArchivo.Substring(0, nombreArchivo.Length - 4);

                // Separar por _
                var partes = baseNombre.Split('_');

                if (partes.Length != 6)
                {
                    return false;
                }

                // Validar ReTraslado
                if (partes[0].ToUpper() != "RETRASLADO")
                {
                    return false;
                }

                // Validar RC.NNNN
                var remito = partes[1].ToUpper();
                if (!remito.StartsWith("RC.") || remito.Length <= 3)
                {
                    return false;
                }

                var numeroRemito = remito.Substring(3);
                if (!int.TryParse(numeroRemito, out _))
                {
                    return false;
                }

                // Validar día (2 dígitos)
                if (partes[2].Length != 2 || !int.TryParse(partes[2], out int dia) || dia < 1 || dia > 31)
                {
                    return false;
                }

                // Validar mes (2 dígitos)
                if (partes[3].Length != 2 || !int.TryParse(partes[3], out int mes) || mes < 1 || mes > 12)
                {
                    return false;
                }

                // Validar año (4 dígitos)
                if (partes[4].Length != 4 || !int.TryParse(partes[4], out _))
                {
                    return false;
                }

                // Validar hora HHMMSS (6 dígitos)
                if (partes[5].Length != 6 || !int.TryParse(partes[5], out _))
                {
                    return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Lee un archivo de recepción y extrae los items
        /// </summary>
        private async Task<List<RecepcionMaterialesItem>> LeerArchivoAsync(string rutaArchivo)
        {
            var items = new List<RecepcionMaterialesItem>();

            try
            {
                var lineas = await File.ReadAllLinesAsync(rutaArchivo);
                if (lineas.Length < 2)
                {
                    return items;
                }

                // Primera línea = encabezado
                var encabezado = lineas[0].Trim();
                var camposEncabezado = encabezado.Split(';');

                // Buscar posiciones de campos requeridos
                var posRemito = BuscarCampo(camposEncabezado, "Remito");
                var posCodInte = BuscarCampo(camposEncabezado, "COD_INTE");
                var posCantSalid = BuscarCampo(camposEncabezado, "CANTSALID");
                var posNroPed = BuscarCampo(camposEncabezado, "Número Pedido");
                var posNumeClie = BuscarCampo(camposEncabezado, "Nume_Clie");

                // Validar que existan todos los campos
                if (posRemito < 0 || posCodInte < 0 || posCantSalid < 0 || posNroPed < 0 || posNumeClie < 0)
                {
                    _logger.LogError("El encabezado no contiene todos los campos requeridos");
                    return items;
                }

                // Procesar líneas de datos (desde la segunda línea)
                for (int i = 1; i < lineas.Length; i++)
                {
                    var linea = lineas[i].Trim();
                    if (string.IsNullOrWhiteSpace(linea))
                    {
                        continue;
                    }

                    var campos = linea.Split(';');

                    // Validar que la línea tenga todas las posiciones necesarias
                    if (campos.Length <= Math.Max(posRemito, Math.Max(posCodInte, Math.Max(posCantSalid, Math.Max(posNroPed, posNumeClie)))))
                    {
                        continue;
                    }

                    // Extraer datos
                    var remito = campos[posRemito].Trim();
                    var codInteStr = campos[posCodInte].Trim();
                    var cantSalidStr = campos[posCantSalid].Trim();
                    var nroPedStr = campos[posNroPed].Trim();
                    var numeClie = campos[posNumeClie].Trim();

                    // Validar datos
                    if (string.IsNullOrWhiteSpace(codInteStr) || !int.TryParse(codInteStr, out int codInte))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(cantSalidStr) || !decimal.TryParse(cantSalidStr, out decimal cantidad))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(nroPedStr) || !long.TryParse(nroPedStr, out long nroPed))
                    {
                        continue;
                    }

                    items.Add(new RecepcionMaterialesItem
                    {
                        Remito = remito,
                        CodInte = codInte,
                        Cantidad = cantidad,
                        NroPed = nroPed,
                        NumeClie = numeClie
                    });
                }

                return items;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al leer archivo {Archivo}", rutaArchivo);
                return items;
            }
        }

        /// <summary>
        /// Busca un campo en el array de campos (case insensitive)
        /// Replica BUSCARCAMPO de VB6
        /// </summary>
        private int BuscarCampo(string[] campos, string nombreCampo)
        {
            var nombreBuscado = NormalizarCampo(nombreCampo);

            for (int i = 0; i < campos.Length; i++)
            {
                var campo = NormalizarCampo(campos[i]);
                if (string.Equals(campo, nombreBuscado, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// Normaliza un nombre de campo (quita espacios y mayúsculas)
        /// Replica NORMALIZARCAMPO de VB6
        /// </summary>
        private string NormalizarCampo(string campo)
        {
            return campo.Trim().ToUpper();
        }

        /// <summary>
        /// Procesa un item de recepción
        /// Genera movimiento de stock y actualiza compras
        /// </summary>
        private async Task ProcesarItemAsync(RecepcionMaterialesItem item, int depositoDestino, ClaimsPrincipal user)
        {
            try
            {
                _logger.LogInformation("Procesando item: Remito={Remito}, CodInte={CodInte}, Cantidad={Cantidad}, NroPed={NroPed}",
                    item.Remito, item.CodInte, item.Cantidad, item.NroPed);

                // 1. Buscar número de orden de compra en DetaOVenta
                var detaOVenta = await _context.Set<DetaOVenta>()
                    .AsNoTracking()
                    .FirstOrDefaultAsync(d => d.Nroped == item.NroPed && d.Codint == item.CodInte);

                int nroOcom = 0;
                if (detaOVenta != null && detaOVenta.NroOcom.HasValue)
                {
                    nroOcom = detaOVenta.NroOcom.Value;
                }

                // 2. Generar movimiento de stock (MOVISTOK)
                await GenerarMovimientoStockAsync(item, depositoDestino, user);

                // 3. Actualizar compras sucursal (ACTUCOMPRASUCU)
                await ActualizarComprasSucursalAsync(item, nroOcom, user);

                _logger.LogInformation("Item procesado exitosamente: CodInte={CodInte}", item.CodInte);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar item: CodInte={CodInte}", item.CodInte);
                throw;
            }
        }

        /// <summary>
        /// Genera movimiento de stock
        /// Replica MOVISTOK de VB6
        /// </summary>
        private async Task GenerarMovimientoStockAsync(RecepcionMaterialesItem item, int depositoDestino, ClaimsPrincipal user)
        {
            // Extraer número de remito del string (ej: RC.1110 -> 1110)
            int numeroRemito = 0;
            if (item.Remito.StartsWith("RC."))
            {
                int.TryParse(item.Remito.Substring(3), out numeroRemito);
            }

            // Crear DTO de movimiento de stock
            var movimiento = new MoviStockDTO
            {
                CodigoInterno = item.CodInte,
                Cantidad = item.Cantidad, // Positivo para ingreso
                Deposito = depositoDestino,
                CodigoMovimiento = "BR", // Breve recepción
                Fecha = DateTime.Now,
                DocumentoPrimarioLargo = item.Remito,
                DocumentoPrimarioCorto = item.Remito,
                DocumentoSecundarioLargo = $"BR.{DateTime.Now:dd/MM/yy}",
                DocumentoSecundarioCorto = $"BR.{DateTime.Now:dd/MM/yy}",
                Referencia = $"Recepción {DateTime.Now:dd/MM/yy}",
                NumeroCliente = string.IsNullOrWhiteSpace(item.NumeClie) ? 0 : int.TryParse(item.NumeClie, out int numCli) ? numCli : 0,
                NumeroPedidoCliente = item.NroPed,
                PrefijoRemito = 1, // Valor por defecto
                NumeroRemito = numeroRemito
            };

            // Registrar movimiento usando el servicio existente (obtiene NumUsuar y CodiEmpr desde el usuario autenticado)
            var resultado = await _moviStockService.RegistrarMovimientoAsync(movimiento, user);

            if (resultado is not OkObjectResult && resultado is not OkResult)
            {
                var detalle = (resultado as ObjectResult)?.Value?.ToString() ?? "Error desconocido";
                throw new Exception($"No se pudo registrar el movimiento en MOVISTO para el código interno {item.CodInte}: {detalle}");
            }
        }

        /// <summary>
        /// Actualiza compras sucursal
        /// Replica ACTUCOMPRASUCU de VB6
        /// </summary>
        private async Task ActualizarComprasSucursalAsync(RecepcionMaterialesItem item, int nroOcom, ClaimsPrincipal user)
        {
            if (nroOcom == 0)
            {
                // Si no hay orden de compra, registrar en TRAZAREC con OCOM = 0
                await RegistrarTrazarecAsync(item, 0, user);
                return;
            }

            // Buscar registros en OCOMDETA para actualizar
            var ocomDetas = await _context.Set<Ocomdeta>()
                .Where(o => o.CodInte == item.CodInte && 
                           o.NumeOcom == nroOcom && 
                           o.CantRec < (o.CantOcom - 0.005m))
                .OrderBy(o => o.FentreSol)
                .ThenBy(o => o.NuOrdItem)
                .ToListAsync();

            decimal cantidadRestante = item.Cantidad;

            foreach (var ocomDeta in ocomDetas)
            {
                if (cantidadRestante <= 0)
                {
                    break;
                }

                var coefUnids = ocomDeta.CoefUnids ?? 1;
                if (coefUnids < 0.0001m) coefUnids = 1;

                var cantidadRecibidaBD = ocomDeta.CantRec ?? 0;
                var cantidadARecibir = cantidadRestante / coefUnids;
                var cantidadCompra = ocomDeta.CantOcom ?? 0;

                decimal sobran = 0;
                var cantidadTotal = cantidadARecibir + cantidadRecibidaBD;

                if (cantidadTotal > cantidadCompra)
                {
                    sobran = cantidadTotal - cantidadCompra;
                    cantidadTotal = cantidadCompra;
                }

                // Actualizar cantidad recibida
                ocomDeta.CantRec = cantidadTotal;

                // Si se completó la compra, actualizar estado
                if (cantidadTotal >= cantidadCompra)
                {
                    ocomDeta.StatOcom = 3;
                }

                await _context.SaveChangesAsync();

                // Registrar en TRAZAREC
                await RegistrarTrazarecAsync(item, nroOcom, user, cantidadARecibir - sobran);

                cantidadRestante = sobran > 0 ? sobran : 0;
            }

            // Si sobra cantidad, registrar en TRAZAREC
            if (cantidadRestante > 0)
            {
                await RegistrarTrazarecAsync(item, nroOcom, user, cantidadRestante);
            }
        }

        /// <summary>
        /// Registra en TRAZAREC
        /// Replica ANOTARECEPCION_TRAZAREC de VB6
        /// </summary>
        private async Task RegistrarTrazarecAsync(RecepcionMaterialesItem item, int nroOcom, ClaimsPrincipal user, decimal? cantidad = null)
        {
            try
            {
                var trazarec = new Trazarec
                {
                    CodiEmpr = 0,
                    NroRemi = int.TryParse(item.Remito.Replace("RC.", ""), out int numRemito) ? numRemito : 0,
                    FechRecep = DateTime.Now,
                    NumeOcom = nroOcom,
                    CantRemi = cantidad ?? item.Cantidad,
                    CodInte = item.CodInte,
                    NroPed = (int)item.NroPed,
                    AddRecord = DateTime.Now,
                    LastUpdate = DateTime.Now,
                    CodiEmprNet = 0,
                    NumUsuar = _usuarioFlexoftService.InterObtenerIdFlexoft(user) ?? 0
                };

                _context.Set<Trazarec>().Add(trazarec);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Registro en TRAZAREC creado: ID={Id}, NroRemi={NroRemi}, CodInte={CodInte}",
                    trazarec.Id, trazarec.NroRemi, trazarec.CodInte);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar en TRAZAREC: CodInte={CodInte}", item.CodInte);
                throw;
            }
        }
    }
}
