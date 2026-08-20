using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace MantenimientoImp.Server.Controllers
{
    [ApiController]
    [Route("/api/BolrecepMovisto")]
    public class BolrecepMovistoController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IMoviStockService _moviStockService;
        private readonly IMoviBrecepenService _moviBrecepenService;
        private readonly FlexcrlService _flexcrlService;

        private const string PROBUS_RECEPCION = "RECEPCIO";
        private const string CLABUS_APRUEBA_CALIDAD = "APRUCALI";

        public BolrecepMovistoController(
            DataContext context,
            IMoviStockService moviStockService,
            IMoviBrecepenService moviBrecepenService,
            FlexcrlService flexcrlService)
        {
            _context = context;
            _moviStockService = moviStockService;
            _moviBrecepenService = moviBrecepenService;
            _flexcrlService = flexcrlService;
        }

        // Aprueba una recepción pendiente y genera los movimientos de stock.
        // La grilla de pendientes agrupa por (recepción + orden de compra), por lo que una misma
        // recepción puede tener varias filas con distintas OC. Para respetar esa separación,
        // este endpoint acepta numeroOrdenCompra (opcional) y procesa SOLO los items de la
        // recepción que corresponden a esa orden de compra. Si es null, procesa toda la recepción.
        [HttpPost("aprobarRecepcionEingresoMovisto/{numeroRecepcion}")]
        public async Task<ActionResult> aprobarRecepcionEingresoMovisto(int numeroRecepcion, [FromBody] List<RecepcionItemDTO> items, [FromQuery] int? numeroOrdenCompra = null)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();

            try
            {
                Console.WriteLine($"[BolrecepMovisto] Iniciando aprobación de recepción {numeroRecepcion}");
                Console.WriteLine($"[BolrecepMovisto] Items recibidos: {items?.Count ?? 0}");

                var numeroRecepcionDecimal = (decimal)numeroRecepcion;

                if (items == null || !items.Any())
                {
                    return BadRequest("No se recibieron items para aprobar");
                }

                // Validar coherencia de cantidades antes de procesar
                var erroresValidacion = new List<string>();
                foreach (var item in items)
                {
                    var suma = (item.CanToApro ?? 0) + (item.CantidadRechazada ?? 0) + (item.CantidadFaltante ?? 0);
                    var diferencia = Math.Abs(item.CantidadNominal - suma);
                    
                    if (diferencia >= 0.0001m)
                    {
                        erroresValidacion.Add($"Item {item.CodigoInterno}: CantNom={item.CantidadNominal}, Suma={suma}, Diferencia={diferencia}");
                    }
                }

                if (erroresValidacion.Any())
                {
                    var mensajeError = "Validación fallida - Cantidades incoherentes:\n" + string.Join("\n", erroresValidacion);
                    Console.WriteLine($"[BolrecepMovisto] ❌ {mensajeError}");
                    return BadRequest(mensajeError);
                }

                // Actualizar la tabla bolrecep
                int itemsActualizados = 0;
                foreach (var item in items)
                {
                    Console.WriteLine($"[BolrecepMovisto] Procesando item CodInt: {item.CodigoInterno} - Apro:{item.CanToApro}, Rech:{item.CantidadRechazada}, Falta:{item.CantidadFaltante}");

                    // Se filtra también por NumOC para no afectar al mismo producto de otra orden de compra
                    // dentro de la misma recepción (evita contaminación cruzada entre OC).
                    var bolrecep = await _context.Bolreceps
                        .Where(b => b.NumeBoRe == numeroRecepcionDecimal && b.CodInte == item.CodigoInterno
                                    && (numeroOrdenCompra == null || b.NumOC == numeroOrdenCompra))
                        .FirstOrDefaultAsync();

                    if (bolrecep != null)
                    {
                        // Actualizamos las cantidades
                        bolrecep.CanToApro = item.CanToApro ?? 0;
                        bolrecep.CanToRech = item.CantidadRechazada ?? 0;
                        bolrecep.CanToFalta = item.CantidadFaltante ?? 0;
                        bolrecep.Status = 1;
                        
                        // Actualizar información de rechazo si hay cantidad rechazada
                        if ((item.CantidadRechazada ?? 0) > 0)
                        {
                            bolrecep.Destirecha = (short?)(item.DestinoRechazo ?? 0);
                            bolrecep.MotiRecha = (short?)(item.MotivoRechazo ?? 0);
                            bolrecep.ObseRecha = item.ObservacionesRechazo;
                            Console.WriteLine($"[BolrecepMovisto] 📝 Rechazo registrado - Destino:{item.DestinoRechazo}, Motivo:{item.MotivoRechazo}");
                        }
                        
                        itemsActualizados++;
                        Console.WriteLine($"[BolrecepMovisto] ✓ Item actualizado: {item.CodigoInterno}");
                    }
                    else
                    {
                        Console.WriteLine($"[BolrecepMovisto] ⚠️ No se encontró bolrecep para CodInt: {item.CodigoInterno}");
                    }
                }

                Console.WriteLine($"[BolrecepMovisto] Items actualizados en Bolrecep: {itemsActualizados}");
                await _context.SaveChangesAsync();

                // Leer la tabla bolrecep para obtener los datos necesarios para los movimientos de stock.
                // Se filtra por NumOC para generar movimientos SOLO de la orden de compra aprobada.
                var itemsRecepcion = await _context.Bolreceps
                    .Where(b => b.NumeBoRe == numeroRecepcionDecimal
                                && (numeroOrdenCompra == null || b.NumOC == numeroOrdenCompra))
                    .ToListAsync();

                bool apruebaCalidad = await ObtenerApruebaCalidadAsync();

                // La actualización de Cant_Rendid solo aplica cuando el circuito está configurado
                // para aprobación de calidad en la solapa intermedia (APRUCALI = true).
                if (apruebaCalidad)
                {
                    var consignacionesActualizadas = await ActualizarCantRendidConsignacionesAsync(itemsRecepcion);
                    Console.WriteLine($"[BolrecepMovisto] Consignaciones actualizadas (Cant_Rendid): {consignacionesActualizadas}");
                }

                // =============================================================
                // BLOQUE DE MOVIMIENTOS DE STOCK
                // Paso A: recorrer todos los materiales de la recepción.
                // Paso B: generar movimientos según el tipo de cantidad:
                //         - Aprobada
                //         - Rechazada
                //         - Faltante / Sobrante
                // =============================================================
                int movimientosCreados = 0;
                foreach (var material in itemsRecepcion)
                {
                    // Datos base de documentos y lote para todos los movimientos del material.
                    int numeroBoleta = (int)Math.Floor(material.NumeBoRe ?? 0m);
                    string remitoProveedor = material.NumRemProv ?? string.Empty;
                    var posicionLote = (int)(material.NroItemOC ?? 0);
                    var loteRecepcion = !string.IsNullOrWhiteSpace(material.IdenLote)
                        ? material.IdenLote
                        : $"BR-{numeroBoleta}-{posicionLote}";
                    
                    // 1) CANTIDAD APROBADA
                    // Si hay cantidad aprobada, se genera ingreso normal de stock.
                    if ((material.CanToApro ?? 0) > 0)
                    {
                        var movAprobado = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = loteRecepcion,
                            CodigoMovimiento = "BR", // BR = Boleta de Recepción
                            DocumentoPrimarioLargo = $"BR-{numeroBoleta}",
                            DocumentoPrimarioCorto = $"BR-{numeroBoleta}",
                            DocumentoSecundarioLargo = remitoProveedor,
                            DocumentoSecundarioCorto = remitoProveedor,
                            Referencia = $"Rto.{remitoProveedor}",
                            Deposito = material.DepoEnt ?? 0,
                            Cantidad = material.CanToApro ?? 0, // Positivo = Ingreso
                            NumeroCliente = material.NumProv.HasValue ? -material.NumProv.Value : 0, // Negativo = Proveedor
                            ValorUnitario = material.ValoUnit ?? 0,
                            Moneda = 0
                        };

                        ActionResult resultApro;
                        if (!apruebaCalidad)
                        {
                            // Sin aprobación de calidad intermedia: impacta stock inmediatamente.
                            resultApro = await _moviStockService.RegistrarMovimientoAsync(movAprobado, User);
                        }
                        else
                        {
                            // Con aprobación de calidad intermedia: el ingreso aprobado queda en BRECEPEN
                            // para su aprobación posterior desde la solapa de calidad.
                            resultApro = await _moviBrecepenService.RegistrarMovimientoAsync(movAprobado, User);

                            var rechFal = Math.Abs((material.CanToRech ?? 0m) + (material.CanToFalta ?? 0m));
                            if (rechFal > 0.005m)
                            {
                                // Cuando hay rechazo/faltante en el mismo material, se registra el movimiento
                                // compensatorio para mantener el circuito equivalente al legado.
                                var movRechFal = new MoviStockDTO
                                {
                                    Fecha = DateTime.Now,
                                    CodigoInterno = material.CodInte ?? 0,
                                    Lote = loteRecepcion,
                                    CodigoMovimiento = "BR",
                                    DocumentoPrimarioLargo = $"BR-{numeroBoleta}",
                                    DocumentoPrimarioCorto = $"BR-{numeroBoleta}",
                                    DocumentoSecundarioLargo = remitoProveedor,
                                    DocumentoSecundarioCorto = remitoProveedor,
                                    Referencia = $"Rto.{remitoProveedor}",
                                    Deposito = material.DepoEnt ?? 0,
                                    Cantidad = rechFal,
                                    NumeroCliente = material.NumProv.HasValue ? -material.NumProv.Value : 0,
                                    ValorUnitario = material.ValoUnit ?? 0,
                                    Moneda = 0
                                };

                                var resultRechFal = await _moviStockService.RegistrarMovimientoAsync(movRechFal, User);
                                if (resultRechFal is OkResult || resultRechFal is OkObjectResult)
                                {
                                    movimientosCreados++;
                                }
                            }
                        }

                        if (resultApro is OkResult || resultApro is OkObjectResult)
                        {
                            movimientosCreados++;
                            Console.WriteLine($"[BolrecepMovisto] ✓ Movimiento APROBADO creado: {material.CanToApro} unidades");
                        }
                    }

                    // 2) CANTIDAD RECHAZADA
                    // Se realizan 2 movimientos: ingreso al depósito destino y salida/devolución.
                    if ((material.CanToRech ?? 0) > 0)
                    {
                        // 2.0) Resolver depósito destino de rechazo (si está configurado).
                        int depositoRechazo = material.DepoEnt ?? 0; // Por defecto el mismo de recepción
                        
                        if (material.Destirecha.HasValue && material.Destirecha > 0)
                        {
                            var destinoRechazo = await _context.DestinoRechazoRecepciones
                                .FirstOrDefaultAsync(d => d.CodigoDestino == material.Destirecha.Value);
                            
                            if (destinoRechazo?.Deposito.HasValue == true && destinoRechazo.Deposito > 0)
                            {
                                depositoRechazo = destinoRechazo.Deposito.Value;
                            }
                        }

                        // 2.1) Movimiento de INGRESO del rechazo al depósito destino.
                        var movRechazoIngreso = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = material.IdenLote ?? string.Empty,
                            CodigoMovimiento = "BR",
                            DocumentoPrimarioLargo = $"BR-{numeroBoleta}",
                            DocumentoPrimarioCorto = $"BR-{numeroBoleta}",
                            DocumentoSecundarioLargo = remitoProveedor,
                            DocumentoSecundarioCorto = remitoProveedor,
                            Referencia = $"RECHAZO - Rto.{remitoProveedor}",
                            Deposito = depositoRechazo,
                            Cantidad = material.CanToRech ?? 0, // Positivo = Ingreso
                            NumeroCliente = material.NumProv.HasValue ? -material.NumProv.Value : 0,
                            ValorUnitario = material.ValoUnit ?? 0
                        };

                        var resultRechazoIng = await _moviStockService.RegistrarMovimientoAsync(movRechazoIngreso, User);
                        if (resultRechazoIng is OkResult || resultRechazoIng is OkObjectResult)
                        {
                            movimientosCreados++;
                            Console.WriteLine($"[BolrecepMovisto] ✓ Movimiento RECHAZO INGRESO creado: {material.CanToRech} unidades en depósito {depositoRechazo}");
                        }

                        // 2.2) Movimiento de SALIDA del rechazo (devolución/egreso).
                        var movRechazoSalida = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = material.IdenLote ?? string.Empty,
                            CodigoMovimiento = "RD", // RD = Rechazo Devolución
                            DocumentoPrimarioLargo = $"RD-{numeroBoleta}",
                            DocumentoPrimarioCorto = $"RD-{numeroBoleta}",
                            DocumentoSecundarioLargo = remitoProveedor,
                            DocumentoSecundarioCorto = remitoProveedor,
                            Referencia = $"RECHAZO - Rto.{remitoProveedor}",
                            Deposito = depositoRechazo,
                            Cantidad = -(material.CanToRech ?? 0), // Negativo = Salida
                            NumeroCliente = material.NumProv.HasValue ? -material.NumProv.Value : 0,
                            ValorUnitario = material.ValoUnit ?? 0
                        };

                        var resultRechazoSal = await _moviStockService.RegistrarMovimientoAsync(movRechazoSalida, User);
                        if (resultRechazoSal is OkResult || resultRechazoSal is OkObjectResult)
                        {
                            movimientosCreados++;
                            Console.WriteLine($"[BolrecepMovisto] ✓ Movimiento RECHAZO SALIDA creado: {material.CanToRech} unidades");
                        }
                    }

                    // 3) CANTIDAD FALTANTE / SOBRANTE
                    // Si la diferencia no es despreciable, se generan movimientos de ida y vuelta
                    // con referencia específica para trazabilidad.
                    if (Math.Abs(material.CanToFalta ?? 0) >= 0.05m)
                    {
                        decimal cantidadFaltante = (material.CanToFalta ?? 0) * -1;
                        string referenciaFaltante = cantidadFaltante > 0 
                            ? $"SOBRANTE - Rto.{remitoProveedor}" 
                            : $"FALTANTE - Rto.{remitoProveedor}";

                        // 3.1) Movimiento de INGRESO de ajuste por faltante/sobrante.
                        var movFaltanteIngreso = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = material.IdenLote ?? string.Empty,
                            CodigoMovimiento = "BR",
                            DocumentoPrimarioLargo = $"BR-{numeroBoleta}",
                            DocumentoPrimarioCorto = $"BR-{numeroBoleta}",
                            DocumentoSecundarioLargo = remitoProveedor,
                            DocumentoSecundarioCorto = remitoProveedor,
                            Referencia = referenciaFaltante,
                            Deposito = material.DepoEnt ?? 0,
                            Cantidad = cantidadFaltante, // Positivo = Ingreso
                            NumeroCliente = material.NumProv.HasValue ? -material.NumProv.Value : 0,
                            ValorUnitario = material.ValoUnit ?? 0
                        };

                        var resultFaltanteIng = await _moviStockService.RegistrarMovimientoAsync(movFaltanteIngreso, User);
                        if (resultFaltanteIng is OkResult || resultFaltanteIng is OkObjectResult)
                        {
                            movimientosCreados++;
                            Console.WriteLine($"[BolrecepMovisto] ✓ Movimiento FALTANTE INGRESO creado: {cantidadFaltante} unidades");
                        }

                        // 3.2) Movimiento de SALIDA complementario para cerrar el ajuste.
                        var movFaltanteSalida = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = material.IdenLote ?? string.Empty,
                            CodigoMovimiento = "BR",
                            DocumentoPrimarioLargo = $"BR-{numeroBoleta}",
                            DocumentoPrimarioCorto = $"BR-{numeroBoleta}",
                            DocumentoSecundarioLargo = remitoProveedor,
                            DocumentoSecundarioCorto = remitoProveedor,
                            Referencia = referenciaFaltante,
                            Deposito = material.DepoEnt ?? 0,
                            Cantidad = -cantidadFaltante, // Negativo = Salida
                            NumeroCliente = material.NumProv.HasValue ? -material.NumProv.Value : 0,
                            ValorUnitario = material.ValoUnit ?? 0
                        };

                        var resultFaltanteSal = await _moviStockService.RegistrarMovimientoAsync(movFaltanteSalida, User);
                        if (resultFaltanteSal is OkResult || resultFaltanteSal is OkObjectResult)
                        {
                            movimientosCreados++;
                            Console.WriteLine($"[BolrecepMovisto] ✓ Movimiento FALTANTE SALIDA creado: {cantidadFaltante} unidades");
                        }
                    }
                }

                Console.WriteLine($"[BolrecepMovisto] Total movimientos de stock creados: {movimientosCreados}");

                // Confirmar la transacción
                await transaction.CommitAsync();

                Console.WriteLine($"[BolrecepMovisto] ✅ Recepción {numeroRecepcion} aprobada exitosamente");
                return Ok(new { mensaje = "Recepción aprobada exitosamente", itemsActualizados, movimientosCreados });
            }
            catch (Exception ex)
            {
                // Revertir la transacción en caso de error
                Console.WriteLine($"[BolrecepMovisto] ❌ Error al aprobar recepción {numeroRecepcion}: {ex.Message}");
                Console.WriteLine($"[BolrecepMovisto] StackTrace: {ex.StackTrace}");
                await transaction.RollbackAsync();
                return BadRequest($"Error al aprobar la recepción: {ex.Message}");
            }
        }

        private async Task<bool> ObtenerApruebaCalidadAsync()
        {
            var control = await _flexcrlService.ObtenerControlAsync(CLABUS_APRUEBA_CALIDAD, PROBUS_RECEPCION);
            var val = (control?.VALCONTROL ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(val))
            {
                return false;
            }

            if (val == "1")
            {
                return true;
            }

            return bool.TryParse(val, out bool parsed) && parsed;
        }

        private async Task<int> ActualizarCantRendidConsignacionesAsync(List<Bolrecep> itemsRecepcion)
        {
            // Paso 0: validación básica de entrada.
            if (itemsRecepcion == null || !itemsRecepcion.Any())
            {
                return 0;
            }

            // Paso 1: tomar solo líneas aprobadas con referencia completa a OC + ítem.
            // CanToApro es la cantidad de producto padre aprobada en recepción.
            var lineasAprobadas = itemsRecepcion
                .Where(x => (x.CanToApro ?? 0m) > 0.005m)
                .Where(x => (x.NumOC ?? 0) > 0)
                .Where(x => (x.NroItemOC ?? 0) > 0)
                .Select(x => new
                {
                    NumeroOrden = x.NumOC!.Value,
                    NumeroItem = x.NroItemOC!.Value,
                    CantidadPadreAprobada = x.CanToApro ?? 0m
                })
                .ToList();

            if (!lineasAprobadas.Any())
            {
                return 0;
            }

            // Paso 2: preparar claves únicas de OC/ítem para reducir consultas.
            var claves = lineasAprobadas
                .Select(x => new { x.NumeroOrden, x.NumeroItem })
                .Distinct()
                .ToList();

            var numerosOrden = claves.Select(x => x.NumeroOrden).Distinct().ToList();

            // Paso 3: traer los ítems de OCOMDETA para obtener la cantidad base del padre (Cant_Ocom).
            var ocomItems = await _context.Ocomdetas
                .Where(d => d.NumeOcom.HasValue && numerosOrden.Contains(d.NumeOcom.Value))
                .Where(d => d.NuOrdItem.HasValue)
                .ToListAsync();

            var ocomMap = ocomItems
                .Where(d => d.NumeOcom.HasValue && d.NuOrdItem.HasValue)
                .GroupBy(d => new { Orden = d.NumeOcom!.Value, Item = d.NuOrdItem!.Value })
                .ToDictionary(g => (g.Key.Orden, g.Key.Item), g => g.First());

            // Paso 4: construir patrón de Nro_Ocom (OC-XXXXXX-YY) para buscar las consignaciones
            // asociadas a cada línea de OC.
            var patronesOcom = claves
                .Select(x => $"OC-{x.NumeroOrden:000000}-{x.NumeroItem:00}")
                .Distinct()
                .ToList();

            // Solo se consideran consignaciones remitidas (Cant_Entreg > 0),
            // porque Cant_Rendid no debe crecer sobre material no entregado.
            var maconsigFiltrado = await _context.Maconsigs
                .Where(m => m.NroOcom != null && patronesOcom.Contains(m.NroOcom))
                .Where(m => (m.CantEntreg ?? 0m) > 0.005m)
                .ToListAsync();

            var maconsigPorNroOcom = maconsigFiltrado
                .GroupBy(m => m.NroOcom ?? string.Empty)
                .ToDictionary(g => g.Key, g => g.ToList());

            int actualizados = 0;

            foreach (var linea in lineasAprobadas)
            {
                // Paso 5: obtener datos del producto padre para calcular proporcionalidad.
                if (!ocomMap.TryGetValue((linea.NumeroOrden, linea.NumeroItem), out var itemPadre))
                {
                    continue;
                }

                var cantidadPadreOrden = itemPadre.CantOcom ?? 0m;
                if (cantidadPadreOrden <= 0.005m)
                {
                    continue;
                }

                var nroOcom = $"OC-{linea.NumeroOrden:000000}-{linea.NumeroItem:00}";
                if (!maconsigPorNroOcom.TryGetValue(nroOcom, out var consignaciones) || consignaciones.Count == 0)
                {
                    continue;
                }

                foreach (var consignacion in consignaciones)
                {
                    // Paso 6: calcular consumo proporcional del componente.
                    // Fórmula: (Cant_Asig del componente / Cant_Ocom del padre) * CantidadPadreAprobada.
                    var cantAsig = consignacion.CantAsig ?? 0m;
                    if (cantAsig <= 0.0001m)
                    {
                        continue;
                    }

                    var consumoProporcional = Math.Round((cantAsig / cantidadPadreOrden) * linea.CantidadPadreAprobada, 4);
                    if (consumoProporcional <= 0.0001m)
                    {
                        continue;
                    }

                    // Paso 7: aplicar tope para no rendir más que lo efectivamente remitido.
                    // Disponible a rendir = Cant_Entreg - Cant_Rendid actual.
                    var cantRendidActual = consignacion.CantRendid ?? 0m;
                    var cantEntreg = consignacion.CantEntreg ?? 0m;
                    var maxRendible = Math.Max(0m, cantEntreg - cantRendidActual);
                    if (maxRendible <= 0.0001m)
                    {
                        continue;
                    }

                    var incremento = Math.Min(consumoProporcional, maxRendible);
                    if (incremento <= 0.0001m)
                    {
                        continue;
                    }

                    // Paso 8: persistir incremento de Cant_Rendid y marca de actualización.
                    consignacion.CantRendid = Math.Round(cantRendidActual + incremento, 4);
                    consignacion.LastUpdate = DateTime.Now;
                    actualizados++;
                }
            }

            return actualizados;
        }
    }
}
