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

        public BolrecepMovistoController(DataContext context, IMoviStockService moviStockService)
        {
            _context = context;
            _moviStockService = moviStockService;
        }

        [HttpPost("aprobarRecepcionEingresoMovisto/{numeroRecepcion}")]
        public async Task<ActionResult> aprobarRecepcionEingresoMovisto(int numeroRecepcion, [FromBody] List<RecepcionItemDTO> items)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();

            try
            {
                Console.WriteLine($"[BolrecepMovisto] Iniciando aprobación de recepción {numeroRecepcion}");
                Console.WriteLine($"[BolrecepMovisto] Items recibidos: {items?.Count ?? 0}");

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

                    var bolrecep = await _context.Bolreceps
                        .Where(b => b.NumeBoRe == numeroRecepcion && b.CodInte == item.CodigoInterno)
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

                // Leer la tabla bolrecep para obtener los datos necesarios para los movimientos de stock
                var itemsRecepcion = await _context.Bolreceps
                    .Where(b => b.NumeBoRe == numeroRecepcion)
                    .ToListAsync();

                // Generar movimientos de stock usando el servicio
                int movimientosCreados = 0;
                foreach (var material in itemsRecepcion)
                {
                    string remitoProveedor = material.NumRemProv ?? string.Empty;
                    
                    // 1. MOVIMIENTO DE CANTIDAD APROBADA (si > 0)
                    if ((material.CanToApro ?? 0) > 0)
                    {
                        var movAprobado = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = material.IdenLote ?? string.Empty,
                            CodigoMovimiento = "BR", // BR = Boleta de Recepción
                            DocumentoPrimarioLargo = $"BR-{material.NumeBoRe}",
                            DocumentoPrimarioCorto = $"BR-{material.NumeBoRe}",
                            DocumentoSecundarioLargo = remitoProveedor,
                            DocumentoSecundarioCorto = remitoProveedor,
                            Referencia = $"Rto.{remitoProveedor}",
                            Deposito = material.DepoEnt ?? 0,
                            Cantidad = material.CanToApro ?? 0, // Positivo = Ingreso
                            NumeroCliente = material.NumProv.HasValue ? -material.NumProv.Value : 0, // Negativo = Proveedor
                            ValorUnitario = material.ValoUnit ?? 0
                        };

                        var resultApro = await _moviStockService.RegistrarMovimientoAsync(movAprobado, User);
                        if (resultApro is OkResult || resultApro is OkObjectResult)
                        {
                            movimientosCreados++;
                            Console.WriteLine($"[BolrecepMovisto] ✓ Movimiento APROBADO creado: {material.CanToApro} unidades");
                        }
                    }

                    // 2. MOVIMIENTOS DE CANTIDAD RECHAZADA (si > 0)
                    if ((material.CanToRech ?? 0) > 0)
                    {
                        // Obtener el depósito del destino de rechazo
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

                        // 2.1. Movimiento de INGRESO del rechazo
                        var movRechazoIngreso = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = material.IdenLote ?? string.Empty,
                            CodigoMovimiento = "BR",
                            DocumentoPrimarioLargo = $"BR-{material.NumeBoRe}",
                            DocumentoPrimarioCorto = $"BR-{material.NumeBoRe}",
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

                        // 2.2. Movimiento de SALIDA del rechazo
                        int numeroBoleta = (int)Math.Floor(material.NumeBoRe ?? 0);
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

                    // 3. MOVIMIENTOS DE CANTIDAD FALTANTE (si diferente de 0)
                    if (Math.Abs(material.CanToFalta ?? 0) >= 0.05m)
                    {
                        decimal cantidadFaltante = (material.CanToFalta ?? 0) * -1;
                        string referenciaFaltante = cantidadFaltante > 0 
                            ? $"SOBRANTE - Rto.{remitoProveedor}" 
                            : $"FALTANTE - Rto.{remitoProveedor}";

                        // 3.1. Movimiento de INGRESO del faltante
                        var movFaltanteIngreso = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = material.IdenLote ?? string.Empty,
                            CodigoMovimiento = "BR",
                            DocumentoPrimarioLargo = $"BR-{material.NumeBoRe}",
                            DocumentoPrimarioCorto = $"BR-{material.NumeBoRe}",
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

                        // 3.2. Movimiento de SALIDA del faltante
                        var movFaltanteSalida = new MoviStockDTO
                        {
                            Fecha = DateTime.Now,
                            CodigoInterno = material.CodInte ?? 0,
                            Lote = material.IdenLote ?? string.Empty,
                            CodigoMovimiento = "BR",
                            DocumentoPrimarioLargo = $"BR-{material.NumeBoRe}",
                            DocumentoPrimarioCorto = $"BR-{material.NumeBoRe}",
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
    }
}
