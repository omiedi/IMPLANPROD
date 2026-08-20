using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Shared.Interfaces;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para gestionar costos de items de órdenes de compra
    /// Implementa lógica de negocio para mostrar y guardar costos con control de acceso
    /// </summary>
    public class OCOMDETACostoService : IOCOMDETACostoService
    {
        private readonly DataContext _context;
        private readonly IUserContextService _userContextService;

        public OCOMDETACostoService(DataContext context, IUserContextService userContextService)
        {
            _context = context;
            _userContextService = userContextService;
        }

        /// <summary>
        /// Obtiene los costos de items de una orden de compra
        /// Hace JOIN entre OCOMDETA y OCOMDETACosto
        /// Si no existe costo, muestra el item con costo NULL
        /// </summary>
        public async Task<List<OCOMDETACostoDto>> GetCostosPorOrdenAsync(int numeOcom)
        {
            // Obtener items de la orden (solo activos, StatOcom < 9)
            var itemsOrden = await _context.Ocomdetas
                .Where(d => d.NumeOcom == numeOcom && d.StatOcom < 9)
                .OrderBy(d => d.NuOrdItem)
                .ToListAsync();

            // Obtener costos existentes para esta orden
            var costosExistentes = await _context.OCOMDETACostos
                .Where(c => c.NumeOcom == numeOcom)
                .ToListAsync();

            // Obtener observaciones del encabezado de la orden (fallback)
            var observacionesEncabezado = await _context.Ocomencas
                .Where(o => o.NumeOcom == numeOcom)
                .Select(o => o.ObseOcom)
                .FirstOrDefaultAsync();

            // Obtener descripciones de monedas
            var monedas = await _context.Tamoneds.ToListAsync();

            // Combinar items con costos
            var resultado = itemsOrden.Select(item =>
            {
                // Buscar costo existente que coincida con NumeOcom + NuOrdItem + CodInte
                var costoExistente = costosExistentes.FirstOrDefault(c =>
                    c.NumeOcom == item.NumeOcom &&
                    c.NuOrdItem == item.NuOrdItem &&
                    c.CodInte == item.CodInte
                );

                // Obtener descripción de moneda
                var moneda = monedas.FirstOrDefault(m => m.CodigoMoneda == item.MoneEmis);

                return new OCOMDETACostoDto
                {
                    // Si existe costo, usar su ID, sino 0 (indica que es nuevo)
                    OCOMDETA_ID = costoExistente?.OCOMDETA_ID ?? 0,
                    
                    // Datos de identificación (vienen de OCOMDETA)
                    NumeOcom = item.NumeOcom,
                    NuOrdItem = item.NuOrdItem,
                    CodInte = item.CodInte,
                    
                    // Datos para mostrar en UI (vienen de OCOMDETA)
                    CodExte = item.CodExte,
                    Descritem = item.Descritem,
                    UniCom = item.UniCom,
                    CantOcom = item.CantOcom,
                    FentreSol = item.FentreSol,
                    MoneEmis = item.MoneEmis,
                    StatOcom = item.StatOcom,
                    
                    // Costo (viene de OCOMDETACosto si existe, sino NULL)
                    PrUnNeto = costoExistente?.PrUnNeto,
                    // Prioridad de observaciones para el modal:
                    // 1) OCOMDETACosto.Observaciones, 2) OCOMENCA.Obse_Ocom
                    Observaciones = !string.IsNullOrWhiteSpace(costoExistente?.Observaciones)
                        ? costoExistente!.Observaciones
                        : observacionesEncabezado,
                    
                    // Descripción de moneda
                    DescripcionMoneda = moneda?.Descripcion ?? "N/D"
                };
            }).ToList();

            return resultado;
        }

        /// <summary>
        /// Guarda o actualiza el costo de un item
        /// Lógica: Si existe registro con mismo NumeOcom + NuOrdItem + CodInte → UPDATE
        ///         Si no existe → INSERT
        /// </summary>
        public async Task<OCOMDETACostoDto> GuardarCostoAsync(OCOMDETACostoDto dto)
        {
            if (!string.IsNullOrEmpty(dto.Observaciones) && dto.Observaciones.Length > 1024)
            {
                throw new Exception("El campo Observaciones no puede superar los 1024 caracteres.");
            }

            // Buscar si ya existe un registro
            var costoExistente = await _context.OCOMDETACostos
                .FirstOrDefaultAsync(c =>
                    c.NumeOcom == dto.NumeOcom &&
                    c.NuOrdItem == dto.NuOrdItem &&
                    c.CodInte == dto.CodInte
                );

            if (costoExistente != null)
            {
                // UPDATE: Actualizar costo existente
                costoExistente.PrUnNeto = dto.PrUnNeto;
                costoExistente.Observaciones = dto.Observaciones;
                costoExistente.CodExte = dto.CodExte;
                costoExistente.Descritem = dto.Descritem;
                costoExistente.LastUpdate = DateTime.Now;
                costoExistente.NumUsuar = _userContextService.GetCurrentUserId();

                _context.OCOMDETACostos.Update(costoExistente);
                await _context.SaveChangesAsync();

                dto.OCOMDETA_ID = costoExistente.OCOMDETA_ID;
            }
            else
            {
                // INSERT: Crear nuevo registro
                // Obtener datos completos del item desde OCOMDETA
                var itemOcomdeta = await _context.Ocomdetas
                    .FirstOrDefaultAsync(d =>
                        d.NumeOcom == dto.NumeOcom &&
                        d.NuOrdItem == dto.NuOrdItem &&
                        d.CodInte == dto.CodInte
                    );

                if (itemOcomdeta == null)
                {
                    throw new Exception($"No se encontró el item {dto.NuOrdItem} en la orden {dto.NumeOcom}");
                }

                // Obtener datos del encabezado
                var ordenEnc = await _context.Ocomencas
                    .FirstOrDefaultAsync(o => o.NumeOcom == dto.NumeOcom);

                var nuevoCosto = new OCOMDETACosto
                {
                    // Identificación
                    NumeOcom = dto.NumeOcom,
                    CodiEmpr = 0,
                    NuOrdItem = dto.NuOrdItem,
                    CodInte = dto.CodInte,
                    
                    // Datos del encabezado
                    FemiOcom = ordenEnc?.FemiOcom,
                    NproOcom = (short?)ordenEnc?.NproOcom,
                    RaSoProv = ordenEnc?.RaSoProv,
                    
                    // Datos del item
                    CodExte = itemOcomdeta.CodExte,
                    Descritem = itemOcomdeta.Descritem,
                    UniMas = itemOcomdeta.UniMas,
                    UnidsOcom = itemOcomdeta.UnidsOcom,
                    CantOcom = itemOcomdeta.CantOcom,
                    UniCom = itemOcomdeta.UniCom,
                    FentreSol = itemOcomdeta.FentreSol,
                    MoneEmis = (short?)itemOcomdeta.MoneEmis,
                    MoneCos = (short?)itemOcomdeta.MoneEmis,
                    StatOcom = itemOcomdeta.StatOcom,
                    
                    // Costo (campo principal editable)
                    PrUnNeto = dto.PrUnNeto,
                    Observaciones = dto.Observaciones,
                    
                    // Auditoría
                    AddRecord = DateTime.Now,
                    NumUsuar = _userContextService.GetCurrentUserId(),
                    CodiEmprNet = 0
                };

                _context.OCOMDETACostos.Add(nuevoCosto);
                await _context.SaveChangesAsync();

                dto.OCOMDETA_ID = nuevoCosto.OCOMDETA_ID;
            }

            return dto;
        }

        /// <summary>
        /// Guarda múltiples costos en una sola operación
        /// Útil para guardar todos los items de una orden
        /// </summary>
        public async Task<List<OCOMDETACostoDto>> GuardarCostosAsync(List<OCOMDETACostoDto> dtos)
        {
            var resultado = new List<OCOMDETACostoDto>();

            foreach (var dto in dtos)
            {
                var dtoGuardado = await GuardarCostoAsync(dto);
                resultado.Add(dtoGuardado);
            }

            return resultado;
        }

        /// <summary>
        /// Guarda costos y registra cambios en NOVEDADES
        /// Implementa la lógica del VB6: ADD_NOVEDAD
        /// - Obtiene precio anterior del mismo producto (último registro por AddRecord)
        /// - Convierte precios a pesos usando cotización de TAMONED
        /// - Guarda en NOVEDADES con observación
        /// </summary>
        public async Task<List<OCOMDETACostoDto>> GuardarCostosConNovedadesAsync(GuardarCostosConNovedadesDto dto)
        {
            // 1. Guardar los costos normalmente
            var costosGuardados = await GuardarCostosAsync(dto.Costos);

            // 2. Si la configuración es "NUNCA", no registrar en NOVEDADES
            if (dto.ConfiguracionACOSREP?.ToUpper() == "NUNCA")
            {
                return costosGuardados;
            }

            // 3. Registrar cambios en NOVEDADES para cada item
            foreach (var costo in dto.Costos)
            {
                await RegistrarNovedadAsync(costo, dto.Observacion ?? string.Empty);
            }

            return costosGuardados;
        }

        /// <summary>
        /// Registra un cambio de precio en la tabla NOVEDADES
        /// Implementa la lógica del VB6: ADD_NOVEDAD
        /// 
        /// LÓGICA DE NEGOCIO:
        /// 1. Busca el precio anterior en NOVEDADES (último PrecNew del producto por código)
        /// 2. Obtiene la cotización anterior para detectar cambios de cotización
        /// 3. Convierte ambos precios a PESOS para comparar (precio * cotización)
        /// 4. Calcula diferencia porcentual entre precio anterior y nuevo (ambos en pesos)
        /// 5. Solo registra si la diferencia es > 0.5%
        /// 
        /// CASOS ESPECIALES:
        /// - Primera vez (sin precio anterior): Registra con PrecAnt = PrecNew
        /// - Cambio de cotización: Aunque el precio en USD sea igual, si la cotización cambió
        ///   y el resultado en pesos es diferente > 0.5%, se registra
        /// - Diferencia <= 0.5%: NO se registra (cambio insignificante)
        /// 
        /// IMPORTANTE:
        /// - PrecAnt y PrecNew SIEMPRE se guardan en PESOS
        /// - Se guarda la cotización actual en CotizaDolar
        /// - Se guarda el número de orden de compra que originó el cambio
        /// </summary>
        /// <param name="costo">DTO con los datos del costo editado</param>
        /// <param name="observacion">Observación del cambio de precio</param>
        private async Task RegistrarNovedadAsync(OCOMDETACostoDto costo, string observacion)
        {
            try
            {
                // ===================================================================
                // VALIDACIÓN 1: Código interno válido
                // ===================================================================
                if (costo.CodInte <= 0)
                {
                    Console.WriteLine($"⚠️ No se registra novedad: CodInte inválido ({costo.CodInte})");
                    return;
                }

                // ===================================================================
                // PASO 1: Obtener moneda y cotización actual
                // ===================================================================
                // Si la moneda es null o 0, usar 1 (pesos)
                short moneda = costo.MoneEmis ?? 1;
                if (moneda < 1) moneda = 1;

                // Obtener cotización ACTUAL de la moneda desde tabla TAMONED
                // Ejemplo: USD = 1050, EUR = 1150, Pesos = 1
                decimal cotizacionActual = await ObtenerCotizacionMonedaAsync(moneda);

                // ===================================================================
                // PASO 2: Obtener precio anterior (en pesos) de NOVEDADES
                // ===================================================================
                // Busca el último registro del producto por código interno
                // Retorna: (PrecNew en pesos, CotizaDolar anterior)
                // Si no existe registro previo, retorna (0, 0)
                var (precioAnteriorPesos, cotizacionAnterior) = await ObtenerPrecioYCotizacionAnteriorAsync((int)(costo.CodInte ?? 0));

                // ===================================================================
                // PASO 3: Convertir precio nuevo a pesos
                // ===================================================================
                // Ejemplo: USD 100 * cotiz 1050 = $105,000 pesos
                decimal precioNuevoPesos = (costo.PrUnNeto ?? 0) * cotizacionActual;

                // ===================================================================
                // PASO 4: Validar diferencia de precio
                // ===================================================================
                // Solo se registra si la diferencia es MAYOR al 0.5%
                // Esto evita registrar cambios insignificantes por redondeos
                decimal diferenciaPorcentaje = 0;
                bool esPrimeraVez = precioAnteriorPesos == 0;

                if (esPrimeraVez)
                {
                    // CASO A: Primera vez (no hay registro previo en NOVEDADES)
                    // Se registra con PrecAnt = PrecNew para establecer el precio base
                    diferenciaPorcentaje = 100;
                    precioAnteriorPesos = precioNuevoPesos;
                    Console.WriteLine($"ℹ️ Primera novedad para CodInte {costo.CodInte}: PrecAnt = PrecNew = {precioNuevoPesos:N4}");
                }
                else
                {
                    // CASO B: Ya existe precio anterior
                    // Calcular diferencia porcentual: |nuevo - anterior| / anterior * 100
                    // Ejemplo: |110,000 - 105,000| / 105,000 * 100 = 4.76%
                    diferenciaPorcentaje = Math.Abs((precioNuevoPesos - precioAnteriorPesos) / precioAnteriorPesos * 100);
                    
                    // Si la diferencia es <= 0.5%, NO registrar (cambio insignificante)
                    if (diferenciaPorcentaje <= 0.5m)
                    {
                        Console.WriteLine($"⚠️ No se registra novedad para CodInte {costo.CodInte}: diferencia {diferenciaPorcentaje:N2}% <= 0.5%");
                        Console.WriteLine($"   PrecAnterior={precioAnteriorPesos:N4} (cotiz={cotizacionAnterior:N4}), PrecNuevo={precioNuevoPesos:N4} (cotiz={cotizacionActual:N4})");
                        return;
                    }
                }

                // ===================================================================
                // PASO 5: Crear registro en NOVEDADES
                // ===================================================================
                // Obtener número de proveedor de la orden de compra desde OCOMENCA
                int? numProv = await ObtenerNumProvOrdenAsync(costo.NumeOcom ?? 0);
                
                // Crear el objeto Novedades con todos los datos
                var novedad = new Novedades
                {
                    Codint = (int)costo.CodInte,              // Código interno del producto
                    CodiEmpr = 0,                              // Siempre 0 según reglas globales
                    FechMod = DateTime.Now,                    // Fecha del cambio
                    PrecAnt = precioAnteriorPesos,             // Precio anterior EN PESOS
                    PrecNew = precioNuevoPesos,                // Precio nuevo EN PESOS
                    NumProv = numProv,                         // Número de proveedor
                    Observacion = observacion.Length > 512 ? observacion.Substring(0, 512) : observacion, // Máx 512 caracteres
                    Moneda = moneda,                           // Código de moneda (1=Pesos, 2=USD, etc)
                    ListaRepos = "COSTO",                     // Fijo: identifica que viene de costos de OC
                    CotizaDolar = cotizacionActual,            // Cotización usada para convertir a pesos
                    NumeOcom = costo.NumeOcom,                 // Número de orden de compra que originó el cambio
                    
                    // Campos de auditoría automáticos
                    AddRecord = DateTime.Now,                  // Fecha de creación del registro
                    NumUsuar = _userContextService.GetCurrentUserId(), // Usuario que realizó el cambio
                    CodiEmprNet = 0                            // Código de empresa (siempre 0)
                };

                _context.Novedades.Add(novedad);
                await _context.SaveChangesAsync();

                Console.WriteLine($"✅ Novedad registrada: CodInte={costo.CodInte}, OrdenCompra={costo.NumeOcom}");
                Console.WriteLine($"   PrecAnt={precioAnteriorPesos:N4}, PrecNew={precioNuevoPesos:N4}, Diferencia={diferenciaPorcentaje:N2}%");
                Console.WriteLine($"   Cotización Anterior={cotizacionAnterior:N4}, Cotización Actual={cotizacionActual:N4}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error al registrar novedad para CodInte {costo.CodInte}: {ex.Message}");
                // No lanzar excepción para no interrumpir el guardado de costos
            }
        }

        /// <summary>
        /// Obtiene la cotización de una moneda desde TAMONED
        /// Si es moneda 1 (pesos), retorna 1
        /// Si no encuentra la moneda, retorna 1
        /// </summary>
        private async Task<decimal> ObtenerCotizacionMonedaAsync(short codMoneda)
        {
            try
            {
                // Si es pesos (código 1), la cotización es 1
                if (codMoneda == 1)
                {
                    return 1m;
                }

                // Buscar la moneda en TAMONED
                var moneda = await _context.Tamoneds
                    .Where(m => m.CodigoMoneda == codMoneda && m.Status == 0)
                    .FirstOrDefaultAsync();

                if (moneda != null && moneda.Cotizacion > 0)
                {
                    return moneda.Cotizacion;
                }

                // Si no se encuentra o no tiene cotización, retornar 1
                Console.WriteLine($"⚠️ No se encontró cotización para moneda {codMoneda}, usando 1");
                return 1m;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error al obtener cotización de moneda {codMoneda}: {ex.Message}");
                return 1m;
            }
        }

        /// <summary>
        /// Obtiene el precio anterior (en pesos) y la cotización anterior de NOVEDADES
        /// 
        /// PROPÓSITO:
        /// - Busca el último registro del producto en NOVEDADES por código interno
        /// - Retorna el PrecNew (que ya está en pesos) como precio anterior
        /// - Retorna también la CotizaDolar para detectar cambios de cotización
        /// 
        /// POR QUÉ SE NECESITA LA COTIZACIÓN ANTERIOR:
        /// Si el precio en USD no cambió pero la cotización sí, el precio en pesos
        /// será diferente y debe registrarse como novedad.
        /// Ejemplo: USD 100 con cotiz 1050 = $105,000
        ///          USD 100 con cotiz 1100 = $110,000 (diferencia 4.76%)
        /// 
        /// RETORNO:
        /// - (precioAnteriorPesos, cotizacionAnterior) si encuentra registro
        /// - (0, 0) si es la primera vez (no hay registro previo)
        /// </summary>
        /// <param name="codInte">Código interno del producto</param>
        /// <returns>Tupla con (precio anterior en pesos, cotización anterior)</returns>
        private async Task<(decimal precioAnteriorPesos, decimal cotizacionAnterior)> ObtenerPrecioYCotizacionAnteriorAsync(int codInte)
        {
            try
            {
                // Buscar la última novedad del producto por código interno
                // Se ordena por AddRecord DESC para obtener el registro más reciente
                // IMPORTANTE: Se busca SOLO por código, no por orden ni item
                var ultimaNovedad = await _context.Novedades
                    .Where(n => n.Codint == codInte)
                    .OrderByDescending(n => n.AddRecord)
                    .FirstOrDefaultAsync();

                if (ultimaNovedad != null)
                {
                    // PrecNew ya está en pesos (se guardó convertido)
                    decimal precioAnterior = ultimaNovedad.PrecNew ?? 0m;
                    // CotizaDolar es la cotización que se usó para convertir
                    decimal cotizacionAnterior = ultimaNovedad.CotizaDolar ?? 1m;
                    
                    Console.WriteLine($"📊 Precio anterior encontrado para CodInte {codInte}: {precioAnterior:N4} pesos (cotiz={cotizacionAnterior:N4})");
                    return (precioAnterior, cotizacionAnterior);
                }

                Console.WriteLine($"ℹ️ No hay precio anterior para CodInte {codInte} (primera vez)");
                return (0m, 0m);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error al obtener precio anterior para CodInte {codInte}: {ex.Message}");
                return (0m, 0m);
            }
        }

        /// <summary>
        /// Obtiene el precio anterior del producto desde NOVEDADES
        /// Busca el último registro ordenado por AddRecord DESC
        /// Si no encuentra, retorna 0
        /// NOTA: Este método ya no se usa para comparación de precios, solo se mantiene por compatibilidad
        /// </summary>
        private async Task<decimal> ObtenerPrecioAnteriorAsync(int codInte)
        {
            try
            {
                var ultimaNovedad = await _context.Novedades
                    .Where(n => n.Codint == codInte)
                    .OrderByDescending(n => n.AddRecord)
                    .FirstOrDefaultAsync();

                if (ultimaNovedad != null && ultimaNovedad.PrecNew.HasValue)
                {
                    return ultimaNovedad.PrecNew.Value;
                }

                return 0m;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error al obtener precio anterior para CodInte {codInte}: {ex.Message}");
                return 0m;
            }
        }

        /// <summary>
        /// Obtiene el número de proveedor de una orden de compra desde OCOMENCA
        /// </summary>
        private async Task<int?> ObtenerNumProvOrdenAsync(int numeOcom)
        {
            try
            {
                var orden = await _context.Ocomencas
                    .Where(o => o.NumeOcom == numeOcom)
                    .FirstOrDefaultAsync();

                return orden?.NproOcom;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error al obtener NproOcom de orden {numeOcom}: {ex.Message}");
                return null;
            }
        }
    }
}
