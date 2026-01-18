using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para manejar la lógica de DETAIVA (detalle de IVA de órdenes de compra)
    /// Replica la funcionalidad del código VB6 para grabar el detalle del IVA
    /// </summary>
    public class DetaivaService
    {
        private readonly DataContext _context;
        private readonly ILogger<DetaivaService> _logger;

        public DetaivaService(DataContext context, ILogger<DetaivaService> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Graba el detalle de IVA en la tabla DETAIVA
        /// Replica la lógica del código VB6: "aqui graba el detalle del IVA"
        /// </summary>
        /// <param name="numeroOrden">Número de orden de compra (NRO& del VB6)</param>
        /// <param name="codigoEmpresa">Código de empresa (CodiEmp% del VB6)</param>
        /// <param name="distribucionIva">Distribución de IVA (equivalente a RIVACOM07 del VB6)</param>
        /// <returns>True si se grabó correctamente</returns>
        public async Task<bool> GrabarDetalleIvaAsync(int numeroOrden, int? codigoEmpresa, DistribucionIvaDTO distribucionIva)
        {
            try
            {
                _logger.LogInformation($"Iniciando grabación de detalle IVA para orden {numeroOrden}");

                // 1. Eliminar registros existentes de DETAIVA para esta orden
                // Equivalente a: STRSQL = "Select * from DETAIVA where Nume_Ocom = " & NRO&
                var detallesExistentes = await _context.Detaivas
                    .Where(d => d.NumeOcom == numeroOrden)
                    .ToListAsync();

                if (detallesExistentes.Any())
                {
                    _context.Detaivas.RemoveRange(detallesExistentes);
                    _logger.LogInformation($"Eliminados {detallesExistentes.Count} registros existentes de DETAIVA para orden {numeroOrden}");
                }

                // 2. Obtener arrays de datos para procesar (equivalente a RIVACOM07 del VB6)
                var basesImponibles = distribucionIva.ObtenerBasesImponibles();
                var importesIva = distribucionIva.ObtenerImportesIva();
                var tasasIva = distribucionIva.ObtenerTasasIva();

                int registrosCreados = 0;

                // 3. Procesar cada porcentaje de IVA (equivalente a For KKII% = 0 To 5)
                // En nuestro caso procesamos índices 0 a 3 (0%, 10.5%, 21%, 27%)
                for (int indice = 0; indice < basesImponibles.Length; indice++)
                {
                    var baseImponible = basesImponibles[indice];
                    var importeIva = importesIva[indice];

                    // Validar si debe crear el registro (equivalente a: If XVALO(.IMPOGRA(KKII%)) >= 0.005 Or XVALO(.IMPOIVA(KKII%)) >= 0.005)
                    if (baseImponible >= 0.005m || importeIva >= 0.005m)
                    {
                        // Crear nuevo registro DETAIVA (equivalente a DETIVA.AddNew)
                        var nuevoDetalle = new Detaiva
                        {
                            // DETIVA!CODIEMPR = CodiEmp%
                            CodiEmpr = codigoEmpresa,
                            
                            // DETIVA!nume_ocom = NRO&
                            NumeOcom = numeroOrden,
                            
                            // DETIVA!NORIT = KKII%
                            Norit = indice,
                            
                            // DETIVA!Tasa_Iva = .TASIVA(KKII%)
                            TasaIva = tasasIva[indice],
                            
                            // DETIVA!Base_Impo = XVALO(.IMPOGRA(KKII%).TEXT)
                            BaseImpo = baseImponible,
                            
                            // DETIVA!Impo_Iva = XVALO(.IMPOIVA(KKII%))
                            ImpoIva = importeIva
                            
                            // NOTA: Los campos de auditoría (AddRecord, LastUpdate, NumUsuar, CodiEmprNet)
                            // se asignan automáticamente por el AuditInterceptor
                        };

                        _context.Detaivas.Add(nuevoDetalle);
                        registrosCreados++;

                        _logger.LogInformation($"Creado registro DETAIVA - Orden: {numeroOrden}, NORIT: {indice}, " +
                                             $"Tasa: {tasasIva[indice]:P2}, Base: {baseImponible:C2}, IVA: {importeIva:C2}");
                    }
                    else
                    {
                        _logger.LogDebug($"Omitido registro DETAIVA - Orden: {numeroOrden}, NORIT: {indice} " +
                                        $"(Base: {baseImponible:C2}, IVA: {importeIva:C2} - menores a 0.005)");
                    }
                }

                // 4. Guardar cambios (equivalente a DETIVA.Update del VB6)
                await _context.SaveChangesAsync();

                _logger.LogInformation($"Grabación de detalle IVA completada exitosamente. " +
                                     $"Orden: {numeroOrden}, Registros creados: {registrosCreados}");

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al grabar detalle IVA para orden {numeroOrden}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Obtiene el detalle de IVA de una orden de compra
        /// </summary>
        /// <param name="numeroOrden">Número de orden de compra</param>
        /// <returns>Lista de registros DETAIVA</returns>
        public async Task<List<Detaiva>> ObtenerDetalleIvaAsync(int numeroOrden)
        {
            try
            {
                var detalles = await _context.Detaivas
                    .Where(d => d.NumeOcom == numeroOrden)
                    .OrderBy(d => d.Norit)
                    .ToListAsync();

                _logger.LogInformation($"Obtenidos {detalles.Count} registros de detalle IVA para orden {numeroOrden}");
                return detalles;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener detalle IVA para orden {numeroOrden}: {ex.Message}");
                return new List<Detaiva>();
            }
        }

        /// <summary>
        /// Convierte una lista de registros DETAIVA a DistribucionIvaDTO
        /// Clasifica dinámicamente según las tasas de IVA encontradas en los datos
        /// </summary>
        /// <param name="detalles">Lista de registros DETAIVA</param>
        /// <returns>DTO con la distribución de IVA</returns>
        public DistribucionIvaDTO ConvertirADistribucionDto(List<Detaiva> detalles)
        {
            var distribucion = new DistribucionIvaDTO();

            _logger.LogInformation($"[ConvertirADistribucionDto] Procesando {detalles.Count} registros DETAIVA");

            // Agrupar por tasa de IVA para clasificación dinámica
            var gruposPorTasa = detalles
                .GroupBy(d => d.TasaIva ?? 0m)
                .OrderBy(g => g.Key)
                .ToList();

            _logger.LogInformation($"[ConvertirADistribucionDto] Encontradas {gruposPorTasa.Count} tasas diferentes de IVA");

            foreach (var grupo in gruposPorTasa)
            {
                var tasaIva = grupo.Key;
                var totalBaseImponible = grupo.Sum(d => d.BaseImpo ?? 0m);
                var totalImporteIva = grupo.Sum(d => d.ImpoIva ?? 0m);

                _logger.LogInformation($"[ConvertirADistribucionDto] Tasa {tasaIva:F2}%: Base {totalBaseImponible:C2}, IVA {totalImporteIva:C2} ({grupo.Count()} registros)");

                // Clasificar dinámicamente según la tasa encontrada
                // Usar un diccionario de tasas estándar para mapeo automático
                var tasasEstandar = new Dictionary<decimal, (Action<decimal> setBase, Action<decimal> setIva, string nombre)>
                {
                    [0m] = (val => distribucion.BaseImponible0 += val, val => { }, "EXENTO (0%)"),
                    [10.5m] = (val => distribucion.BaseImponible105 += val, val => distribucion.ImporteIva105 += val, "10.5%"),
                    [21m] = (val => distribucion.BaseImponible21 += val, val => distribucion.ImporteIva21 += val, "21%"),
                    [27m] = (val => distribucion.BaseImponible27 += val, val => distribucion.ImporteIva27 += val, "27%")
                };

                // Buscar la tasa más cercana con tolerancia de ±0.1
                var tasaEncontrada = tasasEstandar.Keys
                    .FirstOrDefault(tasa => Math.Abs(tasaIva - tasa) < 0.1m);

                if (tasasEstandar.ContainsKey(tasaEncontrada))
                {
                    var (setBase, setIva, nombre) = tasasEstandar[tasaEncontrada];
                    setBase(totalBaseImponible);
                    setIva(totalImporteIva);
                    _logger.LogDebug($"[ConvertirADistribucionDto] Tasa {tasaIva:F2}% clasificada como {nombre}");
                }
                else
                {
                    _logger.LogWarning($"[ConvertirADistribucionDto] Tasa de IVA no estándar encontrada: {tasaIva:F2}% - Clasificando como EXENTO");
                    // Para tasas no estándar, las agregamos como exentas
                    distribucion.BaseImponible0 += totalBaseImponible;
                }
            }

            _logger.LogInformation($"[ConvertirADistribucionDto] Distribución final - " +
                                 $"Exento: {distribucion.BaseImponible0:C2}, " +
                                 $"10.5%: {distribucion.BaseImponible105:C2} (IVA: {distribucion.ImporteIva105:C2}), " +
                                 $"21%: {distribucion.BaseImponible21:C2} (IVA: {distribucion.ImporteIva21:C2}), " +
                                 $"27%: {distribucion.BaseImponible27:C2} (IVA: {distribucion.ImporteIva27:C2})");

            return distribucion;
        }

        /// <summary>
        /// Elimina todos los registros DETAIVA de una orden de compra
        /// </summary>
        /// <param name="numeroOrden">Número de orden de compra</param>
        /// <returns>Cantidad de registros eliminados</returns>
        public async Task<int> EliminarDetalleIvaAsync(int numeroOrden)
        {
            try
            {
                var detalles = await _context.Detaivas
                    .Where(d => d.NumeOcom == numeroOrden)
                    .ToListAsync();

                if (detalles.Any())
                {
                    _context.Detaivas.RemoveRange(detalles);
                    await _context.SaveChangesAsync();
                    
                    _logger.LogInformation($"Eliminados {detalles.Count} registros DETAIVA para orden {numeroOrden}");
                    return detalles.Count;
                }

                return 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al eliminar detalle IVA para orden {numeroOrden}: {ex.Message}");
                throw;
            }
        }
    }
}
