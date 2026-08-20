using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para manejar las operaciones relacionadas con la tabla FLEXCRL
    /// Permite gestionar controles de procesos de negocio
    /// </summary>
    public class FlexcrlService
    {
        private readonly DataContext _context;
        private readonly ILogger<FlexcrlService> _logger;

        /// <summary>
        /// Constructor del servicio FlexcrlService
        /// </summary>
        /// <param name="context">Contexto de base de datos</param>
        /// <param name="logger">Logger para registrar eventos y errores</param>
        public FlexcrlService(DataContext context, ILogger<FlexcrlService> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Actualiza o crea un registro de control de proceso de negocio
        /// Valida si existe la combinación CLABUS + PROBUS:
        /// - Si existe: actualiza el VALCONTROL
        /// - Si no existe: crea un nuevo registro completo
        /// </summary>
        /// <param name="clabus">Clasificación del proceso de negocio</param>
        /// <param name="probus">Proceso de negocio</param>
        /// <param name="valcontrol">Valor de control del proceso</param>
        /// <returns>True si la operación fue exitosa, False en caso contrario</returns>
        public async Task<bool> ActualizarOCrearControlAsync(string clabus, string probus, string valcontrol)
        {
            try
            {
                // Validar parámetros de entrada
                // CLABUS y PROBUS pueden estar vacíos según la lógica VB6 original
                // Solo validamos que no sean null para evitar errores de referencia
                if (clabus == null)
                {
                    clabus = ""; // Convertir null a string vacío
                }

                if (probus == null)
                {
                    probus = ""; // Convertir null a string vacío
                }

                _logger.LogInformation("Procesando control para CLABUS: {Clabus}, PROBUS: {Probus}", clabus, probus);

                // Buscar si existe un registro con la combinación CLABUS + PROBUS
                var registroExistente = await _context.Set<FLEXCRL>()
                    .FirstOrDefaultAsync(f => f.CLABUS == clabus && f.PROBUS == probus);

                if (registroExistente != null)
                {
                    // El registro existe - actualizar VALCONTROL
                    _logger.LogInformation("Registro existente encontrado con ID: {Id}. Actualizando VALCONTROL", registroExistente.ID);
                    
                    registroExistente.VALCONTROL = valcontrol;
                    
                    _context.Set<FLEXCRL>().Update(registroExistente);
                    await _context.SaveChangesAsync();
                    
                    _logger.LogInformation("Registro ID: {Id} actualizado exitosamente", registroExistente.ID);
                    return true;
                }
                else
                {
                    // El registro no existe - crear nuevo registro completo
                    _logger.LogInformation("No se encontró registro existente. Creando nuevo registro");
                    
                    var nuevoRegistro = new FLEXCRL
                    {
                        CLABUS = clabus,
                        PROBUS = probus,
                        VALCONTROL = valcontrol,
                        CodiEmpr = null // Se puede asignar según la lógica de negocio requerida
                    };

                    _context.Set<FLEXCRL>().Add(nuevoRegistro);
                    await _context.SaveChangesAsync();
                    
                    _logger.LogInformation("Nuevo registro creado exitosamente con ID: {Id}", nuevoRegistro.ID);
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar control para CLABUS: {Clabus}, PROBUS: {Probus}", clabus, probus);
                throw new Exception($"Error al procesar control para CLABUS: {clabus}, PROBUS: {probus}: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Obtiene un registro de control específico por CLABUS y PROBUS
        /// </summary>
        /// <param name="clabus">Clasificación del proceso de negocio</param>
        /// <param name="probus">Proceso de negocio</param>
        /// <returns>Registro FLEXCRL si existe, null si no existe</returns>
        public async Task<FLEXCRL?> ObtenerControlAsync(string clabus, string probus)
        {
            try
            {
                var registro = await _context.Set<FLEXCRL>()
                    .AsNoTracking()
                    .FirstOrDefaultAsync(f => f.CLABUS == clabus && f.PROBUS == probus);

                return registro;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener control para CLABUS: {Clabus}, PROBUS: {Probus}", clabus, probus);
                throw new Exception($"Error al obtener control para CLABUS: {clabus}, PROBUS: {probus}: {ex.Message}", ex);
            }
        }

        // TODO: FlexcrlService CÓMO USAR ESTE SERVICIO FlexcrlService
        // 
        // DESDE EL FRONTEND (HTTP):
        // ========================
        // Para actualizar o crear un control desde el frontend, se debe hacer una llamada HTTP PUT al endpoint:
        // PUT /api/flexcrl
        // Body (JSON):
        // {
        //     "clabus": "CLASIFICACION_PROCESO",
        //     "probus": "NOMBRE_PROCESO", 
        //     "valcontrol": "VALOR_DE_CONTROL"
        // }
        //
        // Para obtener un control específico:
        // GET /api/flexcrl/{clabus}/{probus}
        //
        // DESDE UN CONTROLADOR (Inyección de dependencia):
        // ===============================================
        // 1. Inyectar el servicio en el constructor del controlador:
        //    private readonly FlexcrlService _flexcrlService;
        //    
        //    public MiControlador(FlexcrlService flexcrlService)
        //    {
        //        _flexcrlService = flexcrlService;
        //    }
        //
        // 2. Usar el servicio en los métodos del controlador:
        //    // Para actualizar o crear:
        //    var resultado = await _flexcrlService.ActualizarOCrearControlAsync("MI_CLABUS", "MI_PROBUS", "MI_VALOR");
        //    
        //    // Para obtener:
        //    var control = await _flexcrlService.ObtenerControlAsync("MI_CLABUS", "MI_PROBUS");
        //
        // CUÁNDO UTILIZAR:
        // ================
        // - Para gestionar configuraciones dinámicas de procesos de negocio
        // - Para almacenar parámetros de control que pueden cambiar en tiempo de ejecución
        // - Para mantener estados o valores de configuración específicos por proceso
        // - Para implementar sistemas de configuración flexible sin hardcodear valores
    }
}
