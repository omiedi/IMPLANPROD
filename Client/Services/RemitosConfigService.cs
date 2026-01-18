using IMPLANPROD.Client.Repositories;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio para obtener la configuración de remitos
    /// </summary>
    public class RemitosConfigService
    {
        private readonly IRepository _repository;

        public RemitosConfigService(IRepository repository)
        {
            _repository = repository;
        }

        /// <summary>
        /// Obtiene el prefijo configurado para remitos desde FLEXCRL
        /// </summary>
        /// <returns>Prefijo formateado con 4 dígitos (ej: "0001")</returns>
        public async Task<string> ObtenerPrefijoRemitosAsync()
        {
            try
            {
                const string CLABUS_PARAM = ""; // CLABUS vacío según requerimiento
                const string PROBUS_PARAM_PREFIJO = "PREFIREM";

                // Usar "empty" como placeholder cuando CLABUS está vacío para evitar URLs malformadas
                string clabusParam = string.IsNullOrEmpty(CLABUS_PARAM) ? "empty" : CLABUS_PARAM;
                var response = await _repository.Get<FLEXCRL>($"api/flexcrl/{clabusParam}/{PROBUS_PARAM_PREFIJO}");

                if (response.Error || response.Response == null)
                {
                    // Valor por defecto si no existe configuración
                    Console.WriteLine("[RemitosConfigService] No existe configuración de prefijo, usando valor por defecto: 0");
                    return "0000";
                }

                var valorPrefijo = response.Response.VALCONTROL ?? "0";
                
                if (int.TryParse(valorPrefijo, out var prefijo))
                {
                    // Formatear con 4 dígitos completados con ceros a la izquierda
                    var prefijoFormateado = prefijo.ToString("D4");
                    Console.WriteLine($"[RemitosConfigService] Prefijo cargado: {prefijoFormateado}");
                    return prefijoFormateado;
                }
                else
                {
                    Console.WriteLine($"[RemitosConfigService] Error al parsear prefijo '{valorPrefijo}', usando valor por defecto");
                    return "0000";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RemitosConfigService] Error al obtener prefijo: {ex.Message}");
                return "0000"; // Valor por defecto en caso de error
            }
        }

        /// <summary>
        /// Obtiene el valor numérico del prefijo configurado
        /// </summary>
        /// <returns>Valor numérico del prefijo</returns>
        public async Task<int> ObtenerPrefijoRemitosNumericoAsync()
        {
            try
            {
                const string CLABUS_PARAM = ""; // CLABUS vacío según requerimiento
                const string PROBUS_PARAM_PREFIJO = "PREFIREM";

                // Usar "empty" como placeholder cuando CLABUS está vacío para evitar URLs malformadas
                string clabusParam = string.IsNullOrEmpty(CLABUS_PARAM) ? "empty" : CLABUS_PARAM;
                var response = await _repository.Get<FLEXCRL>($"api/flexcrl/{clabusParam}/{PROBUS_PARAM_PREFIJO}");

                if (response.Error || response.Response == null)
                {
                    return 0;
                }

                var valorPrefijo = response.Response.VALCONTROL ?? "0";

                if (int.TryParse(valorPrefijo, out var prefijo))
                {
                    return prefijo;
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RemitosConfigService] Error al obtener prefijo numérico: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Obtiene el tipo de numeración configurado para remitos
        /// </summary>
        /// <returns>Tipo de numeración: "AUTONUME" o "PREIMPRE"</returns>
        public async Task<string> ObtenerTipoNumeracionRemitosAsync()
        {
            try
            {
                const string CLABUS_PARAM_TIPONUMERACION = "NUREMIFA";
                const string PROBUS_PARAM_TIPONUMERACION = ""; // PROBUS vacío según requerimiento

                // Usar "empty" como placeholder cuando PROBUS está vacío para evitar URLs malformadas
                string probusParam = string.IsNullOrEmpty(PROBUS_PARAM_TIPONUMERACION) ? "empty" : PROBUS_PARAM_TIPONUMERACION;
                var response = await _repository.Get<FLEXCRL>($"api/flexcrl/{CLABUS_PARAM_TIPONUMERACION}/{probusParam}");

                if (response.Error || response.Response == null)
                {
                    // Valor por defecto si no existe configuración
                    Console.WriteLine("[RemitosConfigService] No existe configuración de tipo de numeración, usando valor por defecto: AUTONUME");
                    return "AUTONUME";
                }

                var tipoNumeracion = response.Response.VALCONTROL ?? "AUTONUME";
                
                // Validar que sea uno de los valores permitidos
                if (tipoNumeracion == "PREIMPRE" || tipoNumeracion == "AUTONUME")
                {
                    Console.WriteLine($"[RemitosConfigService] Tipo de numeración cargado: {tipoNumeracion}");
                    return tipoNumeracion;
                }
                else
                {
                    Console.WriteLine($"[RemitosConfigService] Valor inválido '{tipoNumeracion}', usando valor por defecto: AUTONUME");
                    return "AUTONUME";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RemitosConfigService] Error al obtener tipo de numeración: {ex.Message}");
                return "AUTONUME"; // Valor por defecto en caso de error
            }
        }

        /// <summary>
        /// Obtiene el último número de remito registrado en MOVISTO
        /// Filtra por fecha mayor o igual a la fecha de nueva numeración configurada
        /// </summary>
        /// <returns>Último número de remito</returns>
        public async Task<int> ObtenerUltimoNumeroRemitoAsync()
        {
            try
            {
                // Obtener la fecha de nueva numeración configurada
                var fechaNuevaNumeracion = await ObtenerFechaNuevaNumeracionAsync();
                
                // Formatear la fecha para la URL (ISO format)
                string fechaParam = fechaNuevaNumeracion.ToString("yyyy-MM-dd");
                
                var response = await _repository.Get<int>($"api/movistock/ultimo-numero-remito?fechaNuevaNumeracion={fechaParam}");
                
                if (response.Error)
                {
                    Console.WriteLine("[RemitosConfigService] Error al obtener último número de remito");
                    return 0;
                }

                Console.WriteLine($"[RemitosConfigService] Último número de remito desde {fechaNuevaNumeracion:dd/MM/yyyy}: {response.Response}");
                return response.Response;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RemitosConfigService] Error al obtener último número de remito: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Obtiene el código de movimiento configurado para remitos desde FLEXCRL
        /// </summary>
        /// <returns>Código de movimiento (ej: "RT")</returns>
        public async Task<string> ObtenerCodigoMovimientoAsync()
        {
            try
            {
                const string CLABUS_PARAM = ""; // CLABUS vacío según requerimiento
                const string PROBUS_PARAM_CODIGOMOVIMIENTO = "CODREFAC";

                // Usar "empty" como placeholder cuando CLABUS está vacío para evitar URLs malformadas
                string clabusParam = string.IsNullOrEmpty(CLABUS_PARAM) ? "empty" : CLABUS_PARAM;
                var response = await _repository.Get<FLEXCRL>($"api/flexcrl/{clabusParam}/{PROBUS_PARAM_CODIGOMOVIMIENTO}");

                if (response.Error || response.Response == null)
                {
                    // Valor por defecto si no existe configuración
                    Console.WriteLine("[RemitosConfigService] No existe configuración de código de movimiento, usando valor por defecto: RT");
                    return "RT";
                }

                var codigoMovimiento = response.Response.VALCONTROL ?? "RT";
                
                // Convertir a mayúsculas y limitar a 2 caracteres
                codigoMovimiento = codigoMovimiento.ToUpper().Trim();
                if (codigoMovimiento.Length > 2)
                {
                    codigoMovimiento = codigoMovimiento.Substring(0, 2);
                }

                Console.WriteLine($"[RemitosConfigService] Código de movimiento cargado: {codigoMovimiento}");
                return codigoMovimiento;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RemitosConfigService] Error al obtener código de movimiento: {ex.Message}");
                return "RT"; // Valor por defecto en caso de error
            }
        }

        /// <summary>
        /// Obtiene la fecha de nueva numeración de remitos desde FLEXCRL
        /// Esta fecha indica a partir de cuándo se debe considerar la numeración de remitos
        /// Si no existe configuración, retorna 01/01/1980 por defecto
        /// </summary>
        /// <returns>Fecha de nueva numeración</returns>
        public async Task<DateTime> ObtenerFechaNuevaNumeracionAsync()
        {
            try
            {
                const string CLABUS_PARAM = ""; // CLABUS vacío según requerimiento
                const string PROBUS_PARAM_FENUREFA = "FENUREFA";
                DateTime fechaPorDefecto = new DateTime(1980, 1, 1);

                // Usar "empty" como placeholder cuando CLABUS está vacío para evitar URLs malformadas
                string clabusParam = string.IsNullOrEmpty(CLABUS_PARAM) ? "empty" : CLABUS_PARAM;
                var response = await _repository.Get<FLEXCRL>($"api/flexcrl/{clabusParam}/{PROBUS_PARAM_FENUREFA}");

                if (response.Error || response.Response == null)
                {
                    // Valor por defecto si no existe configuración
                    Console.WriteLine($"[RemitosConfigService] No existe configuración FENUREFA, usando fecha por defecto: {fechaPorDefecto:dd/MM/yyyy}");
                    return fechaPorDefecto;
                }

                // Intentar parsear la fecha del VALCONTROL
                if (!string.IsNullOrEmpty(response.Response.VALCONTROL))
                {
                    if (DateTime.TryParse(response.Response.VALCONTROL, out DateTime fechaConfig))
                    {
                        Console.WriteLine($"[RemitosConfigService] Fecha de nueva numeración cargada: {fechaConfig:dd/MM/yyyy}");
                        return fechaConfig;
                    }
                }

                Console.WriteLine($"[RemitosConfigService] No se pudo parsear FENUREFA, usando fecha por defecto: {fechaPorDefecto:dd/MM/yyyy}");
                return fechaPorDefecto;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RemitosConfigService] Error al obtener fecha de nueva numeración: {ex.Message}");
                return new DateTime(1980, 1, 1); // Valor por defecto en caso de error
            }
        }

        /// <summary>
        /// Guarda la fecha de nueva numeración de remitos en FLEXCRL
        /// La fecha debe ser mayor o igual a la fecha actual
        /// </summary>
        /// <param name="fecha">Fecha a configurar</param>
        /// <returns>True si se guardó correctamente</returns>
        public async Task<bool> GuardarFechaNuevaNumeracionAsync(DateTime fecha)
        {
            try
            {
                // Validar que la fecha sea mayor o igual a hoy
                if (fecha.Date < DateTime.Now.Date)
                {
                    Console.WriteLine($"[RemitosConfigService] Error: La fecha {fecha:dd/MM/yyyy} es menor a la fecha actual");
                    return false;
                }

                const string CLABUS_PARAM = ""; // CLABUS vacío según requerimiento
                const string PROBUS_PARAM_FENUREFA = "FENUREFA";

                var flexcrl = new FLEXCRL
                {
                    CLABUS = CLABUS_PARAM,
                    PROBUS = PROBUS_PARAM_FENUREFA,
                    VALCONTROL = fecha.ToString("yyyy-MM-dd") // Formato ISO para la base de datos
                };

                // Usar "empty" como placeholder cuando CLABUS está vacío
                string clabusParam = string.IsNullOrEmpty(CLABUS_PARAM) ? "empty" : CLABUS_PARAM;
                var response = await _repository.Put($"api/flexcrl/{clabusParam}/{PROBUS_PARAM_FENUREFA}", flexcrl);

                if (!response.Error)
                {
                    Console.WriteLine($"[RemitosConfigService] Fecha de nueva numeración guardada: {fecha:dd/MM/yyyy}");
                    return true;
                }

                Console.WriteLine($"[RemitosConfigService] Error al guardar fecha de nueva numeración");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RemitosConfigService] Error al guardar fecha de nueva numeración: {ex.Message}");
                return false;
            }
        }
    }
}
