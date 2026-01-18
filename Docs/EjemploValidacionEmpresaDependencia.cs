using IMPLANPROD.Server.Services;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// EJEMPLO DE USO: Validaciones por empresa y dependencia
    /// Este archivo muestra cómo implementar validaciones específicas por empresa/dependencia
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class EjemploValidacionController : ControllerBase
    {
        private readonly IEmpresaDependenciaService _empresaDependenciaService;

        public EjemploValidacionController(IEmpresaDependenciaService empresaDependenciaService)
        {
            _empresaDependenciaService = empresaDependenciaService;
        }

        // ============================================================
        // EJEMPLO 1: Validación simple por empresa
        // ============================================================
        [HttpGet("ejemplo1")]
        public IActionResult Ejemplo1_ValidacionPorEmpresa()
        {
            // Validar si el usuario pertenece a ARCON
            if (_empresaDependenciaService.EsEmpresa("ARCON"))
            {
                // Lógica específica para ARCON
                return Ok("Usuario de ARCON - Lógica especial aplicada");
            }
            else
            {
                // Lógica para otras empresas
                return Ok("Usuario de otra empresa - Lógica estándar");
            }
        }

        // ============================================================
        // EJEMPLO 2: Validación por empresa Y dependencia (TU CASO)
        // ============================================================
        [HttpGet("ejemplo2")]
        public IActionResult Ejemplo2_ValidacionEmpresaYDependencia()
        {
            // TU CASO: Validar si es ARCON y NEUQUEN
            if (_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
            {
                // Lógica específica para ARCON - NEUQUEN
                return Ok(new
                {
                    mensaje = "Usuario de ARCON - NEUQUEN",
                    empresa = _empresaDependenciaService.ObtenerNombreEmpresaActual(),
                    dependencia = _empresaDependenciaService.ObtenerNombreDependenciaActual(),
                    aplicarLogicaEspecial = true
                });
            }
            else
            {
                // Lógica para otros casos
                return Ok(new
                {
                    mensaje = "Usuario de otra empresa/dependencia",
                    empresa = _empresaDependenciaService.ObtenerNombreEmpresaActual(),
                    dependencia = _empresaDependenciaService.ObtenerNombreDependenciaActual(),
                    aplicarLogicaEspecial = false
                });
            }
        }

        // ============================================================
        // EJEMPLO 3: Validación múltiple con switch
        // ============================================================
        [HttpPost("procesar-orden")]
        public IActionResult Ejemplo3_ValidacionMultiple()
        {
            var empresaActual = _empresaDependenciaService.ObtenerNombreEmpresaActual();
            var dependenciaActual = _empresaDependenciaService.ObtenerNombreDependenciaActual();

            // Validación múltiple según empresa y dependencia
            if (_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
            {
                // Caso específico: ARCON - NEUQUEN
                return Ok("Procesando orden con reglas ARCON-NEUQUEN");
            }
            else if (_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "BUENOS AIRES"))
            {
                // Caso específico: ARCON - BUENOS AIRES
                return Ok("Procesando orden con reglas ARCON-BUENOS AIRES");
            }
            else if (_empresaDependenciaService.EsEmpresa("ARCON"))
            {
                // Cualquier otra dependencia de ARCON
                return Ok("Procesando orden con reglas generales de ARCON");
            }
            else
            {
                // Otras empresas
                return Ok("Procesando orden con reglas estándar");
            }
        }

        // ============================================================
        // EJEMPLO 4: Obtener información para logging
        // ============================================================
        [HttpGet("info-usuario")]
        public IActionResult Ejemplo4_InformacionUsuario()
        {
            return Ok(new
            {
                empresa = _empresaDependenciaService.ObtenerNombreEmpresaActual(),
                dependencia = _empresaDependenciaService.ObtenerNombreDependenciaActual(),
                empresaId = _empresaDependenciaService.ObtenerEmpresaIdActual(),
                dependenciaId = _empresaDependenciaService.ObtenerDependenciaIdActual()
            });
        }

        // ============================================================
        // EJEMPLO 5: Validación con lógica compleja
        // ============================================================
        [HttpPost("calcular-precio")]
        public IActionResult Ejemplo5_CalculoPrecioConValidacion([FromBody] decimal precioBase)
        {
            decimal precioFinal = precioBase;
            string razonDescuento = "Sin descuento";

            // Aplicar descuentos según empresa/dependencia
            if (_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
            {
                // Descuento especial para ARCON-NEUQUEN: 15%
                precioFinal = precioBase * 0.85m;
                razonDescuento = "Descuento ARCON-NEUQUEN: 15%";
            }
            else if (_empresaDependenciaService.EsEmpresa("ARCON"))
            {
                // Descuento general ARCON: 10%
                precioFinal = precioBase * 0.90m;
                razonDescuento = "Descuento ARCON general: 10%";
            }
            else if (_empresaDependenciaService.EsDependencia("NEUQUEN"))
            {
                // Descuento por región Neuquén: 5%
                precioFinal = precioBase * 0.95m;
                razonDescuento = "Descuento región Neuquén: 5%";
            }

            return Ok(new
            {
                precioBase,
                precioFinal,
                descuento = precioBase - precioFinal,
                razonDescuento,
                empresa = _empresaDependenciaService.ObtenerNombreEmpresaActual(),
                dependencia = _empresaDependenciaService.ObtenerNombreDependenciaActual()
            });
        }

        // ============================================================
        // EJEMPLO 6: Validación con excepción personalizada
        // ============================================================
        [HttpPost("operacion-restringida")]
        public IActionResult Ejemplo6_OperacionRestringida()
        {
            // Solo permitir esta operación para ARCON-NEUQUEN
            if (!_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
            {
                return Forbid("Esta operación solo está disponible para ARCON-NEUQUEN");
            }

            // Ejecutar operación restringida
            return Ok("Operación ejecutada exitosamente");
        }

        // ============================================================
        // EJEMPLO 7: Validación en servicios (no controladores)
        // ============================================================
        /*
        public class MiServicio
        {
            private readonly IEmpresaDependenciaService _empresaDependenciaService;

            public MiServicio(IEmpresaDependenciaService empresaDependenciaService)
            {
                _empresaDependenciaService = empresaDependenciaService;
            }

            public void ProcesarOrdenCompra(OrdenCompra orden)
            {
                // Validar empresa/dependencia en el servicio
                if (_empresaDependenciaService.EsEmpresaYDependencia("ARCON", "NEUQUEN"))
                {
                    // Lógica especial
                    AplicarReglasEspecialesArconNeuquen(orden);
                }
                else
                {
                    // Lógica estándar
                    AplicarReglasEstandar(orden);
                }
            }
        }
        */
    }
}

/*
═══════════════════════════════════════════════════════════════════════════
RESUMEN DE MÉTODOS DISPONIBLES:
═══════════════════════════════════════════════════════════════════════════

1. EsEmpresa(string nombreEmpresa)
   - Valida si el usuario pertenece a una empresa específica
   - Ejemplo: EsEmpresa("ARCON")

2. EsDependencia(string nombreDependencia)
   - Valida si el usuario pertenece a una dependencia específica
   - Ejemplo: EsDependencia("NEUQUEN")

3. EsEmpresaYDependencia(string nombreEmpresa, string nombreDependencia)
   - Valida ambos criterios simultáneamente
   - Ejemplo: EsEmpresaYDependencia("ARCON", "NEUQUEN")

4. ObtenerNombreEmpresaActual()
   - Retorna el nombre de la empresa del usuario actual
   - Ejemplo: "ARCON"

5. ObtenerNombreDependenciaActual()
   - Retorna el nombre de la dependencia del usuario actual
   - Ejemplo: "NEUQUEN"

6. ObtenerEmpresaIdActual()
   - Retorna el ID de la empresa del usuario actual
   - Ejemplo: 1

7. ObtenerDependenciaIdActual()
   - Retorna el ID de la dependencia del usuario actual
   - Ejemplo: 5

═══════════════════════════════════════════════════════════════════════════
NOTAS IMPORTANTES:
═══════════════════════════════════════════════════════════════════════════

- Las comparaciones de nombres son case-insensitive (no distinguen mayúsculas/minúsculas)
- Los métodos retornan null si el usuario no está autenticado
- Puedes inyectar IEmpresaDependenciaService en cualquier controlador o servicio
- El servicio usa el contexto HTTP actual para obtener el usuario autenticado

═══════════════════════════════════════════════════════════════════════════
*/
