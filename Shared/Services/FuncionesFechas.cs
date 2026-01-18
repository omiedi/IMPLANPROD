using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IMPLANPROD.Shared.Services

{
    public class FuncionesFechas
    {

        // Función que suma dos números
        public DateTime FechaEstimadadeReparacion(DateTime fecha, decimal frecuencia, decimal frecUso, decimal factorDias)
        {
            // Se valida que si faltan datos devuelve la fecha de hoy
            if ((frecuencia + frecUso + factorDias) < 1)
            {
                return DateTime.Now;
            }

            if (factorDias < 1) // Para evitar división por 0
            {
                return DateTime.Now;
            }

            // Calcular los días a sumar
            decimal diasASumar = (frecuencia * frecUso) / factorDias;
            // Sumar los días a la fecha proporcionada
            DateTime fechaEstimadadeReparacion = fecha.AddDays((int)diasASumar);

            // Ajustar la fecha si cae en sábado o domingo
            if (fechaEstimadadeReparacion.DayOfWeek == DayOfWeek.Saturday)
            {
                fechaEstimadadeReparacion = fechaEstimadadeReparacion.AddDays(2); // Pasar al lunes
            }
            else if (fechaEstimadadeReparacion.DayOfWeek == DayOfWeek.Sunday)
            {
                fechaEstimadadeReparacion = fechaEstimadadeReparacion.AddDays(1); // Pasar al lunes
            }
            return fechaEstimadadeReparacion;
        }


    }
}
// TODO EJEMPLO DE COMO USAR UNA FUNCION COMO ESTA EN EL CONTROLADOR
// Controllers/CalculatorController.cs
//using Microsoft.AspNetCore.Mvc;
//using TuProyecto.Services;

//namespace TuProyecto.Controllers
//{
//    [ApiController]
//    [Route("api/[controller]")]
//    public class CalculatorController : ControllerBase
//    {
//        private readonly CalculatorService _calculatorService;

//        public CalculatorController(CalculatorService calculatorService)
//        {
//            _calculatorService = calculatorService;
//        }

//        [HttpGet("sumar")]
//        public ActionResult<double> Sumar(double num1, double num2)
//        {
//            var resultado = _calculatorService.Sumar(num1, num2);
//            return Ok(resultado);
//        }
//    }
//}

// Y ESTO EN PROGRAM 
//builder.Services.AddScoped<FuncionesFechas>();

// Y SI ES EN .RAZOR 
// SE INJECTA @inject FuncionesFechas funcionesFechas
// Y SE USA NORMALMENTE

//FIN EJEMPLO