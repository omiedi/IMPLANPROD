using Azure;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/stock")]
    public class StockController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<StockController> _logger;

        // Almacenamiento temporal para los depósitos seleccionados (en producción debería usar una base de datos o caché)
        private static string _depositosSeleccionados = string.Empty;

        public StockController(DataContext context, ILogger<StockController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Guarda los depósitos seleccionados para consulta de stock
        /// </summary>
        /// <param name="depositosString">String con los IDs de los depósitos separados por coma</param>
        /// <returns>OK si se guardó correctamente</returns>
        [HttpPost("guardardepositos")]
        public ActionResult GuardarDepositosSeleccionados([FromBody] string depositosString)
        {
            try
            {
                _depositosSeleccionados = depositosString;
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al guardar los depósitos seleccionados: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene los depósitos seleccionados para consulta de stock
        /// </summary>
        /// <returns>String con los IDs de los depósitos separados por coma</returns>
        [HttpGet("depositosseleccionados")]
        public ActionResult<string> GetDepositosSeleccionados()
        {
            // Si no hay depósitos seleccionados, devolver una lista vacía en lugar de una cadena vacía
            // para evitar problemas al procesar la respuesta en el cliente
            if (string.IsNullOrEmpty(_depositosSeleccionados))
            {
                return Ok("[]");
            }
            else
            {
                return Ok(_depositosSeleccionados);
            }
        }

        /// <summary>
        /// Obtiene el stock de un producto en un depósito específico
        /// </summary>
        /// <param name="codint">Código interno del producto</param>
        /// <param name="deposito">Código del depósito</param>
        /// <returns>Stock del producto en el depósito</returns>
        [HttpGet("stockdeposito/{codint:int}/{deposito:int}")]
        public async Task<ActionResult<decimal>> GetStockDeposito(int codint, int deposito)
        {
            try
            {
                var stock = await _context.Movistos
                    .Where(m => m.Cod_Inte == codint && m.DepoMovi == deposito)
                    .SumAsync(m => (m.CantIngre ?? 0) - (m.CantSalid ?? 0));

                return Ok(stock);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener el stock del producto {codint} en el depósito {deposito}: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene la estructura multinivel de un producto con sus componentes
        /// </summary>
        /// <param name="codint">Código interno del producto</param>
        /// <param name="niveles">Niveles de profundidad (0 = todos)</param>
        /// <param name="cantidad">Cantidad a simular</param>
        /// <returns>Lista con la estructura multinivel</returns>
        [HttpGet("estructuramultinivel/{codint:int}")]
        public async Task<ActionResult<List<DetalleStockDTO>>> GetEstructuraMultinivel(int codint, int niveles = 0, int cantidad = 1)
        {
            try
            {
                var resultado = new List<DetalleStockDTO>();
                
                // Obtener el producto principal
                var producto = await _context.masters.FirstOrDefaultAsync(m => m.Codint == codint);
                if (producto == null)
                {
                    return NotFound($"No se encontró el producto con código {codint}");
                }
                
                // Agregar el producto principal
                resultado.Add(new DetalleStockDTO
                {
                    CI1 = producto.Codint,
                    CodXEelem = producto.Codigo ?? string.Empty,
                    Descripcion = producto.Descripcion ?? string.Empty,
                    UM = producto.Umedida ?? string.Empty,
                    Uso = 1,
                    CantSimulada = cantidad,
                    StAcumulado = await ObtenerStockTotal(producto.Codint)
                });
                
                // Si se solicitan componentes (niveles > 0 o niveles = 0 que significa todos)
                if (niveles != 1)
                {
                    // Calcular los niveles máximos a procesar
                    int nivelesMaximos = niveles;
                    if (niveles > 0)
                    {
                        nivelesMaximos = niveles - 1;
                    }
                    
                    // Cargar los componentes de forma recursiva
                    await CargarComponentesRecursivos(resultado, codint, 1, nivelesMaximos, cantidad, new List<bool>());
                }
                
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la estructura multinivel");
                return BadRequest($"Error al obtener la estructura multinivel: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Carga los componentes de un producto de forma recursiva
        /// </summary>
        /// <param name="resultado">Lista de resultados donde se agregarán los componentes</param>
        /// <param name="codigoConjunto">Código interno del conjunto</param>
        /// <param name="nivelActual">Nivel actual de recursión</param>
        /// <param name="nivelesMaximos">Niveles máximos a procesar (0 = todos)</param>
        /// <param name="cantidadPadre">Cantidad del componente padre</param>
        private async Task CargarComponentesRecursivos(List<DetalleStockDTO> resultado, int codigoConjunto, int nivelActual, int nivelesMaximos, decimal cantidadPadre, List<bool> prefijos)
        {
            // Obtener los componentes del conjunto actual desde la tabla Formulas
            var componentes = await _context.Formulas
                .Where(f => f.CodIConj == codigoConjunto)
                .OrderBy(f => f.NuorIt)
                .ToListAsync();
                
            foreach (var componente in componentes)
            {
                // Verificar que el código del elemento exista
                if (componente.CodiElem == null)
                {
                    continue;
                }
                
                // Obtener información del master para este componente
                var masterComponente = await _context.masters.FirstOrDefaultAsync(m => m.Codint == componente.CodiElem);
                if (masterComponente == null)
                {
                    continue;
                }
                
                // Calcular uso y cantidad simulada
                decimal usoBruto = componente.UsoBruto ?? 0;
                decimal cantidadSimulada = cantidadPadre * usoBruto;
                
                // Crear indentación simple de 2 espacios por nivel
                string indentacion = new string('-', 2 * nivelActual);
                
                // Agregar el componente al resultado
                resultado.Add(new DetalleStockDTO
                {
                    CI1 = masterComponente.Codint,
                    CodXEelem = masterComponente.Codigo ?? string.Empty,
                    Descripcion = indentacion + (masterComponente.Descripcion ?? string.Empty),
                    UM = masterComponente.Umedida ?? string.Empty,
                    Uso = usoBruto,
                    CantSimulada = cantidadSimulada,
                    StAcumulado = await ObtenerStockTotal(masterComponente.Codint)
                });
                
                // Procesar subcomponentes si no hemos alcanzado el nivel máximo
                if (nivelesMaximos == 0 || nivelActual < nivelesMaximos)
                {
                    await CargarComponentesRecursivos(resultado, masterComponente.Codint, nivelActual + 1, nivelesMaximos, cantidadSimulada, prefijos);
                }
            }
        }
        
        /// <summary>
        /// Obtiene el stock disponible total de un producto (suma de todos los depósitos)
        /// </summary>
        /// <param name="codigoInterno">Código interno del producto</param>
        /// <returns>Stock disponible total</returns>
        [HttpGet("disponible/{codigoInterno:int}")]
        public async Task<decimal> GetStockDisponible(int codigoInterno)
        {
            try
            {
                decimal StTotalacumulado = await ObtenerStockTotal(codigoInterno);
                //var stock = await _context.Movistos
                //    .Where(m => m.Cod_Inte == codigoInterno)
                //    .SumAsync(m => (m.CantIngre ?? 0) - (m.CantSalid ?? 0));

                //_logger.LogInformation($"Stock disponible para producto {codigoInterno}: {stock}");
                return StTotalacumulado;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener stock disponible para producto {codigoInterno}");
                return 0;
            }
        }

        /// <summary>
        /// Obtiene los lotes con saldo disponible para un producto en un depósito específico
        /// </summary>
        /// <param name="codint">Código interno del producto</param>
        /// <param name="depositoId">ID del depósito</param>
        /// <returns>Lista de lotes con saldo disponible</returns>
        [HttpGet("lotes/{codint}/{depositoId}")]
        public async Task<ActionResult<List<LoteDisponibleDTO>>> GetLotesDisponibles(int codint, int depositoId)
        {
            try
            {
                _logger.LogInformation($"[StockController] Obteniendo lotes para CodInt: {codint}, Depósito: {depositoId}");

                // Obtener los lotes con saldo del producto en el depósito especificado
                var lotes = await _context.Movistos
                    .Where(m => m.Cod_Inte == codint && m.DepoMovi == depositoId)
                    .GroupBy(m => m.IdenLote)
                    .Select(g => new LoteDisponibleDTO
                    {
                        Lote = g.Key ?? string.Empty,
                        Saldo = g.Sum(m => (m.CantIngre ?? 0) - (m.CantSalid ?? 0))
                    })
                    .Where(l => l.Saldo > 0 && !string.IsNullOrEmpty(l.Lote))
                    .OrderBy(l => l.Lote)
                    .ToListAsync();

                _logger.LogInformation($"[StockController] Lotes encontrados: {lotes.Count}");

                return Ok(lotes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[StockController] Error al obtener lotes para CodInt: {codint}, Depósito: {depositoId}");
                return BadRequest($"Error al obtener lotes: {ex.Message}");
            }
        }

        /// <summary>
        /// Método auxiliar para obtener el stock total de un producto
        /// </summary>
        private async Task<decimal> ObtenerStockTotal(int codint)
        {
            return await _context.Movistos
                .Where(m => m.Cod_Inte == codint)
                .SumAsync(m => (m.CantIngre ?? 0) - (m.CantSalid ?? 0));
        }
    }
}
