using DocumentFormat.OpenXml.Office2010.Excel;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la gestión de clientes (Deudor12)
    /// </summary>
    [ApiController]
    [Route("/api/deudor12")]
    public class Deudor12Controller : ControllerBase
    {
        private readonly DataContext _context;

        public Deudor12Controller(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene una lista paginada de clientes (Deudor12) con filtro por razón social y status (Stat_Cli)
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, short stat_Cli = 0)
        {
            var queryable = _context.Clientes.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Raso_Cli != null && x.Raso_Cli.ToLower().Contains(pagination.Filter.ToLower()));
            }
            if (stat_Cli >= 0 )
            {
                queryable = queryable.Where(x => x.Stat_Cli >= stat_Cli || x.Stat_Cli == null); // Por defecto activos
            }
            else
            {
                queryable = queryable.Where(x => x.Stat_Cli == stat_Cli);
            }
            return Ok(await queryable
                .OrderBy(x => x.Raso_Cli)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Devuelve el total de páginas para la paginación de clientes, considerando el filtro y el status
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, short stat_Cli = 0)
        {
            var queryable = _context.Clientes.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Raso_Cli != null && x.Raso_Cli.ToLower().Contains(pagination.Filter.ToLower()));
            }
            if (stat_Cli >= 0 )
            {
                queryable = queryable.Where(x => x.Stat_Cli >= stat_Cli || x.Stat_Cli == null);
            }
            else
            {
                queryable = queryable.Where(x => x.Stat_Cli == stat_Cli);
            }
            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene un cliente individual por Id
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var cliente = await _context.Clientes.FirstOrDefaultAsync(x => x.Id == id);
            if (cliente == null)
            {
                return NotFound();
            }
            return Ok(cliente);
        }

        // En el controlador de clientes
        [HttpGet("datos-impresion/{nroCliente}")]
        public async Task<ActionResult<DatosClienteImpresionDTO>> GetDatosClienteImpresion(int nroCliente)
        {
            var cliente = await _context.Clientes.FirstOrDefaultAsync(x => x.Nume_Cli == nroCliente);
            if (cliente == null)
            {
                return NotFound();
            }
            return Ok(cliente);
        }
        /// <summary>
        /// Crea un nuevo cliente
        /// </summary>
        [HttpPost]
        public async Task<IActionResult> PostAsync(Deudor12 deudor)
        {
            // Iniciar una transacción para asegurar la atomicidad de la operación
            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    // Si nume_cli llega vacío o null, asignar el siguiente valor disponible
                    if (deudor.Nume_Cli == 0)
                    {
                        // Buscar el valor máximo actual de nume_cli y sumarle 1
                        int maxNumeCli = await _context.Clientes.MaxAsync(x => (int?)x.Nume_Cli) ?? 0;
                        deudor.Nume_Cli = maxNumeCli + 1;
                    }

                    _context.Add(deudor);
                    await _context.SaveChangesAsync();

                    // Confirmar la transacción si todo salió bien
                    await transaction.CommitAsync();

                    // Devuelve una respuesta HTTP 201 (Created) indicando que el recurso fue creado exitosamente.
                    // Además, incluye en la cabecera 'Location' la URL para consultar el recurso recién creado (usando el método GetAsync y el id asignado).
                    // El cuerpo de la respuesta contiene el objeto creado (deudor) con todos sus datos, incluyendo el Id generado por la base de datos.
                    return CreatedAtAction(nameof(GetAsync), new { id = deudor.Id }, deudor);
                }
                catch (Exception ex)
                {
                    // Si ocurre algún error, deshacer la transacción
                    await transaction.RollbackAsync();

                    // Registrar el error (opcional: puedes agregar logging aquí)

                    // Devolver un error 500 con el mensaje de la excepción
                    return StatusCode(500, $"Ocurrió un error al crear el cliente: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Obtiene un cliente por número de cliente (Nume_Cli)
        /// </summary>
        [HttpGet("numero/{numeroCliente:int}")]
        public async Task<IActionResult> GetByNumeroClienteAsync(int numeroCliente)
        {
            var cliente = await _context.Clientes.FirstOrDefaultAsync(x => x.Nume_Cli == numeroCliente);
            if (cliente == null)
            {
                return NotFound();
            }
            return Ok(cliente);
        }

        /// <summary>
        /// Obtiene lista de clientes para combo (solo activos con Status >= 0)
        /// </summary>
        [HttpGet("combo")]
        public async Task<IActionResult> GetComboAsync()
        {
            var clientes = await _context.Clientes
                .Where(x => x.Stat_Cli >= 0 || x.Stat_Cli == null)
                .Select(x => new { 
                    Nume_Cli = x.Nume_Cli, 
                    Raso_Cli = x.Raso_Cli 
                })
                .OrderBy(x => x.Raso_Cli)
                .ToListAsync();
            
            return Ok(clientes);
        }

        /// <summary>
        /// Actualiza un cliente existente
        /// </summary>
        [HttpPut]
        public async Task<IActionResult> PutAsync(Deudor12 deudor)
        {
            _context.Update(deudor);
            await _context.SaveChangesAsync();
            return Ok(deudor);
        }
       

        
    }
}
