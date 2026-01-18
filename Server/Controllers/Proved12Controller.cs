using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Data;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/Proved12")]
    public class Proved12Controller : ControllerBase
    {
        private readonly DataContext _context;

        public Proved12Controller(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de proveedores para el combo (solo activos)
        /// </summary>
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Proveedores
                .Where(d => d.Stat_Cli >= 0)
                .OrderBy(d => d.Raso_Cli)
                .ToListAsync());
        }

        //// GET: api/Proved12
        //[HttpGet]
        //public async Task<ActionResult<IEnumerable<Proved12>>> GetAll()
        //{
        //    return await _context.Proveedores.ToListAsync();
        //}
        /// <summary>
        /// Obtiene una lista paginada de clientes (Proved12) con filtro por razón social y status (Stat_Cli)
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, short stat_Cli = 0)
        {
            var queryable = _context.Proveedores.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Raso_Cli.ToLower().Contains(pagination.Filter.ToLower()));
            }
            if (stat_Cli == 0 || stat_Cli == 1)
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
            var queryable = _context.Proveedores.AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.Raso_Cli.ToLower().Contains(pagination.Filter.ToLower()));
            }
            if (stat_Cli == 0 || stat_Cli == 1)
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

        // GET: api/Proved12/5
        [HttpGet("{id}")]
        public async Task<ActionResult<Proved12>> Get(int id)
        {
            var entity = await _context.Proveedores.FirstOrDefaultAsync(x => x.Id == id);
            if (entity == null)
                return NotFound();
            return entity;
        }

        /// <summary>
        /// Obtiene un proveedor por su número de cliente (Nume_Cli)
        /// </summary>
        [HttpGet("numero/{numeCli}")]
        public async Task<ActionResult<Proved12>> GetByNumeCli(int numeCli)
        {
            var entity = await _context.Proveedores.FirstOrDefaultAsync(x => x.Nume_Cli == numeCli);
            if (entity == null)
                return NotFound();
            return entity;
        }

        // POST: api/Proved12
        [HttpPost]
        public async Task<ActionResult<Proved12>> Post(Proved12 model)
        {
            // Iniciar una transacción para asegurar la atomicidad de la operación
            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    // Si nume_cli llega vacío o null, asignar el siguiente valor disponible
                    if (model.Nume_Cli == 0)
                    {
                        // Buscar el valor máximo actual de nume_cli y sumarle 1
                        int maxNumeCli = await _context.Proveedores.MaxAsync(x => (int?)x.Nume_Cli) ?? 0;
                        model.Nume_Cli = maxNumeCli + 1;
                        model.Stat_Cli = 0;//si es nuevo proveedor 
                    }

                    _context.Proveedores.Add(model);
                    await _context.SaveChangesAsync();

                    // Confirmar la transacción si todo salió bien
                    await transaction.CommitAsync();

                    // Devuelve una respuesta HTTP 201 (Created) indicando que el recurso fue creado exitosamente.
                    // Además, incluye en la cabecera 'Location' la URL para consultar el recurso recién creado (usando el método Get y el id asignado).
                    // El cuerpo de la respuesta contiene el objeto creado (model) con todos sus datos, incluyendo el Id generado por la base de datos.
                    return CreatedAtAction(nameof(Get), new { id = model.Id }, model);
                }
                catch (Exception ex)
                {
                    // Si ocurre algún error, deshacer la transacción
                    await transaction.RollbackAsync();

                    // Registrar el error (opcional: puedes agregar logging aquí)

                    // Devolver un error 500 con el mensaje de la excepción
                    return StatusCode(500, $"Ocurrió un error al crear el proveedor: {ex.Message}");
                }
            }
        }

        // PUT: api/Proved12/5
        [HttpPut("{id}")]
        public async Task<IActionResult> Put(int id, Proved12 model)
        {
            if (id != model.Id)
                return BadRequest();
            _context.Entry(model).State = EntityState.Modified;
            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!_context.Proveedores.Any(e => e.Id == id))
                    return NotFound();
                else
                    throw;
            }
            return NoContent();
        }

        // DELETE: api/Proved12/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(int id)
        {
            var entity = await _context.Proveedores.FirstOrDefaultAsync(x => x.Id == id);
            if (entity == null)
                return NotFound();
            _context.Proveedores.Remove(entity);
            await _context.SaveChangesAsync();
            return NoContent();
        }
        /// <summary>
        /// Obtiene el campo CTAC_CLI (contacto) de un proveedor específico
        /// Se utiliza para obtener el "Atte. Señor" desde la tabla de proveedores
        /// </summary>
        /// <param name="numeroProveedor">Número del proveedor (Nume_Cli)</param>
        /// <returns>Contenido del campo CTAC_CLI o cadena vacía si no se encuentra</returns>
        [HttpGet("ctac-cli/{numeroProveedor}")]
        public async Task<IActionResult> GetCtacCliAsync(int numeroProveedor)
        {
            try
            {
                Console.WriteLine($"[GetCtacCliAsync] Buscando CTAC_CLI para proveedor: {numeroProveedor}");

                var proveedor = await _context.Proveedores
                    .Where(p => p.Nume_Cli == numeroProveedor && (p.Stat_Cli >= 0 || p.Stat_Cli == null))
                    .FirstOrDefaultAsync();

                if (proveedor != null && !string.IsNullOrWhiteSpace(proveedor.Ctac_Cli))
                {
                    Console.WriteLine($"[GetCtacCliAsync] ✅ Encontrado CTAC_CLI: '{proveedor.Ctac_Cli}'");
                    return Content(proveedor.Ctac_Cli, "text/plain");
                }

                Console.WriteLine($"[GetCtacCliAsync] ⚠️ No se encontró CTAC_CLI para el proveedor {numeroProveedor}");
                return Content(string.Empty, "text/plain");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GetCtacCliAsync] ❌ Error: {ex.Message}");
                return StatusCode(500, $"Error al obtener CTAC_CLI: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza un cliente existente
        /// </summary>
        [HttpPut] //para cambiar el status
        public async Task<IActionResult> PutAsync(Proved12 proveedor)
        {
            _context.Update(proveedor);
            await _context.SaveChangesAsync();
            return Ok(proveedor);
        }

        [HttpGet("buscar")]
        public async Task<IActionResult> BuscarAsync([FromQuery] string termino)
        {
            if (string.IsNullOrWhiteSpace(termino))
            {
                return BadRequest("El término de búsqueda es requerido");
            }

            var proveedores = await _context.Proveedores
                .Where(x => (x.Id.ToString() == termino) ||
                           x.Raso_Cli.ToLower().Contains(termino.ToLower()))
                .Take(10)
                .ToListAsync();

            return Ok(proveedores.Select(p => MapToDTO(p)));
        }
        private ProveedorDTO MapToDTO(Proved12 proveedor)
        {
            return new ProveedorDTO
            {
                Id = proveedor.Id,
                NumeCli = proveedor.Nume_Cli,
                RazonSocial = proveedor.Raso_Cli,
                Cuit = proveedor.Cuit_Cli ?? "",
                Telefono = proveedor.Tele_Cli ?? "",
                Domicilio = $"{proveedor.Domi_Cli} {proveedor.Loca_Cli} {proveedor.Prov_Cli}".Trim(),
                Email = proveedor.Mail_Cli ?? ""
            };
        }
    }

}
