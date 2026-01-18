using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Mvc;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Server.Services;

namespace IMPLANPROD.Server.FuncionesGrales
{
    /// <summary>
    /// Controlador que contiene funciones generales relacionadas con la tabla Master
    /// </summary>
    [ApiController]
    [Route("/api/masterfunciones")]
    public class MasterFuncion : ControllerBase
    {
        private readonly DataContext _context;

        public MasterFuncion(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la unidad de medida de un producto basado en su código interno.
        /// Si la unidad de medida está vacía o no se encuentra el producto, retorna "U." como valor por defecto.
        /// </summary>
        /// <param name="codint">Código interno del producto en la tabla Master</param>
        /// <returns>Unidad de medida del producto o "U." si no tiene una asignada</returns>
        [HttpGet("unimed/{codint}")]
        public async Task<ActionResult<string>> Unimed(int codint)
        {
            // Buscar el producto en la tabla Master por su código interno
            var master = await _context.masters
                .FirstOrDefaultAsync(m => m.Codint == codint);

            if (master == null)
                return Ok("U.");

            // Eliminar espacios al inicio y final, equivalente al TRIM$ de VB6
            string unimed = master.Umedida?.Trim() ?? string.Empty;
            
            // Si está vacío, retornar "U." como valor por defecto
            return Ok(string.IsNullOrEmpty(unimed) ? "U." : unimed);
        }

        /// <summary>
        /// Obtiene la descripción de un producto basado en su código interno.
        /// Si no se encuentra el producto, retorna una cadena vacía.
        /// </summary>
        /// <param name="codint">Código interno del producto en la tabla Master</param>
        /// <returns>Descripción del producto o cadena vacía si no se encuentra</returns>
        [HttpGet("descripcion/{codint}")]
        public async Task<ActionResult<string>> GetDescripcion(int codint)
        {
            var master = await _context.masters
                .FirstOrDefaultAsync(m => m.Codint == codint);

            if (master == null)
                return Ok(string.Empty); //  Devuelve un ActionResult<string>

            string descripcion = master.Descripcion?.Trim() ?? string.Empty;
            return Ok(descripcion); //  Envuelto en Ok(...)
        }
        /// <summary>
        /// Devuelve un subconjunto de datos del producto (MasterFuncionesDTO) basado en su código interno.
        /// </summary>
        /// <param name="codint">Código interno del producto (Codint)</param>
        /// <returns>DTO con campos específicos del producto</returns>
        [HttpGet("funcionesAtributos/{codint}")]
        public async Task<ActionResult<MasterFuncionesDTO>> GetFuncionesDTO(int codint)
        {
            var dto = await _context.masters
                .Where(m => m.Codint == codint)
                .Select(m => new MasterFuncionesDTO
                {
                    DesCompAut = m.DesCompAut,
                    Codigo = m.Codigo,
                    Descripcion = m.Descripcion,
                    Status = m.Status,
                    Umedida = m.Umedida,
                    Timate = m.Timate,
                    Monecos = m.Monecos,
                    Cosuni = m.Cosuni,
                    Provhabi = m.Provhabi,
                    Stomin = m.Stomin,
                    Stomax = m.Stomax,
                    Lorepos = m.Lorepos,
                    Lindivis = m.Lindivis,
                    Grupiva = m.Grupiva,
                    Trazable = m.Trazable
                })
                .FirstOrDefaultAsync();

            if (dto == null)
                return NotFound("Producto no encontrado");

            return Ok(dto);
        }
    }
}
