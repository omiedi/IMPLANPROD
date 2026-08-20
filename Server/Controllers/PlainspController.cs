using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la tabla PLAINSP.
    /// Gestiona las dimensiones según plano del Plan de Inspección y Ensayo
    /// asociadas a un producto identificado por su Cod_Inte (código interno).
    /// Tabla origen VB6: PLAINSP.DAT
    /// </summary>
    [ApiController]
    [Route("api/plainsp")]
    public class PlainspController : ControllerBase
    {
        private readonly DataContext _context;

        public PlainspController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Devuelve todas las filas de PLAINSP correspondientes a un producto,
        /// ordenadas por NroOrden (orden definido por el usuario) y luego por Id
        /// como desempate en caso de NroOrden duplicado.
        /// Usado al abrir el modal "Plan de Inspección y Ensayo" para cargar la grilla.
        /// GET /api/plainsp/by-codinte/{codInte}
        /// </summary>
        /// <param name="codInte">Código interno del producto (Cod_Inte)</param>
        [HttpGet("by-codinte/{codInte:int}")]
        public async Task<ActionResult<List<Plainsp>>> GetByCodInte(int codInte)
        {
            var lista = await _context.Plainsps
                .AsNoTracking()                        // solo lectura, sin tracking de cambios
                .Where(x => x.Cod_Inte == codInte)    // filtra por producto
                .OrderBy(x => x.NroOrden)             // respeta el orden definido por el usuario
                .ThenBy(x => x.Id)                    // desempate por Id de inserción
                .ToListAsync();
            return lista;
        }

        /// <summary>
        /// Guarda (reemplaza) las filas de PLAINSP para un producto.
        /// Estrategia: DELETE + INSERT en una sola transacción implícita de EF Core.
        /// Esto garantiza que el orden visual del usuario quede reflejado
        /// correctamente en el campo NroOrden (1, 2, 3...) sin necesidad de UPDATEs
        /// individuales ni de gestionar IDs existentes.
        /// Usado al presionar "Guardar" en el modal "Plan de Inspección y Ensayo".
        /// POST /api/plainsp/guardar/{codInte}
        /// </summary>
        /// <param name="codInte">Código interno del producto a reemplazar</param>
        /// <param name="items">Lista de filas en el orden actual de la grilla</param>
        [HttpPost("guardar/{codInte:int}")]
        public async Task<ActionResult> Guardar(int codInte, [FromBody] List<Plainsp> items)
        {
            // 1. Obtener y eliminar todos los registros existentes del producto
            var existentes = await _context.Plainsps
                .Where(x => x.Cod_Inte == codInte)
                .ToListAsync();

            if (existentes.Any())
                _context.Plainsps.RemoveRange(existentes);

            // 2. Insertar los nuevos registros asignando NroOrden según posición (base 1)
            for (int i = 0; i < items.Count; i++)
            {
                items[i].Id         = 0;            // forzar INSERT (no UPDATE)
                items[i].Cod_Inte   = codInte;      // asegurar FK correcta
                items[i].NroOrden   = i + 1;        // posición 1-based según orden de la grilla
                items[i].CodiEmpr   = 0;            // regla global del sistema
                items[i].AddRecord  = DateTime.Now; // auditoría: fecha de creación
                items[i].LastUpdate = DateTime.Now; // auditoría: fecha de última modificación
                _context.Plainsps.Add(items[i]);
            }

            // 3. Persistir DELETE + INSERTs en una sola operación
            await _context.SaveChangesAsync();
            return Ok(new { guardados = items.Count });
        }
    }
}
