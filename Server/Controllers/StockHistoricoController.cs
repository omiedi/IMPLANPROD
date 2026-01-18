using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace MantenimientoImp.Server.Controllers
{
    [ApiController]
    [Route("api/stockhistorico")]
    public class StockHistoricoController : ControllerBase
    {
        private readonly DataContext _context;

        public StockHistoricoController(DataContext context)
        {
            _context = context;
        }

        [HttpPost("historico")]
        public async Task<ActionResult<List<MovimientoHistoricoDTO>>> GetHistorico([FromBody] FiltroMovimientosDTO filtro)
        {
            try
            {
                var query = _context.Movistos
                    .Where(m => m.Cod_Inte == filtro.CodigoRepuesto);

                if (filtro.Deposito.HasValue)
                {
                    query = query.Where(m => m.DepoMovi == filtro.Deposito.Value);
                }

                if (filtro.FechaDesde.HasValue)
                {
                    var fechaDesde = filtro.FechaDesde.Value.Date;
                    query = query.Where(m => m.FeReMovi >= fechaDesde);
                }

                if (filtro.FechaHasta.HasValue)
                {
                    var fechaHasta = filtro.FechaHasta.Value.Date.AddDays(1).AddSeconds(-1);
                    query = query.Where(m => m.FeReMovi <= fechaHasta);
                }

                var movimientos = await query
                    .OrderBy(m => m.FeReMovi)
                    .Select(m => new MovimientoHistoricoDTO
                    {
                        Id = (int)m.Cod_Inte,
                        Fecha = m.FeReMovi ?? DateTime.MinValue,
                        TipoMovimiento = m.CodiMovi,
                        DocumentoNumero = m.DoSecLar,
                        Transaccion = m.DoPriLar,
                        Deposito = m.DepoMovi,
                        IdBulto = m.IdenLote,
                        Referencia = m.ReferCor,
                        Entrada = m.CantIngre,
                        Salida = m.CantSalid,
                        Saldo = 0 // Se calculará después
                    })
                    .ToListAsync();

                // Calcular saldos acumulados
                decimal saldoAcumulado = 0;
                foreach (var mov in movimientos)
                {
                    saldoAcumulado += (mov.Entrada ?? 0) - (mov.Salida ?? 0);
                    mov.Saldo = saldoAcumulado;
                }

                return Ok(movimientos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        [HttpGet("historicosimple")]
        public async Task<ActionResult<List<MovimientoHistoricoDTO>>> GetHistoricoSimple(
            [FromQuery] int repuestoId,
            [FromQuery] int? deposito = null,
            [FromQuery] DateTime? fechaDesde = null,
            [FromQuery] DateTime? fechaHasta = null)
        {
            try
            {
                var query = _context.Movistos.AsQueryable();

                // Filtro por código (requerido)
                query = query.Where(m => m.Cod_Inte == repuestoId);

                if (deposito.HasValue)
                {
                    query = query.Where(m => m.DepoMovi == deposito.Value);
                }

                if (fechaDesde.HasValue)
                {
                    var fecha = fechaDesde.Value.Date;
                    query = query.Where(m => m.FeReMovi >= fecha);
                }

                if (fechaHasta.HasValue)
                {
                    var fecha = fechaHasta.Value.Date.AddDays(1).AddSeconds(-1);
                    query = query.Where(m => m.FeReMovi <= fecha);
                }

                var movimientos = await query
                    .OrderBy(m => m.FeReMovi)
                    .Select(m => new MovimientoHistoricoDTO
                    {
                        Id = (int)m.Cod_Inte,
                        Fecha = m.FeReMovi ?? DateTime.Now,
                        TipoMovimiento = m.CodiMovi ?? "",
                        DocumentoNumero = m.DoPriCor ?? "",
                        Transaccion = m.DoSecCor ?? "",
                        Deposito = m.DepoMovi ?? 0,
                        IdBulto = m.IdenLote ?? "",
                        Referencia = m.ReferCor ?? "",
                        Entrada = m.CantIngre,
                        Salida = m.CantSalid,
                        Saldo = 0,
                        CodInte = m.Cod_Inte ?? 0
                    })
                    .ToListAsync();

                decimal saldoAcumulado = 0;
                foreach (var movimiento in movimientos.OrderBy(m => m.Fecha))
                {
                    var entrada = movimiento.Entrada ?? 0;
                    var salida = movimiento.Salida ?? 0;
                    saldoAcumulado += entrada - salida;
                    movimiento.Saldo = saldoAcumulado;
                }


                return Ok(movimientos);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener el histórico simple: {ex.Message}");
                return BadRequest($"Error al obtener el histórico: {ex.Message}");
            }
        }


    }
}
