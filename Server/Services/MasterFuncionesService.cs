using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    public class MasterFuncionesService : IMasterFuncionesService
    {
        private readonly DataContext _context;

        public MasterFuncionesService(DataContext context)
        {
            _context = context;
        }

        public async Task<MasterFuncionesDTO?> ObtenerDatosPorCodintAsync(int codint)
        {
            return await _context.Set<Master>()
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
        }
    }
}
