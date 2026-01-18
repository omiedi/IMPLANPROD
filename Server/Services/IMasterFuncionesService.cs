using IMPLANPROD.Shared.DTOs;

namespace IMPLANPROD.Server.Services
{
    public interface IMasterFuncionesService
    {
        Task<MasterFuncionesDTO?> ObtenerDatosPorCodintAsync(int codint);
    }
}
