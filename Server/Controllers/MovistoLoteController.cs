using CurrieTechnologies.Razor.SweetAlert2;
using IMPLANPROD.Client.Repositories;
using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Server.Utilities;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{

    [ApiController]
    [Route("/api/movistolote")]
    public class movistoloteController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly DocumentNumberGenerator documentNumberGenerator;
        private readonly IUsuarioFlexoftService _UsuarioFlexoftService;
        private readonly ICodiempresa _Codiempresa;



        public movistoloteController(DataContext context, DocumentNumberGenerator documentNumberGenerator, IUsuarioFlexoftService UsuarioFlexoftService, ICodiempresa Codiempresa)
        {
            _context = context;
            this.documentNumberGenerator = documentNumberGenerator;
            _UsuarioFlexoftService = UsuarioFlexoftService;
            _Codiempresa = Codiempresa;
        }
        // ESTO SE DEBE AGREGAR EN EL ENDPOINT DONDE SE DEBE HACER UNA TRANSACCION ******
        //int? idFlexoft = _UsuarioFlexoftService.InterObtenerIdFlexoft(User);
        //short? codigoEmpresa = _Codiempresa.IObtenerCodiempresa(User);
        //        if (!idFlexoft.HasValue)
        //        {
        //            return BadRequest("(Sesion Caducada - Debe cerrar sesión y loguearse nuevamente.");

        //        }
        // ESTA ES LA FORMA DE LLAMADA EN .RAZOR PARA QUE MUESTRE EL MENSAJE QUE RETORNA DESDE EL ENDPOINT  ************
        //var response = await repository.Post("api/stockTransferencia/transferencia", transferencia);
        //    if (response.Error)
        //    {
        //        var mensaje = await response.HttpResponseMessage.Content.ReadAsStringAsync();
        //await sweetAlertService.FireAsync("Error", mensaje, SweetAlertIcon.Error);
        //        return;
        //    }

        public async Task<List<LoteSaldoDTO>> ObtenerLotesConSaldoAsync(int codInte)
        {
            return await _context.Movistos
                .Where(m => m.Cod_Inte == codInte && !string.IsNullOrEmpty(m.IdenLote))
                .GroupBy(m => m.IdenLote)
                .Select(g => new LoteSaldoDTO
                {
                    IdenLote = g.Key,
                    Saldo = g.Sum(m => (m.CantIngre ?? 0) - (m.CantSalid ?? 0))
                })
                .Where(x => x.Saldo > 0)
                .ToListAsync();
        }
    }
}
