using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Control de Cumplimiento de Entregas. Migración de COCUENT09.FRM (proyecto COCUME09, Sub ICUMCLI).
    /// Calcula en memoria lo que VB6 grababa en el archivo de trabajo COCUSEL.DAT.
    /// </summary>
    public class CumplimientoEntregasService
    {
        public const string TipoRemito = "REMITO";
        public const string TipoPedido = "PEDIDO";
        public const string ProbusTolerancia = "COCUENT0";
        public const string ClabusTolerancia = "TOLERANC";
        public const int ToleranciaMaxima = 30;

        private const decimal CantidadMinima = 0.05m;

        private readonly DataContext _context;
        private readonly FlexcrlService _flexcrlService;
        private readonly ILogger<CumplimientoEntregasService> _logger;

        public CumplimientoEntregasService(DataContext context, FlexcrlService flexcrlService, ILogger<CumplimientoEntregasService> logger)
        {
            _context = context;
            _flexcrlService = flexcrlService;
            _logger = logger;
        }

        public async Task<int> ObtenerToleranciaAsync()
        {
            var control = await _flexcrlService.ObtenerControlAsync(ClabusTolerancia, ProbusTolerancia);
            return ParsearEntero(control?.VALCONTROL);
        }

        public Task<bool> GrabarToleranciaAsync(int dias)
        {
            if (dias < 0 || dias > ToleranciaMaxima)
                throw new ArgumentOutOfRangeException(nameof(dias), "Valor Ingresado No Válido (Valores Entre 0 y 30)");

            return _flexcrlService.ActualizarOCrearControlAsync(ClabusTolerancia, ProbusTolerancia, dias.ToString());
        }

        /// <summary>
        /// Depósito de expedición (DEPOEXPE). Se lee con la clave usada por Blazor
        /// (PROBUS='DEPOEXPE') y, si no existe, con la de VB6 (CLABUS='DEPOEXPE').
        /// </summary>
        private async Task<int> ObtenerDepositoExpedicionAsync()
        {
            var control = await _flexcrlService.ObtenerControlAsync("", "DEPOEXPE")
                          ?? await _flexcrlService.ObtenerControlAsync("DEPOEXPE", "");
            return ParsearEntero(control?.VALCONTROL);
        }

        public async Task<CumplimientoEntregaResultadoDTO> ObtenerInformeAsync(
            string tipo, DateTime fechaDesde, DateTime fechaHasta, int? clienteDesde, int? clienteHasta)
        {
            var esRemito = !string.Equals(tipo, TipoPedido, StringComparison.OrdinalIgnoreCase);
            var desde = fechaDesde.Date;
            var hastaExclusivo = fechaHasta.Date.AddDays(1);

            var tolerancia = await ObtenerToleranciaAsync();
            var deposito = await ObtenerDepositoExpedicionAsync();

            var resultado = new CumplimientoEntregaResultadoDTO
            {
                Tipo = esRemito ? TipoRemito : TipoPedido,
                FechaDesde = desde,
                FechaHasta = fechaHasta.Date,
                ToleranciaDias = tolerancia,
                ToleranciaAplicada = esRemito && tolerancia > 0,
                DepositoExpedicion = deposito
            };

            if (deposito <= 0)
                resultado.Advertencia = "No está configurado el depósito de expedición (DEPOEXPE).";

            // VB6: CODIMOVI = 'RT' AND FECHMOV entre fechas AND NUME_CLIE entre clientes AND DEPOMOVI = DEPOEXPE
            var query = _context.Movistos.AsNoTracking()
                .Where(m => m.CodiMovi == "RT"
                         && m.FechMov >= desde && m.FechMov < hastaExclusivo
                         && m.DepoMovi == deposito
                         && m.CantSalid > CantidadMinima);

            if (clienteDesde.HasValue) query = query.Where(m => m.Nume_Clie >= clienteDesde.Value);
            if (clienteHasta.HasValue) query = query.Where(m => m.Nume_Clie <= clienteHasta.Value);

            var movimientos = await query
                .Select(m => new
                {
                    m.Id,
                    m.Nume_Clie,
                    m.DoPriCor,
                    m.DoPriLar,
                    m.FechMov,
                    m.Cod_Inte,
                    m.IdenLote,
                    m.CantSalid,
                    m.NuPedCli,
                    m.NumeOrdPed
                })
                .ToListAsync();

            if (movimientos.Count == 0)
                return resultado;

            var nroPedidos = movimientos.Where(m => (m.NuPedCli ?? 0) > 0).Select(m => m.NuPedCli).Distinct().ToList();
            var codigos = movimientos.Where(m => (m.Cod_Inte ?? 0) > 0).Select(m => m.Cod_Inte!.Value).Distinct().ToList();
            var clientes = movimientos.Select(m => m.Nume_Clie ?? 0).Distinct().ToList();

            var fechasPedido = (await _context.Pedencas.AsNoTracking()
                    .Where(p => nroPedidos.Contains(p.NroPed))
                    .Select(p => new { p.NroPed, p.FechEmis })
                    .ToListAsync())
                .GroupBy(p => p.NroPed!.Value)
                .ToDictionary(g => g.Key, g => g.First().FechEmis);

            var renglonesPedido = await _context.Peddetas.AsNoTracking()
                .Where(p => nroPedidos.Contains(p.NroPed))
                .OrderBy(p => p.PeddetaId)
                .Select(p => new { p.NroPed, p.CodiInt, p.NuOrItem, p.FeSolEnt })
                .ToListAsync();

            var productos = (await _context.masters.AsNoTracking()
                    .Where(m => codigos.Contains(m.Codint))
                    .Select(m => new { m.Codint, m.Codigo, m.Descripcion })
                    .ToListAsync())
                .ToDictionary(m => m.Codint);

            var razones = (await _context.Clientes.AsNoTracking()
                    .Where(c => clientes.Contains(c.Nume_Cli))
                    .Select(c => new { c.Nume_Cli, c.Raso_Cli })
                    .ToListAsync())
                .GroupBy(c => c.Nume_Cli)
                .ToDictionary(g => g.Key, g => g.First().Raso_Cli ?? string.Empty);

            var aplicarTolerancia = esRemito && tolerancia > 0;

            foreach (var mov in movimientos)
            {
                var fechaRemito = mov.FechMov!.Value.Date;
                var nroPedido = mov.NuPedCli ?? 0;
                var codInte = mov.Cod_Inte ?? 0;

                // VB6: sin pedido (NROPEDI < 1) se toma la fecha del remito y el puntaje queda en 100.
                DateTime fechaPedido = fechaRemito;
                DateTime fechaSolicitada = fechaRemito;

                if (nroPedido > 0 && fechasPedido.TryGetValue(nroPedido, out var femi) && femi.HasValue)
                {
                    fechaPedido = femi.Value.Date;

                    // Renglón exacto del pedido por NumeOrdPed (NuOrItem); si no existe,
                    // primer renglón del artículo en el pedido (criterio VB6).
                    var renglon = (mov.NumeOrdPed ?? 0) > 0
                        ? renglonesPedido.FirstOrDefault(r => r.NroPed == nroPedido && r.CodiInt == codInte && r.NuOrItem == mov.NumeOrdPed)
                        : null;
                    renglon ??= renglonesPedido.FirstOrDefault(r => r.NroPed == nroPedido && r.CodiInt == codInte);

                    // VB6: FESOLENT inexistente => plazo solicitado 0.
                    fechaSolicitada = renglon?.FeSolEnt?.Date ?? fechaPedido;
                }

                var plazoSolicitado = DiasEntre(fechaPedido, fechaSolicitada);
                var plazoReal = DiasEntre(fechaPedido, fechaRemito);
                var atraso = plazoReal - plazoSolicitado;

                if (aplicarTolerancia)
                    atraso = atraso >= tolerancia ? atraso - tolerancia : 0;
                if (atraso < 0) atraso = 0;

                decimal puntaje = 100m;
                if (plazoSolicitado + atraso > 0)
                    puntaje = 100m * plazoSolicitado / (plazoSolicitado + atraso);

                productos.TryGetValue(codInte, out var producto);
                razones.TryGetValue(mov.Nume_Clie ?? 0, out var razon);

                resultado.Detalle.Add(new CumplimientoEntregaDTO
                {
                    NumeClie = mov.Nume_Clie ?? 0,
                    RazonSocial = razon ?? string.Empty,
                    NroRemito = NumeroRemito(mov.DoPriCor),
                    RemitoLargo = mov.DoPriLar?.Trim() ?? string.Empty,
                    FechaRemito = fechaRemito,
                    NroPedido = nroPedido,
                    FechaPedido = fechaPedido,
                    FechaSolicitada = fechaSolicitada,
                    CodInte = codInte,
                    Codigo = producto?.Codigo?.Trim() ?? string.Empty,
                    Descripcion = producto?.Descripcion?.Trim() ?? string.Empty,
                    Lote = mov.IdenLote?.Trim() ?? string.Empty,
                    Cantidad = mov.CantSalid ?? 0m,
                    PlazoSolicitado = plazoSolicitado,
                    PlazoReal = plazoReal,
                    Atraso = atraso,
                    Puntaje = Math.Round(puntaje, 1)
                });
            }

            resultado.Detalle = esRemito
                ? resultado.Detalle
                    .OrderBy(d => d.NumeClie).ThenBy(d => d.NroRemito).ThenBy(d => d.RemitoLargo)
                    .ThenBy(d => d.FechaRemito).ThenBy(d => d.Codigo).ToList()
                : resultado.Detalle
                    .OrderBy(d => d.NumeClie).ThenBy(d => d.NroPedido).ThenBy(d => d.FechaPedido)
                    .ThenBy(d => d.Codigo).ThenBy(d => d.FechaRemito).ToList();

            resultado.Clientes = resultado.Detalle
                .GroupBy(d => new { d.NumeClie, d.RazonSocial })
                .Select(g => new CumplimientoEntregaClienteDTO
                {
                    NumeClie = g.Key.NumeClie,
                    RazonSocial = g.Key.RazonSocial,
                    Renglones = g.Count(),
                    Cantidad = g.Sum(d => d.Cantidad),
                    TotalAtraso = g.Sum(d => d.Atraso),
                    TotalPuntaje = g.Sum(d => d.Puntaje),
                    IndiceCumplimiento = Math.Round(g.Average(d => d.Puntaje), 1)
                })
                .OrderBy(c => c.NumeClie)
                .ToList();

            resultado.TotalCantidad = resultado.Detalle.Sum(d => d.Cantidad);
            resultado.TotalAtraso = resultado.Detalle.Sum(d => d.Atraso);
            resultado.TotalPuntaje = resultado.Detalle.Sum(d => d.Puntaje);
            resultado.IndiceCumplimiento = Math.Round(resultado.Detalle.Average(d => d.Puntaje), 1);

            _logger.LogInformation("Cumplimiento de entregas {Tipo}: {Cantidad} renglones ({Desde:d} - {Hasta:d})",
                resultado.Tipo, resultado.Detalle.Count, desde, fechaHasta);

            return resultado;
        }

        /// <summary>VB6 DIENTRE%: días entre dos fechas, 0 si la segunda es anterior.</summary>
        private static int DiasEntre(DateTime desde, DateTime hasta)
            => hasta < desde ? 0 : (hasta - desde).Days;

        /// <summary>VB6: NREMI = Val(Mid(DOPRICOR, 4)), ej. "RT.6449" => 6449.</summary>
        private static long NumeroRemito(string? doPriCor)
        {
            if (string.IsNullOrEmpty(doPriCor) || doPriCor.Length < 4) return 0;
            var digitos = new string(doPriCor.Substring(3).TrimStart().TakeWhile(char.IsDigit).ToArray());
            return long.TryParse(digitos, out var numero) ? numero : 0;
        }

        private static int ParsearEntero(string? valor)
            => int.TryParse(valor?.Trim(), out var numero) ? numero : 0;
    }
}
