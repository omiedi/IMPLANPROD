using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/brecepen")]
    public class BrecepenController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IMoviStockService _moviStockService;

        public BrecepenController(DataContext context, IMoviStockService moviStockService)
        {
            _context = context;
            _moviStockService = moviStockService;
        }

        /// <summary>
        /// Obtiene las recepciones pendientes de aprobación para un producto específico.
        /// Usado en Stock Proyectado de StockInfGeneral.
        ///
        /// Corazón de la consulta (SQL equivalente):
        ///   SELECT IDENLOTE, COD_INTE, SUM(CANTINGRE - CANTSALID) AS CantPendiente
        ///   FROM BRECEPEN
        ///   WHERE COD_INTE = @codint
        ///   GROUP BY IDENLOTE, COD_INTE
        ///   HAVING SUM(CANTINGRE - CANTSALID) > 0
        ///
        /// NOTA EF Core: no se usan cláusulas `let` ni `OrderBy(...).FirstOrDefault()`
        /// dentro del GroupBy porque EF Core no puede traducirlos a SQL
        /// ("Translation of 'Select' which contains grouping parameter without composition").
        /// Se proyectan los agregados directamente en el Select del grupo.
        /// </summary>
        [HttpGet("pendientes-aprobacion-por-producto/{codint}")]
        public async Task<ActionResult<List<StockRecepcionPendienteAprobacionDTO>>> GetPendientesAprobacionPorProducto(int codint)
        {
            try
            {
                // ── Paso 1: agrupar por IDENLOTE + COD_INTE ──────────────────
                // Agregados proyectados directamente (sin `let`), que EF Core sí traduce.
                var grupos = await _context.Brecepens
                    .AsNoTracking()
                    .Where(b => b.COD_INTE == codint)
                    .GroupBy(b => new { b.Idenlote, b.COD_INTE })
                    .Select(g => new
                    {
                        g.Key.Idenlote,
                        g.Key.COD_INTE,
                        CantPendiente = g.Sum(x => (x.Cantingre ?? 0m) - (x.Cantsalid ?? 0m)),
                        // Datos representativos del grupo, resueltos con agregados (traducibles)
                        MaxId = g.Max(x => x.Id),
                        FechaRecepcion = g.Max(x => x.Fechmov),
                        NumeProv = g.Max(x => x.NumeProv),
                        RemitoProveedor = g.Max(x => x.Refercor)
                    })
                    .Where(x => x.CantPendiente > 0m)
                    .ToListAsync();

                if (grupos.Count == 0)
                    return Ok(new List<StockRecepcionPendienteAprobacionDTO>());

                // ── Paso 2: resolver razón social de los proveedores ─────────
                // El número de proveedor de BRECEPEN es Proved12.Nume_Cli (no el Id).
                var numeProvs = grupos
                    .Select(g => g.NumeProv ?? 0)
                    .Where(n => n > 0)
                    .Distinct()
                    .ToList();

                var proveedores = await _context.Proveedores
                    .AsNoTracking()
                    .Where(p => numeProvs.Contains(p.Nume_Cli))
                    .ToDictionaryAsync(p => p.Nume_Cli, p => p.Raso_Cli ?? "");

                // ── Paso 3: armar el DTO ─────────────────────────────────────
                var result = grupos
                    .Select(g =>
                    {
                        var numeProv = g.NumeProv ?? 0;
                        return new StockRecepcionPendienteAprobacionDTO
                        {
                            Numebore = g.MaxId,
                            NumeProv = g.NumeProv,
                            Proveedor = numeProv > 0 && proveedores.ContainsKey(numeProv)
                                ? proveedores[numeProv]
                                : "",
                            CantPendiente = g.CantPendiente,
                            RemitoProveedor = g.RemitoProveedor ?? "",
                            FechaRecepcion = g.FechaRecepcion ?? DateTime.MinValue,
                            Lote = g.Idenlote ?? ""
                        };
                    })
                    .OrderBy(x => x.Numebore)
                    .ToList();

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al obtener recepciones pendientes de aprobación: {ex.Message}");
            }
        }

        // Endpoint que alimenta la grilla de la solapa "Aprueba Almacén".
        // Devuelve solo lo pendiente de aprobación, agrupado por LOTE (IDENLOTE) + producto (COD_INTE).
        // Importante: filtra SOLO los grupos donde (SUM(CANTINGRE) - SUM(CANTSALID de 'AP')) > 0.
        [HttpGet("pendientes-calidad")]
        public async Task<ActionResult<List<BrecepenPendienteCalidadDTO>>> GetPendientesCalidad()
        {
            // Agrupamos por IDENLOTE + COD_INTE (una recepción se identifica por su lote y su producto).
            // Pendiente = CANTIDAD INGRESADA - CANTIDAD APROBADA, donde:
            //   - Ingresada  = SUM(CANTINGRE) del registro original de recepción (CODIMOVI = 'BR').
            //   - Aprobada   = SUM(CANTSALID) de los registros de aprobación (CODIMOVI = 'AP').
            // Los registros 'AP' tienen CANTINGRE = 0 y los 'BR' tienen CANTSALID = 0, por lo que
            // cada aprobación parcial queda registrada como una fila 'AP' independiente.
            var query =
                from b in _context.Brecepens.AsNoTracking()
                join bol in _context.Bolreceps.AsNoTracking()
                    on b.Idenlote equals bol.IdenLote into bolJoin
                from bol in bolJoin.DefaultIfEmpty()
                join m in _context.masters.AsNoTracking()
                    on b.COD_INTE equals m.Codint into masterJoin
                from m in masterJoin.DefaultIfEmpty()
                group new { b, bol, m } by new
                {
                    b.Idenlote,
                    b.COD_INTE,
                    m.Codigo,
                    m.Descripcion
                }
                into g
                select new
                {
                    g.Key.Idenlote,
                    g.Key.COD_INTE,
                    Codigo = g.Key.Codigo,
                    Descripcion = g.Key.Descripcion,
                    // Doprilar representativo del grupo (para mostrar en la grilla)
                    Doprilar = g.Max(x => x.b.Doprilar),
                    // Número de recepción desde BOLRECEP.NumeBoRe
                    NumeroRecepcion = g.Max(x => x.bol != null ? (int?)x.bol.NumeBoRe : null),
                    // Fecha de recepción desde BOLRECEP.FechRecep
                    FechaRecepcion = g.Max(x => x.bol != null ? x.bol.FechRecep : (DateTime?)null),
                    // Depósito de ingreso del registro original de recepción (se usa como valor por
                    // defecto al aprobar; el usuario puede cambiarlo antes de confirmar).
                    // BRECEPEN.Depomovi es short?; lo llevamos a int? para no acoplar el DTO al tipo de columna.
                    DepositoIngreso = g.Max(x => (int?)x.b.Depomovi),
                    // Ingresado: los registros 'AP' tienen CANTINGRE = 0, así que suman 0 aquí.
                    CantidadIngresada = g.Sum(x => x.b.Cantingre ?? 0m),
                    // Aprobado: SUM(CANTSALID) SOLO de los registros CODIMOVI = 'AP'.
                    CantidadAprobada = g.Sum(x =>
                        (x.b.Codimovi ?? "").Trim() == "AP" ? (x.b.Cantsalid ?? 0m) : 0m)
                };

            var result = await query
                .Select(x => new
                {
                    x.Idenlote,
                    x.COD_INTE,
                    x.Codigo,
                    x.Descripcion,
                    x.Doprilar,
                    x.NumeroRecepcion,
                    x.FechaRecepcion,
                    x.DepositoIngreso,
                    CantidadPendiente = x.CantidadIngresada - x.CantidadAprobada
                })
                .Where(x => x.CantidadPendiente > 0)
                .OrderBy(x => x.Doprilar)
                .ThenBy(x => x.Codigo)
                .ToListAsync();

            // Resolvemos la descripción de los depósitos involucrados en una sola consulta,
            // para no golpear la base de datos una vez por fila.
            var codigosDeposito = result
                .Where(x => x.DepositoIngreso.HasValue && x.DepositoIngreso.Value > 0)
                .Select(x => x.DepositoIngreso.Value)
                .Distinct()
                .ToList();

            var nombresDeposito = await _context.Tadeposi
                .Where(d => codigosDeposito.Contains(d.CodigoDepo))
                .ToDictionaryAsync(d => d.CodigoDepo, d => d.Descripcion ?? "");

            var dtoResult = result.Select(x => new BrecepenPendienteCalidadDTO
            {
                Idenlote = x.Idenlote,
                Doprilar = x.Doprilar,
                CodInte = x.COD_INTE,
                Codigo = x.Codigo ?? "SIN DATOS",
                Descripcion = x.Descripcion ?? "SIN DATOS",
                CantidadPendiente = x.CantidadPendiente,
                NumeroRecepcion = x.NumeroRecepcion,
                FechaRecepcion = x.FechaRecepcion,
                DepositoIngreso = x.DepositoIngreso,
                NombreDepositoIngreso = x.DepositoIngreso.HasValue && nombresDeposito.ContainsKey(x.DepositoIngreso.Value)
                    ? nombresDeposito[x.DepositoIngreso.Value]
                    : null
            }).ToList();

            return Ok(dtoResult);
            //SU TRADUCCION
            //            SELECT
            //    b.IDENLOTE,
            //    MAX(b.DOPRILAR) AS DOPRILAR,
            //    m.CODIGO,
            //    m.DESCRIPCION,
            //    SUM(ISNULL(b.CANTINGRE, 0))
            //      - SUM(CASE WHEN LTRIM(RTRIM(b.CODIMOVI)) = 'AP'
            //                 THEN ISNULL(b.CANTSALID, 0) ELSE 0 END) AS CantidadPendiente
            //FROM BRECEPEN b
            //LEFT JOIN MASTER m
            //    ON b.COD_INTE = m.CODINT
            //GROUP BY
            //    b.IDENLOTE,
            //    b.COD_INTE,
            //    m.CODIGO,
            //    m.DESCRIPCION
            //HAVING
            //    (SUM(ISNULL(b.CANTINGRE, 0))
            //     - SUM(CASE WHEN LTRIM(RTRIM(b.CODIMOVI)) = 'AP'
            //                THEN ISNULL(b.CANTSALID, 0) ELSE 0 END)) > 0
            //ORDER BY
            //    DOPRILAR,
            //    m.CODIGO;
        }

        // Consume (aprueba) una cantidad pendiente dentro de BRECEPEN para un lote (IDENLOTE) + producto (COD_INTE).
        // La aprobación NO modifica el registro original: genera un registro nuevo con CODIMOVI = 'AP'.
        [HttpPost("aprobar")]
        public async Task<ActionResult> AprobarCantidad([FromBody] BrecepenAprobarRequestDTO request)
        {
            // ---- Validaciones de entrada ----
            // Sin body no hay nada que aprobar.
            if (request == null)
            {
                return BadRequest("Request inválido");
            }

            // El lote identifica (junto al producto) qué recepción estamos aprobando. Es obligatorio.
            var idenlote = (request.Idenlote ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(idenlote))
            {
                return BadRequest("Idenlote es requerido");
            }

            // Producto (código interno) obligatorio y válido.
            if (request.CodInte <= 0)
            {
                return BadRequest("CodInte es requerido");
            }

            // No tiene sentido aprobar cero o negativo.
            if (request.CantidadAprobar <= 0)
            {
                return BadRequest("CantidadAprobar debe ser mayor a 0");
            }

            // Cargamos TODOS los movimientos del lote + producto (originales 'BR' y aprobaciones 'AP').
            // Necesitamos los 'AP' para calcular correctamente el pendiente ya aprobado.
            var movimientos = await _context.Brecepens
                .Where(x => x.Idenlote.Trim() == idenlote && x.COD_INTE == request.CodInte)
                .OrderBy(x => x.Id)
                .ToListAsync();

            if (movimientos.Count == 0)
            {
                return NotFound("No se encontraron movimientos pendientes para el criterio");
            }

            // pendienteTotal = SUM(CANTINGRE) - SUM(CANTSALID).
            // Los 'AP' aportan (0 - cantsalid) => descuentan lo ya aprobado; los 'BR' aportan (cantingre - 0).
            var pendienteTotal = movimientos.Sum(x => (x.Cantingre ?? 0m) - (x.Cantsalid ?? 0m));
            if (pendienteTotal <= 0)
            {
                return BadRequest("No hay cantidad pendiente para aprobar");
            }

            if (request.CantidadAprobar > pendienteTotal)
            {
                return BadRequest($"Cantidad a aprobar excede el pendiente ({pendienteTotal})");
            }

            // -----------------------------------------------------------------
            // La idea "para mi abuelita":
            // 1) Tenemos uno o varios registros BRECEPEN para ese IDENLOTE + COD_INTE.
            // 2) El usuario aprueba una cantidad.
            // 3) Esa cantidad NO se resta al registro original: se registra como un
            //    NUEVO movimiento de aprobación (CODIMOVI = 'AP') con CANTSALID = lo aprobado.
            // 4) Por cada pedacito de cantidad que consumimos de un registro original:
            //    - Generamos el registro de aprobación 'AP' en BRECEPEN
            //    - Generamos el movimiento real en MOVISTO (ingreso a stock)
            //    - Actualizamos BOLRECEP.CanToApro (cantidad aprobada acumulada)
            //    - Si hay OC, actualizamos OCOMDETA.Cant_Rec con la misma cuenta del VB6
            // Todo esto debe quedar consistente: usamos una transacción, si falla una parte no guardamos nada.
            // -----------------------------------------------------------------
            await using var tx = await _context.Database.BeginTransactionAsync();

            try
            {
                // "restante" = cantidad que todavía nos falta repartir/aprobar.
                // Arranca con lo que pidió el usuario y se va descontando en cada vuelta.
                var restante = request.CantidadAprobar;

                // Recorremos todos los movimientos del lote + producto.
                // Vamos "consumiendo" el pendiente de cada registro original hasta cubrir lo pedido.
                foreach (var brecepen in movimientos)
                {
                    // Si ya repartimos todo lo que había que aprobar, cortamos el recorrido.
                    if (restante <= 0)
                    {
                        break;
                    }

                    // Pendiente de ESTE registro puntual = lo que ingresó menos lo que ya salió.
                    // Los registros de aprobación ('AP') tienen CANTINGRE = 0, por lo que su
                    // pendienteMov da negativo y se saltan (no se vuelven a aprobar).
                    var pendienteMov = (brecepen.Cantingre ?? 0m) - (brecepen.Cantsalid ?? 0m);
                    if (pendienteMov <= 0)
                    {
                        continue;
                    }

                    // "Consumir" = cuánto aprobamos de este registro particular.
                    // Es el menor entre lo que le queda pendiente al registro y lo que aún nos falta repartir.
                    var consumir = Math.Min(pendienteMov, restante);

                    // Depósito de ingreso a usar: si el usuario eligió uno explícitamente en la
                    // confirmación de aprobación, prevalece sobre el depósito original del registro.
                    // BRECEPEN.Depomovi es short?, por eso convertimos el valor elegido por el usuario.
                    short? depositoIngreso = (request.DepositoIngreso.HasValue && request.DepositoIngreso.Value > 0)
                        ? (short)request.DepositoIngreso.Value
                        : brecepen.Depomovi;

                    // 1) NO modificamos el registro original (CODIMOVI = 'BR').
                    //    En su lugar generamos un NUEVO registro de aprobación (CODIMOVI = 'AP'),
                    //    idéntico al original pero con:
                    //      - CANTINGRE = 0
                    //      - CANTSALID = cantidad aprobada (consumir)
                    //      - FECHMOV / FEREMOVI = ahora
                    //    De esta forma el pendiente se sigue calculando como
                    //    CANTINGRE(original) - SUM(CANTSALID) y además queda registro
                    //    histórico de cuándo y cuánto se aprobó en cada aprobación parcial.
                    var ahora = DateTime.Now;
                    var aprobacion = new IMPLANPROD.Shared.Entities.Brecepen
                    {
                        Fechmov = ahora,
                        Feremovi = ahora,
                        COD_INTE = brecepen.COD_INTE,
                        Idenlote = brecepen.Idenlote,
                        Codimovi = "AP",
                        Doprilar = brecepen.Doprilar,
                        Doseclar = brecepen.Doseclar,
                        Dopricor = brecepen.Dopricor,
                        Doseccor = brecepen.Doseccor,
                        Refercor = brecepen.Refercor,
                        Valount = brecepen.Valount,
                        Moneval = brecepen.Moneval,
                        Depomovi = depositoIngreso,
                        NumeClie = brecepen.NumeClie,
                        NumeProv = brecepen.NumeProv,
                        Cantingre = 0m,
                        Cantsalid = consumir,
                        ClaveHp = brecepen.ClaveHp,
                        CodiEmprNet = brecepen.CodiEmprNet,
                        NumUsuar = brecepen.NumUsuar,
                        AddRecord = ahora,
                        LastUpdate = ahora
                    };
                    _context.Brecepens.Add(aprobacion);

                    // 2) Generar el movimiento real de stock en MOVISTO (ingreso al depósito).
                    // Copiamos los datos identificatorios desde el registro original de recepción.
                    // Nota: MoviStockDTO usa NumeroCliente < 0 para indicar que es un PROVEEDOR;
                    // en BRECEPEN el proveedor viene en positivo (NUME_PROV), por eso lo negamos.
                    var movi = new MoviStockDTO
                    {
                        Referencia = brecepen.Refercor,                                  // Remito / referencia del comprobante
                        Fecha = brecepen.Fechmov ?? DateTime.Now,                        // Fecha del movimiento original
                        Lote = brecepen.Idenlote,                                        // Lote de la recepción
                        CodigoMovimiento = (brecepen.Codimovi ?? string.Empty).Trim(),   // Código de movimiento original
                        DocumentoPrimarioLargo = brecepen.Doprilar,                      // Documento primario (largo)
                        DocumentoPrimarioCorto = brecepen.Dopricor,                      // Documento primario (corto)
                        DocumentoSecundarioLargo = brecepen.Doseclar,                    // Documento secundario (largo)
                        DocumentoSecundarioCorto = brecepen.Doseccor,                    // Documento secundario (corto)
                        Deposito = depositoIngreso ?? 0,                                 // Depósito destino (editable por el usuario al aprobar)
                        NumeroCliente = -Math.Abs(brecepen.NumeProv ?? 0),               // Proveedor en negativo (convención MOVISTO)
                        CodigoInterno = brecepen.COD_INTE,                               // Producto (código interno)
                        Cantidad = consumir                                              // Cantidad que estamos aprobando ahora
                    };

                    // Registramos el movimiento. Si falla, deshacemos TODO (rollback) y devolvemos el error.
                    var resultMovisto = await _moviStockService.RegistrarMovimientoAsync(movi, User);
                    if (resultMovisto is BadRequestObjectResult bad)
                    {
                        await tx.RollbackAsync();
                        return bad;
                    }

                    // 3) Actualizar BOLRECEP.CanToApro (cantidad aprobada acumulada del bolón de recepción)
                    //    y de paso leer los datos que necesitamos para actualizar la Orden de Compra.
                    // Condición pedida: bolrecep.idenlote = brecepen.idenlote AND bolrecep.cod_inte = brecepen.cod_inte
                    var bol = await _context.Bolreceps
                        .Where(b => b.IdenLote == brecepen.Idenlote && b.CodInte == brecepen.COD_INTE)
                        .FirstOrDefaultAsync();

                    // Datos que sacamos del BOLRECEP para (más abajo) actualizar la OC. Valores por defecto por si no hay bolrecep.
                    decimal? cantNomi = null;   // Cantidad nominal del ítem
                    short numOitOc = 0;         // Número de ítem dentro de la OC
                    int nuOc = 0;               // Número de Orden de Compra (0 = sin OC)

                    if (bol != null)
                    {
                        // Sumamos lo recién aprobado a lo ya aprobado del bolrecep.
                        bol.CanToApro = (bol.CanToApro ?? 0m) + consumir;
                        bol.LastUpdate = DateTime.Now;

                        // Guardamos los datos para el cálculo de la OC.
                        cantNomi = bol.CantNom;
                        numOitOc = bol.NroItemOC ?? 0;
                        nuOc = bol.NumOC ?? 0;
                    }

                    // 4) Si hay OC, actualizar OCOMDETA (equivalente VB6)
                    // VB6:
                    //   CAPRO = CANMODIF
                    //   COEFUNI = !Coef_Unids (si < 0.0001 => 1)
                    //   CANTORECI = !Cant_Rec - (CANTNOMI/COEFUNI) + (CAPRO/COEFUNI)
                    //   !Cant_Rec = CANTORECI
                    // Solo entramos si el bolrecep tiene una OC asociada (nuOc > 0) y conocemos la cantidad nominal.
                    if (nuOc > 0 && cantNomi.HasValue)
                    {
                        // CAPRO = cantidad aprobada en esta vuelta (equivalente a CANMODIF del VB6).
                        var capro = consumir;

                        // Buscamos el ítem de OC que corresponda a esta OC y a este producto, y que aún esté vigente (STATOCOM < 8).
                        // - Si NroItemOC == 0, el VB6 busca el primer item con Cant_Rec >= (CantNom - 0.005)
                        // - Si NroItemOC != 0, busca ese NuOrd_Item puntual.
                        var queryOcom = _context.Ocomdetas
                            .Where(d => d.NumeOcom == nuOc)                          // Misma Orden de Compra
                            .Where(d => (int)(d.CodInte ?? 0) == brecepen.COD_INTE)  // Mismo producto
                            .Where(d => (d.StatOcom ?? 0) < 8);                      // Ítem no cerrado/anulado

                        if (numOitOc == 0)
                        {
                            // Sin nº de ítem: tomamos el primero que todavía tenga saldo suficiente (con tolerancia de 0.005).
                            var umbral = Math.Abs((cantNomi.Value - 0.005m));
                            queryOcom = queryOcom.Where(d => (d.CantRec ?? 0m) >= umbral);
                        }
                        else
                        {
                            // Con nº de ítem: buscamos exactamente ese renglón de la OC.
                            queryOcom = queryOcom.Where(d => (d.NuOrdItem ?? 0) == numOitOc);
                        }

                        var ocomItem = await queryOcom.FirstOrDefaultAsync();
                        if (ocomItem != null)
                        {
                            // Solo ajustamos si lo aprobado es significativo (> 0.05), igual que el VB6.
                            if (capro > 0.05m)
                            {
                                // COEFUNI = coeficiente de unidades; si es prácticamente cero lo forzamos a 1 para no dividir por 0.
                                var coefUni = ocomItem.CoefUnids ?? 0m;
                                if (coefUni < 0.0001m)
                                {
                                    coefUni = 1m;
                                }

                                // Recalculamos el pendiente a recibir de la OC (misma fórmula del VB6):
                                // CantOReci = Cant_Rec - (CantNom / CoefUni) + (CApro / CoefUni)
                                var cantRecActual = ocomItem.CantRec ?? 0m;
                                var cantOrecu = cantRecActual - (cantNomi.Value / coefUni) + (capro / coefUni);
                                ocomItem.CantRec = cantOrecu;
                                ocomItem.LastUpdate = DateTime.Now;
                            }
                        }
                    }

                    // Descontamos de "restante" lo que acabamos de aprobar en este registro
                    // y seguimos con el siguiente movimiento (si todavía falta repartir).
                    restante -= consumir;
                }

                // Si todo salió bien, guardamos los cambios (nuevo registro 'AP', BOLRECEP, OCOMDETA)
                // y confirmamos la transacción de forma atómica.
                await _context.SaveChangesAsync();
                await tx.CommitAsync();
                return Ok();
            }
            catch (Exception ex)
            {
                await tx.RollbackAsync();
                return BadRequest($"Error al aprobar calidad: {ex.Message}");
            }
        }
    }
}
