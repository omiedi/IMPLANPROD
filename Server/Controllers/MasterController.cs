using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la gestión de clientes (Deudor12)
    /// </summary>
    [ApiController]
    [Route("/api/master")]
    public class MasterController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<MasterController> _logger;
        private readonly ISyncService _syncService;

        public MasterController(DataContext context, ILogger<MasterController> logger, ISyncService syncService)
        {
            _context = context;
            _logger = logger;
            _syncService = syncService;
        }

        /// <summary>
        /// Obtiene una lista paginada de productos (Master) con filtro combinado por '+'
        /// </summary>
        [HttpGet("search")]
        public async Task<ActionResult<List<Master>>> SearchMaster([FromQuery] string filter, [FromQuery] int pageSize = 100)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filter))
                {
                    return new List<Master>();
                }

                // Validar pageSize para evitar consultas excesivas
                if (pageSize > 500) pageSize = 500;
                if (pageSize < 10) pageSize = 10;

                // Filtro combinado con '+': cada término separado por '+' actúa como un AND.
                // Ej: "858 + ADAPTADOR" busca registros que contengan "858" Y "ADAPTADOR"
                // en código y/o descripción.
                var terms = filter
                    .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Trim().ToLower())
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .ToList();

                if (!terms.Any())
                {
                    return new List<Master>();
                }

                var queryable = _context.masters
                    .AsNoTracking()
                    .Where(x => (x.Status ?? 0) >= 0);

                foreach (var term in terms)
                {
                    var like = $"%{term}%";
                    queryable = queryable.Where(x =>
                        EF.Functions.Like((x.Descripcion ?? string.Empty).ToLower(), like) ||
                        EF.Functions.Like((x.Codigo ?? string.Empty).ToLower(), like) ||
                        EF.Functions.Like(x.Codint.ToString(), like)
                    );
                }

                var results = await queryable
                    .OrderBy(x => x.Descripcion)
                    .Take(pageSize)
                    .Select(x => new Master
                    {
                        Codint = x.Codint,
                        Codigo = x.Codigo ?? string.Empty,
                        Descripcion = x.Descripcion ?? string.Empty,
                        Umedida = x.Umedida ?? string.Empty,
                        Status = x.Status,
                        Cosuni = x.Cosuni,
                        Monecos = x.Monecos,
                        Trazable = x.Trazable
                    })
                    .ToListAsync();

                return Ok(results);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en la búsqueda de productos");
                return StatusCode(500, $"Error en la búsqueda: {ex.Message}");
            }
        }

        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination, short status = 0)
        {
            try
            {
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                // var queryable = _context.masters.AsQueryable();
                var queryable = _context.masters.AsNoTracking().AsQueryable();


                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var terms = pagination.Filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(x =>
                            EF.Functions.Like((x.Descripcion ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((x.Codigo ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like(x.Codint.ToString(), like)
                        );
                    }
                }

                if (status >= 0)
                {
                    queryable = queryable.Where(x => x.Status >= status || x.Status == null);
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == status);
                }

                if (!await queryable.AnyAsync())
                {
                    return Ok(new List<Master>());
                }

                var result = await queryable
                    .OrderBy(x => x.Descripcion)
                    .Paginate(pagination)
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la lista de productos");
                return StatusCode(500, $"Error al obtener la lista de productos: {ex.Message}");
            }
        }

        [HttpGet("stock-index")]
        public async Task<ActionResult<List<MasterStockListItemDTO>>> GetStockIndexAsync([FromQuery] PaginationDTO pagination)
        {
            try
            {
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                var queryable = _context.masters
                    .AsNoTracking()
                    .Where(x => (x.Status ?? 0) >= 0);

                // Filtro combinado con '+': cada término separado por '+' actúa como un AND.
                // Ej: "858 + ADAPTADOR" busca registros que contengan "858" Y "ADAPTADOR"
                // en código y/o descripción.
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var terms = pagination.Filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(x =>
                            EF.Functions.Like((x.Descripcion ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((x.Codigo ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like(x.Codint.ToString(), like)
                        );
                    }
                }

                var result = await queryable
                    .OrderBy(x => x.Descripcion)
                    .Select(x => new MasterStockListItemDTO
                    {
                        Codint = x.Codint,
                        Codigo = x.Codigo ?? string.Empty,
                        Descripcion = x.Descripcion ?? string.Empty,
                        Umedida = x.Umedida ?? string.Empty
                    })
                    .Paginate(pagination)
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la lista de productos (StockIndex)");
                return StatusCode(500, $"Error al obtener la lista de productos: {ex.Message}");
            }
        }

        [HttpGet("stock-index/totalPages")]
        public async Task<ActionResult<int>> GetStockIndexPagesAsync([FromQuery] PaginationDTO pagination)
        {
            try
            {
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                var queryable = _context.masters
                    .AsNoTracking()
                    .Where(x => (x.Status ?? 0) >= 0);

                // Filtro combinado con '+': cada término separado por '+' actúa como un AND.
                // Debe replicar exactamente la lógica del endpoint de listado para que las
                // páginas calculadas coincidan con los resultados filtrados.
                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var terms = pagination.Filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(x =>
                            EF.Functions.Like((x.Descripcion ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((x.Codigo ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like(x.Codint.ToString(), like)
                        );
                    }
                }

                var count = await queryable.CountAsync();
                var totalPages = count == 0
                    ? 1
                    : (int)Math.Ceiling(count / (double)pagination.RecordsNumber);

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular totalPages (StockIndex)");
                return StatusCode(500, $"Error al calcular el total de páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Devuelve el total de páginas para la paginación de producto, considerando el filtro y el status
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, short status = 0)
        {
            try
            {
                // Validar que RecordsNumber sea mayor que cero
                if (pagination.RecordsNumber <= 0)
                {
                    return BadRequest("El número de registros por página debe ser mayor que cero.");
                }

                var queryable = _context.masters.AsQueryable();

                if (!string.IsNullOrWhiteSpace(pagination.Filter))
                {
                    var terms = pagination.Filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(x =>
                            EF.Functions.Like((x.Descripcion ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((x.Codigo ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like(x.Codint.ToString(), like)
                        );
                    }
                }

                if (status >= 0)
                {
                    queryable = queryable.Where(x => x.Status >= status || x.Status == null);
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == status);
                }

                // Obtener el conteo total de registros que cumplen con el filtro
                double count = await queryable.CountAsync();

                // Si no hay registros, devolver 1 página (para mostrar el mensaje "No hay registros")
                if (count == 0)
                {
                    return Ok(1);
                }

                // Calcular el número total de páginas
                double totalPages = Math.Ceiling(count / pagination.RecordsNumber);

                // Asegurar que siempre haya al menos una página
                totalPages = Math.Max(1, totalPages);

                return Ok(totalPages);
            }
            catch (Exception ex)
            {
                // Log del error para debugging
                _logger.LogError(ex, "Error al calcular el total de páginas");
                return StatusCode(500, $"Error al calcular el total de páginas: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene un producto individual por Id
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var producto = await _context.masters.FirstOrDefaultAsync(x => x.Id == id);
            if (producto == null)
            {
                return NotFound();
            }
            return Ok(producto);
        }
        /// <summary>
        /// Obtiene un producto individual por Codint
        /// </summary>
        [HttpGet("xCodint{codint:int}")]
        public async Task<IActionResult> GetxCodintAsync(int codint)
        {
            var producto = await _context.masters.FirstOrDefaultAsync(x => x.Codint == codint);
            if (producto == null)
            {
                return NotFound();
            }
            return Ok(producto);
        }

        /// <summary>
        /// Obtiene un producto por su código externo (Codigo) exacto.
        /// Usado por el buscador editable de /StockInfGeneral.
        /// </summary>
        /// <param name="codigo">Código externo del producto</param>
        /// <returns>El producto Master si existe, NotFound si no</returns>
        [HttpGet("por-codigo/{codigo}")]
        public async Task<IActionResult> GetPorCodigoAsync(string codigo)
        {
            if (string.IsNullOrWhiteSpace(codigo))
            {
                return NotFound();
            }

            var producto = await _context.masters
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.Codigo == codigo.Trim());

            if (producto == null)
            {
                return NotFound();
            }

            return Ok(producto);
        }

        /// <summary>
        /// Crea un nuevo producto
        /// </summary>
        [HttpPost]
        public async Task<IActionResult> PostAsync(Master master)
        {
            // Validar que el código no exista
            if (await _context.masters.AnyAsync(x => x.Codigo == master.Codigo))
            {
                return BadRequest($"Ya existe un producto con el código {master.Codigo}");
            }

            // Iniciar una transacción para asegurar la atomicidad de la operación
            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    // Si codint llega en 0, asignar el siguiente valor disponible
                  // VER PIE DE PAGINA SI ESTO DA ERROR SEGURAMENTE LA SECUENCIA NO ESTE BIEN CONFIGURADA EN LA BASE DE DATOS, DEBE HABER UNA SECUENCIA LLAMADA Seq_MasterCodint
                    if (master.Codint == 0)
                    {
                        var connection = _context.Database.GetDbConnection();

                        using var command = connection.CreateCommand();
                        command.CommandText = "SELECT NEXT VALUE FOR Seq_MasterCodint";
                        command.Transaction = _context.Database.CurrentTransaction?.GetDbTransaction();

                        var result = await command.ExecuteScalarAsync();

                        master.Codint = Convert.ToInt32(result);
                    }

                        _context.Add(master);
                    await _context.SaveChangesAsync();

                    // Confirmar la transacción si todo salió bien
                    await transaction.CommitAsync();

                    // Generar archivo de sincronización si está habilitado
                    try
                    {
                        await _syncService.GenerarArchivoSyncMasterAsync(master.Codint);
                    }
                    catch (Exception exSync)
                    {
                        _logger.LogWarning(exSync, "Error al generar archivo de sincronización (no afecta la operación principal)");
                    }

                    // Devuelve una respuesta HTTP 201 (Created) con la URL y el objeto creado
                    //return CreatedAtAction(nameof(GetAsync), new { codint = master.Codint }, master);
                    return Ok(master);
                }
                catch (Exception ex)
                {
                     await transaction.RollbackAsync();

                     return StatusCode(500, $"Ocurrió un error al crear el producto: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Crea un nuevo producto
        /// </summary>
        //[HttpPost]
        //public async Task<IActionResult> PostAsync(Master master)
        //{
        //    // Validar que el código no exista
        //    if (await _context.masters.AnyAsync(x => x.Codigo == master.Codigo))
        //    {
        //        return BadRequest($"Ya existe un producto con el código {master.Codigo}");
        //    }

        //    using (var transaction = await _context.Database.BeginTransactionAsync())
        //    {
        //        try
        //        {
        //            // Si codint llega en 0, asignar el siguiente valor disponible
        //            if (master.Codint == 0)
        //            { // VER PIE DE PAGINA SI ESTO DA ERROR SEGURAMENTE LA SECUENCIA NO ESTE BIEN CONFIGURADA EN LA BASE DE DATOS, DEBE HABER UNA SECUENCIA LLAMADA Seq_MasterCodint
        //                master.Codint = await _context.Database
        //                .SqlQuery<int>($"SELECT NEXT VALUE FOR Seq_MasterCodint")
        //                .FirstAsync();
        //            }

        //            _context.masters.Add(master);
        //            await _context.SaveChangesAsync();

        //            await transaction.CommitAsync();

        //            // Generar archivo de sincronización
        //            try
        //            {
        //                await _syncService.GenerarArchivoSyncMasterAsync(master.Codint);
        //            }
        //            catch (Exception exSync)
        //            {
        //                _logger.LogWarning(exSync, "Error al generar archivo de sincronización (no afecta la operación principal)");
        //            }

        //            // IMPORTANTE: usar el mismo parámetro que el GET
        //            return CreatedAtAction(nameof(GetAsync), new { codint = master.Codint }, master);
        //        }
        //        catch (Exception ex)
        //        {
        //            await transaction.RollbackAsync();

        //            return StatusCode(500, $"Ocurrió un error al crear el producto: {ex.Message}");
        //        }
        //    }
        //}
        [HttpPost("alta-x-copia")]
        public async Task<IActionResult> AltaXCopiaAsync([FromBody] MasterAltaXCopiaDTO dto)
        {
            if (dto == null)
            {
                return BadRequest("Solicitud inválida.");
            }

            if (string.IsNullOrWhiteSpace(dto.NuevoCodigo))
            {
                return BadRequest("El código a dar de alta es requerido.");
            }

            if (string.IsNullOrWhiteSpace(dto.NuevaDescripcion))
            {
                return BadRequest("La descripción es requerida.");
            }

            var nuevoCodigo = dto.NuevoCodigo.Trim();
            var nuevaDescripcion = dto.NuevaDescripcion.Trim();

            if (await _context.masters.AnyAsync(x => x.Codigo == nuevoCodigo))
            {
                return BadRequest($"Ya existe un producto con el código {nuevoCodigo}");
            }

            var baseMaster = await _context.masters.AsNoTracking().FirstOrDefaultAsync(x => x.Codint == dto.BaseCodint);
            if (baseMaster == null)
            {
                return NotFound($"No se encontró el código base con Codint {dto.BaseCodint}.");
            }

            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    // VER PIE DE PAGINA SI ESTO DA ERROR SEGURAMENTE LA SECUENCIA NO ESTE BIEN CONFIGURADA EN LA BASE DE DATOS, DEBE HABER UNA SECUENCIA LLAMADA Seq_MasterCodint

                    var connection = _context.Database.GetDbConnection();

                    using var command = connection.CreateCommand();
                    command.CommandText = "SELECT NEXT VALUE FOR Seq_MasterCodint";
                    command.Transaction = _context.Database.CurrentTransaction?.GetDbTransaction();

                    var result = await command.ExecuteScalarAsync();
                    var newCodint = Convert.ToInt32(result);

                    var nuevo = new Master
                    {
                        Codint = newCodint,
                        Ferevis = baseMaster.Ferevis,
                        CodiEmpr = baseMaster.CodiEmpr ?? (short)0,
                        Observa = baseMaster.Observa,
                        Especificacion = baseMaster.Especificacion,
                        DesCompAut = baseMaster.DesCompAut,
                        DescriComp = baseMaster.DescriComp,
                        DepoEntra = baseMaster.DepoEntra,
                        DepoCon = baseMaster.DepoCon,
                        Codigo = nuevoCodigo,
                        Descripcion = nuevaDescripcion,
                        Descridos = baseMaster.Descridos,
                        Status = baseMaster.Status,
                        Alias = baseMaster.Alias,
                        Umedida = dto.NuevaUnidadMedida ?? baseMaster.Umedida,
                        Timate = dto.NuevoTipoMaterial ?? baseMaster.Timate,
                        Critrep = baseMaster.Critrep,
                        Monecos = baseMaster.Monecos,
                        Fecosto = baseMaster.Fecosto,
                        PorDescCos = baseMaster.PorDescCos,
                        Cosuni = baseMaster.Cosuni,
                        Pornacio = baseMaster.Pornacio,
                        PorComVen = baseMaster.PorComVen,
                        Provhabi = baseMaster.Provhabi,
                        Stomin = baseMaster.Stomin,
                        Stomax = baseMaster.Stomax,
                        Lorepos = baseMaster.Lorepos,
                        Lindivis = baseMaster.Lindivis,
                        Pesuni = baseMaster.Pesuni,
                        PesMetro = baseMaster.PesMetro,
                        Color = baseMaster.Color,
                        Talle = baseMaster.Talle,
                        Descompo = baseMaster.Descompo,
                        Grupiva = baseMaster.Grupiva,
                        Grucoven = baseMaster.Grucoven,
                        CLinea = baseMaster.CLinea,
                        Cfamilia = baseMaster.Cfamilia,
                        NPlano = baseMaster.NPlano,
                        Revision = baseMaster.Revision,
                        Trazable = baseMaster.Trazable,
                        Revi_Formu = baseMaster.Revi_Formu,
                        FechApFor = baseMaster.FechApFor,
                        UsuApFor = baseMaster.UsuApFor,
                        Cod_Interna = baseMaster.Cod_Interna,
                        Bien_Cap = baseMaster.Bien_Cap,
                        Codigo_Ncm = baseMaster.Codigo_Ncm,
                        Codigo_Sec = baseMaster.Codigo_Sec,
                        Pro_Umed = baseMaster.Pro_Umed,
                        Iva_Id = baseMaster.Iva_Id,
                        ArchiPdf = baseMaster.ArchiPdf,
                        Familia = baseMaster.Familia,
                        Lindiven = baseMaster.Lindiven,
                        CanDeci = baseMaster.CanDeci,
                        ConsuProm = baseMaster.ConsuProm,
                        CosCal = baseMaster.CosCal,
                        FeCosCal = baseMaster.FeCosCal,
                        Stat_Movi = baseMaster.Stat_Movi,
                        Clas_ATC = baseMaster.Clas_ATC,
                        Vers_BC = baseMaster.Vers_BC,
                        FeVig_BC = baseMaster.FeVig_BC,
                        FeRev_BC = baseMaster.FeRev_BC,
                        CodiEmprNet = baseMaster.CodiEmprNet,
                        NumUsuar = baseMaster.NumUsuar,
                        RowGuid = Guid.NewGuid(),
                        AddRecord = baseMaster.AddRecord,
                        LastUpdate = baseMaster.LastUpdate
                    };

                    _context.masters.Add(nuevo);
                    await _context.SaveChangesAsync();

                    var formulasBase = await _context.Formulas
                        .AsNoTracking()
                        .Where(f => f.CodIConj == dto.BaseCodint)
                        .ToListAsync();

                    if (formulasBase.Count > 0)
                    {
                        var nuevasFormulas = formulasBase.Select(f => new Formulas
                        {
                            CodIConj = newCodint,
                            CodiEmpr = f.CodiEmpr,
                            CodXConj = nuevoCodigo,
                            CodiElem = f.CodiElem,
                            CodXElem = f.CodXElem,
                            NuorIt = f.NuorIt,
                            UsoBruto = f.UsoBruto,
                            Unis_Form = f.Unis_Form,
                            Lar_Corte = f.Lar_Corte,
                            Pladispo = f.Pladispo,
                            RefeComp = f.RefeComp,
                            Marborra = f.Marborra,
                            UMED_FORM = f.UMED_FORM,
                            COMPONENTE = f.COMPONENTE,
                            MODUXBB = f.MODUXBB,
                            CORTE = f.CORTE,
                            VALOBADA_1 = f.VALOBADA_1,
                            AddRecord = f.AddRecord,
                            LastUpdate = f.LastUpdate,
                            CodiEmprNet = f.CodiEmprNet,
                            NumUsuar = f.NumUsuar
                        }).ToList();

                        _context.Formulas.AddRange(nuevasFormulas);
                        await _context.SaveChangesAsync();
                    }

                    var secopersBase = await _context.Secopers
                        .AsNoTracking()
                        .Where(s => s.CodIConj == dto.BaseCodint)
                        .ToListAsync();

                    if (secopersBase.Count > 0)
                    {
                        var nuevasSecopers = secopersBase.Select(s => new Secoper
                        {
                            CodIConj = newCodint,
                            CodiEmpr = s.CodiEmpr,
                            CodXConj = nuevoCodigo,
                            NumeOper = s.NumeOper,
                            AltOper = s.AltOper,
                            DescOper = s.DescOper,
                            StatOper = s.StatOper,
                            CodMaq0 = s.CodMaq0,
                            CodMaq1 = s.CodMaq1,
                            CodMaq2 = s.CodMaq2,
                            CodMaq3 = s.CodMaq3,
                            CodMaq4 = s.CodMaq4,
                            CodMaq5 = s.CodMaq5,
                            HorsPrep = s.HorsPrep,
                            HorsFijo = s.HorsFijo,
                            MinuStan = s.MinuStan,
                            HorsMHer = s.HorsMHer,
                            HorsLimp = s.HorsLimp,
                            PieCiclo = s.PieCiclo,
                            PorDemor = s.PorDemor,
                            Cadencia = s.Cadencia,
                            CoefSolp = s.CoefSolp,
                            LoteStan = s.LoteStan,
                            CantiOps = s.CantiOps,
                            Cantiops2 = s.Cantiops2,
                            CategOps = s.CategOps,
                            Categops2 = s.Categops2,
                            ProgCNum = s.ProgCNum,
                            HoProFab = s.HoProFab,
                            ObseOper = s.ObseOper,
                            ProTerOp = s.ProTerOp,
                            ProvOper = s.ProvOper,
                            PreUnid = s.PreUnid,
                            MonePrec = s.MonePrec,
                            FechPrec = s.FechPrec,
                            Nuoper_Prep = s.Nuoper_Prep,
                            Catoper_Prep = s.Catoper_Prep,
                            Fech_Revis = s.Fech_Revis,
                            Usua_Revis = s.Usua_Revis,
                            rowguid = Guid.NewGuid(),
                            SecuenciaFabricacion = s.SecuenciaFabricacion,
                            OperacionTerceros = s.OperacionTerceros,
                            NuOrdenManual = s.NuOrdenManual,
                            AddRecord = s.AddRecord,
                            LastUpdate = s.LastUpdate,
                            CodiEmprNet = s.CodiEmprNet,
                            NumUsuar = s.NumUsuar
                        }).ToList();

                        _context.Secopers.AddRange(nuevasSecopers);
                        await _context.SaveChangesAsync();
                    }

                    await transaction.CommitAsync();

                    try
                    {
                        await _syncService.GenerarArchivoSyncMasterAsync(nuevo.Codint);
                    }
                    catch (Exception exSync)
                    {
                        _logger.LogWarning(exSync, "Error al generar archivo de sincronización (no afecta la operación principal)");
                    }

                    return Ok(nuevo);
                }
                catch (Exception ex)
                {
                    await transaction.RollbackAsync();
                    _logger.LogError(ex, "Error en Alta x Copia");
                    return StatusCode(500, $"Ocurrió un error al crear por copia: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Actualiza un producto existente
        /// </summary>
        [HttpPut]
        public async Task<IActionResult> PutAsync(Master master)
        {
            // Iniciar una transacción para asegurar la atomicidad de la operación
            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    // Verificar si el producto existe
                    var productoExistente = await _context.masters.AnyAsync(x => x.Id == master.Id);
                    if (!productoExistente)
                    {
                        return NotFound($"No se encontró el producto con ID {master.Id}");
                    }

                    _context.Update(master);
                    await _context.SaveChangesAsync();

                    // Confirmar la transacción si todo salió bien
                    await transaction.CommitAsync();

                    // Generar archivo de sincronización si está habilitado
                    try
                    {
                        await _syncService.GenerarArchivoSyncMasterAsync(master.Codint);
                    }
                    catch (Exception exSync)
                    {
                        _logger.LogWarning(exSync, "Error al generar archivo de sincronización (no afecta la operación principal)");
                    }

                    return Ok(master);
                }
                catch (Exception ex)
                {
                    // Si ocurre algún error, deshacer la transacción
                    await transaction.RollbackAsync();

                    // Devolver un error 500 con el mensaje de la excepción
                    return StatusCode(500, $"Ocurrió un error al actualizar el producto: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Cambia el estado de un producto
        /// </summary>
        [HttpPut("cambiaStatus/{id:int}/{status:int}")]
        public async Task<IActionResult> CambiaStatusAsync(int id, short status)
        {
            // Iniciar una transacción para asegurar la atomicidad de la operación
            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    var master = await _context.masters.FirstOrDefaultAsync(x => x.Id == id);
                    if (master == null)
                    {
                        return NotFound($"No se encontró el producto con ID {id}");
                    }

                    master.Status = status;
                    _context.Update(master);
                    await _context.SaveChangesAsync();

                    // Confirmar la transacción si todo salió bien
                    await transaction.CommitAsync();
                    return NoContent();
                }
                catch (Exception ex)
                {
                    // Si ocurre algún error, deshacer la transacción
                    await transaction.RollbackAsync();

                    // Devolver un error 500 con el mensaje de la excepción
                    return StatusCode(500, $"Ocurrió un error al cambiar el estado del producto: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Actualiza el campo Especificacion de un producto
        /// </summary>
        /// <param name="id">ID del producto</param>
        /// <param name="especificacion">Texto de especificación</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut("{id}/especificacion")]
        public async Task<IActionResult> UpdateEspecificacionAsync(int id, [FromBody] string especificacion)
        {
            try
            {
                var master = await _context.masters.FirstOrDefaultAsync(m => m.Id == id);
                if (master == null)
                {
                    return NotFound($"No se encontró el producto con ID {id}");
                }

                master.Especificacion = especificacion;
                await _context.SaveChangesAsync();

                return Ok(new { Message = "Especificación actualizada correctamente" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al actualizar la especificación del producto {id}");
                return StatusCode(500, $"Error al actualizar la especificación: {ex.Message}");
            }
        }

        /// <summary>
        /// Busca productos/repuestos en la tabla MASTER por código y descripción
        /// Retorna RepuestoDTO para compatibilidad con el sistema de recepciones
        /// </summary>
        /// <param name="busqueda">Término de búsqueda</param>
        /// <returns>Lista de productos que coinciden con los criterios de búsqueda</returns>
        [HttpGet("buscar")]
        public async Task<ActionResult<List<RepuestoDTO>>> Buscar([FromQuery] string? busqueda)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(busqueda))
                {
                    return BadRequest("El término de búsqueda no puede estar vacío");
                }

                // Buscar en la tabla MASTER por código o descripción
                var productos = await _context.masters
                    .Where(m => (m.Status >= 0 || m.Status == null) &&
                               (m.Codigo!.Contains(busqueda) || 
                                m.Descripcion!.Contains(busqueda)))
                    .OrderBy(m => m.Descripcion)
                    .Take(25)
                    .Select(m => new RepuestoDTO
                    {
                        Codint = m.Codint,
                        CodigoExterno = m.Codigo ?? "",
                        Descritem = m.Descripcion ?? "",
                        Uni_Com = m.Umedida ?? "",
                        UnidadMedidaDescripcion = m.Umedida ?? ""
                    })
                    .ToListAsync();

                _logger.LogInformation($"Productos encontrados en MASTER: {productos.Count}");
                return Ok(productos);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en búsqueda de productos: {ex.Message}");
                _logger.LogError(ex, "Error al buscar productos en MASTER");
                return StatusCode(500, $"Error interno al realizar la búsqueda: {ex.Message}");
            }
        }
        [HttpGet("GetMaster1Dto")]
        public async Task<IActionResult> GetMaster1Dto([FromQuery] short status = 0, [FromQuery] string? filter = null)
        {
            try
            {
                var queryable = _context.masters.AsNoTracking().AsQueryable();

                if (status >= 0)
                {
                    queryable = queryable.Where(x => x.Status >= status || x.Status == null);
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == status);
                }

                // 🔎 Agregamos el filtro textual si fue enviado
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    string filtro = filter.Trim().ToLower();
                    queryable = queryable.Where(x =>
                        x.Descripcion!.ToLower().Contains(filtro) ||
                        x.Codigo!.ToLower().Contains(filtro)
                    // podés agregar más campos si querés
                    );
                }

                if (!await queryable.AnyAsync())
                {
                    return Ok(new List<Master1DTO>());
                }

                var result = await queryable
                    .OrderBy(x => x.Descripcion)
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener la lista de productos");
                return StatusCode(500, $"Error al obtener la lista de productos: {ex.Message}");
            }
        }

        [HttpGet("cantidad-registros")]
        public async Task<ActionResult<int>> GetCantidadRegistros([FromQuery] string? filter = "", short status = 0)
        {
            try
            {
                var queryable = _context.masters.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    var terms = filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(x =>
                            EF.Functions.Like((x.Descripcion ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((x.Codigo ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like(x.Codint.ToString(), like)
                        );
                    }
                }

                if (status >= 0)
                {
                    queryable = queryable.Where(x => x.Status >= status || x.Status == null);
                }
                else
                {
                    queryable = queryable.Where(x => x.Status == status);
                }

                int count = await queryable.CountAsync();
                return Ok(count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener cantidad de registros");
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }
        //listqado de pantalla
        [HttpGet("GetMasterListado")]
        public async Task<IActionResult> GetMasterListado([FromQuery] short status = 0, [FromQuery] string? filter = null)
        {
            try
            {
                // Carga los tipos de material en memoria
                var tatimats = await _context.Tatimat
                    .AsNoTracking()
                    .ToListAsync();

                var queryable = _context.masters
                    .AsNoTracking()
                    .Where(x => status >= 0 ? x.Status >= status || x.Status == null : x.Status == status);

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    var terms = filter
                        .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim().ToLower())
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();

                    foreach (var term in terms)
                    {
                        var like = $"%{term}%";
                        queryable = queryable.Where(x =>
                            EF.Functions.Like((x.Descripcion ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like((x.Codigo ?? string.Empty).ToLower(), like) ||
                            EF.Functions.Like(x.Codint.ToString(), like)
                        );
                    }
                }

                var mastersFiltrados = await queryable
                    .OrderBy(x => x.Descripcion)
                    .ToListAsync();

                // Proyección en memoria usando los datos ya cargados
                var resultado = mastersFiltrados.Select(m => new MasterListadoIndexDTO
                {
                    Id = m.Id.ToString(),
                    Codint = m.Codint.ToString(),
                    Codigo = m.Codigo,
                    Descripcion = m.Descripcion,
                    Umedida = m.Umedida,
                    Timate = m.Timate.ToString(),
                    TipoMaterialDescripcion = m.Timate != null
                        ? tatimats.FirstOrDefault(t => t.CodigoTatimat == m.Timate)?.Descripcion ?? "-"
                        : "-"
                }).OrderBy(x => x.Descripcion).ToList();

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener MasterListadoDTO");
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }

        /// <summary>
        /// Valida si un producto puede ser dado de baja verificando si está en fórmulas
        /// </summary>
        /// <param name="codint">Código interno del producto</param>
        /// <returns>Validación con lista de códigos padres si está en fórmulas</returns>
        [HttpGet("validar-baja/{codint:int}")]
        public async Task<ActionResult<ValidacionBajaMasterDTO>> ValidarBajaAsync(int codint)
        {
            try
            {
                // Buscar fórmulas donde CodIConj (producto padre) tenga este producto como componente (CodiElem)
                var formulasPadres = await _context.Formulas
                    .AsNoTracking()
                    .Where(f => f.CodiElem == codint)
                    .Select(f => f.CodIConj)
                    .Distinct()
                    .ToListAsync();

                var resultado = new ValidacionBajaMasterDTO
                {
                    PuedeDarDeBaja = !formulasPadres.Any()
                };

                if (!resultado.PuedeDarDeBaja)
                {
                    // Obtener los códigos externos de los productos padres
                    var codigosPadres = await _context.masters
                        .AsNoTracking()
                        .Where(m => formulasPadres.Contains(m.Codint))
                        .Select(m => m.Codigo ?? string.Empty)
                        .ToListAsync();

                    resultado.CodigosPadres = codigosPadres;
                }

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al validar baja de producto");
                return StatusCode(500, $"Error al validar baja: {ex.Message}");
            }
        }

        /// <summary>
        /// Realiza baja lógica de un producto (Status = -1)
        /// </summary>
        /// <param name="codint">Código interno del producto</param>
        /// <returns>Resultado de la operación</returns>
        [HttpPut("baja-logica/{codint:int}")]
        public async Task<IActionResult> BajaLogicaAsync(int codint)
        {
            try
            {
                var master = await _context.masters.FirstOrDefaultAsync(x => x.Codint == codint);
                if (master == null)
                {
                    return NotFound($"No se encontró el producto con Codint {codint}");
                }

                master.Status = -1;
                _context.Update(master);
                await _context.SaveChangesAsync();

                return Ok(new { message = "Producto dado de baja exitosamente" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al realizar baja lógica de producto");
                return StatusCode(500, $"Error al realizar baja lógica: {ex.Message}");
            }
        }
    }
}



//CORRECCIONES DE LA TABLA MASTER
//SE CORRIO EL SCRIPT PARA CREAR LA TABLA EN ARCON2 CON EL ESQUEMA Y LOS DATOS Y LA GENERO DE 0 CORRECTAMENTE
//LUEGO SE AGREGO UN CAMPO IDENTITY
//ALTER TABLE MASTER ADD ID INT IDENTITY(1,1);

//ESTO PARA CAMBIAR LOS DATOS DE FLOAT A DECIMAL VERIFICAR SI EL CAMBIO NO AFECTA AL SISTEMA FLEXOFT
//DECLARE @sql NVARCHAR(MAX);

//SELECT @sql = STRING_AGG('ALTER TABLE MASTER ALTER COLUMN ' + COLUMN_NAME + ' DECIMAL(18,2);', ' ')
//FROM INFORMATION_SCHEMA.COLUMNS
//WHERE TABLE_NAME = 'MASTER' AND DATA_TYPE = 'float'; --Cambia 'float' si tu tipo es 'double'

//EXEC sp_executesql @sql;
//   ***** OJO VER LOS CAMPOS QUE USAN 4 DECIMALES ***** SI NO CAMBIARLOS MANUALMENTE, TAMBIEN CORREGIR LA ENTIDAD
// EN *** ARCON **** SE USO DE ESTA MANERA PARA PASAR A DECIMAL LOS FLOAT
//DECLARE @sql NVARCHAR(MAX) = '';

//SELECT @sql = @sql + 'ALTER TABLE MASTER ALTER COLUMN ' + COLUMN_NAME + ' DECIMAL(18,2); ' 
//FROM INFORMATION_SCHEMA.COLUMNS 
//WHERE TABLE_NAME = 'MASTER' AND DATA_TYPE = 'float';

//EXEC sp_executesql @sql;


// **** ---- COMO IMPLEMENTAR LA SECUENCIA EN SQL SERVER SI NO ESTA CREADA ---- ****
//1 PARA OBETENER EL VALOR MAXIMO
//SELECT MAX(Codint) FROM master

//2 CREAR LA SECUENCIA CON EL VALOR MAXIMO OBTENIDO + 1 PARA EVITAR CONFLICTOS, EN ESTE CASO SE OBTUVO 8401, POR LO QUE LA SECUENCIA INICIA EN 8402
/*CREATE SEQUENCE Seq_MasterCodint
AS INT
START WITH 8402
INCREMENT BY 1;
*/

//PARA PROBAR LA SECUENCIA SI SE CREO BIEN
/*SELECT NEXT VALUE FOR Seq_MasterCodint

//PARA ELIMINAR LA SECUENCIA SI SE NECESITA RECREARLA O POR CUALQUIER MOTIVO, PERO CUIDADO CON LOS DATOS EXISTENTES QUE USEN LA SECUENCIA
-DROP SEQUENCE Seq_MasterCodint;  SI SE CREO MAL O SE NECESITA RECREAR POR ALGUNA RAZON PERO LUEGO SE DEBE VOLVER A CREAR SI ESTA EN USO EN EL CODIGO
*/
