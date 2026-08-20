using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Helpers;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("/api/formulas")]
    public class FormulasController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ISyncService _syncService;
        private readonly ILogger<FormulasController> _logger;
        private readonly FlexcrlService _flexcrlService;
        private readonly IDerechosUsuarioService _derechosService;

        public FormulasController(DataContext context, ISyncService syncService, ILogger<FormulasController> logger, FlexcrlService flexcrlService, IDerechosUsuarioService derechosService)
        {
            _context = context;
            _syncService = syncService;
            _logger = logger;
            _flexcrlService = flexcrlService;
            _derechosService = derechosService;
        }

        /// <summary>
        /// Obtiene la lista de familias para el combo (solo activos)
        /// </summary>
        [AllowAnonymous]
        [HttpGet("combo")]
        public async Task<ActionResult> GetCombo()
        {
            return Ok(await _context.Formulas
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene la lista paginada de fórmulas con filtros
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetAsync([FromQuery] PaginationDTO pagination)
        {
            var queryable = _context.Formulas
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => 
                    (x.CodXConj != null && x.CodXConj.ToLower().Contains(pagination.Filter.ToLower())) ||
                    (x.CodXElem != null && x.CodXElem.ToLower().Contains(pagination.Filter.ToLower())));
            }
            return Ok(await queryable
                .OrderBy(x => x.CodIConj)
                .ThenBy(x => x.NuorIt)
                .Paginate(pagination)
                .ToListAsync());
        }

        /// <summary>
        /// Obtiene todos los componentes de un producto terminado
        /// </summary>//***** ESTO LO USA MASTER COMPONENTS ****
        /// 
        [HttpGet("conjunto/{codigoConjunto:int}")]
        public async Task<IActionResult> GetComponentesConjuntoAsync(int codigoConjunto)
        {
            try
            {
                var componentes = await _context.Formulas
             .Where(f => f.CodIConj == codigoConjunto)
             .OrderBy(f => f.NuorIt)
             .ToListAsync();

                return Ok(componentes);
            }
            catch (Exception ex)
            {
                throw;
            } 
         
        }

        /// <summary>
        /// Guarda o actualiza los componentes de un producto terminado
        /// </summary>
        [HttpPost("conjunto/{codigoConjunto:int}")]
        public async Task<IActionResult> SaveComponentesConjuntoAsync(int codigoConjunto, [FromBody] List<Formulas> componentes)
        {
            using var transaction = await _context.Database.BeginTransactionAsync(); // Iniciar transacción

            try
            {
                // Eliminar componentes existentes
                var componentesExistentes = await _context.Formulas
                    .Where(f => f.CodIConj == codigoConjunto)
                    .ToListAsync();

                _context.Formulas.RemoveRange(componentesExistentes);

                // Agregar nuevos componentes
                foreach (var componente in componentes)
                {
                    componente.Id = 0; // Asegurar que se cree como nuevo
                    componente.CodIConj = codigoConjunto; // Asignar el código correcto
                    var cantidad = componente.UsoBruto ?? componente.Unis_Form ?? 0m;
                    componente.UsoBruto = cantidad;
                    componente.Unis_Form = cantidad;
                    _context.Formulas.Add(componente);
                }

                await _context.SaveChangesAsync(); // Guardar cambios
                await transaction.CommitAsync();  // Confirmar la transacción

                // Generar archivo de sincronización si está habilitado
                try
                {
                    await _syncService.GenerarArchivoSyncFormulasAsync(codigoConjunto);
                }
                catch (Exception syncEx)
                {
                    _logger.LogError(syncEx, $"Error al generar archivo de sincronización para Formulas CodIConj={codigoConjunto}");
                }

                return Ok();
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync(); // Deshacer los cambios si hay un error
                return BadRequest($"Error al guardar los componentes: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el total de páginas para la paginación
        /// </summary>
        [HttpGet("totalPages")]
        public async Task<ActionResult> GetPages([FromQuery] PaginationDTO pagination, int estado = 0)
        {
            var queryable = _context.Formulas.AsQueryable();

            // Aplicar filtro si existe
            if (!string.IsNullOrWhiteSpace(pagination.Filter))
            {
                queryable = queryable.Where(x => x.CodXElem!.ToLower().Contains(pagination.Filter.ToLower()));
            }

            double count = await queryable.CountAsync();
            double totalPages = Math.Ceiling(count / pagination.RecordsNumber);
            return Ok(totalPages);
        }

        /// <summary>
        /// Obtiene una Formulas por su ID
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetAsync(int id)
        {
            var formula = await _context.Formulas
                .FirstOrDefaultAsync(x => x.Id == id);
            if (formula == null)
            {
                return NotFound();
            }
            return Ok(formula);
        }

        /// <summary>
        /// Crea una nueva formula
        /// </summary>
        [HttpPost]
        public async Task<ActionResult> PostAsync(Formulas formula)
        {
            try
            {
                // Verificar si ya existe una formula con el mismo código
                var existingFormula = await _context.Formulas
                    .FirstOrDefaultAsync(x => x.CodXElem == formula.CodXElem);

                if (existingFormula != null)
                {
                    return BadRequest($"Ya existe una formula con el código {formula.CodXElem}");
                }

                _context.Add(formula);
                await _context.SaveChangesAsync();
                return Ok(formula);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }
        /// <summary>
        /// Actualiza una formula existente
        /// </summary>
        [HttpPut]
        public async Task<ActionResult> PutAsync(Formulas formula)
        {
            try
            {
                // Verificar si ya existe una formula con el mismo código (excluyendo la actual)
                var existingFormula = await _context.Formulas
                    .FirstOrDefaultAsync(x => x.CodXElem == formula.CodXElem && x.Id != formula.Id);

                if (existingFormula != null)
                {
                    return BadRequest($"Ya existe una formula con el código {formula.CodXElem}");
                }

                _context.Update(formula);
                await _context.SaveChangesAsync();
                return Ok(formula);
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }
        }

        /// <summary>
        /// Obtiene la implosión de un producto: lista de productos (conjuntos) en los que el ítem es componente.
        /// Equivalente a la consulta VB6:
        /// SELECT codiconj, codielem, usobruto, codint, descripcion, codigo, status, umedida
        /// FROM FORMULAS INNER JOIN MASTER ON codiconj=codint
        /// WHERE codielem = {codigoInternoProducto} AND status > 0
        /// </summary>
        /// <param name="codigoInternoProducto">Código interno del ítem (CodIElem)</param>
        /// <returns>Lista de ImplosionItemDTO con uso bruto ajustado según PORMILES</returns>
        [HttpGet("implosion/{codigoInternoProducto:int}")]
        public async Task<IActionResult> GetImplosionAsync(int codigoInternoProducto)
        {
            try
            {
                // Obtener el control PORMILES del usuario actual
                // Equivalente a HABIPORMI% en VB6
                bool habilitarPorMiles = false;
                try
                {
                    var controlPormiles = await _flexcrlService.ObtenerControlAsync("AMASTO", "PORMILES");
                    if (controlPormiles != null)
                    {
                        var valor = controlPormiles.VALCONTROL?.Trim();
                        habilitarPorMiles = valor == "1" || valor?.ToUpper() == "SI";
                    }
                }
                catch
                {
                    // Si no se puede obtener el control, se asume deshabilitado
                    habilitarPorMiles = false;
                }

                // Buscar fórmulas donde el ítem actual es componente (CodIElem)
                // y unir con Master para obtener datos del conjunto (CodIConj = Codint)
                var query = from f in _context.Formulas
                            join m in _context.masters on f.CodIConj equals m.Codint
                            where f.CodiElem == codigoInternoProducto && (m.Status ?? 0) > 0
                            orderby f.CodIConj, f.NuorIt
                            select new { f, m };

                var resultados = await query.ToListAsync();

                if (!resultados.Any())
                {
                    return NotFound($"No se encontró implosión para el producto con código interno {codigoInternoProducto}");
                }

                var lista = new List<ImplosionItemDTO>();
                foreach (var r in resultados)
                {
                    decimal usoBruto = r.f.UsoBruto ?? 0;
                    bool ajustadoPorMil = false;

                    // Lógica PORMILES: si está habilitado, el uso bruto > 0 y < 0.005, multiplicar por 1000
                    if (habilitarPorMiles && usoBruto > 0 && usoBruto < 0.005m)
                    {
                        usoBruto = 1000 * usoBruto;
                        ajustadoPorMil = true;
                    }

                    lista.Add(new ImplosionItemDTO
                    {
                        CodIConj = r.f.CodIConj ?? 0,
                        Codigo = r.m.Codigo ?? string.Empty,
                        Descripcion = r.m.Descripcion ?? "Sin descripción",
                        Umedida = r.m.Umedida ?? string.Empty,
                        UsoBruto = usoBruto,
                        AjustadoPorMil = ajustadoPorMil
                    });
                }

                return Ok(lista);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener implosión para CodIElem={CodIElem}", codigoInternoProducto);
                return BadRequest($"Error al obtener la implosión: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene la fórmula de componentes de un producto con stock y cálculos adicionales
        /// </summary>
        [HttpGet("componentes/{codigoInternoProducto:int}")]
        public async Task<IActionResult> GetComponentesConStockAsync(int codigoInternoProducto)
        {
            try
            {
                // Obtener los componentes de la fórmula
                var componentes = await _context.Formulas
                    .Where(f => f.CodIConj == codigoInternoProducto)
                    .OrderBy(f => f.NuorIt)
                    .ToListAsync();

                if (componentes == null || !componentes.Any())
                {
                    return NotFound($"No se encontró fórmula para el producto con código interno {codigoInternoProducto}");
                }

                // Obtener todos los códigos de elementos para verificar repetidos
                var codigosElementos = componentes.Select(c => c.CodiElem).ToList();
                var codigosRepetidos = codigosElementos.GroupBy(c => c).Where(g => g.Count() > 1).Select(g => g.Key).ToList();

                // Crear lista de DTOs con información adicional
                var resultado = new List<FormulaComponenteDTO>();

                foreach (var componente in componentes)
                {
                    if (componente.CodiElem == null) continue;

                    // Buscar el master correspondiente para obtener la descripción
                    var master = await _context.masters
                        .FirstOrDefaultAsync(m => m.Codint == componente.CodiElem);

                  

                    // Crear el DTO con la información
                    var componenteDTO = new FormulaComponenteDTO
                    {
                        CodIElem = componente.CodiElem ?? 0,
                        CodXElem = componente.CodXElem ?? string.Empty,
                        Descripcion = master?.Descripcion ?? "Sin descripción",
                        UsoBruto = componente.UsoBruto ?? 0,
                        EsRepetido = codigosRepetidos.Contains(componente.CodiElem),
                        NecesidadTotal = 0, // Se calculará en el frontend con la cantidad a fabricar
                        CantidadAUtilizar = 0 // Se inicializará en el frontend
                    };

                    resultado.Add(componenteDTO);
                }
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener los componentes: {ex.Message}");
            }
        }

        /// <summary>
        /// Elimina una formula (borrado físico)
        /// </summary>
        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var formula = await _context.Formulas.FirstOrDefaultAsync(x => x.Id == id);
            if (formula == null)
            {
                return NotFound();
            }
            _context.Remove(formula);
            await _context.SaveChangesAsync();
            return NoContent();
        }

        /// <summary>
        /// Marca una formula como inactiva (borrado lógico)
        /// </summary>
        //[HttpPut("baja/{id:int}")]
        //public async Task<IActionResult> FamiliaStatusMenosUnoAsync(int id)
        //{
        //    var formula = await _context.Formulas.FirstOrDefaultAsync(x => x.Id == id);
        //    if (formula == null)
        //    {
        //        return NotFound();
        //    }
        //    formula.Status = -1;
        //    _context.Update(formula);
        //    await _context.SaveChangesAsync();
        //    return NoContent();
        //}
        //______________________________________________________________________
        //[HttpDelete("cambiaStatus/{id:int}/{status:int}")]
        //public async Task<IActionResult> MarcaStatusMenosUnoAsync(int id, int status = 0)
        //{
        //    var familia = await _context.Familias.FirstOrDefaultAsync(x => x.Id == id);
        //    if (familia == null)
        //    {
        //        return NotFound();
        //    }

        //    // Cambiar el campo status 
        //    familia.Status = status;

        //    // Guardar los cambios en la base de datos
        //    await _context.SaveChangesAsync();

        //    return NoContent();
        //}

        /// <summary>
        /// Valida si se debe bloquear la modificación de la fórmula según las reglas de negocio
        /// Adaptado del código VB6 DebeBloquearModificacionFormula
        /// </summary>
        /// <param name="codiInt">Código interno del producto</param>
        /// <param name="numUsuar">IdFlexoft del usuario activo, usado para validar el derecho FORSUPLI en DER_ACCE</param>
        /// <returns>Resultado de la validación</returns>
        [HttpGet("validar-bloqueo/{codiInt:int}")]
        public async Task<ActionResult<FormulaValidationResultDTO>> ValidarBloqueoModificacionFormula(int codiInt, [FromQuery] int? numUsuar)
        {
            var resultado = new FormulaValidationResultDTO { DebeBloquear = false };

            try
            {
                // PASO 1: Código inválido
                if (codiInt <= 0)
                {
                    return Ok(resultado);
                }

                // PASO 2: Si no hay límite configurado, no se controla
                var limiteModFormulaDiasConfig = await _flexcrlService.ObtenerControlAsync("PEDIDO", "LIMMOEST");
                int limiteDias = 0;
                if (limiteModFormulaDiasConfig != null && int.TryParse(limiteModFormulaDiasConfig.VALCONTROL, out int valorLimite))
                {
                    limiteDias = valorLimite;
                }

                if (limiteDias <= 0)
                {
                    return Ok(resultado);
                }

                // PASO 3: Si el usuario tiene el derecho puntual FORSUPLI en DER_ACCE, no se controla
                // Equivale al VB6: DERACCE%("FORSUPLI/NOMESS", "MODIFICAR FORMULA CON LIMITE CUMPLIDO") >= 2
                // El nivel 2 corresponde a DER_WRITE = 1 (derecho de escritura)
                if (numUsuar.HasValue && numUsuar.Value > 0)
                {
                    var nivelAcceso = await _derechosService.VerificarDerechoAsync(
                        (short)numUsuar.Value,
                        "FORSUPLI",
                        "MODIFICAR FORMULA CON LIMITE CUMPLIDO");

                    if (nivelAcceso >= 2)
                    {
                        return Ok(resultado);
                    }
                }

                // PASO 4: Obtener el modo de bloqueo configurado
                var modoBloqueoConfig = await _flexcrlService.ObtenerControlAsync("PEDIDO", "MODOBLOQUEO");
                int modoBloqueo = 1; // Por defecto: Primer nivel
                if (modoBloqueoConfig != null && int.TryParse(modoBloqueoConfig.VALCONTROL, out int valorModo))
                {
                    modoBloqueo = valorModo;
                }

                // Fecha del servidor de base de datos (nunca la fecha local del cliente)
                var fechaServidor = await ObtenerFechaServidorAsync();

                // PASO 5: Valida el propio artículo
                var ultimoPedido = await ObtenerUltimoPedidoAsync(codiInt);

                if (ultimoPedido != null)
                {
                    var diasDesdeUltimoPed = (fechaServidor.Date - ultimoPedido.FechEmis.Date).Days;

                    if (diasDesdeUltimoPed > limiteDias)
                    {
                        resultado.DebeBloquear = true;
                        resultado.Mensaje = $"El producto tiene su último pedido hace {diasDesdeUltimoPed} días. (Límite permitido: {limiteDias} días)";
                        resultado.FechaUltimoPedido = ultimoPedido.FechEmis;
                        resultado.DiasDesdeUltimoPedido = diasDesdeUltimoPed;
                        resultado.NroPedido = ultimoPedido.NroPed;
                        resultado.NroCliente = ultimoPedido.NroClie;
                        resultado.RazonSocial = ultimoPedido.RaSoClie;
                        return Ok(resultado);
                    }
                }

                // PASO 6: Valida los padres según el modo de bloqueo
                if (modoBloqueo >= 2)
                {
                    var listaPadres = await ObtenerListaPadresAsync(codiInt, modoBloqueo);

                    if (listaPadres.Count > 0)
                    {
                        var resultadoPadres = await TienePadresConPedidoVencidoAsync(listaPadres, limiteDias, fechaServidor);

                        if (resultadoPadres != null)
                        {
                            resultado.DebeBloquear = true;
                            resultado.Mensaje = resultadoPadres.Mensaje;
                            resultado.CodigoPadre = resultadoPadres.CodigoPadre;
                            resultado.DescripcionPadre = resultadoPadres.DescripcionPadre;
                            resultado.FechaUltimoPedido = resultadoPadres.FechaUltimoPedido;
                            resultado.DiasDesdeUltimoPedido = resultadoPadres.DiasDesdeUltimoPedido;
                            resultado.NroPedido = resultadoPadres.NroPedido;
                            resultado.NroCliente = resultadoPadres.NroCliente;
                            resultado.RazonSocial = resultadoPadres.RazonSocial;
                            resultado.EstructuraFormula = await ConstruirEstructuraFormulaAsync(codiInt, resultadoPadres.CodigoPadre);
                        }
                    }
                }

                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al validar bloqueo de fórmula para CodInt={CodInt}", codiInt);
                return BadRequest($"Error al validar bloqueo: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene la fecha actual del servidor de base de datos (GETDATE()).
        /// Equivalente a la función NOWSERVER del VB6.
        /// Nunca debe usarse la fecha local de la máquina cliente para estas validaciones.
        /// </summary>
        private async Task<DateTime> ObtenerFechaServidorAsync()
        {
            return await _context.Database
                .SqlQuery<DateTime>($"SELECT GETDATE() AS Value")
                .SingleAsync();
        }

        /// <summary>
        /// Obtiene el último pedido de venta registrado para un producto, con los datos del cliente
        /// Adaptado del código VB6 DiasDesdeUltimoPedido
        /// </summary>
        /// <param name="codiInt">Código interno del producto</param>
        /// <returns>Datos del último pedido, o null si el producto no tiene pedidos</returns>
        private async Task<UltimoPedidoInfo?> ObtenerUltimoPedidoAsync(int codiInt)
        {
            return await _context.Peddetas
                .Where(d => d.CodiInt == codiInt
                            && d.Status >= 0
                            && d.Status < 8
                            && d.FechEmis != null)
                .OrderByDescending(d => d.FechEmis)
                .Select(d => new UltimoPedidoInfo
                {
                    FechEmis = d.FechEmis!.Value,
                    NroPed = d.NroPed,
                    NroClie = d.NroClie,
                    RaSoClie = d.RaSoClie
                })
                .FirstOrDefaultAsync();
        }

        /// <summary>
        /// Datos del último pedido de venta de un producto, usados para informar el motivo del bloqueo
        /// </summary>
        private sealed class UltimoPedidoInfo
        {
            public DateTime FechEmis { get; set; }
            public int? NroPed { get; set; }
            public int? NroClie { get; set; }
            public string? RaSoClie { get; set; }
        }

        /// <summary>
        /// Obtiene la lista de padres de un producto según el modo de bloqueo configurado
        /// Adaptado del código VB6 ObtenerListaPadres
        /// </summary>
        /// <param name="codiInt">Código interno del producto</param>
        /// <param name="modoBloqueo">Modo de bloqueo: 1=Primer nivel, 2=Primer y segundo nivel, 3=Todos los niveles</param>
        /// <returns>Lista de códigos internos de los padres a validar</returns>
        private async Task<List<int>> ObtenerListaPadresAsync(int codiInt, int modoBloqueo)
        {
            // Modo 2: Primer y segundo nivel (hasta 2 niveles hacia arriba en el árbol de estructura)
            // Modo 3: Todos los niveles (recursivo, toda la cadena de ancestros)

            if (modoBloqueo == 2)
            {
                // ESCENARIO B: Validar hasta 2 niveles hacia arriba
                // Nivel 1: Padres directos del producto activo
                var padresNivel1 = await ObtenerPadresDirectosAsync(codiInt);

                // Nivel 2: Padres de los padres (abuelos del producto activo)
                var padresNivel2 = new List<int>();
                foreach (var padre in padresNivel1)
                {
                    var padresDePadre = await ObtenerPadresDirectosAsync(padre);
                    padresNivel2.AddRange(padresDePadre);
                }

                // Combinar niveles 1 y 2, eliminando duplicados
                return padresNivel1.Concat(padresNivel2).Distinct().ToList();
            }
            else // Modo 3: Todos los niveles recursivamente
            {
                // ESCENARIO C: Validar toda la genealogía recursivamente
                // Recorre todos los ancestros hasta la raíz del árbol
                return await ObtenerPadresRecursivoAsync(codiInt, new HashSet<int>());
            }
        }

        /// <summary>
        /// Obtiene los padres directos de un producto (nivel 1 en el árbol de estructura)
        /// </summary>
        /// <param name="codiInt">Código interno del producto</param>
        /// <returns>Lista de códigos de padres directos</returns>
        private async Task<List<int>> ObtenerPadresDirectosAsync(int codiInt)
        {
            return await _context.Formulas
                .Where(f => f.CodiElem == codiInt && f.CodIConj != null)
                .Select(f => f.CodIConj!.Value)
                .Distinct()
                .ToListAsync();
        }

        /// <summary>
        /// Obtiene todos los padres recursivamente (para modo 3 - Todos los niveles)
        /// Implementa búsqueda recursiva completa de toda la cadena de ancestros
        /// </summary>
        /// <param name="codiInt">Código interno del producto</param>
        /// <param name="visitados">HashSet para evitar ciclos infinitos y duplicados</param>
        /// <returns>Lista de todos los códigos de padres en la cadena de ancestros</returns>
        private async Task<List<int>> ObtenerPadresRecursivoAsync(int codiInt, HashSet<int> visitados)
        {
            // Obtener padres directos del producto actual
            var padres = await ObtenerPadresDirectosAsync(codiInt);
            var resultado = new List<int>();

            foreach (var padre in padres)
            {
                // Si este padre no ha sido visitado, procesarlo
                if (!visitados.Contains(padre))
                {
                    visitados.Add(padre);
                    resultado.Add(padre);

                    // Recursivamente obtener padres de este padre (continuar hacia arriba en el árbol)
                    var padresRecursivos = await ObtenerPadresRecursivoAsync(padre, visitados);
                    resultado.AddRange(padresRecursivos);
                }
            }

            return resultado;
        }

        /// <summary>
        /// Valida si alguno de los padres tiene su último pedido fuera del límite de días permitido
        /// Adaptado del código VB6 TienePadresConPedidoVencido
        /// </summary>
        /// <param name="padres">Códigos internos de los padres a validar</param>
        /// <param name="limiteDias">Límite de días configurado en FLEXCRL (PEDIDO/LIMMOEST)</param>
        /// <param name="fechaServidor">Fecha obtenida del servidor de base de datos</param>
        /// <returns>Resultado con el detalle del padre vencido, o null si ninguno está vencido</returns>
        private async Task<FormulaValidationResultDTO?> TienePadresConPedidoVencidoAsync(List<int> padres, int limiteDias, DateTime fechaServidor)
        {
            if (padres == null || padres.Count == 0)
            {
                return null;
            }

            // Pedidos de todos los padres, con los datos del cliente (PEDDETA.NroClie / PEDDETA.RaSoClie)
            var pedidosPadres = await (from d in _context.Peddetas
                                       join m in _context.masters on d.CodiInt equals (int?)m.Codint
                                       where d.CodiInt != null
                                             && padres.Contains(d.CodiInt.Value)
                                             && d.Status >= 0
                                             && d.Status < 8
                                             && d.FechEmis != null
                                       select new
                                       {
                                           CodiInt = d.CodiInt!.Value,
                                           FechEmis = d.FechEmis!.Value,
                                           d.NroPed,
                                           d.NroClie,
                                           d.RaSoClie,
                                           m.Codigo,
                                           m.Descripcion
                                       })
                                       .ToListAsync();

            // Último pedido de cada padre (equivale al GROUP BY con MAX(Fechemis) del VB6)
            // y luego el HAVING DATEDIFF(DAY, MAX(Fechemis), GETDATE()) > limiteDias
            // ordenado por la fecha más antigua (ORDER BY MAX(Fechemis) ASC)
            var padreVencido = pedidosPadres
                .GroupBy(x => x.CodiInt)
                .Select(g => g.OrderByDescending(x => x.FechEmis).First())
                .Where(x => (fechaServidor.Date - x.FechEmis.Date).Days > limiteDias)
                .OrderBy(x => x.FechEmis)
                .FirstOrDefault();

            if (padreVencido == null)
            {
                return null;
            }

            var dias = (fechaServidor.Date - padreVencido.FechEmis.Date).Days;

            return new FormulaValidationResultDTO
            {
                DebeBloquear = true,
                Mensaje = $"No se puede modificar la fórmula. El producto padre: {padreVencido.Codigo} - {padreVencido.Descripcion} tiene su último pedido del día: {padreVencido.FechEmis:dd/MM/yyyy}. Hace {dias} días. (Límite permitido: {limiteDias} días)",
                CodigoPadre = padreVencido.Codigo,
                DescripcionPadre = padreVencido.Descripcion,
                FechaUltimoPedido = padreVencido.FechEmis,
                DiasDesdeUltimoPedido = dias,
                NroPedido = padreVencido.NroPed,
                NroCliente = padreVencido.NroClie,
                RazonSocial = padreVencido.RaSoClie,
                EstructuraFormula = await ConstruirEstructuraFormulaAsync(padreVencido.CodiInt, padreVencido.Codigo)
            };
        }

        /// <summary>
        /// Construye la cadena de estructura de fórmula desde el padre vencido hasta el código activo
        /// Ejemplo: "PADRE > HIJO1 > HIJO2 > HIJO3 (CÓDIGO ACTIVO)"
        /// Usa búsqueda en profundidad (DFS) desde el padre vencido hacia abajo para encontrar el camino correcto
        /// </summary>
        /// <param name="codiIntActivo">Código interno del producto activo (el que se está editando)</param>
        /// <param name="codigoPadreVencido">Código del producto padre que está vencido</param>
        /// <returns>Cadena de estructura o null si no se pudo construir</returns>
        private async Task<string?> ConstruirEstructuraFormulaAsync(int codiIntActivo, string? codigoPadreVencido)
        {
            if (string.IsNullOrEmpty(codigoPadreVencido))
            {
                return null;
            }

            // Obtener el código interno del padre vencido
            var codiIntPadreVencido = await _context.masters
                .Where(m => m.Codigo == codigoPadreVencido)
                .Select(m => m.Codint)
                .FirstOrDefaultAsync();

            if (codiIntPadreVencido == 0)
            {
                return null;
            }

            // Obtener el código del producto activo
            var productoActivo = await _context.masters
                .Where(m => m.Codint == codiIntActivo)
                .Select(m => m.Codigo)
                .FirstOrDefaultAsync();

            if (string.IsNullOrEmpty(productoActivo))
            {
                return null;
            }

            // Buscar el camino desde el padre vencido hacia abajo hasta el código activo usando DFS
            var camino = await BuscarCaminoHaciaActivoAsync(codiIntPadreVencido, codiIntActivo, new HashSet<int>());

            if (camino == null || camino.Count == 0)
            {
                return null;
            }

            // Construir la cadena de códigos
            var codigos = new List<string> { codigoPadreVencido };
            foreach (var codiInt in camino)
            {
                var codigo = await _context.masters
                    .Where(m => m.Codint == codiInt)
                    .Select(m => m.Codigo)
                    .FirstOrDefaultAsync();
                if (!string.IsNullOrEmpty(codigo))
                {
                    codigos.Add(codigo);
                }
            }

            var resultado = string.Join(" > ", codigos);
            return $"{resultado} (CÓDIGO ACTIVO)";
        }

        /// <summary>
        /// Busca el camino desde un nodo hacia abajo hasta encontrar el código activo usando DFS
        /// </summary>
        /// <param name="codiIntActual">Código interno del nodo actual</param>
        /// <param name="codiIntObjetivo">Código interno del producto activo (objetivo)</param>
        /// <param name="visitados">HashSet para evitar ciclos</param>
        /// <returns>Lista de códigos internos en el camino hacia el objetivo, o null si no se encontró</returns>
        private async Task<List<int>?> BuscarCaminoHaciaActivoAsync(int codiIntActual, int codiIntObjetivo, HashSet<int> visitados)
        {
            if (codiIntActual == codiIntObjetivo)
            {
                return new List<int>(); // Encontrado, camino vacío (el último nodo es el objetivo)
            }

            if (visitados.Contains(codiIntActual))
            {
                return null; // Ciclo detectado
            }

            visitados.Add(codiIntActual);

            // Obtener hijos directos del nodo actual (donde CodIConj = codiIntActual)
            var hijos = await _context.Formulas
                .Where(f => f.CodIConj == codiIntActual && f.CodiElem != null)
                .Select(f => f.CodiElem!.Value)
                .Distinct()
                .ToListAsync();

            foreach (var hijo in hijos)
            {
                var camino = await BuscarCaminoHaciaActivoAsync(hijo, codiIntObjetivo, visitados);
                if (camino != null)
                {
                    // Encontramos el camino a través de este hijo
                    var resultado = new List<int> { hijo };
                    resultado.AddRange(camino);
                    return resultado;
                }
            }

            return null; // No se encontró camino a través de ningún hijo
        }
    }
}
