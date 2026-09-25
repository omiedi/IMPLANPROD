using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;
using System.Text;
using System.Text.Json;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio de sincronización entre Central y Sucursal
    /// </summary>
    public class SyncService : ISyncService
    {
        private readonly DataContext _context;
        private readonly IConfiguration _configuration;
        private readonly ILogger<SyncService> _logger;
        private readonly AnotacionesPedidoService _anotacionesService;

        public SyncService(DataContext context, IConfiguration configuration, ILogger<SyncService> logger, AnotacionesPedidoService anotacionesService)
        {
            _context = context;
            _configuration = configuration;
            _logger = logger;
            _anotacionesService = anotacionesService;
        }

        /// <summary>
        /// Genera un archivo JSON de sincronización para un conjunto de Formulas
        /// </summary>
        public async Task<bool> GenerarArchivoSyncFormulasAsync(int codiConj)
        {
            try
            {
                var config = ObtenerConfiguracion();
                if (!config.GeneraScriptSync || !config.GeneraProductos)
                {
                    return false;
                }

                var rutaIntercambio = config.RutaIntercambio;
                if (string.IsNullOrEmpty(rutaIntercambio) || !Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio no válida: {rutaIntercambio}");
                    return false;
                }

                var componentes = await _context.Formulas
                    .Where(f => f.CodIConj == codiConj)
                    .Select(f => new FormulaItemDTO
                    {
                        CodIConj = f.CodIConj,
                        CodiEmpr = f.CodiEmpr,
                        CodXConj = f.CodXConj,
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
                        VALOBADA_1 = f.VALOBADA_1
                    })
                    .ToListAsync();

                if (componentes.Count == 0)
                {
                    _logger.LogWarning($"No se encontraron componentes para el conjunto {codiConj}");
                    return false;
                }

                var syncDto = new SyncFormulasDTO
                {
                    Tabla = "FORMULAS",
                    Operacion = "UPSERT",
                    FechaGeneracion = DateTime.Now,
                    CodIConj = codiConj,
                    Componentes = componentes
                };

                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var nombreArchivo = $"FORMULAS_{codiConj}_{timestamp}.json";
                var rutaArchivo = Path.Combine(rutaIntercambio, nombreArchivo);

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = null
                };

                var json = JsonSerializer.Serialize(syncDto, options);
                await File.WriteAllTextAsync(rutaArchivo, json);

                _logger.LogInformation($"Archivo de sincronización generado: {nombreArchivo}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al generar archivo de sincronización para Formulas {codiConj}");
                return false;
            }
        }

        /// <summary>
        /// Genera un archivo JSON de sincronización para un conjunto de SecOper
        /// </summary>
        public async Task<bool> GenerarArchivoSyncSecOperAsync(int codiConj)
        {
            try
            {
                var config = ObtenerConfiguracion();
                if (!config.GeneraScriptSync || !config.GeneraProductos)
                {
                    return false;
                }

                var rutaIntercambio = config.RutaIntercambio;
                if (string.IsNullOrEmpty(rutaIntercambio) || !Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio no válida: {rutaIntercambio}");
                    return false;
                }

                var operaciones = await _context.Secopers
                    .Where(s => s.CodIConj == codiConj)
                    .Select(s => new SecOperItemDTO
                    {
                        CodIConj = s.CodIConj,
                        CodiEmpr = s.CodiEmpr,
                        CodXConj = s.CodXConj,
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
                        rowguid = s.rowguid,
                        SecuenciaFabricacion = s.SecuenciaFabricacion,
                        OperacionTerceros = s.OperacionTerceros,
                        NuOrdenManual = s.NuOrdenManual
                    })
                    .ToListAsync();

                if (operaciones.Count == 0)
                {
                    _logger.LogWarning($"No se encontraron operaciones para el conjunto {codiConj}");
                    return false;
                }

                var syncDto = new SyncSecOperDTO
                {
                    Tabla = "SECOPER",
                    Operacion = "UPSERT",
                    FechaGeneracion = DateTime.Now,
                    CodIConj = codiConj,
                    Operaciones = operaciones
                };

                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var nombreArchivo = $"SECOPER_{codiConj}_{timestamp}.json";
                var rutaArchivo = Path.Combine(rutaIntercambio, nombreArchivo);

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = null
                };

                var json = JsonSerializer.Serialize(syncDto, options);
                await File.WriteAllTextAsync(rutaArchivo, json);

                _logger.LogInformation($"Archivo de sincronización generado: {nombreArchivo}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al generar archivo de sincronización para SecOper {codiConj}");
                return false;
            }
        }

        /// <summary>
        /// Genera archivo JSON de sincronización para un registro Master (Central)
        /// </summary>
        public async Task<bool> GenerarArchivoSyncMasterAsync(int codint)
        {
            try
            {
                var config = ObtenerConfiguracion();

                // Logging de diagnóstico: mostrar el estado de la configuración
                // para facilitar la detección de por qué no se generan archivos.
                _logger.LogInformation(
                    "GenerarArchivoSyncMasterAsync(Codint={Codint}) - Config: GeneraScriptSync={GeneraScriptSync}, GeneraProductos={GeneraProductos}, RutaIntercambio='{RutaIntercambio}'",
                    codint, config.GeneraScriptSync, config.GeneraProductos, config.RutaIntercambio);

                if (!config.GeneraScriptSync || !config.GeneraProductos)
                {
                    _logger.LogWarning(
                        "No se genera archivo de sincronización para Master.Codint={Codint} porque la configuración lo deshabilita " +
                        "(GeneraScriptSync={GeneraScriptSync}, GeneraProductos={GeneraProductos}). " +
                        "Verifique la tabla SyncConfig en la base de datos o la sección SyncConfig en appsettings.json.",
                        codint, config.GeneraScriptSync, config.GeneraProductos);
                    return false;
                }

                var rutaIntercambio = config.RutaIntercambio;
                if (string.IsNullOrEmpty(rutaIntercambio) || !Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio no válida o inexistente: '{rutaIntercambio}'");
                    return false;
                }

                var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == codint);
                if (master == null)
                {
                    _logger.LogWarning($"No se encontró Master con Codint={codint}");
                    return false;
                }

                var syncDto = new SyncMasterDTO
                {
                    Tabla = "MASTER",
                    Operacion = "UPSERT",
                    Fecha = DateTime.Now,
                    Datos = master
                };

                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var nombreArchivo = $"MASTER_{master.Codint}_{timestamp}.json";
                var rutaCompleta = Path.Combine(rutaIntercambio, nombreArchivo);

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = null
                };

                var json = JsonSerializer.Serialize(syncDto, options);
                await File.WriteAllTextAsync(rutaCompleta, json);

                _logger.LogInformation($"Archivo de sincronización generado: {nombreArchivo}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al generar archivo de sincronización para Master.Codint={codint}");
                return false;
            }
        }

        /// <summary>
        /// Procesa todos los archivos de sincronización pendientes en la carpeta de intercambio (Sucursal)
        /// 
        /// Estrategia de procesamiento:
        /// 1. Busca archivos JSON de tres tipos: MASTER_*.json, FORMULAS_*.json, SECOPER_*.json
        /// 2. Ordena por fecha de creación (más antiguos primero) para mantener el orden de cambios
        /// 3. Para cada archivo:
        ///    - Verifica si ya fue procesado exitosamente (tabla SyncProcesados)
        ///    - Si no fue procesado: deserializa y aplica cambios a la BD
        ///    - Registra resultado en SyncProcesados (EXITOSO o ERROR)
        /// 4. Si un archivo falla, continúa con el siguiente (no falla todo el proceso)
        /// 5. Archivos fallidos se mueven a subcarpeta ERROR_MASTER/ para evitar reprocesamiento
        /// 
        /// Estrategias de UPSERT:
        /// - Master: Usa MERGE de SQL Server (actualiza si existe, inserta si no)
        /// - Formulas/SecOper: DELETE + INSERT (reemplaza todo el conjunto del CodIConj)
        /// 
        /// Por qué DELETE+INSERT en Formulas/SecOper:
        /// - Simplifica la lógica al reemplazar todos los componentes/operaciones de un conjunto
        /// - Evita conflictos de orden (NuorIt) y permite cambios en la estructura del conjunto
        /// - Usa transacción para garantizar atomicidad (todo o nada)
        /// 
        /// Auditoría:
        /// - Cada archivo procesado se registra en SyncProcesados
        /// - Permite evitar reprocesar archivos ya exitosos
        /// - Permite diagnosticar errores de sincronización
        /// - Archivos fallidos se mueven a carpeta ERROR_MASTER/ para revisión manual
        /// </summary>
        public async Task<SyncResultDTO> ProcesarArchivosSyncAsync(string rutaIntercambio)
        {
            var resultado = new SyncResultDTO
            {
                Exitoso = true,
                ArchivosProcessados = 0,
                RegistrosSincronizados = 0,
                Errores = 0,
                DetalleErrores = new List<string>()
            };

            try
            {
                if (string.IsNullOrWhiteSpace(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = "Ruta de intercambio no configurada";
                    return resultado;
                }

                if (!Directory.Exists(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = $"La ruta de intercambio no existe: {rutaIntercambio}";
                    return resultado;
                }

                var archivosMaster = Directory.GetFiles(rutaIntercambio, "MASTER_*.json");
                var archivosFormulas = Directory.GetFiles(rutaIntercambio, "FORMULAS_*.json");
                var archivosSecOper = Directory.GetFiles(rutaIntercambio, "SECOPER_*.json");
                
                var archivos = archivosMaster.Concat(archivosFormulas).Concat(archivosSecOper)
                    .OrderBy(f => File.GetCreationTime(f))
                    .ToList();

                if (archivos.Count == 0)
                {
                    resultado.Mensaje = "No hay archivos pendientes de sincronización";
                    return resultado;
                }

                foreach (var archivoPath in archivos)
                {
                    var nombreArchivo = Path.GetFileName(archivoPath);

                    var yaProcessado = await _context.SyncProcesados
                        .AnyAsync(s => s.Archivo == nombreArchivo && s.Estado == "EXITOSO");

                    if (yaProcessado)
                    {
                        _logger.LogInformation($"Archivo ya procesado: {nombreArchivo}");
                        continue;
                    }

                    try
                    {
                        var json = await File.ReadAllTextAsync(archivoPath);

                        if (nombreArchivo.StartsWith("MASTER_"))
                        {
                            var syncDto = JsonSerializer.Deserialize<SyncMasterDTO>(json);

                            if (syncDto == null || syncDto.Datos == null)
                            {
                                throw new Exception("Archivo JSON inválido o sin datos");
                            }

                            await ProcesarMasterUpsertAsync(syncDto.Datos);
                            await RegistrarProcesamientoAsync(nombreArchivo, "MASTER", "UPSERT", "EXITOSO", null, syncDto.Datos.Codint);

                            resultado.ArchivosProcessados++;
                            resultado.RegistrosSincronizados++;
                        }
                        else if (nombreArchivo.StartsWith("FORMULAS_"))
                        {
                            var syncDto = JsonSerializer.Deserialize<SyncFormulasDTO>(json);

                            if (syncDto == null || syncDto.Componentes == null || syncDto.Componentes.Count == 0)
                            {
                                throw new Exception("Archivo JSON inválido o sin componentes");
                            }

                            await ProcesarFormulasUpsertAsync(syncDto);
                            await RegistrarProcesamientoAsync(nombreArchivo, "FORMULAS", "UPSERT", "EXITOSO", null, syncDto.CodIConj);

                            resultado.ArchivosProcessados++;
                            resultado.RegistrosSincronizados += syncDto.Componentes.Count;
                        }
                        else if (nombreArchivo.StartsWith("SECOPER_"))
                        {
                            var syncDto = JsonSerializer.Deserialize<SyncSecOperDTO>(json);

                            if (syncDto == null || syncDto.Operaciones == null || syncDto.Operaciones.Count == 0)
                            {
                                throw new Exception("Archivo JSON inválido o sin operaciones");
                            }

                            await ProcesarSecOperUpsertAsync(syncDto);
                            await RegistrarProcesamientoAsync(nombreArchivo, "SECOPER", "UPSERT", "EXITOSO", null, syncDto.CodIConj);

                            resultado.ArchivosProcessados++;
                            resultado.RegistrosSincronizados += syncDto.Operaciones.Count;
                        }

                        _logger.LogInformation($"Archivo procesado exitosamente: {nombreArchivo}");
                    }
                    catch (Exception ex)
                    {
                        resultado.Errores++;
                        var mensajeError = $"Error en {nombreArchivo}: {ex.Message}";
                        resultado.DetalleErrores.Add(mensajeError);

                        var tabla = nombreArchivo.StartsWith("MASTER_") ? "MASTER" : 
                                    nombreArchivo.StartsWith("FORMULAS_") ? "FORMULAS" : "SECOPER";
                        await RegistrarProcesamientoAsync(nombreArchivo, tabla, "UPSERT", "ERROR", ex.Message, null);

                        // Mover archivo fallido a carpeta ERROR_MASTER para evitar reprocesamiento
                        try
                        {
                            var carpetaError = Path.Combine(rutaIntercambio, "ERROR_MASTER");
                            if (!Directory.Exists(carpetaError))
                            {
                                Directory.CreateDirectory(carpetaError);
                            }
                            var rutaDestino = Path.Combine(carpetaError, nombreArchivo);
                            File.Move(archivoPath, rutaDestino, overwrite: true);
                            _logger.LogInformation($"Archivo movido a carpeta ERROR_MASTER: {nombreArchivo}");
                        }
                        catch (Exception moveEx)
                        {
                            _logger.LogError(moveEx, $"No se pudo mover archivo a carpeta ERROR_MASTER: {nombreArchivo}");
                        }

                        _logger.LogError(ex, $"Error al procesar archivo: {nombreArchivo}");
                    }
                }

                resultado.Mensaje = $"Procesamiento completado. Archivos: {resultado.ArchivosProcessados}, Registros: {resultado.RegistrosSincronizados}, Errores: {resultado.Errores}";
                resultado.Exitoso = resultado.Errores == 0;

                return resultado;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error general al procesar archivos de sincronización");
                resultado.Exitoso = false;
                resultado.Mensaje = $"Error general: {ex.Message}";
                return resultado;
            }
        }

        /// <summary>
        /// Procesa un UPSERT de Master usando MERGE de SQL Server
        /// </summary>
        private async Task ProcesarMasterUpsertAsync(Master master)
        {
            var sql = @"
                MERGE Master AS target
                USING (SELECT @Codint AS Codint) AS source
                ON target.Codint = source.Codint
                WHEN MATCHED THEN
                    UPDATE SET
                        Ferevis = @Ferevis,
                        CodiEmpr = @CodiEmpr,
                        Observa = @Observa,
                        Especificacion = @Especificacion,
                        DesCompAut = @DesCompAut,
                        DescriComp = @DescriComp,
                        DepoEntra = @DepoEntra,
                        DepoCon = @DepoCon,
                        Codigo = @Codigo,
                        Descripcion = @Descripcion,
                        Descridos = @Descridos,
                        Status = @Status,
                        Alias = @Alias,
                        Umedida = @Umedida,
                        Timate = @Timate,
                        Critrep = @Critrep,
                        Monecos = @Monecos,
                        Fecosto = @Fecosto,
                        PorDescCos = @PorDescCos,
                        Cosuni = @Cosuni,
                        Pornacio = @Pornacio,
                        PorComVen = @PorComVen,
                        Provhabi = @Provhabi,
                        Stomin = @Stomin,
                        Stomax = @Stomax,
                        Lorepos = @Lorepos,
                        Lindivis = @Lindivis,
                        Pesuni = @Pesuni,
                        PesMetro = @PesMetro,
                        Color = @Color,
                        Talle = @Talle,
                        Descompo = @Descompo,
                        Grupiva = @Grupiva,
                        Grucoven = @Grucoven,
                        CLinea = @CLinea,
                        Cfamilia = @Cfamilia,
                        NPlano = @NPlano,
                        Revision = @Revision,
                        Trazable = @Trazable,
                        Revi_Formu = @Revi_Formu,
                        FechApFor = @FechApFor,
                        UsuApFor = @UsuApFor,
                        Cod_Interna = @Cod_Interna,
                        Bien_Cap = @Bien_Cap,
                        Codigo_Ncm = @Codigo_Ncm,
                        Codigo_Sec = @Codigo_Sec,
                        Pro_Umed = @Pro_Umed,
                        Iva_Id = @Iva_Id,
                        ArchiPdf = @ArchiPdf,
                        Familia = @Familia,
                        Lindiven = @Lindiven,
                        CanDeci = @CanDeci,
                        ConsuProm = @ConsuProm,
                        CosCal = @CosCal,
                        FeCosCal = @FeCosCal,
                        Stat_Movi = @Stat_Movi,
                        Clas_ATC = @Clas_ATC,
                        Vers_BC = @Vers_BC,
                        FeVig_BC = @FeVig_BC,
                        FeRev_BC = @FeRev_BC,
                        CantApro = @CantApro,
                        LeTiCom = @LeTiCom,
                        PiezasXCaja = @PiezasXCaja,
                        M3Caja = @M3Caja,
                        KiloNetoCaja = @KiloNetoCaja,
                        KiloBrutoCaja = @KiloBrutoCaja,
                        Autonomia = @Autonomia,
                        VALOBADA_1 = @VALOBADA_1,
                        VALOBADA_2 = @VALOBADA_2,
                        VALOBADA_3 = @VALOBADA_3,
                        Autono_Proy = @Autono_Proy,
                        UsuPrePes = @UsuPrePes,
                        UsuPreDol = @UsuPreDol,
                        MANO_OBRA_REAL = @MANO_OBRA_REAL,
                        Tamano_Etiqueta = @Tamano_Etiqueta,
                        CueContaI = @CueContaI,
                        Molde = @Molde,
                        Consumo_X_Corte = @Consumo_X_Corte,
                        GRAMAJE = @GRAMAJE,
                        ANCHO = @ANCHO,
                        GR_ACTU = @GR_ACTU,
                        Cod_Arba = @Cod_Arba,
                        MICRONES = @MICRONES,
                        VALOBADA_4 = @VALOBADA_4,
                        VALOBADA_5 = @VALOBADA_5,
                        VALOBADA_6 = @VALOBADA_6,
                        LastUpdate = GETDATE()
                WHEN NOT MATCHED THEN
                    INSERT (Codint, Ferevis, CodiEmpr, Observa, Especificacion, DesCompAut, DescriComp, 
                            DepoEntra, DepoCon, Codigo, Descripcion, Descridos, Status, Alias, Umedida, 
                            Timate, Critrep, Monecos, Fecosto, PorDescCos, Cosuni, Pornacio, PorComVen, 
                            Provhabi, Stomin, Stomax, Lorepos, Lindivis, Pesuni, PesMetro, Color, Talle, 
                            Descompo, Grupiva, Grucoven, CLinea, Cfamilia, NPlano, Revision, Trazable, 
                            Revi_Formu, FechApFor, UsuApFor, Cod_Interna, Bien_Cap, Codigo_Ncm, Codigo_Sec, 
                            Pro_Umed, Iva_Id, ArchiPdf, Familia, Lindiven, CanDeci, ConsuProm, CosCal, 
                            FeCosCal, Stat_Movi, Clas_ATC, Vers_BC, FeVig_BC, FeRev_BC, CantApro, LeTiCom, 
                            PiezasXCaja, M3Caja, KiloNetoCaja, KiloBrutoCaja, Autonomia, VALOBADA_1, 
                            VALOBADA_2, VALOBADA_3, Autono_Proy, UsuPrePes, UsuPreDol, MANO_OBRA_REAL, 
                            Tamano_Etiqueta, CueContaI, Molde, Consumo_X_Corte, GRAMAJE, ANCHO, GR_ACTU, 
                            Cod_Arba, MICRONES, VALOBADA_4, VALOBADA_5, VALOBADA_6, AddRecord, CodiEmprNet, NumUsuar)
                    VALUES (@Codint, @Ferevis, @CodiEmpr, @Observa, @Especificacion, @DesCompAut, @DescriComp, 
                            @DepoEntra, @DepoCon, @Codigo, @Descripcion, @Descridos, @Status, @Alias, @Umedida, 
                            @Timate, @Critrep, @Monecos, @Fecosto, @PorDescCos, @Cosuni, @Pornacio, @PorComVen, 
                            @Provhabi, @Stomin, @Stomax, @Lorepos, @Lindivis, @Pesuni, @PesMetro, @Color, @Talle, 
                            @Descompo, @Grupiva, @Grucoven, @CLinea, @Cfamilia, @NPlano, @Revision, @Trazable, 
                            @Revi_Formu, @FechApFor, @UsuApFor, @Cod_Interna, @Bien_Cap, @Codigo_Ncm, @Codigo_Sec, 
                            @Pro_Umed, @Iva_Id, @ArchiPdf, @Familia, @Lindiven, @CanDeci, @ConsuProm, @CosCal, 
                            @FeCosCal, @Stat_Movi, @Clas_ATC, @Vers_BC, @FeVig_BC, @FeRev_BC, @CantApro, @LeTiCom, 
                            @PiezasXCaja, @M3Caja, @KiloNetoCaja, @KiloBrutoCaja, @Autonomia, @VALOBADA_1, 
                            @VALOBADA_2, @VALOBADA_3, @Autono_Proy, @UsuPrePes, @UsuPreDol, @MANO_OBRA_REAL, 
                            @Tamano_Etiqueta, @CueContaI, @Molde, @Consumo_X_Corte, @GRAMAJE, @ANCHO, @GR_ACTU, 
                            @Cod_Arba, @MICRONES, @VALOBADA_4, @VALOBADA_5, @VALOBADA_6, GETDATE(), @CodiEmprNet, @NumUsuar);";

            var parameters = new[]
            {
                new SqlParameter("@Codint", master.Codint),
                new SqlParameter("@Ferevis", (object?)master.Ferevis ?? DBNull.Value),
                new SqlParameter("@CodiEmpr", (object?)master.CodiEmpr ?? DBNull.Value),
                new SqlParameter("@Observa", (object?)master.Observa ?? DBNull.Value),
                new SqlParameter("@Especificacion", (object?)master.Especificacion ?? DBNull.Value),
                new SqlParameter("@DesCompAut", (object?)master.DesCompAut ?? DBNull.Value),
                new SqlParameter("@DescriComp", (object?)master.DescriComp ?? DBNull.Value),
                new SqlParameter("@DepoEntra", (object?)master.DepoEntra ?? DBNull.Value),
                new SqlParameter("@DepoCon", (object?)master.DepoCon ?? DBNull.Value),
                new SqlParameter("@Codigo", (object?)master.Codigo ?? DBNull.Value),
                new SqlParameter("@Descripcion", (object?)master.Descripcion ?? DBNull.Value),
                new SqlParameter("@Descridos", (object?)master.Descridos ?? DBNull.Value),
                new SqlParameter("@Status", (object?)master.Status ?? DBNull.Value),
                new SqlParameter("@Alias", (object?)master.Alias ?? DBNull.Value),
                new SqlParameter("@Umedida", (object?)master.Umedida ?? DBNull.Value),
                new SqlParameter("@Timate", (object?)master.Timate ?? DBNull.Value),
                new SqlParameter("@Critrep", (object?)master.Critrep ?? DBNull.Value),
                new SqlParameter("@Monecos", (object?)master.Monecos ?? DBNull.Value),
                new SqlParameter("@Fecosto", (object?)master.Fecosto ?? DBNull.Value),
                new SqlParameter("@PorDescCos", (object?)master.PorDescCos ?? DBNull.Value),
                new SqlParameter("@Cosuni", (object?)master.Cosuni ?? DBNull.Value),
                new SqlParameter("@Pornacio", (object?)master.Pornacio ?? DBNull.Value),
                new SqlParameter("@PorComVen", (object?)master.PorComVen ?? DBNull.Value),
                new SqlParameter("@Provhabi", (object?)master.Provhabi ?? DBNull.Value),
                new SqlParameter("@Stomin", (object?)master.Stomin ?? DBNull.Value),
                new SqlParameter("@Stomax", (object?)master.Stomax ?? DBNull.Value),
                new SqlParameter("@Lorepos", (object?)master.Lorepos ?? DBNull.Value),
                new SqlParameter("@Lindivis", (object?)master.Lindivis ?? DBNull.Value),
                new SqlParameter("@Pesuni", (object?)master.Pesuni ?? DBNull.Value),
                new SqlParameter("@PesMetro", (object?)master.PesMetro ?? DBNull.Value),
                new SqlParameter("@Color", (object?)master.Color ?? DBNull.Value),
                new SqlParameter("@Talle", (object?)master.Talle ?? DBNull.Value),
                new SqlParameter("@Descompo", (object?)master.Descompo ?? DBNull.Value),
                new SqlParameter("@Grupiva", (object?)master.Grupiva ?? DBNull.Value),
                new SqlParameter("@Grucoven", (object?)master.Grucoven ?? DBNull.Value),
                new SqlParameter("@CLinea", (object?)master.CLinea ?? DBNull.Value),
                new SqlParameter("@Cfamilia", (object?)master.Cfamilia ?? DBNull.Value),
                new SqlParameter("@NPlano", (object?)master.NPlano ?? DBNull.Value),
                new SqlParameter("@Revision", (object?)master.Revision ?? DBNull.Value),
                new SqlParameter("@Trazable", (object?)master.Trazable ?? DBNull.Value),
                new SqlParameter("@Revi_Formu", (object?)master.Revi_Formu ?? DBNull.Value),
                new SqlParameter("@FechApFor", (object?)master.FechApFor ?? DBNull.Value),
                new SqlParameter("@UsuApFor", (object?)master.UsuApFor ?? DBNull.Value),
                new SqlParameter("@Cod_Interna", (object?)master.Cod_Interna ?? DBNull.Value),
                new SqlParameter("@Bien_Cap", (object?)master.Bien_Cap ?? DBNull.Value),
                new SqlParameter("@Codigo_Ncm", (object?)master.Codigo_Ncm ?? DBNull.Value),
                new SqlParameter("@Codigo_Sec", (object?)master.Codigo_Sec ?? DBNull.Value),
                new SqlParameter("@Pro_Umed", (object?)master.Pro_Umed ?? DBNull.Value),
                new SqlParameter("@Iva_Id", (object?)master.Iva_Id ?? DBNull.Value),
                new SqlParameter("@ArchiPdf", (object?)master.ArchiPdf ?? DBNull.Value),
                new SqlParameter("@Familia", (object?)master.Familia ?? DBNull.Value),
                new SqlParameter("@Lindiven", (object?)master.Lindiven ?? DBNull.Value),
                new SqlParameter("@CanDeci", (object?)master.CanDeci ?? DBNull.Value),
                new SqlParameter("@ConsuProm", (object?)master.ConsuProm ?? DBNull.Value),
                new SqlParameter("@CosCal", (object?)master.CosCal ?? DBNull.Value),
                new SqlParameter("@FeCosCal", (object?)master.FeCosCal ?? DBNull.Value),
                new SqlParameter("@Stat_Movi", (object?)master.Stat_Movi ?? DBNull.Value),
                new SqlParameter("@Clas_ATC", (object?)master.Clas_ATC ?? DBNull.Value),
                new SqlParameter("@Vers_BC", (object?)master.Vers_BC ?? DBNull.Value),
                new SqlParameter("@FeVig_BC", (object?)master.FeVig_BC ?? DBNull.Value),
                new SqlParameter("@FeRev_BC", (object?)master.FeRev_BC ?? DBNull.Value),
                new SqlParameter("@CantApro", (object?)master.CantApro ?? DBNull.Value),
                new SqlParameter("@LeTiCom", (object?)master.LeTiCom ?? DBNull.Value),
                new SqlParameter("@PiezasXCaja", (object?)master.PiezasXCaja ?? DBNull.Value),
                new SqlParameter("@M3Caja", (object?)master.M3Caja ?? DBNull.Value),
                new SqlParameter("@KiloNetoCaja", (object?)master.KiloNetoCaja ?? DBNull.Value),
                new SqlParameter("@KiloBrutoCaja", (object?)master.KiloBrutoCaja ?? DBNull.Value),
                new SqlParameter("@Autonomia", (object?)master.Autonomia ?? DBNull.Value),
                new SqlParameter("@VALOBADA_1", (object?)master.VALOBADA_1 ?? DBNull.Value),
                new SqlParameter("@VALOBADA_2", (object?)master.VALOBADA_2 ?? DBNull.Value),
                new SqlParameter("@VALOBADA_3", (object?)master.VALOBADA_3 ?? DBNull.Value),
                new SqlParameter("@Autono_Proy", (object?)master.Autono_Proy ?? DBNull.Value),
                new SqlParameter("@UsuPrePes", (object?)master.UsuPrePes ?? DBNull.Value),
                new SqlParameter("@UsuPreDol", (object?)master.UsuPreDol ?? DBNull.Value),
                new SqlParameter("@MANO_OBRA_REAL", (object?)master.MANO_OBRA_REAL ?? DBNull.Value),
                new SqlParameter("@Tamano_Etiqueta", (object?)master.Tamano_Etiqueta ?? DBNull.Value),
                new SqlParameter("@CueContaI", (object?)master.CueContaI ?? DBNull.Value),
                new SqlParameter("@Molde", (object?)master.Molde ?? DBNull.Value),
                new SqlParameter("@Consumo_X_Corte", (object?)master.Consumo_X_Corte ?? DBNull.Value),
                new SqlParameter("@GRAMAJE", (object?)master.GRAMAJE ?? DBNull.Value),
                new SqlParameter("@ANCHO", (object?)master.ANCHO ?? DBNull.Value),
                new SqlParameter("@GR_ACTU", (object?)master.GR_ACTU ?? DBNull.Value),
                new SqlParameter("@Cod_Arba", (object?)master.Cod_Arba ?? DBNull.Value),
                new SqlParameter("@MICRONES", (object?)master.MICRONES ?? DBNull.Value),
                new SqlParameter("@VALOBADA_4", (object?)master.VALOBADA_4 ?? DBNull.Value),
                new SqlParameter("@VALOBADA_5", (object?)master.VALOBADA_5 ?? DBNull.Value),
                new SqlParameter("@VALOBADA_6", (object?)master.VALOBADA_6 ?? DBNull.Value),
                new SqlParameter("@CodiEmprNet", (object?)master.CodiEmprNet ?? DBNull.Value),
                new SqlParameter("@NumUsuar", (object?)master.NumUsuar ?? DBNull.Value)
            };

            await _context.Database.ExecuteSqlRawAsync(sql, parameters);
        }

        /// <summary>
        /// Procesa un UPSERT de Formulas usando MERGE de SQL Server
        /// Primero elimina todos los componentes del conjunto y luego inserta los nuevos
        /// </summary>
        private async Task ProcesarFormulasUpsertAsync(SyncFormulasDTO syncDto)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            
            try
            {
                // Eliminar componentes existentes del conjunto
                var deleteSql = "DELETE FROM Formulas WHERE CodIConj = @CodIConj";
                await _context.Database.ExecuteSqlRawAsync(deleteSql, 
                    new SqlParameter("@CodIConj", syncDto.CodIConj));

                // Insertar cada componente
                foreach (var componente in syncDto.Componentes)
                {
                    var insertSql = @"
                        INSERT INTO Formulas (
                            CodIConj, CodiEmpr, CodXConj, CodiElem, CodXElem, NuorIt,
                            UsoBruto, Unis_Form, Lar_Corte, Pladispo, RefeComp, Marborra,
                            UMED_FORM, COMPONENTE, MODUXBB, CORTE, VALOBADA_1
                        ) VALUES (
                            @CodIConj, @CodiEmpr, @CodXConj, @CodiElem, @CodXElem, @NuorIt,
                            @UsoBruto, @Unis_Form, @Lar_Corte, @Pladispo, @RefeComp, @Marborra,
                            @UMED_FORM, @COMPONENTE, @MODUXBB, @CORTE, @VALOBADA_1
                        )";

                    var parameters = new[]
                    {
                        new SqlParameter("@CodIConj", (object?)componente.CodIConj ?? DBNull.Value),
                        new SqlParameter("@CodiEmpr", (object?)componente.CodiEmpr ?? DBNull.Value),
                        new SqlParameter("@CodXConj", (object?)componente.CodXConj ?? DBNull.Value),
                        new SqlParameter("@CodiElem", (object?)componente.CodiElem ?? DBNull.Value),
                        new SqlParameter("@CodXElem", (object?)componente.CodXElem ?? DBNull.Value),
                        new SqlParameter("@NuorIt", (object?)componente.NuorIt ?? DBNull.Value),
                        new SqlParameter("@UsoBruto", (object?)componente.UsoBruto ?? DBNull.Value),
                        new SqlParameter("@Unis_Form", (object?)componente.Unis_Form ?? DBNull.Value),
                        new SqlParameter("@Lar_Corte", (object?)componente.Lar_Corte ?? DBNull.Value),
                        new SqlParameter("@Pladispo", (object?)componente.Pladispo ?? DBNull.Value),
                        new SqlParameter("@RefeComp", (object?)componente.RefeComp ?? DBNull.Value),
                        new SqlParameter("@Marborra", (object?)componente.Marborra ?? DBNull.Value),
                        new SqlParameter("@UMED_FORM", (object?)componente.UMED_FORM ?? DBNull.Value),
                        new SqlParameter("@COMPONENTE", (object?)componente.COMPONENTE ?? DBNull.Value),
                        new SqlParameter("@MODUXBB", (object?)componente.MODUXBB ?? DBNull.Value),
                        new SqlParameter("@CORTE", (object?)componente.CORTE ?? DBNull.Value),
                        new SqlParameter("@VALOBADA_1", (object?)componente.VALOBADA_1 ?? DBNull.Value)
                    };

                    await _context.Database.ExecuteSqlRawAsync(insertSql, parameters);
                }

                await transaction.CommitAsync();
                _logger.LogInformation($"Sincronización de Formulas completada para CodIConj={syncDto.CodIConj}, {syncDto.Componentes.Count} componentes");
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, $"Error al procesar Formulas para CodIConj={syncDto.CodIConj}");
                throw;
            }
        }

        /// <summary>
        /// Procesa un UPSERT de SecOper usando MERGE de SQL Server
        /// Primero elimina todas las operaciones del conjunto y luego inserta las nuevas
        /// </summary>
        private async Task ProcesarSecOperUpsertAsync(SyncSecOperDTO syncDto)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            
            try
            {
                // Eliminar operaciones existentes del conjunto
                var deleteSql = "DELETE FROM SECOPER WHERE CodIConj = @CodIConj";
                await _context.Database.ExecuteSqlRawAsync(deleteSql, 
                    new SqlParameter("@CodIConj", syncDto.CodIConj));

                // Insertar cada operación
                foreach (var operacion in syncDto.Operaciones)
                {
                    var insertSql = @"
                        INSERT INTO SECOPER (
                            CodIConj, CodiEmpr, CodXConj, NumeOper, AltOper, DescOper, StatOper,
                            CodMaq0, CodMaq1, CodMaq2, CodMaq3, CodMaq4, CodMaq5,
                            HorsPrep, HorsFijo, MinuStan, HorsMHer, HorsLimp, PieCiclo, PorDemor,
                            Cadencia, CoefSolp, LoteStan, CantiOps, Cantiops2, CategOps, Categops2,
                            ProgCNum, HoProFab, ObseOper, ProTerOp, ProvOper, PreUnid, MonePrec,
                            FechPrec, Nuoper_Prep, Catoper_Prep, Fech_Revis, Usua_Revis, rowguid,
                            SecuenciaFabricacion, OperacionTerceros, NuOrdenManual
                        ) VALUES (
                            @CodIConj, @CodiEmpr, @CodXConj, @NumeOper, @AltOper, @DescOper, @StatOper,
                            @CodMaq0, @CodMaq1, @CodMaq2, @CodMaq3, @CodMaq4, @CodMaq5,
                            @HorsPrep, @HorsFijo, @MinuStan, @HorsMHer, @HorsLimp, @PieCiclo, @PorDemor,
                            @Cadencia, @CoefSolp, @LoteStan, @CantiOps, @Cantiops2, @CategOps, @Categops2,
                            @ProgCNum, @HoProFab, @ObseOper, @ProTerOp, @ProvOper, @PreUnid, @MonePrec,
                            @FechPrec, @Nuoper_Prep, @Catoper_Prep, @Fech_Revis, @Usua_Revis, @rowguid,
                            @SecuenciaFabricacion, @OperacionTerceros, @NuOrdenManual
                        )";

                    var parameters = new[]
                    {
                        new SqlParameter("@CodIConj", operacion.CodIConj),
                        new SqlParameter("@CodiEmpr", (object?)operacion.CodiEmpr ?? DBNull.Value),
                        new SqlParameter("@CodXConj", (object?)operacion.CodXConj ?? DBNull.Value),
                        new SqlParameter("@NumeOper", (object?)operacion.NumeOper ?? DBNull.Value),
                        new SqlParameter("@AltOper", (object?)operacion.AltOper ?? DBNull.Value),
                        new SqlParameter("@DescOper", (object?)operacion.DescOper ?? DBNull.Value),
                        new SqlParameter("@StatOper", (object?)operacion.StatOper ?? DBNull.Value),
                        new SqlParameter("@CodMaq0", (object?)operacion.CodMaq0 ?? DBNull.Value),
                        new SqlParameter("@CodMaq1", (object?)operacion.CodMaq1 ?? DBNull.Value),
                        new SqlParameter("@CodMaq2", (object?)operacion.CodMaq2 ?? DBNull.Value),
                        new SqlParameter("@CodMaq3", (object?)operacion.CodMaq3 ?? DBNull.Value),
                        new SqlParameter("@CodMaq4", (object?)operacion.CodMaq4 ?? DBNull.Value),
                        new SqlParameter("@CodMaq5", (object?)operacion.CodMaq5 ?? DBNull.Value),
                        new SqlParameter("@HorsPrep", (object?)operacion.HorsPrep ?? DBNull.Value),
                        new SqlParameter("@HorsFijo", (object?)operacion.HorsFijo ?? DBNull.Value),
                        new SqlParameter("@MinuStan", (object?)operacion.MinuStan ?? DBNull.Value),
                        new SqlParameter("@HorsMHer", (object?)operacion.HorsMHer ?? DBNull.Value),
                        new SqlParameter("@HorsLimp", (object?)operacion.HorsLimp ?? DBNull.Value),
                        new SqlParameter("@PieCiclo", (object?)operacion.PieCiclo ?? DBNull.Value),
                        new SqlParameter("@PorDemor", (object?)operacion.PorDemor ?? DBNull.Value),
                        new SqlParameter("@Cadencia", (object?)operacion.Cadencia ?? DBNull.Value),
                        new SqlParameter("@CoefSolp", (object?)operacion.CoefSolp ?? DBNull.Value),
                        new SqlParameter("@LoteStan", (object?)operacion.LoteStan ?? DBNull.Value),
                        new SqlParameter("@CantiOps", (object?)operacion.CantiOps ?? DBNull.Value),
                        new SqlParameter("@Cantiops2", (object?)operacion.Cantiops2 ?? DBNull.Value),
                        new SqlParameter("@CategOps", (object?)operacion.CategOps ?? DBNull.Value),
                        new SqlParameter("@Categops2", (object?)operacion.Categops2 ?? DBNull.Value),
                        new SqlParameter("@ProgCNum", (object?)operacion.ProgCNum ?? DBNull.Value),
                        new SqlParameter("@HoProFab", (object?)operacion.HoProFab ?? DBNull.Value),
                        new SqlParameter("@ObseOper", (object?)operacion.ObseOper ?? DBNull.Value),
                        new SqlParameter("@ProTerOp", (object?)operacion.ProTerOp ?? DBNull.Value),
                        new SqlParameter("@ProvOper", (object?)operacion.ProvOper ?? DBNull.Value),
                        new SqlParameter("@PreUnid", (object?)operacion.PreUnid ?? DBNull.Value),
                        new SqlParameter("@MonePrec", (object?)operacion.MonePrec ?? DBNull.Value),
                        new SqlParameter("@FechPrec", (object?)operacion.FechPrec ?? DBNull.Value),
                        new SqlParameter("@Nuoper_Prep", (object?)operacion.Nuoper_Prep ?? DBNull.Value),
                        new SqlParameter("@Catoper_Prep", (object?)operacion.Catoper_Prep ?? DBNull.Value),
                        new SqlParameter("@Fech_Revis", (object?)operacion.Fech_Revis ?? DBNull.Value),
                        new SqlParameter("@Usua_Revis", (object?)operacion.Usua_Revis ?? DBNull.Value),
                        new SqlParameter("@rowguid", operacion.rowguid),
                        new SqlParameter("@SecuenciaFabricacion", (object?)operacion.SecuenciaFabricacion ?? DBNull.Value),
                        new SqlParameter("@OperacionTerceros", (object?)operacion.OperacionTerceros ?? DBNull.Value),
                        new SqlParameter("@NuOrdenManual", (object?)operacion.NuOrdenManual ?? DBNull.Value)
                    };

                    await _context.Database.ExecuteSqlRawAsync(insertSql, parameters);
                }

                await transaction.CommitAsync();
                _logger.LogInformation($"Sincronización de SecOper completada para CodIConj={syncDto.CodIConj}, {syncDto.Operaciones.Count} operaciones");
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                _logger.LogError(ex, $"Error al procesar SecOper para CodIConj={syncDto.CodIConj}");
                throw;
            }
        }

        /// <summary>
        /// Registra el procesamiento de un archivo en la tabla de auditoría
        /// </summary>
        public async Task RegistrarProcesamientoAsync(string archivo, string tabla, string operacion, string estado, string? error = null, int? codintAfectado = null)
        {
            var registro = new SyncProcesado
            {
                Archivo = archivo,
                Tabla = tabla,
                Operacion = operacion,
                FechaProcesamiento = DateTime.Now,
                Estado = estado,
                Error = error,
                CodintAfectado = codintAfectado,
                Codiempr = 0
            };

            _context.SyncProcesados.Add(registro);
            await _context.SaveChangesAsync();
        }

        /// <summary>
        /// Obtiene la configuración de sincronización desde la base de datos
        /// Si no existe, la crea con valores por defecto desde appsettings
        /// </summary>
        public SyncConfigDTO ObtenerConfiguracion()
        {
            var configDb = _context.SyncConfigs.FirstOrDefault();
            
            if (configDb == null)
            {
                return new SyncConfigDTO
                {
                    GeneraScriptSync = _configuration.GetValue<bool>("SyncConfig:GeneraScriptSync"),
                    RutaIntercambio = _configuration.GetValue<string>("SyncConfig:RutaIntercambio") ?? "",
                    RutaIntercambioRemitos = _configuration.GetValue<string>("SyncConfig:RutaIntercambioRemitos") ?? "",
                    RutaIntercambioReparaciones = _configuration.GetValue<string>("SyncConfig:RutaIntercambioReparaciones") ?? "",
                    RutaIntercambioAjustes = _configuration.GetValue<string>("SyncConfig:RutaIntercambioAjustes") ?? "",
                    RutaIntercambioPedidosReposicion = _configuration.GetValue<string>("SyncConfig:RutaIntercambioPedidosReposicion") ?? "",
                    GeneraProductos = _configuration.GetValue<bool>("SyncConfig:GeneraProductos"),
                    ProcesaProductos = _configuration.GetValue<bool>("SyncConfig:ProcesaProductos"),
                    GeneraRemitos = _configuration.GetValue<bool>("SyncConfig:GeneraRemitos"),
                    ProcesaRemitos = _configuration.GetValue<bool>("SyncConfig:ProcesaRemitos"),
                    GeneraReparaciones = _configuration.GetValue<bool>("SyncConfig:GeneraReparaciones"),
                    ProcesaReparaciones = _configuration.GetValue<bool>("SyncConfig:ProcesaReparaciones"),
                    GeneraAjustes = _configuration.GetValue<bool>("SyncConfig:GeneraAjustes"),
                    ProcesaAjustes = _configuration.GetValue<bool>("SyncConfig:ProcesaAjustes"),
                    GeneraPedidosReposicion = _configuration.GetValue<bool>("SyncConfig:GeneraPedidosReposicion"),
                    ProcesaPedidosReposicion = _configuration.GetValue<bool>("SyncConfig:ProcesaPedidosReposicion")
                };
            }

            return new SyncConfigDTO
            {
                GeneraScriptSync = configDb.GeneraScriptSync,
                RutaIntercambio = configDb.RutaIntercambio ?? "",
                RutaIntercambioRemitos = configDb.RutaIntercambioRemitos ?? "",
                RutaIntercambioReparaciones = configDb.RutaIntercambioReparaciones ?? "",
                RutaIntercambioAjustes = configDb.RutaIntercambioAjustes ?? "",
                RutaIntercambioPedidosReposicion = configDb.RutaIntercambioPedidosReposicion ?? "",
                GeneraProductos = configDb.GeneraProductos,
                ProcesaProductos = configDb.ProcesaProductos,
                GeneraRemitos = configDb.GeneraRemitos,
                ProcesaRemitos = configDb.ProcesaRemitos,
                GeneraReparaciones = configDb.GeneraReparaciones,
                ProcesaReparaciones = configDb.ProcesaReparaciones,
                GeneraAjustes = configDb.GeneraAjustes,
                ProcesaAjustes = configDb.ProcesaAjustes,
                GeneraPedidosReposicion = configDb.GeneraPedidosReposicion,
                ProcesaPedidosReposicion = configDb.ProcesaPedidosReposicion
            };
        }

        /// <summary>
        /// Guarda la configuración de sincronización en la base de datos
        /// </summary>
        public async Task<bool> GuardarConfiguracionAsync(SyncConfigDTO config)
        {
            try
            {
                var configDb = await _context.SyncConfigs.FirstOrDefaultAsync();
                
                if (configDb == null)
                {
                    configDb = new SyncConfig
                    {
                        GeneraScriptSync = config.GeneraScriptSync,
                        RutaIntercambio = config.RutaIntercambio ?? "",
                        RutaIntercambioRemitos = config.RutaIntercambioRemitos ?? "",
                        RutaIntercambioReparaciones = config.RutaIntercambioReparaciones ?? "",
                        RutaIntercambioAjustes = config.RutaIntercambioAjustes ?? "",
                        RutaIntercambioPedidosReposicion = config.RutaIntercambioPedidosReposicion ?? "",
                        GeneraProductos = config.GeneraProductos,
                        ProcesaProductos = config.ProcesaProductos,
                        GeneraRemitos = config.GeneraRemitos,
                        ProcesaRemitos = config.ProcesaRemitos,
                        GeneraReparaciones = config.GeneraReparaciones,
                        ProcesaReparaciones = config.ProcesaReparaciones,
                        GeneraAjustes = config.GeneraAjustes,
                        ProcesaAjustes = config.ProcesaAjustes,
                        GeneraPedidosReposicion = config.GeneraPedidosReposicion,
                        ProcesaPedidosReposicion = config.ProcesaPedidosReposicion,
                        LastUpdate = DateTime.Now
                    };
                    _context.SyncConfigs.Add(configDb);
                }
                else
                {
                    configDb.GeneraScriptSync = config.GeneraScriptSync;
                    configDb.RutaIntercambio = config.RutaIntercambio ?? "";
                    configDb.RutaIntercambioRemitos = config.RutaIntercambioRemitos ?? "";
                    configDb.RutaIntercambioReparaciones = config.RutaIntercambioReparaciones ?? "";
                    configDb.RutaIntercambioAjustes = config.RutaIntercambioAjustes ?? "";
                    configDb.RutaIntercambioPedidosReposicion = config.RutaIntercambioPedidosReposicion ?? "";
                    configDb.GeneraProductos = config.GeneraProductos;
                    configDb.ProcesaProductos = config.ProcesaProductos;
                    configDb.GeneraRemitos = config.GeneraRemitos;
                    configDb.ProcesaRemitos = config.ProcesaRemitos;
                    configDb.GeneraReparaciones = config.GeneraReparaciones;
                    configDb.ProcesaReparaciones = config.ProcesaReparaciones;
                    configDb.GeneraAjustes = config.GeneraAjustes;
                    configDb.ProcesaAjustes = config.ProcesaAjustes;
                    configDb.GeneraPedidosReposicion = config.GeneraPedidosReposicion;
                    configDb.ProcesaPedidosReposicion = config.ProcesaPedidosReposicion;
                    configDb.LastUpdate = DateTime.Now;
                    _context.SyncConfigs.Update(configDb);
                }
                
                await _context.SaveChangesAsync();
                
                _logger.LogInformation("Configuración de sincronización guardada exitosamente en base de datos");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al guardar configuración de sincronización en base de datos");
                return false;
            }
        }

        /// <summary>
        /// Transforma el número de remito del formato del TXT al formato de MOVISTO
        /// Formato TXT: RT-X-0001-00018694
        /// Formato MOVISTO: RT.18694 (para DoPriCor, DoPriLar, IdenLote)
        /// Formato numérico: 18694 (para DOPRINUM)
        /// Formato NUMPVTA: 1 (posición 6-10 del TXT)
        /// </summary>
        private (string corto, int numerico, int numPvta) TransformarNumeroRemito(string numeroRemitoTxt)
        {
            try
            {
                // Extraer las últimas 8 posiciones para el número numérico
                if (numeroRemitoTxt.Length >= 8)
                {
                    var ultimas8 = numeroRemitoTxt.Substring(numeroRemitoTxt.Length - 8);
                    if (int.TryParse(ultimas8, out int numerico))
                    {
                        // Formato corto: RT.18694
                        var corto = $"RT.{numerico}";
                        
                        // Extraer posición 6-10 para NUMPVTA (índice 5-9 en 0-based)
                        var numPvta = 0;
                        if (numeroRemitoTxt.Length >= 10)
                        {
                            var posicion6a10 = numeroRemitoTxt.Substring(5, 5).Trim();
                            int.TryParse(posicion6a10, out numPvta);
                        }
                        
                        return (corto, numerico, numPvta);
                    }
                }
                
                // Si no se puede parsear, retornar valores por defecto
                return ("", 0, 0);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al transformar número de remito: {numeroRemitoTxt}");
                return ("", 0, 0);
            }
        }

        /// <summary>
        /// Transforma el número de orden de reparación del formato del TXT al formato de MOVISTO
        /// Formato TXT: OR.1234 u OR-1234
        /// Formato MOVISTO: OR.1234 u OR-1234 (identificador completo para DoPriCor, DoPriLar, DoSecCor, DoSecLar, IdenLote)
        /// Formato numérico: 1234 (para DOPRINUM)
        /// Formato NUMPVTA: 0 (no aplica para reparaciones)
        /// </summary>
        private (string identificador, int numerico, int numPvta) TransformarNumeroReparacion(string numeroReparacionTxt)
        {
            try
            {
                numeroReparacionTxt = numeroReparacionTxt.Trim();

                // Extraer el prefijo (OR. u OR-) y el valor numérico
                if (numeroReparacionTxt.StartsWith("OR.", StringComparison.OrdinalIgnoreCase) ||
                    numeroReparacionTxt.StartsWith("OR-", StringComparison.OrdinalIgnoreCase))
                {
                    var parteNumerica = numeroReparacionTxt.Substring(3).Trim();
                    if (int.TryParse(parteNumerica, out int numerico))
                    {
                        // Normalizar el identificador a OR.XXXX (usando punto)
                        var identificador = $"OR.{numerico}";
                        return (identificador, numerico, 0);
                    }
                }

                // Si no se puede parsear, retornar valores por defecto
                return (numeroReparacionTxt, 0, 0);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al transformar número de reparación: {numeroReparacionTxt}");
                return (numeroReparacionTxt, 0, 0);
            }
        }

        /// <summary>
        /// Transforma el número de ajuste de stock del formato del TXT al formato de MOVISTO
        /// Formato TXT: AJ.1234 u AJ-1234
        /// Formato MOVISTO: AJ.1234 u AJ-1234 (identificador completo para DoPriCor, DoPriLar, IdenLote)
        /// Formato numérico: 1234 (para DOPRINUM)
        /// Formato NUMPVTA: 0 (no aplica para ajustes)
        /// </summary>
        private (string identificador, int numerico, int numPvta) TransformarNumeroAjuste(string numeroAjusteTxt)
        {
            try
            {
                numeroAjusteTxt = numeroAjusteTxt.Trim();

                // Extraer el prefijo (AJ. u AJ-) y el valor numérico
                if (numeroAjusteTxt.StartsWith("AJ.", StringComparison.OrdinalIgnoreCase) ||
                    numeroAjusteTxt.StartsWith("AJ-", StringComparison.OrdinalIgnoreCase))
                {
                    var parteNumerica = numeroAjusteTxt.Substring(3).Trim();
                    if (int.TryParse(parteNumerica, out int numerico))
                    {
                        // Normalizar el identificador a AJ.XXXX (usando punto)
                        var identificador = $"AJ.{numerico}";
                        return (identificador, numerico, 0);
                    }
                }

                // Si no se puede parsear, retornar valores por defecto
                return (numeroAjusteTxt, 0, 0);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al transformar número de ajuste: {numeroAjusteTxt}");
                return (numeroAjusteTxt, 0, 0);
            }
        }

        /// <summary>
        /// Parsea un archivo TXT de remito y genera un objeto SyncRemitoDTO
        /// El archivo tiene delimitador punto y coma (;) y espacios de relleno
        /// </summary>
        private SyncRemitoDTO? ParsearArchivoRemitoTxt(string rutaArchivo, int numUsuar)
        {
            try
            {
                // Leer el archivo con codificación ISO-8859-1 (Latin1) para corregir caracteres especiales
                var lineas = File.ReadAllLines(rutaArchivo, Encoding.GetEncoding("ISO-8859-1"));
                if (lineas.Length < 2)
                {
                    _logger.LogWarning($"Archivo de remito sin suficientes líneas: {rutaArchivo}");
                    return null;
                }

                // Primera línea: encabezado (se ignora)
                // Segunda línea en adelante: datos de detalle
                var syncDto = new SyncRemitoDTO
                {
                    FechaGeneracion = DateTime.Now,
                    NumeroRemito = ""
                };

                foreach (var linea in lineas.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(linea))
                        continue;

                    // Dividir por punto y coma y limpiar espacios
                    var campos = linea.Split(';').Select(c => c.Trim()).ToArray();
                    
                    if (campos.Length < 22)
                    {
                        _logger.LogWarning($"Línea de remito con campos insuficientes: {linea}");
                        continue;
                    }

                    // Mapear campos del TXT al DTO
                    var item = new RemitoItemDTO
                    {
                        // Campos numéricos del TXT
                        CantIngre = decimal.TryParse(campos[8], out var cantIngre) ? cantIngre : 0,
                        CantSalid = decimal.TryParse(campos[7], out var cantSalid) ? cantSalid : 0,
                        CodInte = int.TryParse(campos[3], out var codInte) ? codInte : null,
                        NuorItem = int.TryParse(campos[5], out var nuorItem) ? nuorItem : null,
                        Nume_Clie = int.TryParse(campos[12], out var numeClie) ? numeClie : null,
                        
                        // Campos de texto del TXT
                        CodExte = campos[4],
                        DescriLar = campos[6],
                        // R.SOCIAL está en la posición 18 del TXT (ver FOREMIT07.frm en VB6).
                        // Antes se leía campos[16] que corresponde a VENDED, no a la razón social.
                        ReferCor = campos.Length > 18 ? campos[18] : "",
                        
                        // Fecha de movimiento del remito (FECHMOV)
                        FeReMovi = DateTime.TryParseExact(campos[13], "dd/MM/yy", null, 
                            System.Globalization.DateTimeStyles.None, out var fechMov) ? fechMov : null,
                        
                        // Transformar número de remito
                        DoPriCor = "",
                        DoPriLar = "",
                        DOPRINUM = 0,
                        NUMPVTA = 0
                    };

                    // Transformar número de remito del campo REMITO (posición 0)
                    var numeroRemitoTxt = campos[0];
                    var (corto, numerico, numPvta) = TransformarNumeroRemito(numeroRemitoTxt);
                    item.DoPriCor = corto;
                    item.DoPriLar = corto;
                    item.DOPRINUM = numerico;
                    item.NUMPVTA = numPvta;
                    syncDto.NumeroRemito = corto;

                    // Valores por defecto que dependen del usuario activo
                    item.NumUsuar = numUsuar;
                    item.FechMov = DateTime.Now;
                    item.FHor_Verif = DateTime.Now;

                    syncDto.Movimientos.Add(item);
                }

                if (syncDto.Movimientos.Count == 0)
                {
                    _logger.LogWarning($"Archivo de remito sin movimientos válidos: {rutaArchivo}");
                    return null;
                }

                return syncDto;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al parsear archivo de remito: {rutaArchivo}");
                return null;
            }
        }

        /// <summary>
        /// Parsea un archivo TXT de orden de reparación y genera un objeto SyncRemitoDTO
        /// El archivo tiene delimitador punto y coma (;) y espacios de relleno
        /// El identificador comienza con OR. u OR-
        /// </summary>
        private SyncRemitoDTO? ParsearArchivoReparacionTxt(string rutaArchivo, int numUsuar)
        {
            try
            {
                // Leer el archivo con codificación ISO-8859-1 (Latin1) para corregir caracteres especiales
                var lineas = File.ReadAllLines(rutaArchivo, Encoding.GetEncoding("ISO-8859-1"));
                if (lineas.Length < 2)
                {
                    _logger.LogWarning($"Archivo de reparación sin suficientes líneas: {rutaArchivo}");
                    return null;
                }

                // Primera línea: encabezado (se ignora)
                // Segunda línea en adelante: datos de detalle
                var syncDto = new SyncRemitoDTO
                {
                    FechaGeneracion = DateTime.Now,
                    NumeroRemito = ""
                };

                foreach (var linea in lineas.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(linea))
                        continue;

                    // Dividir por punto y coma y limpiar espacios
                    var campos = linea.Split(';').Select(c => c.Trim()).ToArray();

                    if (campos.Length < 22)
                    {
                        _logger.LogWarning($"Línea de reparación con campos insuficientes: {linea}");
                        continue;
                    }

                    // Mapear campos del TXT al DTO
                    var item = new RemitoItemDTO
                    {
                        // Campos numéricos del TXT
                        CantIngre = decimal.TryParse(campos[8], out var cantIngre) ? cantIngre : 0,
                        CantSalid = decimal.TryParse(campos[7], out var cantSalid) ? cantSalid : 0,
                        CodInte = int.TryParse(campos[3], out var codInte) ? codInte : null,
                        NuorItem = int.TryParse(campos[5], out var nuorItem) ? nuorItem : null,
                        Nume_Clie = int.TryParse(campos[12], out var numeClie) ? numeClie : null,

                        // Campos de texto del TXT
                        CodExte = campos[4],
                        DescriLar = campos[6],
                        // R.SOCIAL está en la posición 18 del TXT (ver FOREMIT07.frm en VB6).
                        ReferCor = campos.Length > 18 ? campos[18] : "",

                        // Fecha de movimiento de la reparación (FECHMOV)
                        FechMov = DateTime.TryParseExact(campos[13], "dd/MM/yy", null,
                            System.Globalization.DateTimeStyles.None, out var fechMov) ? fechMov : null,

                        // Transformar número de reparación
                        DoPriCor = "",
                        DoPriLar = "",
                        DoSecCor = "",
                        DoSecLar = "",
                        DOPRINUM = 0,
                        NUMPVTA = 0
                    };

                    // Transformar número de reparación del campo ORDEN (posición 0)
                    var numeroReparacionTxt = campos[0];
                    var (identificador, numerico, numPvta) = TransformarNumeroReparacion(numeroReparacionTxt);
                    item.DoPriCor = identificador;
                    item.DoPriLar = identificador;
                    item.DoSecCor = identificador;
                    item.DoSecLar = identificador;
                    item.DOPRINUM = numerico;
                    item.NUMPVTA = numPvta;
                    syncDto.NumeroRemito = identificador;

                    // Valores por defecto que dependen del usuario activo
                    item.NumUsuar = numUsuar;
                    item.FHor_Verif = DateTime.Now;

                    syncDto.Movimientos.Add(item);
                }

                if (syncDto.Movimientos.Count == 0)
                {
                    _logger.LogWarning($"Archivo de reparación sin movimientos válidos: {rutaArchivo}");
                    return null;
                }

                return syncDto;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al parsear archivo de reparación: {rutaArchivo}");
                return null;
            }
        }

        /// <summary>
        /// Parsea un archivo TXT de ajuste de stock y genera un objeto SyncRemitoDTO
        /// El archivo tiene delimitador punto y coma (;) y espacios de relleno
        /// El identificador comienza con AJ. u AJ-
        /// </summary>
        private SyncRemitoDTO? ParsearArchivoAjusteTxt(string rutaArchivo, int numUsuar)
        {
            try
            {
                // Leer el archivo con codificación ISO-8859-1 (Latin1) para corregir caracteres especiales
                var lineas = File.ReadAllLines(rutaArchivo, Encoding.GetEncoding("ISO-8859-1"));
                if (lineas.Length < 2)
                {
                    _logger.LogWarning($"Archivo de ajuste sin suficientes líneas: {rutaArchivo}");
                    return null;
                }

                // Primera línea: encabezado (se ignora)
                // Segunda línea en adelante: datos de detalle
                var syncDto = new SyncRemitoDTO
                {
                    FechaGeneracion = DateTime.Now,
                    NumeroRemito = ""
                };

                foreach (var linea in lineas.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(linea))
                        continue;

                    // Dividir por punto y coma y limpiar espacios
                    var campos = linea.Split(';').Select(c => c.Trim()).ToArray();

                    if (campos.Length < 22)
                    {
                        _logger.LogWarning($"Línea de ajuste con campos insuficientes: {linea}");
                        continue;
                    }

                    // Mapear campos del TXT al DTO
                    var item = new RemitoItemDTO
                    {
                        // Campos numéricos del TXT
                        CantIngre = decimal.TryParse(campos[8], out var cantIngre) ? cantIngre : 0,
                        CantSalid = decimal.TryParse(campos[7], out var cantSalid) ? cantSalid : 0,
                        CodInte = int.TryParse(campos[3], out var codInte) ? codInte : null,
                        NuorItem = int.TryParse(campos[5], out var nuorItem) ? nuorItem : null,
                        Nume_Clie = int.TryParse(campos[12], out var numeClie) ? numeClie : null,

                        // Campos de texto del TXT
                        CodExte = campos[4],
                        DescriLar = campos[6],
                        // R.SOCIAL está en la posición 18 del TXT (ver FOREMIT07.frm en VB6).
                        ReferCor = campos.Length > 18 ? campos[18] : "",

                        // Fecha de movimiento del ajuste (FECHMOV)
                        FechMov = DateTime.TryParseExact(campos[13], "dd/MM/yy", null,
                            System.Globalization.DateTimeStyles.None, out var fechMov) ? fechMov : null,

                        // Transformar número de ajuste
                        DoPriCor = "",
                        DoPriLar = "",
                        DoSecCor = "",
                        DoSecLar = "",
                        DOPRINUM = 0,
                        NUMPVTA = 0
                    };

                    // Transformar número de ajuste del campo AJUSTE (posición 0)
                    var numeroAjusteTxt = campos[0];
                    var (identificador, numerico, numPvta) = TransformarNumeroAjuste(numeroAjusteTxt);
                    item.DoPriCor = identificador;
                    item.DoPriLar = identificador;
                    item.DoSecCor = "";
                    item.DoSecLar = "";
                    item.DOPRINUM = numerico;
                    item.NUMPVTA = numPvta;
                    syncDto.NumeroRemito = identificador;

                    // Valores por defecto que dependen del usuario activo
                    item.NumUsuar = numUsuar;
                    item.FHor_Verif = DateTime.Now;

                    syncDto.Movimientos.Add(item);
                }

                if (syncDto.Movimientos.Count == 0)
                {
                    _logger.LogWarning($"Archivo de ajuste sin movimientos válidos: {rutaArchivo}");
                    return null;
                }

                return syncDto;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al parsear archivo de ajuste: {rutaArchivo}");
                return null;
            }
        }

        /// <summary>
        /// Genera un archivo JSON de sincronización de remitos desde un archivo TXT
        /// Mueve el TXT a PROCESADOS_TXT si tiene éxito, o a ERROR_REMITOS_TXT si falla
        /// </summary>
        public async Task<bool> GenerarArchivoSyncRemitoDesdeTxtAsync(string rutaArchivoTxt, int numUsuar)
        {
            try
            {
                var config = ObtenerConfiguracion();
                // Permitir conversión tanto en Central (ProcesaRemitos) como en Sucursal (GeneraRemitos)
                if (!config.GeneraScriptSync || (!config.GeneraRemitos && !config.ProcesaRemitos))
                {
                    _logger.LogWarning("Conversión TXT→JSON no permitida: GeneraScriptSync={GeneraScriptSync}, GeneraRemitos={GeneraRemitos}, ProcesaRemitos={ProcesaRemitos}",
                        config.GeneraScriptSync, config.GeneraRemitos, config.ProcesaRemitos);
                    return false;
                }

                var rutaIntercambio = config.RutaIntercambioRemitos;
                if (string.IsNullOrEmpty(rutaIntercambio) || !Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio de remitos no válida: {rutaIntercambio}");
                    return false;
                }

                // Crear carpetas PROCESADOS_TXT y ERROR_REMITOS_TXT si no existen
                var rutaProcesadosTxt = Path.Combine(rutaIntercambio, "PROCESADOS_TXT");
                var rutaErrorRemitosTxt = Path.Combine(rutaIntercambio, "ERROR_REMITOS_TXT");
                
                if (!Directory.Exists(rutaProcesadosTxt))
                {
                    Directory.CreateDirectory(rutaProcesadosTxt);
                }
                if (!Directory.Exists(rutaErrorRemitosTxt))
                {
                    Directory.CreateDirectory(rutaErrorRemitosTxt);
                }

                // Parsear el archivo TXT
                var syncDto = ParsearArchivoRemitoTxt(rutaArchivoTxt, numUsuar);
                if (syncDto == null)
                {
                    // Mover a ERROR_REMITOS_TXT si falla el parseo
                    var nombreArchivo = Path.GetFileName(rutaArchivoTxt);
                    var rutaError = Path.Combine(rutaErrorRemitosTxt, nombreArchivo);
                    File.Move(rutaArchivoTxt, rutaError, overwrite: true);
                    _logger.LogWarning($"Archivo TXT movido a ERROR_REMITOS_TXT por error de parseo: {nombreArchivo}");
                    return false;
                }

                // Generar nombre de archivo JSON
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var nombreArchivoJson = $"REMITO_{syncDto.NumeroRemito}_{timestamp}.json";
                var rutaCompletaJson = Path.Combine(rutaIntercambio, nombreArchivoJson);

                // Serializar a JSON
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = null
                };

                var json = JsonSerializer.Serialize(syncDto, options);
                await File.WriteAllTextAsync(rutaCompletaJson, json);

                _logger.LogInformation($"Archivo de sincronización de remito generado: {nombreArchivoJson}");

                // Mover el TXT original a PROCESADOS_TXT
                var nombreArchivoTxt = Path.GetFileName(rutaArchivoTxt);
                var rutaProcesado = Path.Combine(rutaProcesadosTxt, nombreArchivoTxt);
                File.Move(rutaArchivoTxt, rutaProcesado, overwrite: true);
                _logger.LogInformation($"Archivo TXT movido a PROCESADOS_TXT: {nombreArchivoTxt}");

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al generar archivo de sincronización para remito: {rutaArchivoTxt}");
                
                // Mover a ERROR_REMITOS_TXT si hay error general
                try
                {
                    var config = ObtenerConfiguracion();
                    var rutaErrorRemitosTxt = Path.Combine(config.RutaIntercambioRemitos, "ERROR_REMITOS_TXT");
                    if (!Directory.Exists(rutaErrorRemitosTxt))
                    {
                        Directory.CreateDirectory(rutaErrorRemitosTxt);
                    }
                    var nombreArchivo = Path.GetFileName(rutaArchivoTxt);
                    var rutaError = Path.Combine(rutaErrorRemitosTxt, nombreArchivo);
                    File.Move(rutaArchivoTxt, rutaError, overwrite: true);
                    _logger.LogWarning($"Archivo TXT movido a ERROR_REMITOS_TXT por error general: {nombreArchivo}");
                }
                catch (Exception moveEx)
                {
                    _logger.LogError(moveEx, "Error al mover archivo TXT a ERROR_REMITOS_TXT");
                }
                
                return false;
            }
        }

        /// <summary>
        /// Genera un archivo JSON de sincronización de remitos directamente desde datos de movimiento
        /// Se usa cuando IMPLANPROD genera el remito (sin archivo TXT)
        /// </summary>
        public async Task<bool> GenerarArchivoSyncRemitoDesdeDatosAsync(List<RemitoItemDTO> movimientos, int numUsuar)
        {
            try
            {
                var config = ObtenerConfiguracion();
                if (!config.GeneraScriptSync || !config.GeneraRemitos)
                {
                    return false;
                }

                var rutaIntercambio = config.RutaIntercambioRemitos;
                if (string.IsNullOrEmpty(rutaIntercambio) || !Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio de remitos no válida: {rutaIntercambio}");
                    return false;
                }

                if (movimientos == null || !movimientos.Any())
                {
                    _logger.LogWarning("No hay movimientos para generar archivo de sincronización");
                    return false;
                }

                // Asignar valores por defecto que dependen del usuario activo
                foreach (var movimiento in movimientos)
                {
                    movimiento.NumUsuar = movimiento.NumUsuar ?? numUsuar;
                    movimiento.FechMov = DateTime.Now;
                    movimiento.FHor_Verif = DateTime.Now;
                }

                // Obtener el número de remito del primer movimiento
                var numeroRemito = movimientos.First().DoPriLar ?? "";

                // Crear el DTO de sincronización
                var syncDto = new SyncRemitoDTO
                {
                    FechaGeneracion = DateTime.Now,
                    NumeroRemito = numeroRemito,
                    Movimientos = movimientos
                };

                // Generar nombre de archivo JSON
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var nombreArchivoJson = $"REMITO_{syncDto.NumeroRemito}_{timestamp}.json";
                var rutaCompletaJson = Path.Combine(rutaIntercambio, nombreArchivoJson);

                // Serializar a JSON
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = null
                };

                var json = JsonSerializer.Serialize(syncDto, options);
                await File.WriteAllTextAsync(rutaCompletaJson, json);

                _logger.LogInformation($"Archivo de sincronización de remito generado desde datos: {nombreArchivoJson} con {movimientos.Count} items");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al generar archivo de sincronización desde datos");
                return false;
            }
        }

        /// <summary>
        /// Genera un archivo JSON de sincronización de remitos directamente desde un solo movimiento
        /// Sobrecarga para compatibilidad con casos de un solo item
        /// </summary>
        public async Task<bool> GenerarArchivoSyncRemitoDesdeDatosAsync(RemitoItemDTO movimiento, int numUsuar)
        {
            return await GenerarArchivoSyncRemitoDesdeDatosAsync(new List<RemitoItemDTO> { movimiento }, numUsuar);
        }

        /// <summary>
        /// Genera un archivo JSON de sincronización de órdenes de reparación desde un archivo TXT
        /// Mueve el TXT a PROCESADOS_TXT si tiene éxito, o a ERROR_REPARACIONES_TXT si falla
        /// </summary>
        public async Task<bool> GenerarArchivoSyncReparacionDesdeTxtAsync(string rutaArchivoTxt, int numUsuar)
        {
            try
            {
                var config = ObtenerConfiguracion();
                // Permitir conversión tanto en Central (ProcesaReparaciones) como en Sucursal (GeneraReparaciones)
                if (!config.GeneraScriptSync || (!config.GeneraReparaciones && !config.ProcesaReparaciones))
                {
                    _logger.LogWarning("Conversión TXT→JSON no permitida para reparaciones: GeneraScriptSync={GeneraScriptSync}, GeneraReparaciones={GeneraReparaciones}, ProcesaReparaciones={ProcesaReparaciones}",
                        config.GeneraScriptSync, config.GeneraReparaciones, config.ProcesaReparaciones);
                    return false;
                }

                var rutaIntercambio = config.RutaIntercambioReparaciones;
                if (string.IsNullOrEmpty(rutaIntercambio) || !Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio de reparaciones no válida: {rutaIntercambio}");
                    return false;
                }

                // Crear carpetas PROCESADOS_TXT y ERROR_REPARACIONES_TXT si no existen
                var rutaProcesadosTxt = Path.Combine(rutaIntercambio, "PROCESADOS_TXT");
                var rutaErrorReparacionesTxt = Path.Combine(rutaIntercambio, "ERROR_REPARACIONES_TXT");

                if (!Directory.Exists(rutaProcesadosTxt))
                {
                    Directory.CreateDirectory(rutaProcesadosTxt);
                }
                if (!Directory.Exists(rutaErrorReparacionesTxt))
                {
                    Directory.CreateDirectory(rutaErrorReparacionesTxt);
                }

                // Parsear el archivo TXT
                var syncDto = ParsearArchivoReparacionTxt(rutaArchivoTxt, numUsuar);
                if (syncDto == null)
                {
                    // Mover a ERROR_REPARACIONES_TXT si falla el parseo
                    var nombreArchivo = Path.GetFileName(rutaArchivoTxt);
                    var rutaError = Path.Combine(rutaErrorReparacionesTxt, nombreArchivo);
                    File.Move(rutaArchivoTxt, rutaError, overwrite: true);
                    _logger.LogWarning($"Archivo TXT de reparación movido a ERROR_REPARACIONES_TXT por error de parseo: {nombreArchivo}");
                    return false;
                }

                // Generar nombre de archivo JSON
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var nombreArchivoJson = $"REPARACION_{syncDto.NumeroRemito}_{timestamp}.json";
                var rutaCompletaJson = Path.Combine(rutaIntercambio, nombreArchivoJson);

                // Serializar a JSON
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = null
                };

                var json = JsonSerializer.Serialize(syncDto, options);
                await File.WriteAllTextAsync(rutaCompletaJson, json);

                _logger.LogInformation($"Archivo de sincronización de reparación generado: {nombreArchivoJson}");

                // Mover el TXT original a PROCESADOS_TXT
                var nombreArchivoTxt = Path.GetFileName(rutaArchivoTxt);
                var rutaProcesado = Path.Combine(rutaProcesadosTxt, nombreArchivoTxt);
                File.Move(rutaArchivoTxt, rutaProcesado, overwrite: true);
                _logger.LogInformation($"Archivo TXT de reparación movido a PROCESADOS_TXT: {nombreArchivoTxt}");

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al generar archivo de sincronización para reparación: {rutaArchivoTxt}");

                // Mover a ERROR_REPARACIONES_TXT si hay error general
                try
                {
                    var config = ObtenerConfiguracion();
                    var rutaErrorReparacionesTxt = Path.Combine(config.RutaIntercambioReparaciones, "ERROR_REPARACIONES_TXT");
                    if (!Directory.Exists(rutaErrorReparacionesTxt))
                    {
                        Directory.CreateDirectory(rutaErrorReparacionesTxt);
                    }
                    var nombreArchivo = Path.GetFileName(rutaArchivoTxt);
                    var rutaError = Path.Combine(rutaErrorReparacionesTxt, nombreArchivo);
                    File.Move(rutaArchivoTxt, rutaError, overwrite: true);
                    _logger.LogWarning($"Archivo TXT de reparación movido a ERROR_REPARACIONES_TXT por error general: {nombreArchivo}");
                }
                catch (Exception moveEx)
                {
                    _logger.LogError(moveEx, "Error al mover archivo TXT de reparación a ERROR_REPARACIONES_TXT");
                }

                return false;
            }
        }

        /// <summary>
        /// Genera un archivo JSON de sincronización de ajustes de stock desde un archivo TXT
        /// Mueve el TXT a PROCESADOS_TXT si tiene éxito, o a ERROR_AJUSTES_TXT si falla
        /// </summary>
        public async Task<bool> GenerarArchivoSyncAjusteDesdeTxtAsync(string rutaArchivoTxt, int numUsuar)
        {
            try
            {
                var config = ObtenerConfiguracion();
                // Permitir conversión tanto en Central (ProcesaAjustes) como en Sucursal (GeneraAjustes)
                if (!config.GeneraScriptSync || (!config.GeneraAjustes && !config.ProcesaAjustes))
                {
                    _logger.LogWarning("Conversión TXT→JSON no permitida para ajustes: GeneraScriptSync={GeneraScriptSync}, GeneraAjustes={GeneraAjustes}, ProcesaAjustes={ProcesaAjustes}",
                        config.GeneraScriptSync, config.GeneraAjustes, config.ProcesaAjustes);
                    return false;
                }

                var rutaIntercambio = config.RutaIntercambioAjustes;
                if (string.IsNullOrEmpty(rutaIntercambio) || !Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio de ajustes no válida: {rutaIntercambio}");
                    return false;
                }

                // Crear carpetas PROCESADOS_TXT y ERROR_AJUSTES_TXT si no existen
                var rutaProcesadosTxt = Path.Combine(rutaIntercambio, "PROCESADOS_TXT");
                var rutaErrorAjustesTxt = Path.Combine(rutaIntercambio, "ERROR_AJUSTES_TXT");

                if (!Directory.Exists(rutaProcesadosTxt))
                {
                    Directory.CreateDirectory(rutaProcesadosTxt);
                }
                if (!Directory.Exists(rutaErrorAjustesTxt))
                {
                    Directory.CreateDirectory(rutaErrorAjustesTxt);
                }

                // Parsear el archivo TXT
                var syncDto = ParsearArchivoAjusteTxt(rutaArchivoTxt, numUsuar);
                if (syncDto == null)
                {
                    // Mover a ERROR_AJUSTES_TXT si falla el parseo
                    var nombreArchivo = Path.GetFileName(rutaArchivoTxt);
                    var rutaError = Path.Combine(rutaErrorAjustesTxt, nombreArchivo);
                    File.Move(rutaArchivoTxt, rutaError, overwrite: true);
                    _logger.LogWarning($"Archivo TXT de ajuste movido a ERROR_AJUSTES_TXT por error de parseo: {nombreArchivo}");
                    return false;
                }

                // Generar nombre de archivo JSON
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var nombreArchivoJson = $"AJUSTE_{syncDto.NumeroRemito}_{timestamp}.json";
                var rutaCompletaJson = Path.Combine(rutaIntercambio, nombreArchivoJson);

                // Serializar a JSON
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = null
                };

                var json = JsonSerializer.Serialize(syncDto, options);
                await File.WriteAllTextAsync(rutaCompletaJson, json);

                _logger.LogInformation($"Archivo de sincronización de ajuste generado: {nombreArchivoJson}");

                // Mover el TXT original a PROCESADOS_TXT
                var nombreArchivoTxt = Path.GetFileName(rutaArchivoTxt);
                var rutaProcesado = Path.Combine(rutaProcesadosTxt, nombreArchivoTxt);
                File.Move(rutaArchivoTxt, rutaProcesado, overwrite: true);
                _logger.LogInformation($"Archivo TXT de ajuste movido a PROCESADOS_TXT: {nombreArchivoTxt}");

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al generar archivo de sincronización para ajuste: {rutaArchivoTxt}");

                // Mover a ERROR_AJUSTES_TXT si hay error general
                try
                {
                    var config = ObtenerConfiguracion();
                    var rutaErrorAjustesTxt = Path.Combine(config.RutaIntercambioAjustes, "ERROR_AJUSTES_TXT");
                    if (!Directory.Exists(rutaErrorAjustesTxt))
                    {
                        Directory.CreateDirectory(rutaErrorAjustesTxt);
                    }
                    var nombreArchivo = Path.GetFileName(rutaArchivoTxt);
                    var rutaError = Path.Combine(rutaErrorAjustesTxt, nombreArchivo);
                    File.Move(rutaArchivoTxt, rutaError, overwrite: true);
                    _logger.LogWarning($"Archivo TXT de ajuste movido a ERROR_AJUSTES_TXT por error general: {nombreArchivo}");
                }
                catch (Exception moveEx)
                {
                    _logger.LogError(moveEx, "Error al mover archivo TXT de ajuste a ERROR_AJUSTES_TXT");
                }

                return false;
            }
        }

        /// <summary>
        /// Procesa todos los archivos JSON de órdenes de reparación pendientes en la carpeta de intercambio
        /// Se usa en la Central cuando ProcesaReparaciones = true
        /// Genera INSERT en la tabla MOVISTO
        /// </summary>
        public async Task<SyncResultDTO> ProcesarArchivosReparacionesAsync(string rutaIntercambio, int numUsuar)
        {
            var resultado = new SyncResultDTO
            {
                Exitoso = true,
                ArchivosProcessados = 0,
                RegistrosSincronizados = 0,
                Errores = 0,
                DetalleErrores = new List<string>()
            };

            try
            {
                if (string.IsNullOrWhiteSpace(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = "Ruta de intercambio de reparaciones no configurada";
                    return resultado;
                }

                if (!Directory.Exists(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = $"Ruta de intercambio de reparaciones no existe: {rutaIntercambio}";
                    return resultado;
                }

                // Buscar archivos JSON de reparaciones
                var archivosReparaciones = Directory.GetFiles(rutaIntercambio, "REPARACION_*.json");

                // Ordenar por fecha de creación (más antiguos primero)
                var archivos = archivosReparaciones.OrderBy(f => File.GetCreationTime(f)).ToList();

                if (archivos.Count == 0)
                {
                    resultado.Mensaje = "No hay archivos de reparaciones pendientes de sincronización";
                    return resultado;
                }

                foreach (var archivoPath in archivos)
                {
                    var nombreArchivo = Path.GetFileName(archivoPath);

                    // Verificar si ya fue procesado exitosamente
                    var yaProcesado = await _context.SyncProcesados
                        .AnyAsync(s => s.Archivo == nombreArchivo && s.Estado == "EXITOSO");

                    if (yaProcesado)
                    {
                        _logger.LogInformation($"Archivo de reparación ya procesado exitosamente: {nombreArchivo}");
                        continue;
                    }

                    try
                    {
                        // Leer y deserializar el archivo JSON
                        var json = await File.ReadAllTextAsync(archivoPath);
                        var syncDto = JsonSerializer.Deserialize<SyncRemitoDTO>(json);

                        if (syncDto == null || syncDto.Movimientos.Count == 0)
                        {
                            resultado.Errores++;
                            resultado.DetalleErrores.Add($"Archivo de reparación vacío o inválido: {nombreArchivo}");
                            await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "ERROR", "Archivo vacío o inválido", null);
                            continue;
                        }

                        // Procesar cada movimiento (INSERT en MOVISTO)
                        foreach (var movimiento in syncDto.Movimientos)
                        {
                            await ProcesarMovistoInsertAsync(movimiento, numUsuar);
                        }

                        resultado.ArchivosProcessados++;
                        resultado.RegistrosSincronizados += syncDto.Movimientos.Count;

                        // Registrar procesamiento exitoso
                        await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "EXITOSO", null, null);

                        _logger.LogInformation($"Archivo de reparación procesado exitosamente: {nombreArchivo}");

                        // Mover archivo procesado a carpeta PROCESADOS
                        try
                        {
                            var carpetaProcesados = Path.Combine(rutaIntercambio, "PROCESADOS");
                            if (!Directory.Exists(carpetaProcesados))
                            {
                                Directory.CreateDirectory(carpetaProcesados);
                            }
                            var rutaDestino = Path.Combine(carpetaProcesados, nombreArchivo);
                            File.Move(archivoPath, rutaDestino, overwrite: true);
                            _logger.LogInformation($"Archivo de reparación movido a PROCESADOS: {nombreArchivo}");
                        }
                        catch (Exception moveEx)
                        {
                            _logger.LogError(moveEx, $"No se pudo mover archivo de reparación a carpeta PROCESADOS: {nombreArchivo}");
                        }
                    }
                    catch (Exception ex)
                    {
                        resultado.Errores++;
                        var mensajeError = $"Error en {nombreArchivo}: {ex.Message}";
                        resultado.DetalleErrores.Add(mensajeError);

                        await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "ERROR", ex.Message, null);

                        // Mover archivo fallido a carpeta ERROR_REPARACIONES
                        try
                        {
                            var carpetaError = Path.Combine(rutaIntercambio, "ERROR_REPARACIONES");
                            if (!Directory.Exists(carpetaError))
                            {
                                Directory.CreateDirectory(carpetaError);
                            }
                            var rutaDestino = Path.Combine(carpetaError, nombreArchivo);
                            File.Move(archivoPath, rutaDestino, overwrite: true);
                        }
                        catch (Exception moveEx)
                        {
                            _logger.LogError(moveEx, $"No se pudo mover archivo de reparación a carpeta ERROR_REPARACIONES: {nombreArchivo}");
                        }
                    }
                }

                resultado.Mensaje = $"Procesamiento de reparaciones completado: {resultado.ArchivosProcessados} archivos, {resultado.RegistrosSincronizados} movimientos";
                return resultado;
            }
            catch (Exception ex)
            {
                resultado.Exitoso = false;
                resultado.Mensaje = $"Error general al procesar archivos de reparaciones: {ex.Message}";
                _logger.LogError(ex, "Error general al procesar archivos de reparaciones");
                return resultado;
            }
        }

        /// <summary>
        /// Procesa todos los archivos JSON de ajustes de stock pendientes en la carpeta de intercambio
        /// Se usa en la Central cuando ProcesaAjustes = true
        /// Genera INSERT en la tabla MOVISTO
        /// </summary>
        public async Task<SyncResultDTO> ProcesarArchivosAjustesAsync(string rutaIntercambio, int numUsuar)
        {
            var resultado = new SyncResultDTO
            {
                Exitoso = true,
                ArchivosProcessados = 0,
                RegistrosSincronizados = 0,
                Errores = 0,
                DetalleErrores = new List<string>()
            };

            try
            {
                if (string.IsNullOrWhiteSpace(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = "Ruta de intercambio de ajustes no configurada";
                    return resultado;
                }

                if (!Directory.Exists(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = $"Ruta de intercambio de ajustes no existe: {rutaIntercambio}";
                    return resultado;
                }

                // Buscar archivos JSON de ajustes
                var archivosAjustes = Directory.GetFiles(rutaIntercambio, "AJUSTE_*.json");

                // Ordenar por fecha de creación (más antiguos primero)
                var archivos = archivosAjustes.OrderBy(f => File.GetCreationTime(f)).ToList();

                if (archivos.Count == 0)
                {
                    resultado.Mensaje = "No hay archivos de ajustes pendientes de sincronización";
                    return resultado;
                }

                foreach (var archivoPath in archivos)
                {
                    var nombreArchivo = Path.GetFileName(archivoPath);

                    // Verificar si ya fue procesado exitosamente
                    var yaProcesado = await _context.SyncProcesados
                        .AnyAsync(s => s.Archivo == nombreArchivo && s.Estado == "EXITOSO");

                    if (yaProcesado)
                    {
                        _logger.LogInformation($"Archivo de ajuste ya procesado exitosamente: {nombreArchivo}");
                        continue;
                    }

                    try
                    {
                        // Leer y deserializar el archivo JSON
                        var json = await File.ReadAllTextAsync(archivoPath);
                        var syncDto = JsonSerializer.Deserialize<SyncRemitoDTO>(json);

                        if (syncDto == null || syncDto.Movimientos.Count == 0)
                        {
                            resultado.Errores++;
                            resultado.DetalleErrores.Add($"Archivo de ajuste vacío o inválido: {nombreArchivo}");
                            await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "ERROR", "Archivo vacío o inválido", null);
                            continue;
                        }

                        // Procesar cada movimiento (INSERT en MOVISTO)
                        foreach (var movimiento in syncDto.Movimientos)
                        {
                            await ProcesarMovistoInsertAsync(movimiento, numUsuar);
                        }

                        resultado.ArchivosProcessados++;
                        resultado.RegistrosSincronizados += syncDto.Movimientos.Count;

                        // Registrar procesamiento exitoso
                        await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "EXITOSO", null, null);

                        _logger.LogInformation($"Archivo de ajuste procesado exitosamente: {nombreArchivo}");

                        // Mover archivo procesado a carpeta PROCESADOS
                        try
                        {
                            var carpetaProcesados = Path.Combine(rutaIntercambio, "PROCESADOS");
                            if (!Directory.Exists(carpetaProcesados))
                            {
                                Directory.CreateDirectory(carpetaProcesados);
                            }
                            var rutaDestino = Path.Combine(carpetaProcesados, nombreArchivo);
                            File.Move(archivoPath, rutaDestino, overwrite: true);
                            _logger.LogInformation($"Archivo de ajuste movido a PROCESADOS: {nombreArchivo}");
                        }
                        catch (Exception moveEx)
                        {
                            _logger.LogError(moveEx, $"No se pudo mover archivo de ajuste a carpeta PROCESADOS: {nombreArchivo}");
                        }
                    }
                    catch (Exception ex)
                    {
                        resultado.Errores++;
                        var mensajeError = $"Error en {nombreArchivo}: {ex.Message}";
                        resultado.DetalleErrores.Add(mensajeError);

                        await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "ERROR", ex.Message, null);

                        // Mover archivo fallido a carpeta ERROR_AJUSTES
                        try
                        {
                            var carpetaError = Path.Combine(rutaIntercambio, "ERROR_AJUSTES");
                            if (!Directory.Exists(carpetaError))
                            {
                                Directory.CreateDirectory(carpetaError);
                            }
                            var rutaDestino = Path.Combine(carpetaError, nombreArchivo);
                            File.Move(archivoPath, rutaDestino, overwrite: true);
                        }
                        catch (Exception moveEx)
                        {
                            _logger.LogError(moveEx, $"No se pudo mover archivo de ajuste a carpeta ERROR_AJUSTES: {nombreArchivo}");
                        }
                    }
                }

                resultado.Mensaje = $"Procesamiento de ajustes completado: {resultado.ArchivosProcessados} archivos, {resultado.RegistrosSincronizados} movimientos";
                return resultado;
            }
            catch (Exception ex)
            {
                resultado.Exitoso = false;
                resultado.Mensaje = $"Error general al procesar archivos de ajustes: {ex.Message}";
                _logger.LogError(ex, "Error general al procesar archivos de ajustes");
                return resultado;
            }
        }

        /// <summary>
        /// Procesa todos los archivos JSON de remitos pendientes en la carpeta de intercambio
        /// Se usa en la Central cuando ProcesaRemitos = true
        /// Genera INSERT en la tabla MOVISTO
        /// </summary>
        public async Task<SyncResultDTO> ProcesarArchivosRemitosAsync(string rutaIntercambio, int numUsuar)
        {
            var resultado = new SyncResultDTO
            {
                Exitoso = true,
                ArchivosProcessados = 0,
                RegistrosSincronizados = 0,
                Errores = 0,
                DetalleErrores = new List<string>()
            };

            try
            {
                if (string.IsNullOrWhiteSpace(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = "Ruta de intercambio de remitos no configurada";
                    return resultado;
                }

                if (!Directory.Exists(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = $"Ruta de intercambio de remitos no existe: {rutaIntercambio}";
                    return resultado;
                }

                // Buscar archivos JSON de remitos
                var archivosRemitos = Directory.GetFiles(rutaIntercambio, "REMITO_*.json");
                
                // Ordenar por fecha de creación (más antiguos primero)
                var archivos = archivosRemitos.OrderBy(f => File.GetCreationTime(f)).ToList();

                if (archivos.Count == 0)
                {
                    resultado.Mensaje = "No hay archivos de remitos pendientes de sincronización";
                    return resultado;
                }

                foreach (var archivoPath in archivos)
                {
                    var nombreArchivo = Path.GetFileName(archivoPath);

                    // Verificar si ya fue procesado exitosamente
                    var yaProcesado = await _context.SyncProcesados
                        .AnyAsync(s => s.Archivo == nombreArchivo && s.Estado == "EXITOSO");

                    if (yaProcesado)
                    {
                        _logger.LogInformation($"Archivo de remito ya procesado exitosamente: {nombreArchivo}");
                        continue;
                    }

                    try
                    {
                        // Leer y deserializar el archivo JSON
                        var json = await File.ReadAllTextAsync(archivoPath);
                        var syncDto = JsonSerializer.Deserialize<SyncRemitoDTO>(json);

                        if (syncDto == null || syncDto.Movimientos.Count == 0)
                        {
                            resultado.Errores++;
                            resultado.DetalleErrores.Add($"Archivo de remito vacío o inválido: {nombreArchivo}");
                            await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "ERROR", "Archivo vacío o inválido", null);
                            continue;
                        }

                        // Procesar cada movimiento (INSERT en MOVISTO)
                        foreach (var movimiento in syncDto.Movimientos)
                        {
                            await ProcesarMovistoInsertAsync(movimiento, numUsuar);
                        }

                        resultado.ArchivosProcessados++;
                        resultado.RegistrosSincronizados += syncDto.Movimientos.Count;

                        // Registrar procesamiento exitoso
                        await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "EXITOSO", null, null);

                        _logger.LogInformation($"Archivo de remito procesado exitosamente: {nombreArchivo}");

                        // Mover archivo procesado a carpeta PROCESADOS
                        try
                        {
                            var carpetaProcesados = Path.Combine(rutaIntercambio, "PROCESADOS");
                            if (!Directory.Exists(carpetaProcesados))
                            {
                                Directory.CreateDirectory(carpetaProcesados);
                            }
                            var rutaDestino = Path.Combine(carpetaProcesados, nombreArchivo);
                            File.Move(archivoPath, rutaDestino, overwrite: true);
                            _logger.LogInformation($"Archivo movido a PROCESADOS: {nombreArchivo}");
                        }
                        catch (Exception moveEx)
                        {
                            _logger.LogError(moveEx, $"No se pudo mover archivo a carpeta PROCESADOS: {nombreArchivo}");
                        }
                    }
                    catch (Exception ex)
                    {
                        resultado.Errores++;
                        var mensajeError = $"Error en {nombreArchivo}: {ex.Message}";
                        resultado.DetalleErrores.Add(mensajeError);

                        await RegistrarProcesamientoAsync(nombreArchivo, "MOVISTO", "INSERT", "ERROR", ex.Message, null);

                        // Mover archivo fallido a carpeta ERROR_REMITOS
                        try
                        {
                            var carpetaError = Path.Combine(rutaIntercambio, "ERROR_REMITOS");
                            if (!Directory.Exists(carpetaError))
                            {
                                Directory.CreateDirectory(carpetaError);
                            }
                            var rutaDestino = Path.Combine(carpetaError, nombreArchivo);
                            File.Move(archivoPath, rutaDestino, overwrite: true);
                        }
                        catch (Exception moveEx)
                        {
                            _logger.LogError(moveEx, $"No se pudo mover archivo a carpeta ERROR_REMITOS: {nombreArchivo}");
                        }
                    }
                }

                resultado.Mensaje = $"Procesamiento completado: {resultado.ArchivosProcessados} archivos, {resultado.RegistrosSincronizados} movimientos";
                return resultado;
            }
            catch (Exception ex)
            {
                resultado.Exitoso = false;
                resultado.Mensaje = $"Error general al procesar archivos de remitos: {ex.Message}";
                _logger.LogError(ex, "Error general al procesar archivos de remitos");
                return resultado;
            }
        }

        /// <summary>
        /// Procesa un movimiento de stock individual generando un INSERT en MOVISTO
        /// Mapea RemitoItemDTO a la entidad Movisto existente
        /// </summary>
        private async Task ProcesarMovistoInsertAsync(RemitoItemDTO movimiento, int numUsuar)
        {
            try
            {
                // Detectar si es una orden de reparación (DoPriLar comienza con OR- u OR.)
                // o un ajuste de stock (DoPriLar comienza con AJ- u AJ.)
                var doPriLar = movimiento.DoPriLar ?? "";
                var esReparacion = doPriLar.StartsWith("OR-", StringComparison.OrdinalIgnoreCase) ||
                                   doPriLar.StartsWith("OR.", StringComparison.OrdinalIgnoreCase);
                var esAjuste = doPriLar.StartsWith("AJ-", StringComparison.OrdinalIgnoreCase) ||
                               doPriLar.StartsWith("AJ.", StringComparison.OrdinalIgnoreCase);

                // Si ReferCor está vacío, buscar la razón social en Deudor12
                string referCor = movimiento.ReferCor ?? "";
                if (string.IsNullOrWhiteSpace(referCor) && movimiento.Nume_Clie.HasValue)
                {
                    var deudor = await _context.Clientes
                        .Where(d => d.Nume_Cli == movimiento.Nume_Clie.Value)
                        .Select(d => d.Raso_Cli)
                        .FirstOrDefaultAsync();

                    if (!string.IsNullOrWhiteSpace(deudor))
                    {
                        referCor = deudor;
                        _logger.LogInformation("ReferCor obtenido de Deudor12 para cliente {Cliente}: {RazonSocial}", movimiento.Nume_Clie, referCor);
                    }
                }

                // Para reparaciones: DoPriCor, DoPriLar, DoSecCor, DoSecLar e IdenLote deben tener el mismo identificador
                // Para ajustes: DoPriCor, DoPriLar e IdenLote deben tener el mismo identificador, DoSecCor y DoSecLar vacíos
                var identificador = doPriLar;
                if (esReparacion)
                {
                    _logger.LogInformation("Procesando orden de reparación: {Identificador}", identificador);
                }
                else if (esAjuste)
                {
                    _logger.LogInformation("Procesando ajuste de stock: {Identificador}", identificador);
                }

                // Crear la entidad MOVISTO a partir del DTO de sincronización
                var movisto = new Movisto
                {
                    // Campos mapeados del archivo TXT
                    CantIngre = movimiento.CantIngre ?? 0,
                    CantSalid = movimiento.CantSalid ?? 0,
                    Cod_Exte = movimiento.CodExte,
                    Cod_Inte = movimiento.CodInte,
                    Descri_Lar = movimiento.DescriLar,
                    Descrip = !string.IsNullOrEmpty(movimiento.DescriLar)
                        ? movimiento.DescriLar.Substring(0, Math.Min(64, movimiento.DescriLar.Length))
                        : "",
                    DoPriCor = (esReparacion || esAjuste) ? identificador : movimiento.DoPriCor,
                    DoPriLar = (esReparacion || esAjuste) ? identificador : movimiento.DoPriLar,
                    DoSecCor = esAjuste ? "" : (esReparacion ? identificador : movimiento.DoSecCor),
                    DoSecLar = esAjuste ? "" : (esReparacion ? identificador : movimiento.DoSecLar),
                    DOPRINUM = movimiento.DOPRINUM,
                    Nume_Clie = movimiento.Nume_Clie,
                    NUMPVTA = (short?)movimiento.NUMPVTA,
                    NuorItem = movimiento.NuorItem,
                    ReferCor = referCor,
                    FeReMovi = (esReparacion || esAjuste) ? DateTime.Now : movimiento.FeReMovi,

                    // Campos con valores por defecto
                    TipoCam = movimiento.TipoCam,
                    TipoTran = (short?)movimiento.TipoTran,
                    ValoUnit = movimiento.ValoUnit,
                    ValoUnitIva = movimiento.ValoUnitIva,
                    ValTotIt = movimiento.ValTotIt,
                    ValTotItIva = movimiento.ValTotItIva,
                    VerificadoPor = (short?)movimiento.VerificadoPor,
                    CantSaldo = movimiento.CantSaldo,
                    CodiEmpr = movimiento.CodiEmpr,
                    FechMov = movimiento.FechMov,
                    FHor_Verif = movimiento.FHor_Verif,
                    ID_OPERACION = (short?)movimiento.ID_OPERACION,
                    IdCalidad = movimiento.IdCalidad,
                    MoneVal = (short?)movimiento.MoneVal,
                    Nume_Prov = movimiento.Nume_Prov,
                    NumeOrdPed = (short?)movimiento.NumeOrdPed,
                    NuPedCli = movimiento.NuPedCli,
                    NumUsuar = movimiento.NumUsuar ?? numUsuar,
                    CodiMovi = esReparacion ? "OR" : (esAjuste ? "AJ" : movimiento.CodiMovi),
                    DepoMovi = 10, // Siempre 10 en la migración
                    IdenLote = !string.IsNullOrEmpty(identificador)
                        ? identificador.Substring(0, Math.Min(16, identificador.Length))
                        : "",

                    // Campos de auditoría automáticos
                    AddRecord = DateTime.Now,
                    LastUpdate = DateTime.Now,
                    CodiEmprNet = movimiento.CodiEmpr
                };

                // Agregar al contexto y guardar
                _context.Movistos.Add(movisto);
                await _context.SaveChangesAsync();
                
                _logger.LogInformation($"Movimiento de stock insertado: Id={movisto.Id}, CodInte={movimiento.CodInte}, CantSalid={movisto.CantSalid}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al procesar movimiento: CodInte={movimiento.CodInte}");
                throw;
            }
        }

        /// <summary>
        /// Lee un archivo Ocompra_Sucursal.txt y lo convierte a DTO
        /// Este archivo se genera en la sucursal para sincronizar pedidos de reposición con la central
        ///
        /// Formato del archivo según código VB6:
        /// - Nombre: Ocompra_Sucursal AAAA_MM_DD HHMMSS.TXT
        /// - Delimitador: punto y coma (;)
        /// - Primera línea: encabezado con nombres de columnas
        /// - Líneas siguientes: datos de cada registro
        /// - Última línea opcional: mensaje con formato ||MENSAJE||;texto
        ///
        /// Estructura de cada línea de datos (12 campos separados por ;):
        /// 1. SUCURSAL (8 chars) - Número de sucursal
        /// 2. CODIGO (8 chars) - Código interno del producto (Codint)
        /// 3. NRO.PED. (8 chars) - Número de pedido
        /// 4. CTE.PED. (8 chars) - Cliente del pedido
        /// 5. USUARIO (8 chars) - Usuario que generó el pedido
        /// 6. OCOMPRA (8 chars) - Número de orden de compra
        /// 7. CANTIDAD (10 chars) - Cantidad solicitada
        /// 8. SOLENTRE (10 chars) - Fecha solicitud de entrega (dd/MM/yyyy)
        /// 9. VERS-COMP (8 chars) - Versión
        /// 10. FEMISION (8 chars) - Fecha de emisión (dd/MM/yy)
        /// 11. O.COMPRA CLIENTE (16 chars) - Orden de compra del cliente
        /// 12. R.SOCIAL CTE.PEDIDO (32 chars) - Razón social del cliente
        ///
        /// NOTA PARA DESARROLLADORES JUNIOR:
        /// - Este método es ROJO (solo lectura): no modifica la base de datos
        /// - Usa File.ReadAllLinesAsync para lectura asíncrona del archivo
        /// - El formato del archivo viene del sistema VB6 original, no cambiarlo
        /// - Los campos tienen anchos fijos pero están separados por punto y coma
        /// - La línea de mensaje es opcional y solo aparece si hay observaciones
        /// </summary>
        public async Task<OcompraSucursalArchivoResultadoDTO> LeerArchivoOcompraSucursalAsync(string rutaArchivo)
        {
            var resultado = new OcompraSucursalArchivoResultadoDTO
            {
                NombreArchivo = Path.GetFileName(rutaArchivo)
            };

            try
            {
                // ====================================================================
                // PASO 1: Extraer fecha de generación del nombre del archivo
                // ====================================================================
                // El nombre del archivo tiene el formato: Ocompra_Sucursal AAAA_MM_DD HHMMSS.TXT
                // Ejemplo: Ocompra_Sucursal 2024_08_06 103045.TXT
                // Extraemos la fecha y hora para auditoría y trazabilidad
                var nombreSinExtension = Path.GetFileNameWithoutExtension(rutaArchivo);
                var partes = nombreSinExtension.Split(' ');
                if (partes.Length >= 3)
                {
                    resultado.FechaGeneracion = $"{partes[1]} {partes[2]}";
                }

                // ====================================================================
                // PASO 2: Leer todas las líneas del archivo
                // ====================================================================
                // File.ReadAllLinesAsync lee todo el archivo en memoria de forma asíncrona
                // Esto es adecuado para archivos pequeños/medianos (menos de 10MB)
                // Para archivos muy grandes, considerar leer línea por línea con StreamReader
                var lineas = await File.ReadAllLinesAsync(rutaArchivo);

                if (lineas.Length == 0)
                {
                    _logger.LogWarning($"Archivo vacío: {rutaArchivo}");
                    return resultado;
                }

                // ====================================================================
                // PASO 3: Procesar cada línea del archivo
                // ====================================================================
                foreach (var linea in lineas)
                {
                    // Saltar líneas vacías
                    if (string.IsNullOrWhiteSpace(linea))
                        continue;

                    // ====================================================================
                    // CASO 1: Línea de mensaje (última línea opcional)
                    // ====================================================================
                    // Formato: ||MENSAJE||;texto del mensaje
                    // Esta línea solo aparece si hay observaciones sobre el pedido
                    // El mensaje puede tener hasta 1000 caracteres según código VB6
                    if (linea.StartsWith("||MENSAJE||"))
                    {
                        // Extraer el mensaje después del delimitador ";"
                        var partesMensaje = linea.Split(';');
                        if (partesMensaje.Length > 1)
                        {
                            resultado.Mensaje = partesMensaje[1].Trim();
                        }
                        continue; // Pasar a la siguiente línea
                    }

                    // ====================================================================
                    // CASO 2: Línea de encabezado (primera línea)
                    // ====================================================================
                    // La primera línea contiene los nombres de las columnas
                    // Comienza con "SUCURSAL" y no contiene datos, solo títulos
                    // La saltamos porque no necesitamos procesarla
                    if (linea.StartsWith("SUCURSAL"))
                    {
                        continue; // Saltar encabezado
                    }

                    // ====================================================================
                    // CASO 3: Línea de datos (registros de pedidos)
                    // ====================================================================
                    // Cada línea de datos tiene 12 campos separados por punto y coma (;)
                    // Parseamos cada campo y lo asignamos al DTO correspondiente
                    var campos = linea.Split(';');
                    if (campos.Length >= 12)
                    {
                        var registro = new OcompraSucursalArchivoDTO
                        {
                            // Campo 1: SUCURSAL - Número de sucursal (ej: 1, 2, 3)
                            Sucursal = campos[0].Trim(),
                            // Campo 2: CODIGO - Código interno del producto (Codint de masters)
                            Codigo = campos[1].Trim(),
                            // Campo 3: NRO.PED. - Número de pedido generado
                            NroPedido = campos[2].Trim(),
                            // Campo 4: CTE.PED. - Cliente del pedido
                            CtePedido = campos[3].Trim(),
                            // Campo 5: USUARIO - Usuario que generó el pedido (USNBR% Mod 100 en VB6)
                            Usuario = campos[4].Trim(),
                            // Campo 6: OCOMPRA - Número de orden de compra de sucursal
                            OCompra = campos[5].Trim(),
                            // Campo 7: CANTIDAD - Cantidad solicitada del producto
                            Cantidad = campos[6].Trim(),
                            // Campo 8: SOLENTRE - Fecha solicitud de entrega (formato dd/MM/yyyy)
                            SolEntre = campos[7].Trim(),
                            // Campo 9: VERS-COMP - Versión de la orden de compra
                            Version = campos[8].Trim(),
                            // Campo 10: FEMISION - Fecha de emisión (formato dd/MM/yy)
                            Femision = campos[9].Trim(),
                            // Campo 11: O.COMPRA CLIENTE - Orden de compra del cliente (opcional)
                            OCompraCliente = campos[10].Trim(),
                            // Campo 12: R.SOCIAL CTE.PEDIDO - Razón social del cliente
                            RSocialCtePedido = campos[11].Trim()
                        };

                        resultado.Registros.Add(registro);
                    }
                }

                _logger.LogInformation($"Archivo Ocompra_Sucursal leído: {resultado.NombreArchivo}, Registros: {resultado.Registros.Count}");
                return resultado;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al leer archivo Ocompra_Sucursal: {rutaArchivo}");
                throw; // Re-lanzar la excepción para que el llamador pueda manejarla
            }
        }

        /// <summary>
        /// Procesa los registros de un archivo Ocompra_Sucursal.txt
        /// Inserta los registros en la tabla OC_SUCURSAL según la lógica VB6 original
        ///
        /// Estrategia de procesamiento según código VB6:
        /// 1. MARCAR REGISTROS EXISTENTES: Antes de procesar, marca MARBORRA = 1 para todos
        ///    los registros existentes con la misma O_COMPRA_SUC y SUCURSAL (equivalente a VUELTA% = 0)
        /// 2. INSERTAR NUEVOS REGISTROS: Inserta todos los registros del archivo (SIEMPRE INSERT, nunca UPDATE)
        ///    según código VB6: ".AddNew" sin condicional
        /// 3. ELIMINAR REGISTROS MARCADOS: Elimina los registros con MARBORRA = 1 (los antiguos)
        ///    para reemplazar completamente los datos anteriores con los del archivo
        /// 4. GRABAR MENSAJES: Si el archivo tiene un mensaje (||MENSAJE||), lo graba en MENSA_SUCURSAL
        ///    según código VB6: "Call GrabarMensaje"
        /// 5. Usa transacción para garantizar atomicidad (todo o nada)
        ///
        /// NOTA PARA DESARROLLADORES JUNIOR:
        /// - Este método MODIFICA la base de datos (INSERT/DELETE)
        /// - Usa Entity Framework Core para operaciones de base de datos
        /// - La transacción garantiza que si falla un registro, se deshacen todos los cambios
        /// - La lógica es diferente a un UPSERT normal: primero marca, luego inserta, luego elimina
        /// - Esto es para garantizar que los datos del archivo reemplacen completamente los anteriores
        /// - Marborra = 1 significa "marcado para borrar", 0 significa activo
        /// </summary>
        public async Task<bool> ProcesarRegistrosOcompraSucursalAsync(OcompraSucursalArchivoResultadoDTO archivoResultado, int numUsuar)
        {
            // ====================================================================
            // INICIAR TRANSACCIÓN
            // ====================================================================
            // La transacción garantiza atomicidad: o se guardan todos los registros,
            // o no se guarda ninguno si hay algún error.
            // Esto es crítico para mantener la integridad de los datos.
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                // ====================================================================
                // PASO 1: MARCAR REGISTROS EXISTENTES CON MARBORRA = 1
                // ====================================================================
                // Según código VB6: "LA PRIMERA VUELTA ANTES DE GRABAR LE MARCA CAMPO MARBORRA = 1
                // POR SI EXISTIA EL REGISTRO LUEGO AL FINAL LOS ELIMINA"
                // Esto marca todos los registros existentes para la misma O_COMPRA_SUC y SUCURSAL
                // para luego eliminarlos después de insertar los nuevos.

                // Obtener todos los OCompraSuc y Sucursal únicos del archivo
                var grupos = archivoResultado.Registros
                    .Where(r => int.TryParse(r.Sucursal, out int s) && int.TryParse(r.OCompra, out int o))
                    .Select(r => new { Sucursal = int.Parse(r.Sucursal), OCompraSuc = int.Parse(r.OCompra) })
                    .Distinct()
                    .ToList();

                foreach (var grupo in grupos)
                {
                    // Marcar todos los registros existentes con MARBORRA = 1
                    var registrosAMarcar = await _context.O_Sucursals
                        .Where(o => o.OCompraSuc == grupo.OCompraSuc && o.Sucursal == grupo.Sucursal && o.Marborra != -1)
                        .ToListAsync();

                    foreach (var reg in registrosAMarcar)
                    {
                        reg.Marborra = 1; // Marcar para borrar
                        reg.LastUpdate = DateTime.Now;
                        reg.NumUsuar = numUsuar;
                    }

                    _logger.LogInformation($"Marcados {registrosAMarcar.Count} registros con MARBORRA=1: OCompraSuc={grupo.OCompraSuc}, Sucursal={grupo.Sucursal}");
                }

                await _context.SaveChangesAsync();

                // ====================================================================
                // PASO 2: INSERTAR NUEVOS REGISTROS (SIEMPRE INSERT)
                // ====================================================================
                // Según código VB6: ".AddNew" sin condicional - SIEMPRE inserta, nunca actualiza
                // Esto reemplaza completamente los datos anteriores con los del archivo

                foreach (var registro in archivoResultado.Registros)
                {
                    // ====================================================================
                    // VALIDACIÓN Y PARSEO DE CAMPOS NUMÉRICOS
                    // ====================================================================
                    // Los campos del archivo vienen como string, hay que convertirlos
                    // Si un campo no se puede convertir, se registra un warning y se salta el registro

                    // Parsear Sucursal (debe ser un entero válido)
                    if (!int.TryParse(registro.Sucursal, out int sucursal))
                    {
                        _logger.LogWarning($"Sucursal inválida: {registro.Sucursal}");
                        continue; // Saltar este registro y continuar con el siguiente
                    }

                    // Parsear Codint (código interno del producto)
                    if (!int.TryParse(registro.Codigo, out int codint))
                    {
                        _logger.LogWarning($"Código inválido: {registro.Codigo}");
                        continue;
                    }

                    // Parsear OCompraSuc (número de orden de compra)
                    if (!int.TryParse(registro.OCompra, out int oCompraSuc))
                    {
                        _logger.LogWarning($"OCompra inválido: {registro.OCompra}");
                        continue;
                    }

                    // Parsear Cantidad (debe ser un decimal válido)
                    if (!decimal.TryParse(registro.Cantidad, out decimal cantidad))
                    {
                        _logger.LogWarning($"Cantidad inválida: {registro.Cantidad}");
                        continue;
                    }

                    // Parsear NroPedido (número de pedido)
                    int nroPedido = 0;
                    if (!string.IsNullOrWhiteSpace(registro.NroPedido))
                    {
                        int.TryParse(registro.NroPedido, out nroPedido);
                    }

                    // Parsear CtePedido (cliente del pedido)
                    int ctePedido = 0;
                    if (!string.IsNullOrWhiteSpace(registro.CtePedido))
                    {
                        int.TryParse(registro.CtePedido, out ctePedido);
                    }

                    // Parsear Usuario (usuario que generó el pedido)
                    int usuario = 0;
                    if (!string.IsNullOrWhiteSpace(registro.Usuario))
                    {
                        int.TryParse(registro.Usuario, out usuario);
                    }

                    // Parsear Version (versión de la orden de compra)
                    int version = 0;
                    if (!string.IsNullOrWhiteSpace(registro.Version))
                    {
                        int.TryParse(registro.Version, out version);
                    }

                    // ====================================================================
                    // PARSEO DE FECHAS (formato VB6: dd/MM/yyyy o dd/MM/yy)
                    // ====================================================================
                    // DateTime.TryParse es flexible y acepta ambos formatos
                    // Si la fecha es inválida, se guarda como null
                    DateTime? fechaSoliciCompra = null;
                    if (!string.IsNullOrWhiteSpace(registro.SolEntre))
                    {
                        if (DateTime.TryParse(registro.SolEntre, out DateTime fechaSol))
                        {
                            fechaSoliciCompra = fechaSol;
                        }
                    }

                    DateTime? fechaEmision = null;
                    if (!string.IsNullOrWhiteSpace(registro.Femision))
                    {
                        if (DateTime.TryParse(registro.Femision, out DateTime fechaEmi))
                        {
                            fechaEmision = fechaEmi;
                        }
                    }

                    // ====================================================================
                    // INSERTAR NUEVO REGISTRO (SIEMPRE INSERT, NUNCA UPDATE)
                    // ====================================================================
                    // Según código VB6: ".AddNew" sin condicional - SIEMPRE inserta
                    // Creamos un nuevo registro con todos los campos requeridos
                    // Marborra = 0 significa que el registro está activo (no borrado)
                    // CodiEmprNet = 0 según las reglas del sistema (valor por defecto)
                    var nuevoRegistro = new Oc_Sucursal
                    {
                        OCompraSuc = oCompraSuc,
                        Sucursal = (short)sucursal, // Convertir a short porque el campo es SMALLINT
                        Codint = codint,
                        Cantidad = cantidad,
                        FechaSoliciCompra = fechaSoliciCompra,
                        Fechemisi = fechaEmision, // FECHEMISI = CVDAT(FEMISI%)
                        Marborra = 0, // 0 = activo, -1 = borrado
                        Status = 0, // 0 = activo según VB6
                        AddRecord = DateTime.Now,
                        LastUpdate = DateTime.Now,
                        CodiEmprNet = 0,
                        NumUsuar = numUsuar,
                        // Campos adicionales según VB6
                        NropedSuc = nroPedido, // NROPED_SUC = NROPEDIDOSUCURSAL
                        NclienPedido = ctePedido, // NCLIEN_PEDIDO = CLIENTEPEDIDO
                        Usuario = usuario, // USUARIO = USUARIOX%
                        NpedGeneradoEnAvellaneda = 0, // NPED_GENERADO_EN_AVELLANEDA
                        VersionCompraSuc = (short)version, // VERSION_COMPRA_SUC (es short)
                        OComCliente = registro.OCompraCliente?.Substring(0, Math.Min(16, registro.OCompraCliente?.Length ?? 0)), // OCOMPRACLIENTE$
                        RazCteSuc = registro.RSocialCtePedido?.Substring(0, Math.Min(32, registro.RSocialCtePedido?.Length ?? 0)) // RAZ_CTE_SUC
                    };

                    _context.O_Sucursals.Add(nuevoRegistro);
                    _logger.LogInformation($"Insertado OC_SUCURSAL: OCompraSuc={oCompraSuc}, Sucursal={sucursal}, Codint={codint}");
                }

                await _context.SaveChangesAsync();

                // ====================================================================
                // PASO 3: ELIMINAR REGISTROS CON MARBORRA = 1
                // ====================================================================
                // Según código VB6: "LUEGO AL FINAL LOS ELIMINA"
                // Eliminamos los registros que marcamos en el paso 1
                foreach (var grupo in grupos)
                {
                    var registrosAEliminar = await _context.O_Sucursals
                        .Where(o => o.OCompraSuc == grupo.OCompraSuc && o.Sucursal == grupo.Sucursal && o.Marborra == 1)
                        .ToListAsync();

                    _context.O_Sucursals.RemoveRange(registrosAEliminar);
                    _logger.LogInformation($"Eliminados {registrosAEliminar.Count} registros con MARBORRA=1: OCompraSuc={grupo.OCompraSuc}, Sucursal={grupo.Sucursal}");
                }

                await _context.SaveChangesAsync();

                // ====================================================================
                // PASO 4: GRABAR MENSAJES EN MENSA_SUCURSAL
                // ====================================================================
                // Según código VB6: "Call GrabarMensaje(CLTESUCUR, NUOC&, ACUMUMENXAJE$, USUARIOX%)"
                // Si el archivo tiene un mensaje, lo grabamos en la tabla MENSA_SUCURSAL
                if (!string.IsNullOrWhiteSpace(archivoResultado.Mensaje))
                {
                    // Grabar el mensaje para cada grupo (OCompraSuc + Sucursal)
                    foreach (var grupo in grupos)
                    {
                        // Verificar si ya existe un mensaje para este grupo
                        var mensajeExistente = await _context.MensaSucursals
                            .FirstOrDefaultAsync(m => m.OCompraSuc == grupo.OCompraSuc && m.Sucursal == grupo.Sucursal);

                        if (mensajeExistente != null)
                        {
                            // Actualizar mensaje existente
                            mensajeExistente.Mensaje = archivoResultado.Mensaje;
                            mensajeExistente.Usuario = numUsuar;
                            mensajeExistente.Status = 0;
                            mensajeExistente.LastUpdate = DateTime.Now;
                            mensajeExistente.NumUsuar = numUsuar;
                            _logger.LogInformation($"Actualizado MENSA_SUCURSAL: OCompraSuc={grupo.OCompraSuc}, Sucursal={grupo.Sucursal}");
                        }
                        else
                        {
                            // Insertar nuevo mensaje
                            var nuevoMensaje = new MensaSucursal
                            {
                                Sucursal = grupo.Sucursal,
                                OCompraSuc = grupo.OCompraSuc,
                                Usuario = numUsuar,
                                Status = 0,
                                Mensaje = archivoResultado.Mensaje,
                                CodiEmpr = 0,
                                AddRecord = DateTime.Now,
                                LastUpdate = DateTime.Now,
                                CodiEmprNet = 0,
                                NumUsuar = numUsuar
                            };

                            _context.MensaSucursals.Add(nuevoMensaje);
                            _logger.LogInformation($"Insertado MENSA_SUCURSAL: OCompraSuc={grupo.OCompraSuc}, Sucursal={grupo.Sucursal}");
                        }
                    }

                    await _context.SaveChangesAsync();
                }

                // ====================================================================
                // GUARDAR CAMBIOS Y CONFIRMAR TRANSACCIÓN
                // ====================================================================
                // SaveChanges() persiste todos los cambios en la base de datos
                // Commit() confirma la transacción, haciendo los cambios permanentes
                await transaction.CommitAsync();

                _logger.LogInformation($"Procesamiento completado: {archivoResultado.Registros.Count} registros");
                return true;
            }
            catch (Exception ex)
            {
                // ====================================================================
                // ERROR: DESHACER TRANSACCIÓN
                // ====================================================================
                // Si algo falla, hacemos rollback para dejar la BD en el estado original
                await transaction.RollbackAsync();
                _logger.LogError(ex, $"Error al procesar registros Ocompra_Sucursal");
                throw; // Re-lanzar la excepción para que el llamador pueda manejarla
            }
        }

        /// <summary>
        /// Procesa todos los archivos Ocompra_Sucursal*.txt en la ruta de intercambio
        /// Este método se llama desde la central para procesar archivos enviados por la sucursal
        ///
        /// Estrategia de procesamiento:
        /// 1. Busca archivos con el patrón Ocompra_Sucursal*.txt en la ruta de intercambio
        /// 2. Ordena por fecha de creación (más antiguos primero) para mantener el orden cronológico
        /// 3. Para cada archivo:
        ///    - Verifica si ya fue procesado exitosamente (tabla SyncProcesados)
        ///    - Si no fue procesado: lee el archivo y procesa sus registros
        ///    - Registra resultado en SyncProcesados (EXITOSO o ERROR)
        /// 4. Si un archivo falla, continúa con el siguiente (no falla todo el proceso)
        /// 5. Retorna un resumen con cantidad de archivos procesados, registros sincronizados y errores
        ///
        /// NOTA PARA DESARROLLADORES JUNIOR:
        /// - Este método es el punto de entrada principal para procesar archivos de sincronización
        /// - Usa la tabla SyncProcesados para evitar reprocesar archivos ya exitosos
        /// - Si un archivo falla, no detiene el procesamiento de los demás (resiliencia)
        /// - El orden cronológico es importante para mantener la secuencia de cambios
        /// - El resultado incluye estadísticas para mostrar al usuario en el frontend
        /// </summary>
        public async Task<SyncResultDTO> ProcesarArchivosOcompraSucursalAsync(string rutaIntercambio, int numUsuar)
        {
            // ====================================================================
            // INICIALIZAR RESULTADO
            // ====================================================================
            // SyncResultDTO contiene estadísticas del procesamiento:
            // - Exitoso: indica si el proceso general fue exitoso
            // - ArchivosProcessados: cantidad de archivos procesados exitosamente
            // - RegistrosSincronizados: total de registros insertados/actualizados
            // - Errores: cantidad de archivos que fallaron
            // - DetalleErrores: lista de mensajes de error para cada archivo fallido
            var resultado = new SyncResultDTO
            {
                Exitoso = true,
                ArchivosProcessados = 0,
                RegistrosSincronizados = 0,
                Errores = 0,
                DetalleErrores = new List<string>()
            };

            try
            {
                // ====================================================================
                // VALIDAR RUTA DE INTERCAMBIO
                // ====================================================================
                // La ruta debe estar configurada y ser accesible
                // Si no está configurada, retornamos error inmediatamente
                if (string.IsNullOrWhiteSpace(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = "Ruta de intercambio de pedidos de reposición no configurada";
                    return resultado;
                }

                // La ruta debe existir en el sistema de archivos
                if (!Directory.Exists(rutaIntercambio))
                {
                    resultado.Exitoso = false;
                    resultado.Mensaje = $"Ruta de intercambio no existe: {rutaIntercambio}";
                    return resultado;
                }

                // ====================================================================
                // BUSCAR ARCHIVOS PENDIENTES DE PROCESAMIENTO
                // ====================================================================
                // Buscamos archivos con el patrón Ocompra_Sucursal*.txt
                // El asterisco (*) permite cualquier fecha/hora en el nombre
                var archivos = Directory.GetFiles(rutaIntercambio, "Ocompra_Sucursal*.txt");

                // ====================================================================
                // ORDENAR ARCHIVOS POR FECHA DE CREACIÓN
                // ====================================================================
                // Es importante procesar los archivos en orden cronológico
                // para mantener la secuencia correcta de los cambios
                // Los más antiguos se procesan primero
                archivos = archivos.OrderBy(f => File.GetCreationTime(f)).ToArray();

                // ====================================================================
                // VERIFICAR SI HAY ARCHIVOS PARA PROCESAR
                // ====================================================================
                if (archivos.Length == 0)
                {
                    resultado.Mensaje = "No hay archivos Ocompra_Sucursal pendientes de procesamiento";
                    return resultado;
                }

                // ====================================================================
                // PROCESAR CADA ARCHIVO
                // ====================================================================
                foreach (var archivoPath in archivos)
                {
                    var nombreArchivo = Path.GetFileName(archivoPath);

                    try
                    {
                        // ====================================================================
                        // VERIFICAR SI EL ARCHIVO YA FUE PROCESADO
                        // ====================================================================
                        // Consultamos la tabla SyncProcesados para ver si este archivo
                        // ya fue procesado exitosamente anteriormente
                        // Esto evita reprocesar archivos y duplicar datos
                        var yaProcesado = await _context.SyncProcesados
                            .AnyAsync(s => s.Archivo == nombreArchivo && s.Estado == "EXITOSO");

                        if (yaProcesado)
                        {
                            _logger.LogInformation($"Archivo Ocompra_Sucursal ya procesado exitosamente: {nombreArchivo}");
                            continue; // Saltar este archivo y continuar con el siguiente
                        }

                        // ====================================================================
                        // LEER EL ARCHIVO
                        // ====================================================================
                        // Llamamos al método que lee el archivo TXT y lo convierte a DTO
                        // Este método NO modifica la base de datos, solo lee el archivo
                        var archivoResultado = await LeerArchivoOcompraSucursalAsync(archivoPath);

                        // ====================================================================
                        // PROCESAR LOS REGISTROS EN LA BASE DE DATOS
                        // ====================================================================
                        // Llamamos al método que hace INSERT/UPDATE en la tabla OC_SUCURSAL
                        // Este método SÍ modifica la base de datos (usa transacción)
                        var procesado = await ProcesarRegistrosOcompraSucursalAsync(archivoResultado, numUsuar);

                        if (procesado)
                        {
                            // ====================================================================
                            // GENERAR ARCHIVO ARCHIVO_GENERO_PEDIDO
                            // ====================================================================
                            // Según código VB6: después de procesar exitosamente el archivo Ocompra_Sucursal.txt,
                            // se genera un archivo ARCHIVO_GENERO_PEDIDO para notificar a la sucursal
                            // Este archivo se genera con los datos de OC_SUCURSAL procesados
                            try
                            {
                                await GenerarArchivoGeneroPedidoDesdeOcSucursalAsync(archivoResultado, numUsuar);
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError(ex, "Error al generar archivo ARCHIVO_GENERO_PEDIDO");
                                // No fallamos el proceso si no se puede generar el archivo
                            }

                            // ====================================================================
                            // REGISTRAR PROCESAMIENTO EXITOSO
                            // ====================================================================
                            // Guardamos un registro en SyncProcesados para auditoría
                            // Esto permite saber qué archivos ya fueron procesados exitosamente
                            var syncProcesado = new SyncProcesado
                            {
                                Archivo = nombreArchivo,
                                Tabla = "OC_SUCURSAL",
                                Operacion = "UPSERT",
                                FechaProcesamiento = DateTime.Now,
                                Estado = "EXITOSO",
                                Error = archivoResultado.Mensaje // Mensaje opcional del archivo
                            };

                            _context.SyncProcesados.Add(syncProcesado);
                            await _context.SaveChangesAsync();

                            // ====================================================================
                            // ELIMINAR ARCHIVO DESPUÉS DE PROCESAR EXITOSAMENTE
                            // ====================================================================
                            // Según código VB6: "EL ARCHIVO SI SE GUARDO DE MODO EXITOSO DEBE ELIMINARSE"
                            // Eliminamos el archivo para evitar reprocesarlo
                            try
                            {
                                File.Delete(archivoPath);
                                _logger.LogInformation($"Archivo Ocompra_Sucursal eliminado exitosamente: {nombreArchivo}");
                            }
                            catch (Exception ex)
                            {
                                _logger.LogWarning(ex, $"No se pudo eliminar el archivo Ocompra_Sucursal: {nombreArchivo}");
                                // No fallamos el proceso si no se puede eliminar el archivo
                            }

                            // Actualizar estadísticas del resultado
                            resultado.ArchivosProcessados++;
                            resultado.RegistrosSincronizados += archivoResultado.Registros.Count;

                            _logger.LogInformation($"Archivo Ocompra_Sucursal procesado exitosamente: {nombreArchivo}");
                        }
                        else
                        {
                            // Si el procesamiento falló, incrementamos contador de errores
                            resultado.Errores++;
                            resultado.DetalleErrores.Add($"Error al procesar archivo: {nombreArchivo}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // ====================================================================
                        // ERROR AL PROCESAR UN ARCHIVO
                        // ====================================================================
                        // Si un archivo falla, NO detenemos el procesamiento de los demás
                        // Registramos el error y continuamos con el siguiente archivo
                        resultado.Errores++;
                        resultado.DetalleErrores.Add($"Error al procesar archivo {Path.GetFileName(archivoPath)}: {ex.Message}");
                        _logger.LogError(ex, $"Error al procesar archivo Ocompra_Sucursal: {archivoPath}");

                        // Registrar el error en SyncProcesados
                        var syncProcesado = new SyncProcesado
                        {
                            Archivo = Path.GetFileName(archivoPath),
                            Tabla = "OC_SUCURSAL",
                            Operacion = "UPSERT",
                            FechaProcesamiento = DateTime.Now,
                            Estado = "ERROR",
                            Error = ex.Message
                        };

                        _context.SyncProcesados.Add(syncProcesado);
                        await _context.SaveChangesAsync();
                    }
                }

                // ====================================================================
                // RETORNAR RESULTADO FINAL
                // ====================================================================
                resultado.Mensaje = $"Procesamiento completado: {resultado.ArchivosProcessados} archivos, {resultado.RegistrosSincronizados} registros";
                return resultado;
            }
            catch (Exception ex)
            {
                // ====================================================================
                // ERROR GENERAL DEL PROCESO
                // ====================================================================
                // Si hay un error general (ej: error de conexión a BD), marcamos como fallido
                resultado.Exitoso = false;
                resultado.Mensaje = $"Error general al procesar archivos Ocompra_Sucursal: {ex.Message}";
                _logger.LogError(ex, "Error general al procesar archivos Ocompra_Sucursal");
                return resultado;
            }
        }

        /// <summary>
        /// Genera el archivo ARCHIVO_GENERO_PEDIDO después de generar un pedido
        /// Este archivo se envía a la sucursal para notificar que el pedido fue generado
        /// Basado en código VB6: GENERAR_ARCHIVO_GENERO_PEDIDO
        /// 
        /// Formato del nombre: ARCHIVO_GENERO_PEDIDO DDMMYYYY HHMMSS.TXT
        /// Ejemplo: ARCHIVO_GENERO_PEDIDO 06082026 153045.TXT
        /// </summary>
        public async Task<bool> GenerarArchivoGeneroPedidoAsync(int nroPedido, int nroClie, int nroCompra, string accion, int numUsuar)
        {
            try
            {
                // ====================================================================
                // OBTENER RUTA DE INTERCAMBIO
                // ====================================================================
                var config = ObtenerConfiguracion();
                var rutaIntercambio = config.RutaIntercambioPedidosReposicion;

                if (string.IsNullOrWhiteSpace(rutaIntercambio))
                {
                    _logger.LogWarning("Ruta de intercambio de pedidos de reposición no configurada. No se puede generar ARCHIVO_GENERO_PEDIDO");
                    return false;
                }

                if (!Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio no existe: {rutaIntercambio}. No se puede generar ARCHIVO_GENERO_PEDIDO");
                    return false;
                }

                // ====================================================================
                // GENERAR NOMBRE DEL ARCHIVO
                // ====================================================================
                // Según VB6: A$ = Date, A$ = REPLACAR(A$, "/", "")
                //            B$ = Time, B$ = REPLACAR(B$, ":", ""), B$ = REPLACAR(B$, ".", "")
                //            NARCHI2% = FILEOPEN%(Ruta$ + "\ARCHIVO_GENERO_PEDIDO " + A$ + " " + B$ + ".TXT", 2, 256)
                var fecha = DateTime.Now.ToString("ddMMyyyy");
                var hora = DateTime.Now.ToString("HHmmss");
                var nombreArchivo = $"ARCHIVO_GENERO_PEDIDO {fecha} {hora}.TXT";
                var rutaCompleta = Path.Combine(rutaIntercambio, nombreArchivo);

                // ====================================================================
                // OBTENER DETALLE DEL PEDIDO (PEDDETA)
                // ====================================================================
                // Según VB6: SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
                //            SQLX$ = SQLX$ + " WHERE NROPED = " & NROPED&
                //            SQLX$ = SQLX$ + " AND  NroClie = " & NCLIE
                var peddetas = await _context.Peddetas
                    .Where(p => p.NroPed == nroPedido && p.NroClie == nroClie && p.MarBorra != -1)
                    .ToListAsync();

                if (!peddetas.Any())
                {
                    _logger.LogWarning($"No se encontraron detalles de pedido para NROPED={nroPedido}, NROCLIE={nroClie}");
                    return false;
                }

                // ====================================================================
                // GENERAR ARCHIVO
                // ====================================================================
                using var writer = new StreamWriter(rutaCompleta, false, System.Text.Encoding.Default);

                // ====================================================================
                // ESCRIBIR ENCABEZADO
                // ====================================================================
                // Según VB6: TX$ = AJUSTI$("NCLIENTE", 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("N.PEDIDO", 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("N.COMPRA", 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("EMISION", 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("CI1%", 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("CODIGOEXT", 24): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("DESCRIPCION", 32): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("CANTIPEDIDA", 10): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("CANTIENTRE", 10): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("ACCION", 10): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$("USUARIO", 8): TX$ = TX$
                var encabezado = $"{Ajustar("NCLIENTE", 8)};" +
                                 $"{Ajustar("N.PEDIDO", 8)};" +
                                 $"{Ajustar("N.COMPRA", 8)};" +
                                 $"{Ajustar("EMISION", 8)};" +
                                 $"{Ajustar("CI1%", 8)};" +
                                 $"{Ajustar("CODIGOEXT", 24)};" +
                                 $"{Ajustar("DESCRIPCION", 32)};" +
                                 $"{Ajustar("CANTIPEDIDA", 10)};" +
                                 $"{Ajustar("CANTIENTRE", 10)};" +
                                 $"{Ajustar("ACCION", 10)};" +
                                 $"{Ajustar("USUARIO", 8)}";

                writer.WriteLine(encabezado);

                // ====================================================================
                // ESCRIBIR DETALLES
                // ====================================================================
                // Según VB6: TX$ = AJUSTI$(SAFETEXT$(NCLIE), 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(SAFETEXT$(NROPED&), 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(SAFETEXT$(NUOC&), 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(FECHATEX$(FEMI%), 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(XVALO(!CodiInt), 8): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(SAFETEXT$(!CODIEXT), 24): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(SAFETEXT$(!Descrip), 32): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(FORMATNUM$(XVALO(!CanPed), "F10.1"), 10): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(FORMATNUM$(XVALO(!CANTENT), "F10.1"), 10): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(SAFETEXT$(ACCION$), 10): TX$ = TX$ + ";"
                //            TX$ = TX$ + AJUSTI$(SAFETEXT$(NUSER1%), 8): TX$ = TX$
                // NUSER1% = USNBR% Mod 100
                var usuarioMod100 = numUsuar % 100;
                var fechaEmision = peddetas.FirstOrDefault()?.FechEmis ?? DateTime.Now;
                var fechaEmisionStr = fechaEmision.ToString("dd/MM/yy");

                foreach (var peddeta in peddetas)
                {
                    var linea = $"{Ajustar(nroClie.ToString(), 8)};" +
                                $"{Ajustar(nroPedido.ToString(), 8)};" +
                                $"{Ajustar(nroCompra.ToString(), 8)};" +
                                $"{Ajustar(fechaEmisionStr, 8)};" +
                                $"{Ajustar(peddeta.CodiInt?.ToString() ?? "", 8)};" +
                                $"{Ajustar(peddeta.CodiExt ?? "", 24)};" +
                                $"{Ajustar(peddeta.Descrip ?? "", 32)};" +
                                $"{Ajustar(FormatoDecimal(peddeta.CanPed ?? 0), 10)};" +
                                $"{Ajustar(FormatoDecimal(peddeta.CantEnt ?? 0), 10)};" +
                                $"{Ajustar(accion, 10)};" +
                                $"{Ajustar(usuarioMod100.ToString(), 8)}";

                    writer.WriteLine(linea);
                }

                writer.Close();

                _logger.LogInformation($"Archivo ARCHIVO_GENERO_PEDIDO generado exitosamente: {nombreArchivo}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar archivo ARCHIVO_GENERO_PEDIDO");
                return false;
            }
        }

        /// <summary>
        /// Ajusta un texto a un ancho específico (rellenando con espacios a la derecha)
        /// Equivalente a AJUSTI$ en VB6
        /// </summary>
        private string Ajustar(string texto, int ancho)
        {
            if (texto == null) texto = "";
            if (texto.Length > ancho) texto = texto.Substring(0, ancho);
            return texto.PadRight(ancho);
        }

        /// <summary>
        /// Formatea un decimal como string con formato F10.1
        /// Equivalente a FORMATNUM$(XVALO(!CanPed), "F10.1") en VB6
        /// </summary>
        private string FormatoDecimal(decimal valor)
        {
            return valor.ToString("F1").PadLeft(10);
        }

        /// <summary>
        /// Genera el archivo ARCHIVO_GENERO_PEDIDO directamente desde los datos de OC_SUCURSAL
        /// Este método se llama después de procesar exitosamente el archivo Ocompra_Sucursal.txt
        /// No genera pedidos en la base de datos, solo crea el archivo de notificación
        /// </summary>
        private async Task<bool> GenerarArchivoGeneroPedidoDesdeOcSucursalAsync(OcompraSucursalArchivoResultadoDTO archivoResultado, int numUsuar)
        {
            try
            {
                // ====================================================================
                // OBTENER RUTA DE INTERCAMBIO
                // ====================================================================
                var config = ObtenerConfiguracion();
                var rutaIntercambio = config.RutaIntercambioPedidosReposicion;

                if (string.IsNullOrWhiteSpace(rutaIntercambio))
                {
                    _logger.LogWarning("Ruta de intercambio de pedidos de reposición no configurada. No se puede generar ARCHIVO_GENERO_PEDIDO");
                    return false;
                }

                if (!Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio no existe: {rutaIntercambio}. No se puede generar ARCHIVO_GENERO_PEDIDO");
                    return false;
                }

                // ====================================================================
                // GENERAR NOMBRE DEL ARCHIVO
                // ====================================================================
                var fecha = DateTime.Now.ToString("ddMMyyyy");
                var hora = DateTime.Now.ToString("HHmmss");
                var nombreArchivo = $"ARCHIVO_GENERO_PEDIDO {fecha} {hora}.TXT";
                var rutaCompleta = Path.Combine(rutaIntercambio, nombreArchivo);

                // ====================================================================
                // GENERAR ARCHIVO
                // ====================================================================
                using var writer = new StreamWriter(rutaCompleta, false, System.Text.Encoding.Default);

                // ====================================================================
                // ESCRIBIR ENCABEZADO
                // ====================================================================
                var encabezado = $"{Ajustar("NCLIENTE", 8)};" +
                                 $"{Ajustar("N.PEDIDO", 8)};" +
                                 $"{Ajustar("N.COMPRA", 8)};" +
                                 $"{Ajustar("EMISION", 8)};" +
                                 $"{Ajustar("CI1%", 8)};" +
                                 $"{Ajustar("CODIGOEXT", 24)};" +
                                 $"{Ajustar("DESCRIPCION", 32)};" +
                                 $"{Ajustar("CANTIPEDIDA", 10)};" +
                                 $"{Ajustar("CANTIENTRE", 10)};" +
                                 $"{Ajustar("ACCION", 10)};" +
                                 $"{Ajustar("USUARIO", 8)}";

                writer.WriteLine(encabezado);

                // ====================================================================
                // ESCRIBIR DETALLES DESDE OC_SUCURSAL
                // ====================================================================
                var usuarioMod100 = numUsuar % 100;
                var fechaEmisionStr = DateTime.Now.ToString("dd/MM/yy");

                // Agrupar por Sucursal y OCompraSuc
                var grupos = archivoResultado.Registros
                    .Where(r => int.TryParse(r.Sucursal, out int s) && int.TryParse(r.OCompra, out int o))
                    .GroupBy(r => new { Sucursal = int.Parse(r.Sucursal), OCompraSuc = int.Parse(r.OCompra) });

                foreach (var grupo in grupos)
                {
                    var nroClie = grupo.Key.Sucursal;
                    var nroCompra = grupo.Key.OCompraSuc;

                    // Buscar si ya existe un pedido generado para esta OC
                    var nroPedido = await _context.O_Sucursals
                        .Where(o => o.Sucursal == nroClie && o.OCompraSuc == nroCompra && o.Marborra != -1)
                        .Select(o => o.NpedGeneradoEnAvellaneda)
                        .FirstOrDefaultAsync();

                    var nroPedidoStr = nroPedido?.ToString() ?? "PENDIENTE";

                    foreach (var registro in grupo)
                    {
                        // Obtener código externo y descripción desde masters
                        var codExt = "";
                        var descripcion = "";
                        if (int.TryParse(registro.Codigo, out int codInt) && codInt > 0)
                        {
                            var master = await _context.masters
                                .FirstOrDefaultAsync(m => m.Codint == codInt);
                            codExt = master?.Codigo ?? "";
                            descripcion = master?.Descripcion ?? "";
                        }

                        // Parsear cantidad
                        decimal cantidad = 0;
                        decimal.TryParse(registro.Cantidad, out cantidad);

                        var linea = $"{Ajustar(nroClie.ToString(), 8)};" +
                                    $"{Ajustar(nroPedidoStr, 8)};" +
                                    $"{Ajustar(nroCompra.ToString(), 8)};" +
                                    $"{Ajustar(fechaEmisionStr, 8)};" +
                                    $"{Ajustar(registro.Codigo ?? "", 8)};" +
                                    $"{Ajustar(codExt, 24)};" +
                                    $"{Ajustar(descripcion, 32)};" +
                                    $"{Ajustar(FormatoDecimal(cantidad), 10)};" +
                                    $"{Ajustar(FormatoDecimal(0), 10)};" +
                                    $"{Ajustar("GENEPEDI", 10)};" +
                                    $"{Ajustar(usuarioMod100.ToString(), 8)}";

                        writer.WriteLine(linea);
                    }
                }

                writer.Close();

                _logger.LogInformation($"Archivo ARCHIVO_GENERO_PEDIDO generado exitosamente: {nombreArchivo}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar archivo ARCHIVO_GENERO_PEDIDO desde OC_SUCURSAL");
                return false;
            }
        }
    }
}
