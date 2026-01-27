using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
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

        public SyncService(DataContext context, IConfiguration configuration, ILogger<SyncService> logger)
        {
            _context = context;
            _configuration = configuration;
            _logger = logger;
        }

        /// <summary>
        /// Genera un archivo JSON de sincronización para un conjunto de Formulas
        /// </summary>
        public async Task<bool> GenerarArchivoSyncFormulasAsync(int codiConj)
        {
            try
            {
                var config = ObtenerConfiguracion();
                if (!config.GeneraScriptSync || !config.EsCentral)
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
                if (!config.GeneraScriptSync || !config.EsCentral)
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
                if (!config.GeneraScriptSync || !config.EsCentral)
                {
                    return false;
                }

                var rutaIntercambio = config.RutaIntercambio;
                if (string.IsNullOrEmpty(rutaIntercambio) || !Directory.Exists(rutaIntercambio))
                {
                    _logger.LogWarning($"Ruta de intercambio no válida: {rutaIntercambio}");
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
                    EsCentral = _configuration.GetValue<bool>("SyncConfig:EsCentral")
                };
            }
            
            return new SyncConfigDTO
            {
                GeneraScriptSync = configDb.GeneraScriptSync,
                RutaIntercambio = configDb.RutaIntercambio ?? "",
                EsCentral = configDb.EsCentral
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
                        EsCentral = config.EsCentral,
                        LastUpdate = DateTime.Now
                    };
                    _context.SyncConfigs.Add(configDb);
                }
                else
                {
                    configDb.GeneraScriptSync = config.GeneraScriptSync;
                    configDb.RutaIntercambio = config.RutaIntercambio ?? "";
                    configDb.EsCentral = config.EsCentral;
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
    }
}
