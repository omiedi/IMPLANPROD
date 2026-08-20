using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para gestión de derechos de usuarios
    /// Implementa la lógica del código VB6 para manejo de permisos
    /// </summary>
    public class DerechosUsuarioService : IDerechosUsuarioService
    {
        private readonly DataContext _context;
        private readonly ILogger<DerechosUsuarioService> _logger;

        public DerechosUsuarioService(DataContext context, ILogger<DerechosUsuarioService> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todos los derechos de un usuario específico
        /// Replica la consulta SQL del código VB6:
        /// SELECT I.COD_APLI, D.NUM_USER, I.DES_APLI, D.DER_READ, D.DER_WRITE
        /// FROM INDI_APLI I LEFT JOIN DER_ACCE D ON D.COD_APLI = I.COD_APLI
        /// WHERE D.NUM_USER = @numUser OR D.NUM_USER IS NULL
        /// ORDER BY I.COD_APLI
        /// </summary>
        public async Task<List<DerechoUsuarioDTO>> ObtenerDerechosUsuarioAsync(short numUser)
        {
            try
            {
                _logger.LogInformation($"Obteniendo derechos para usuario {numUser}");

                // LEFT JOIN entre INDI_APLI y DER_ACCE
                var query = from aplicacion in _context.IndiAplis
                            join derecho in _context.DerAcces.Where(d => d.NumUser == numUser)
                            on aplicacion.CodApli equals derecho.CodApli into derechosGroup
                            from derecho in derechosGroup.DefaultIfEmpty()
                            orderby aplicacion.CodApli
                            select new DerechoUsuarioDTO
                            {
                                CodApli = aplicacion.CodApli,
                                DesApli = aplicacion.DesApli ?? string.Empty,
                                NumUser = derecho != null ? derecho.NumUser : (short?)null,
                                DerRead = derecho != null ? derecho.DerRead : (short)0,
                                DerWrite = derecho != null ? derecho.DerWrite : (short)0,
                                IdDerAcce = derecho != null ? derecho.Id : (int?)null
                            };

                var resultado = await query.ToListAsync();

                _logger.LogInformation($"Se obtuvieron {resultado.Count} aplicaciones para el usuario {numUser}");

                return resultado;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener derechos del usuario {numUser}");
                throw;
            }
        }

        /// <summary>
        /// Guarda o actualiza los derechos de un usuario
        /// </summary>
        public async Task<bool> GuardarDerechosUsuarioAsync(short numUser, List<DerechoUsuarioDTO> derechos)
        {
            try
            {
                _logger.LogInformation($"Guardando derechos para usuario {numUser}");

                foreach (var derecho in derechos)
                {
                    // Buscar si ya existe un registro para este usuario y aplicación
                    var derechoExistente = await _context.DerAcces
                        .FirstOrDefaultAsync(d => d.NumUser == numUser && d.CodApli == derecho.CodApli);

                    if (derechoExistente != null)
                    {
                        // Actualizar registro existente
                        derechoExistente.DerRead = derecho.DerRead;
                        derechoExistente.DerWrite = derecho.DerWrite;
                        _context.DerAcces.Update(derechoExistente);
                    }
                    else
                    {
                        // Crear nuevo registro solo si tiene algún derecho asignado
                        if (derecho.DerRead == 1 || derecho.DerWrite == 1)
                        {
                            var nuevoDerecho = new DerAcce
                            {
                                NumUser = numUser,
                                CodApli = derecho.CodApli,
                                DerRead = derecho.DerRead,
                                DerWrite = derecho.DerWrite
                            };
                            await _context.DerAcces.AddAsync(nuevoDerecho);
                        }
                    }
                }

                await _context.SaveChangesAsync();
                _logger.LogInformation($"Derechos guardados correctamente para usuario {numUser}");

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al guardar derechos del usuario {numUser}");
                return false;
            }
        }

        /// <summary>
        /// Verifica si un usuario tiene derecho de lectura sobre una aplicación
        /// </summary>
        public async Task<bool> TieneDerechoLecturaAsync(short numUser, string codApli)
        {
            try
            {
                var derecho = await _context.DerAcces
                    .FirstOrDefaultAsync(d => d.NumUser == numUser && d.CodApli == codApli);

                return derecho != null && derecho.DerRead == 1;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al verificar derecho de lectura para usuario {numUser}, aplicación {codApli}");
                return false;
            }
        }

        /// <summary>
        /// Verifica si un usuario tiene derecho de escritura sobre una aplicación
        /// </summary>
        public async Task<bool> TieneDerechoEscrituraAsync(short numUser, string codApli)
        {
            try
            {
                var derecho = await _context.DerAcces
                    .FirstOrDefaultAsync(d => d.NumUser == numUser && d.CodApli == codApli);

                return derecho != null && derecho.DerWrite == 1;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al verificar derecho de escritura para usuario {numUser}, aplicación {codApli}");
                return false;
            }
        }

        /// <summary>
        /// Verifica los derechos de un usuario sobre una aplicación (replica funcionalidad DERACCE del VB6)
        /// Retorna:
        /// 0 = Sin acceso (crea el registro en INDI_APLI si no existe)
        /// 1 = Solo lectura
        /// 2 = Lectura y escritura (o solo escritura)
        /// 
        /// Lógica VB6 replicada:
        /// - Si el usuario tiene perfil ADMINISTRADOR, retorna directamente 2 (acceso completo)
        /// - Si no existe el registro en INDI_APLI, lo crea automáticamente
        /// - Verifica los derechos en DER_ACCE para el usuario
        /// - Retorna el nivel de acceso según los permisos
        /// </summary>
        public async Task<int> VerificarDerechoAsync(short numUser, string codApli, string desApli)
        {
            try
            {
                // PASO 1: Verificar si el usuario tiene perfil ADMINISTRADOR
                // Buscar el usuario por IdFlexoft (que es numUser) y obtener su perfil
                var usuario = await _context.Usuarios
                    .Include(u => u.Perfil)
                    .FirstOrDefaultAsync(u => u.IdFlexoft == numUser);

                if (usuario != null && usuario.Perfil != null)
                {
                    // Verificar si el perfil es ADMINISTRADOR (case-insensitive)
                    if (usuario.Perfil.Nombre.Equals("ADMINISTRADOR", StringComparison.OrdinalIgnoreCase))
                    {
                        _logger.LogInformation($"Usuario {numUser} tiene perfil ADMINISTRADOR - Retornando acceso completo (2) para '{codApli}'");
                        return 2; // Acceso completo para administradores
                    }
                }

                // Truncar a los límites de la BD (como en VB6: Left$(MDY$, 32) y Left$(APM$, 128))
                var codApliTruncado = codApli.Length > 32 ? codApli.Substring(0, 32) : codApli;
                var desApliTruncado = desApli.Length > 128 ? desApli.Substring(0, 128) : desApli;

                _logger.LogInformation($"Verificando derecho para usuario {numUser}, aplicación '{codApliTruncado}'");

                // 1. Verificar si existe el registro en INDI_APLI
                var aplicacionExiste = await _context.IndiAplis
                    .AnyAsync(i => i.CodApli == codApliTruncado && i.DesApli == desApliTruncado);

                // 2. Si no existe, crear el registro en INDI_APLI (replica lógica VB6)
                if (!aplicacionExiste)
                {
                    _logger.LogInformation($"Aplicación '{codApliTruncado}' no existe en INDI_APLI, creando registro...");

                    // Verificar si existe con el mismo código pero diferente descripción
                    var aplicacionPorCodigo = await _context.IndiAplis
                        .FirstOrDefaultAsync(i => i.CodApli == codApliTruncado);

                    if (aplicacionPorCodigo == null)
                    {
                        // No existe, crear nuevo registro
                        var nuevaAplicacion = new IndiApli
                        {
                            CodApli = codApliTruncado,
                            DesApli = desApliTruncado
                        };

                        await _context.IndiAplis.AddAsync(nuevaAplicacion);
                        await _context.SaveChangesAsync();

                        _logger.LogInformation($"Registro creado en INDI_APLI: '{codApliTruncado}' - '{desApliTruncado}'");
                    }
                    else
                    {
                        // Existe con el mismo código, actualizar descripción
                        aplicacionPorCodigo.DesApli = desApliTruncado;
                        _context.IndiAplis.Update(aplicacionPorCodigo);
                        await _context.SaveChangesAsync();

                        _logger.LogInformation($"Registro actualizado en INDI_APLI: '{codApliTruncado}' - '{desApliTruncado}'");
                    }
                }

                // 3. Verificar los derechos del usuario en DER_ACCE
                var derecho = await _context.DerAcces
                    .FirstOrDefaultAsync(d => d.NumUser == numUser && d.CodApli == codApliTruncado);

                // 4. Si no tiene registro en DER_ACCE, crearlo con permisos en 0 y retornar 0 (sin acceso)
                if (derecho == null)
                {
                    _logger.LogInformation($"Usuario {numUser} no tiene derechos asignados para '{codApliTruncado}' - Creando DER_ACCE con permisos en 0");

                    var nuevoDerecho = new DerAcce
                    {
                        NumUser = numUser,
                        CodApli = codApliTruncado,
                        DerRead = 0,
                        DerWrite = 0,
                        CodiEmprNet = 0,
                        NumUsuar = numUser,
                        AddRecord = DateTime.Now,
                        LastUpdate = DateTime.Now,
                    };

                    await _context.DerAcces.AddAsync(nuevoDerecho);
                    await _context.SaveChangesAsync();

                    _logger.LogInformation($"DER_ACCE creado para usuario {numUser}, aplicación '{codApliTruncado}' con DER_READ=0 y DER_WRITE=0");
                    return 0;
                }

                // 5. Determinar nivel de acceso según los permisos
                // Si tiene escritura (con o sin lectura) = 2
                if (derecho.DerWrite == 1)
                {
                    _logger.LogInformation($"Usuario {numUser} tiene derechos de ESCRITURA para '{codApliTruncado}' - Retornando 2");
                    return 2;
                }

                // Si solo tiene lectura = 1
                if (derecho.DerRead == 1)
                {
                    _logger.LogInformation($"Usuario {numUser} tiene derechos de SOLO LECTURA para '{codApliTruncado}' - Retornando 1");
                    return 1;
                }

                // Si no tiene ningún derecho activado = 0
                _logger.LogInformation($"Usuario {numUser} no tiene derechos activos para '{codApliTruncado}' - Retornando 0 (sin acceso)");
                return 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al verificar derecho para usuario {numUser}, aplicación {codApli}");
                // En caso de error, retornar 0 (sin acceso) por seguridad
                return 0;
            }
        }
    }
}
