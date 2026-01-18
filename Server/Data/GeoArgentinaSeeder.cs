using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Data
{
    /// <summary>
    /// Clase para cargar datos geográficos de Argentina en la base de datos
    /// </summary>
    public class GeoArgentinaSeeder
    {
        private readonly DataContext _context;
        private readonly GeorefService _georefService;
        private readonly ILogger<GeoArgentinaSeeder> _logger;

        /// <summary>
        /// Constructor del seeder de datos geográficos de Argentina
        /// </summary>
        /// <param name="context">Contexto de la base de datos</param>
        /// <param name="georefService">Servicio para obtener datos de la API Georef</param>
        /// <param name="logger">Logger para registrar eventos</param>
        public GeoArgentinaSeeder(DataContext context, GeorefService georefService, ILogger<GeoArgentinaSeeder> logger)
        {
            _context = context;
            _georefService = georefService;
            _logger = logger;
        }

        /// <summary>
        /// Carga los datos geográficos de Argentina en la base de datos
        /// </summary>
        public async Task SeedGeoDataAsync()
        {
            try
            {
                // Verificar si ya existe el país Argentina
                var argentina = await _context.Countries
                    .FirstOrDefaultAsync(c => c.Name.ToLower() == "argentina");

                if (argentina == null)
                {
                    // Crear el país Argentina si no existe
                    argentina = new Country
                    {
                        Name = "Argentina",
                        status = 0
                    };

                    _context.Countries.Add(argentina);
                    await _context.SaveChangesAsync();
                    _logger.LogInformation("País Argentina creado con ID: {Id}", argentina.Id);
                }

                // Obtener todas las provincias de Argentina
                var provincias = await _georefService.GetProvinciasAsync();
                _logger.LogInformation("Se obtuvieron {Count} provincias de Argentina", provincias.Count);

                foreach (var provincia in provincias)
                {
                    // Verificar si la provincia ya existe
                    var existingState = await _context.States
                        .FirstOrDefaultAsync(s => s.CountryId == argentina.Id &&
                                                 s.Name.ToLower() == provincia.Nombre.ToLower());

                    if (existingState == null)
                    {
                        // Crear la provincia si no existe
                        var state = new State
                        {
                            Name = provincia.Nombre,
                            CountryId = argentina.Id,
                            status = 0
                           
                        };

                        _context.States.Add(state);
                        await _context.SaveChangesAsync();
                        _logger.LogInformation("Provincia {Nombre} creada con ID: {Id}", provincia.Nombre, state.Id);

                        // --- Eliminada la carga de localidades por decisión del usuario ---
                        //_logger.LogInformation("Intentando cargar localidades para la provincia {Provincia}", provincia.Nombre);
                        //var localidades = await _georefService.GetLocalidadesByProvinciaAsync(state.ExternalId);
                        //_logger.LogInformation("Se obtuvieron {Count} localidades para la provincia {Provincia}", localidades.Count, provincia.Nombre);
                        //var localidadesAgrupadas = localidades.GroupBy(l => l.Nombre.Trim()).Select(g => g.First()).ToList();
                        //foreach (var localidad in localidadesAgrupadas)
                        //{
                        //    var existingCity = await _context.Cities.FirstOrDefaultAsync(c => c.StateId == state.Id && c.Name.ToLower() == localidad.Nombre.ToLower());
                        //    if (existingCity == null)
                        //    {
                        //        var city = new City
                        //        {
                        //            Name = localidad.Nombre,
                        //            StateId = state.Id,
                        //            status = 0
                        //        };
                        //        _context.Cities.Add(city);
                        //    }
                        //}
                        //await _context.SaveChangesAsync();
                        //_logger.LogInformation("Localidades de {Provincia} guardadas correctamente", provincia.Nombre);
                    }
                    else
                    {
                        _logger.LogInformation("La provincia {Nombre} ya existe con ID: {Id}",
                            provincia.Nombre, existingState.Id);
                    }
                }

                _logger.LogInformation("Carga de datos geográficos de Argentina completada con éxito");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al cargar datos geográficos de Argentina");
                throw;
            }
        }
    }
}
