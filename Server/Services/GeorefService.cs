using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Threading.Tasks;
using IMPLANPROD.Shared.Entities;
using Microsoft.Extensions.Logging;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para consumir la API de Georef Argentina y obtener datos geográficos
    /// </summary>
    public class GeorefService
    {
        private readonly HttpClient _httpClient;
        private readonly ILogger<GeorefService> _logger;
        private const string BaseUrl = "https://apis.datos.gob.ar/georef/api/";

        /// <summary>
        /// Constructor del servicio Georef
        /// </summary>
        /// <param name="httpClient">Cliente HTTP para realizar peticiones</param>
        /// <param name="logger">Logger para registrar eventos</param>
        public GeorefService(HttpClient httpClient, ILogger<GeorefService> logger)
        {
            _httpClient = httpClient;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todas las provincias de Argentina
        /// </summary>
        /// <returns>Lista de provincias</returns>
        public async Task<List<ProvinciasResponse.Provincia>> GetProvinciasAsync()
        {
            try
            {
                var response = await _httpClient.GetFromJsonAsync<ProvinciasResponse>($"{BaseUrl}provincias?campos=id,nombre");
                return response?.Provincias ?? new List<ProvinciasResponse.Provincia>();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener provincias de Argentina");
                return new List<ProvinciasResponse.Provincia>();
            }
        }

        /// <summary>
        /// Obtiene todas las localidades de una provincia específica
        /// </summary>
        /// <param name="provinciaId">ID de la provincia</param>
        /// <returns>Lista de localidades</returns>
        public async Task<List<LocalidadesResponse.Localidad>> GetLocalidadesByProvinciaAsync(string provinciaId)
        {
            try
            {
                _logger.LogInformation($"Obteniendo localidades para la provincia ID: {provinciaId}");

                int offset = 0;
                int limit = 100; 
                var allLocalidades = new List<LocalidadesResponse.Localidad>();
                bool hasMore = true;
                int intentos = 0;
                const int maxIntentos = 3;

                while (hasMore && intentos < maxIntentos)
                {
                    string url = $"{BaseUrl}localidades?provincia={provinciaId}&campos=id,nombre,departamento.nombre,municipio.nombre&max={limit}&offset={offset}";
                    _logger.LogInformation($"Consultando URL: {url}");

                    try
                    {
                        var response = await _httpClient.GetAsync(url);
                        if (!response.IsSuccessStatusCode)
                        {
                            _logger.LogWarning($"Respuesta HTTP no exitosa: {(int)response.StatusCode} - {response.ReasonPhrase}");
                            break;
                        }

                        var jsonContent = await response.Content.ReadAsStringAsync();
                        _logger.LogInformation($"Respuesta recibida (primeros 200 chars): {jsonContent.Substring(0, Math.Min(jsonContent.Length, 200))}...");

                        var options = new JsonSerializerOptions
                        {
                            PropertyNameCaseInsensitive = true,
                            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                        };

                        var localidadesResponse = JsonSerializer.Deserialize<LocalidadesResponse>(jsonContent, options);

                        if (localidadesResponse?.Localidades == null)
                        {
                            _logger.LogWarning($"La respuesta de localidades es nula para la provincia {provinciaId}");
                            hasMore = false;
                        }
                        else if (localidadesResponse.Localidades.Count == 0)
                        {
                            _logger.LogWarning($"No se encontraron localidades para la provincia {provinciaId} en offset {offset}");
                            hasMore = false;
                        }
                        else
                        {
                            _logger.LogInformation($"Se obtuvieron {localidadesResponse.Localidades.Count} localidades para la provincia {provinciaId} en offset {offset}");

                            foreach (var localidad in localidadesResponse.Localidades)
                            {
                                if (string.IsNullOrEmpty(localidad.Nombre))
                                {
                                    localidad.Nombre = "Sin nombre";
                                }
                                else
                                {
                                    localidad.Nombre = localidad.Nombre.Trim();
                                }
                            }

                            allLocalidades.AddRange(localidadesResponse.Localidades);
                            offset += limit;

                            if (localidadesResponse.Localidades.Count < limit)
                            {
                                _logger.LogInformation($"Se alcanzó el final de los datos para la provincia {provinciaId}");
                                hasMore = false;
                            }
                        }

                        intentos = 0;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"Error al obtener localidades de la provincia {provinciaId} en offset {offset}");
                        intentos++;
                        if (intentos >= maxIntentos)
                        {
                            _logger.LogWarning($"Se alcanzó el máximo de intentos ({maxIntentos}) para la provincia {provinciaId}");
                        }
                    }
                }

                _logger.LogInformation($"Total localidades encontradas para provincia {provinciaId}: {allLocalidades.Count}");
                return allLocalidades;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error general al obtener localidades de la provincia {provinciaId}");
                return new List<LocalidadesResponse.Localidad>();
            }
        }

        /// <summary>
        /// Obtiene los departamentos de una provincia específica
        /// </summary>
        /// <param name="provinciaId">ID de la provincia</param>
        /// <returns>Lista de departamentos</returns>
        public async Task<List<DepartamentosResponse.Departamento>> GetDepartamentosByProvinciaAsync(string provinciaId)
        {
            try
            {
                var response = await _httpClient.GetFromJsonAsync<DepartamentosResponse>(
                    $"{BaseUrl}departamentos?provincia={provinciaId}&campos=id,nombre&max=500");
                
                return response?.Departamentos ?? new List<DepartamentosResponse.Departamento>();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener departamentos de la provincia {provinciaId}");
                return new List<DepartamentosResponse.Departamento>();
            }
        }

        /// <summary>
        /// Clases para deserializar las respuestas de la API
        /// </summary>
        public class ProvinciasResponse
        {
            public List<Provincia> Provincias { get; set; } = new List<Provincia>();

            public class Provincia
            {
                public string Id { get; set; } = string.Empty;
                public string Nombre { get; set; } = string.Empty;
            }
        }

        public class DepartamentosResponse
        {
            public List<Departamento> Departamentos { get; set; } = new List<Departamento>();

            public class Departamento
            {
                public string Id { get; set; } = string.Empty;
                public string Nombre { get; set; } = string.Empty;
            }
        }

        public class LocalidadesResponse
        {
            public List<Localidad> Localidades { get; set; } = new List<Localidad>();

            public class Localidad
            {
                public string Id { get; set; } = string.Empty;
                public string Nombre { get; set; } = string.Empty;
                public Departamento Departamento { get; set; } = new Departamento();
                public Municipio Municipio { get; set; } = new Municipio();
            }

            public class Departamento
            {
                public string Nombre { get; set; } = string.Empty;
            }

            public class Municipio
            {
                public string Nombre { get; set; } = string.Empty;
            }
        }
    }
}
