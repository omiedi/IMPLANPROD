using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.JSInterop;

namespace IMPLANPROD.Client.Repositories
{
    /// <summary>
    /// Implementación del repositorio para interactuar con la API
    /// </summary>
    public class Repository : IRepository
    {
        private readonly HttpClient _httpClient;
        private readonly IJSRuntime _jsRuntime;

        public Repository(HttpClient httpClient, IJSRuntime jsRuntime)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _jsRuntime = jsRuntime ?? throw new ArgumentNullException(nameof(jsRuntime));
        }

        private async Task<string> GetBearerTokenAsync()
        {
            return await _jsRuntime.InvokeAsync<string>("sessionStorage.getItem", "token") ?? string.Empty;
        }

        /// <summary>
        /// Verifica si la respuesta incluye un token renovado por el middleware del servidor
        /// (TokenRenewalMiddleware). Si existe el header "X-Renewed-Token", lo guarda en
        /// sessionStorage y actualiza el header de autorización del HttpClient.
        /// Esto implementa la renovación de token con actividad (sliding expiration):
        /// mientras el usuario haga llamadas al backend, el token se renueva automáticamente
        /// y la sesión no expira.
        /// </summary>
        private async Task TryCaptureRenewedTokenAsync(HttpResponseMessage response)
        {
            try
            {
                if (response.Headers.TryGetValues("X-Renewed-Token", out var renewedTokens))
                {
                    var newToken = renewedTokens.FirstOrDefault();
                    if (!string.IsNullOrEmpty(newToken))
                    {
                        await _jsRuntime.InvokeVoidAsync("sessionStorage.setItem", "token", newToken);
                        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", newToken);
                        Console.WriteLine("[Repository] Token renovado automáticamente por el servidor (sliding expiration)");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Repository] Error al capturar token renovado: {ex.Message}");
            }
        }

        public async Task SetTokenAsync(string token)
        {
            await _jsRuntime.InvokeVoidAsync("sessionStorage.setItem", "token", token);
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        public async Task<string> GetTokenAsync()
        {
            return await _jsRuntime.InvokeAsync<string>("sessionStorage.getItem", "token") ?? string.Empty;
        }

        public async Task RemoveTokenAsync()
        {
            await _jsRuntime.InvokeVoidAsync("sessionStorage.removeItem", "token");
            _httpClient.DefaultRequestHeaders.Authorization = null;
        }

        public async Task SetUserInfoAsync(object userInfo)
        {
            var json = JsonSerializer.Serialize(userInfo);
            await _jsRuntime.InvokeVoidAsync("sessionStorage.setItem", "userInfo", json);
        }

        public async Task<T?> GetUserInfoAsync<T>() where T : class
        {
            try
            {
                var json = await _jsRuntime.InvokeAsync<string>("sessionStorage.getItem", "userInfo");
                
                if (string.IsNullOrEmpty(json))
                {
                    Console.WriteLine("⚠️ GetUserInfoAsync: sessionStorage.userInfo está vacío o null");
                    return null;
                }
                
                Console.WriteLine($"📦 GetUserInfoAsync: JSON recuperado (primeros 100 chars): {json.Substring(0, Math.Min(100, json.Length))}...");
                
                var result = JsonSerializer.Deserialize<T>(json);
                Console.WriteLine($"✅ GetUserInfoAsync: Deserialización exitosa - Tipo: {typeof(T).Name}");
                
                return result;
            }
            catch (JSException jsEx)
            {
                Console.WriteLine($"❌ GetUserInfoAsync - Error JavaScript: {jsEx.Message}");
                throw new Exception($"Error al acceder a sessionStorage: {jsEx.Message}", jsEx);
            }
            catch (JsonException jsonEx)
            {
                Console.WriteLine($"❌ GetUserInfoAsync - Error de deserialización JSON: {jsonEx.Message}");
                throw new Exception($"Error al deserializar UserInfo: {jsonEx.Message}", jsonEx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ GetUserInfoAsync - Error inesperado: {ex.GetType().Name} - {ex.Message}");
                throw;
            }
        }

        public async Task RemoveUserInfoAsync()
        {
            await _jsRuntime.InvokeVoidAsync("sessionStorage.removeItem", "userInfo");
        }

        public async Task<HttpResponseWrapper<T>> Get<T>(string url)
        {
            var token = await GetBearerTokenAsync();
            if (!string.IsNullOrEmpty(token))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            var responseHTTP = await _httpClient.GetAsync(url);
            await TryCaptureRenewedTokenAsync(responseHTTP);
            if (responseHTTP.IsSuccessStatusCode)
            {
                var response = await UnserializeAnswer<T>(responseHTTP);
                return new HttpResponseWrapper<T>(response, false, responseHTTP);
            }

            return new HttpResponseWrapper<T>(default, true, responseHTTP);
        }

        public async Task<HttpResponseWrapper<byte[]>> GetBytes(string url)
        {
            var token = await GetBearerTokenAsync();
            if (!string.IsNullOrEmpty(token))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            var responseHTTP = await _httpClient.GetAsync(url);
            await TryCaptureRenewedTokenAsync(responseHTTP);
            if (responseHTTP.IsSuccessStatusCode)
            {
                var bytes = await responseHTTP.Content.ReadAsByteArrayAsync();
                return new HttpResponseWrapper<byte[]>(bytes, false, responseHTTP);
            }

            return new HttpResponseWrapper<byte[]>(default, true, responseHTTP);
        }

        public async Task<HttpResponseWrapper<object>> Post<T>(string url, T model)
        {
            var token = await GetBearerTokenAsync();
            if (!string.IsNullOrEmpty(token))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            var messageJSON = JsonSerializer.Serialize(model);
            var messageContent = new StringContent(messageJSON, Encoding.UTF8, "application/json");
            var responseHttp = await _httpClient.PostAsync(url, messageContent);
            await TryCaptureRenewedTokenAsync(responseHttp);
            return new HttpResponseWrapper<object>(null, !responseHttp.IsSuccessStatusCode, responseHttp);
        }

        public async Task<HttpResponseWrapper<TResponse>> Post<T, TResponse>(string url, T model)
        {
            var token = await GetBearerTokenAsync();
            if (!string.IsNullOrEmpty(token))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            var messageJSON = JsonSerializer.Serialize(model);
            var messageContent = new StringContent(messageJSON, Encoding.UTF8, "application/json");
            var responseHttp = await _httpClient.PostAsync(url, messageContent);
            await TryCaptureRenewedTokenAsync(responseHttp);
            if (responseHttp.IsSuccessStatusCode)
            {
                var response = await UnserializeAnswer<TResponse>(responseHttp);
                return new HttpResponseWrapper<TResponse>(response, false, responseHttp);
            }
            return new HttpResponseWrapper<TResponse>(default, true, responseHttp);
        }

        public async Task<HttpResponseWrapper<TResponse>> PostFormData<TResponse>(string url, MultipartFormDataContent content)
        {
            var token = await GetBearerTokenAsync();
            if (!string.IsNullOrEmpty(token))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            var responseHttp = await _httpClient.PostAsync(url, content);
            await TryCaptureRenewedTokenAsync(responseHttp);
            if (responseHttp.IsSuccessStatusCode)
            {
                var response = await UnserializeAnswer<TResponse>(responseHttp);
                return new HttpResponseWrapper<TResponse>(response, false, responseHttp);
            }
            return new HttpResponseWrapper<TResponse>(default, true, responseHttp);
        }

        public async Task<HttpResponseWrapper<object>> Delete(string url)
        {
            var token = await GetBearerTokenAsync();
            if (!string.IsNullOrEmpty(token))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            var responseHTTP = await _httpClient.DeleteAsync(url);
            await TryCaptureRenewedTokenAsync(responseHTTP);
            return new HttpResponseWrapper<object>(null, !responseHTTP.IsSuccessStatusCode, responseHTTP);
        }

        public async Task<HttpResponseWrapper<object>> Put<T>(string url, T model)
        {
            var token = await GetBearerTokenAsync();
            if (!string.IsNullOrEmpty(token))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            var messageJSON = JsonSerializer.Serialize(model);
            var messageContent = new StringContent(messageJSON, Encoding.UTF8, "application/json");
            var responseHttp = await _httpClient.PutAsync(url, messageContent);
            await TryCaptureRenewedTokenAsync(responseHttp);
            return new HttpResponseWrapper<object>(null, !responseHttp.IsSuccessStatusCode, responseHttp);
        }

        public async Task<HttpResponseWrapper<TResponse>> Put<T, TResponse>(string url, T model)
        {
            var token = await GetBearerTokenAsync();
            if (!string.IsNullOrEmpty(token))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }

            var messageJSON = JsonSerializer.Serialize(model);
            var messageContent = new StringContent(messageJSON, Encoding.UTF8, "application/json");
            var responseHttp = await _httpClient.PutAsync(url, messageContent);
            await TryCaptureRenewedTokenAsync(responseHttp);
            if (responseHttp.IsSuccessStatusCode)
            {
                var response = await UnserializeAnswer<TResponse>(responseHttp);
                return new HttpResponseWrapper<TResponse>(response, false, responseHttp);
            }
            return new HttpResponseWrapper<TResponse>(default, true, responseHttp);
        }

        private async Task<T> UnserializeAnswer<T>(HttpResponseMessage httpResponse)
        {
            var responseString = await httpResponse.Content.ReadAsStringAsync();

            // Si el caller pidió string, devolvemos el texto tal cual (caso típico: endpoint que responde texto plano).
            // Esto evita excepciones cuando el backend devuelve 3813 (sin comillas JSON).
            if (typeof(T) == typeof(string))
            {
                object result = responseString;

                // Si vino como string JSON ("3813"), intentamos deserializar para quitar comillas escapadas.
                if (!string.IsNullOrWhiteSpace(responseString) && responseString.StartsWith("\"") && responseString.EndsWith("\""))
                {
                    try
                    {
                        result = JsonSerializer.Deserialize<string>(responseString) ?? string.Empty;
                    }
                    catch
                    {
                        result = responseString;
                    }
                }

                return (T)result;
            }

            return JsonSerializer.Deserialize<T>(responseString, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            })!;
        }
    }
}