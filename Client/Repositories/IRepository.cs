namespace IMPLANPROD.Client.Repositories
{
    /// <summary>
    /// Interfaz que define las operaciones básicas para interactuar con la API
    /// </summary>
    public interface IRepository
    {
        Task<HttpResponseWrapper<T>> Get<T>(string url);
        Task<HttpResponseWrapper<byte[]>> GetBytes(string url);
        Task<HttpResponseWrapper<object>> Post<T>(string url, T model);
        Task<HttpResponseWrapper<TResponse>> Post<T, TResponse>(string url, T model);
        Task<HttpResponseWrapper<TResponse>> PostFormData<TResponse>(string url, MultipartFormDataContent content);
        Task<HttpResponseWrapper<object>> Delete(string url);
        Task<HttpResponseWrapper<object>> Put<T>(string url, T model);
        Task<HttpResponseWrapper<TResponse>> Put<T, TResponse>(string url, T model);

        // Métodos para manejo de autenticación
        Task SetTokenAsync(string token);
        Task<string> GetTokenAsync();
        Task RemoveTokenAsync();
        
        Task SetUserInfoAsync(object userInfo);
        Task<T?> GetUserInfoAsync<T>() where T : class;
        Task RemoveUserInfoAsync();
    }
}