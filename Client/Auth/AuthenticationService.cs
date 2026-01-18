using System.Security.Claims;
using System.Text.Json;
using IMPLANPROD.Client.Repositories;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Components.Authorization;

namespace IMPLANPROD.Client.Auth
{
    public class AuthenticationService : AuthenticationStateProvider
    {
        private readonly IRepository _repository;

        public AuthenticationService(IRepository repository)
        {
            _repository = repository;
        }

        public override async Task<AuthenticationState> GetAuthenticationStateAsync()
        {
            var token = await _repository.GetTokenAsync();

            if (string.IsNullOrEmpty(token))
            {
                return new AuthenticationState(new ClaimsPrincipal(new ClaimsIdentity()));
            }

            var userInfo = await _repository.GetUserInfoAsync<UserInfo>();
            if (userInfo == null)
            {
                return new AuthenticationState(new ClaimsPrincipal(new ClaimsIdentity()));
            }

            var claims = new List<Claim>
            {
                new Claim(ClaimTypes.Name, userInfo.NombreUsuario),
                new Claim(ClaimTypes.Email, userInfo.Email),
                new Claim(ClaimTypes.Role, userInfo.PerfilNombre),
                new Claim("EmpresaId", userInfo.EmpresaId.ToString()),
                new Claim("PerfilId", userInfo.PerfilId.ToString())
            };

            var identity = new ClaimsIdentity(claims, "JwtAuth");
            return new AuthenticationState(new ClaimsPrincipal(identity));
        }

        public async Task LoginAsync(string token, UserInfo userInfo)
        {
            await _repository.SetTokenAsync(token);
            await _repository.SetUserInfoAsync(userInfo);

            var claims = new List<Claim>
            {
                new Claim(ClaimTypes.Name, userInfo.NombreUsuario),
                new Claim(ClaimTypes.Email, userInfo.Email),
                new Claim(ClaimTypes.Role, userInfo.PerfilNombre),
                new Claim("EmpresaId", userInfo.EmpresaId.ToString()),
                new Claim("PerfilId", userInfo.PerfilId.ToString())
            };

            var identity = new ClaimsIdentity(claims, "JwtAuth");
            var user = new ClaimsPrincipal(identity);
            NotifyAuthenticationStateChanged(Task.FromResult(new AuthenticationState(user)));
        }

        public async Task LogoutAsync()
        {
            await _repository.RemoveTokenAsync();
            await _repository.RemoveUserInfoAsync();
            NotifyAuthenticationStateChanged(Task.FromResult(new AuthenticationState(new ClaimsPrincipal(new ClaimsIdentity()))));
        }
    }

    public class UserInfo
    {
        public int Id { get; set; }
        public string Email { get; set; } = null!;
        public string NombreUsuario { get; set; } = null!;
        public int? PerfilId { get; set; }
        public string PerfilNombre { get; set; } = null!;
        public int EmpresaId { get; set; }
        public string EmpresaNombre { get; set; } = null!;
        public int? IdFlexoft { get; set; }
        public int? Codiempresa { get; set; }
    }
}
