//using Microsoft.AspNetCore.Components.Web;
//using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
//using IMPLANPROD.Client;
//using Blazored.Modal;
//using CurrieTechnologies.Razor.SweetAlert2;

//var builder = WebAssemblyHostBuilder.CreateDefault(args);
//builder.RootComponents.Add<App>("#app");
//builder.RootComponents.Add<HeadOutlet>("head::after");

//builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });

//await builder.Build().RunAsync();

using Blazored.Modal;
using Blazored.Toast;
using CurrieTechnologies.Razor.SweetAlert2;
using IMPLANPROD.Client;
using IMPLANPROD.Client.Auth;
using IMPLANPROD.Client.Helpers;
using IMPLANPROD.Client.Repositories;
using IMPLANPROD.Client.Services;
using IMPLANPROD.Shared.Services;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");
builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddScoped<IRepository, Repository>();
builder.Services.AddAuthorizationCore();
builder.Services.AddScoped<AnotacionesParserService>();
builder.Services.AddSingleton<FiltrosPedidosService>();
builder.Services.AddScoped<CambiosNoGuardadosService>();
builder.Services.AddScoped<AuthErrorHandlerService>();
builder.Services.AddScoped<RemitosConfigService>();
builder.Services.AddScoped<AuthenticationStateProvider, AuthenticationService>();
builder.Services.AddScoped<AuthenticationService>(sp => 
    (AuthenticationService)sp.GetRequiredService<AuthenticationStateProvider>());
builder.Services.AddScoped<DependenciaService>();
builder.Services.AddScoped<InactivityService>();
builder.Services.AddSweetAlert2();
builder.Services.AddBlazoredModal();
builder.Services.AddBlazoredToast();
builder.Services.AddScoped<FuncionesFechas>();
builder.Services.AddScoped<DerechosUsuarioClientService>();

await builder.Build().RunAsync();
