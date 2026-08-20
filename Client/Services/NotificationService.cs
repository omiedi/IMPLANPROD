using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using IMPLANPROD.Client.Repositories;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.DTOs.Notificaciones;

namespace IMPLANPROD.Client.Services;

public class NotificationService
{
    public event Action? OnChange;

    private readonly IRepository _repository;

    private readonly List<NotificacionItemDTO> _items = new();

    public NotificationService(IRepository repository)
    {
        _repository = repository;
    }

    public IReadOnlyList<NotificacionItemDTO> Items => _items;

    public int GetUnreadCountForUser(int userId)
    {
        if (userId <= 0)
        {
            return 0;
        }

        return _items.Count(x => !x.Leida);
    }

    public IReadOnlyList<NotificacionItemDTO> GetItemsForUser(int userId)
    {
        if (userId <= 0)
        {
            return Array.Empty<NotificacionItemDTO>();
        }

        return _items.ToList();
    }

    public async Task RefreshAsync(PaginationDTO? pagination = null)
    {
        var p = pagination ?? new PaginationDTO { Page = 1, RecordsNumber = 50 };
        var url = $"api/notificaciones?Page={p.Page}&RecordsNumber={p.RecordsNumber}&Filter={Uri.EscapeDataString(p.Filter ?? string.Empty)}";

        var response = await _repository.Get<List<NotificacionItemDTO>>(url);
        if (response.Error || response.Response == null)
        {
            return;
        }

        _items.Clear();
        _items.AddRange(response.Response);
        OnChange?.Invoke();
    }

    public async Task<int> GetUnreadCountFromServerAsync()
    {
        var response = await _repository.Get<int>("api/notificaciones/unreadCount");
        if (response.Error)
        {
            return 0;
        }

        return response.Response;
    }

    public async Task<int?> AddForAllAsync(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return null;
        }

        var dto = new CrearNotificacionRequestDTO
        {
            Titulo = message.Trim(),
            EsParaTodos = true
        };

        var response = await _repository.Post<CrearNotificacionRequestDTO, int>("api/notificaciones", dto);
        if (response.Error)
        {
            return null;
        }

        await RefreshAsync();
        return response.Response;
    }

    public async Task<int?> AddForUsersAsync(string message, IEnumerable<int> targetUserIds)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return null;
        }

        var list = targetUserIds?.Where(x => x > 0).Distinct().ToList() ?? new List<int>();
        if (list.Count == 0)
        {
            return null;
        }

        var dto = new CrearNotificacionRequestDTO
        {
            Titulo = message.Trim(),
            EsParaTodos = false,
            UsuarioIds = list
        };

        var response = await _repository.Post<CrearNotificacionRequestDTO, int>("api/notificaciones", dto);
        if (response.Error)
        {
            return null;
        }

        await RefreshAsync();
        return response.Response;
    }

    public async Task MarkAsReadAsync(int notificacionId)
    {
        if (notificacionId <= 0)
        {
            return;
        }

        var response = await _repository.Post("api/notificaciones/markRead", new MarcarLeidaRequestDTO { NotificacionId = notificacionId });
        if (!response.Error)
        {
            await RefreshAsync();
        }
    }

    public async Task<HttpResponseWrapper<object>> HideForMeAsync(int notificacionId)
    {
        if (notificacionId <= 0)
        {
            return new HttpResponseWrapper<object>(null, true, new HttpResponseMessage(System.Net.HttpStatusCode.BadRequest));
        }

        var response = await _repository.Post("api/notificaciones/hide", new OcultarNotificacionRequestDTO { NotificacionId = notificacionId });
        if (!response.Error)
        {
            await RefreshAsync();
        }

        return response;
    }

    public async Task<NotificacionDetalleDTO?> GetDetailAsync(int notificacionId)
    {
        if (notificacionId <= 0)
        {
            return null;
        }

        var response = await _repository.Get<NotificacionDetalleDTO>($"api/notificaciones/{notificacionId}/detail");
        if (response.Error)
        {
            return null;
        }

        return response.Response;
    }
}
