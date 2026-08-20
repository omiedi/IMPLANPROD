---
description: Modal de notificaciones (campanita)
---

# Modal de Notificaciones

Este modal permite a cada usuario ver sus notificaciones persistentes, abrir el detalle, marcarlas como leídas y ocultarlas ("borrar para mí").

---

## > Instructivo para el Usuario

<details>
<summary><strong>1. Funcionamiento</strong></summary>

- **Abrir el modal**
  - Hacé click en la campanita.
- **Leer una notificación**
  - Click en el título para abrir/cerrar el detalle.
- **Marcar como leída**
  - Click en el ícono de ojo (marca como leída sin abrir el detalle).
- **Borrar para mí**
  - Click en el ícono de cesto (la oculta solo para tu usuario).

</details>

<details>
<summary><strong>2. Verificación</strong></summary>

- **Contador**
  - El badge muestra cuántas notificaciones tenés sin leer.
- **Refrescar**
  - Usá el botón `Refrescar` para volver a cargar desde el servidor.

</details>

<details>
<summary><strong>3. Solución de problemas comunes</strong></summary>

### Problema: no desaparece al borrar para mí
**Solución:**
- Confirmar que el servidor esté actualizado.
- Confirmar que existe la tabla `NotificacionOcultas` en la base.

### Problema: no marca como leída
**Solución:**
- Reintentar y revisar el mensaje de error.
- Confirmar que existe la tabla `NotificacionLeidas` en la base.

</details>

---

## > Instructivo para el Desarrollador

<details>
<summary><strong>1. Arquitectura del Sistema</strong></summary>

- **Client**
  - `Client/Shared/Components/NotificationsModal.razor`
  - `Client/Services/NotificationService.cs`
- **Server**
  - `Server/Controllers/NotificacionesController.cs`
  - `Server/Services/NotificacionesService.cs`
- **Data**
  - Tablas:
    - `Notificaciones`
    - `NotificacionLeidas`
    - `NotificacionOcultas`
    - `NotificacionDestinatarios`

</details>

<details>
<summary><strong>2. Flujo de Funcionamiento</strong></summary>

```
Modal abre
    ↓
GET api/notificaciones
    ↓
Click ojo
    ↓
POST api/notificaciones/markRead
    ↓
Refresh

Click cesto
    ↓
POST api/notificaciones/hide
    ↓
Refresh
```

</details>

<details>
<summary><strong>3. Endpoints principales</strong></summary>

```http
GET    /api/notificaciones
GET    /api/notificaciones/unreadCount
POST   /api/notificaciones/markRead
POST   /api/notificaciones/hide
GET    /api/notificaciones/{id}/detail
```

</details>

---

## 📊 Resumen Técnico

| Componente | Ubicación | Propósito |
|------------|-----------|-----------|
| Modal | Client/Shared/Components/NotificationsModal.razor | UI de notificaciones del usuario |
| Service | Server/Services/NotificacionesService.cs | Lógica de listado/lecturas/ocultar |
| Controller | Server/Controllers/NotificacionesController.cs | API pública de notificaciones |
