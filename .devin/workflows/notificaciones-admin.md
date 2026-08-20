---
description: CRUD Admin de notificaciones manuales
---

# Notificaciones (Admin)

Pantallas administrativas para crear/editar/borrar notificaciones manuales y asignar destinatarios (todos o usuarios específicos).

---

## > Instructivo para el Usuario

<details>
<summary><strong>1. Crear una notificación</strong></summary>

1. Ir a `Configuración > Usuarios > Notificaciones (Admin)`.
2. Click en `Nueva`.
3. Completar:
   - **Título** (obligatorio)
   - **Resumen** y **Detalle HTML** (opcional)
   - **Para todos** o seleccionar **Destinatarios**
4. Click en `Guardar`.

</details>

<details>
<summary><strong>2. Editar / Borrar</strong></summary>

- En el listado, usar `Editar` para modificar título/destinatarios/expiración.
- Usar `Borrar` para eliminar la notificación y sus destinatarios.

</details>

<details>
<summary><strong>3. Solución de problemas comunes</strong></summary>

### Problema: no deja guardar sin destinatarios
**Solución:**
- Si no marcás **Para todos**, tenés que tildar al menos 1 usuario.

</details>

---

## > Instructivo para el Desarrollador

<details>
<summary><strong>1. Componentes principales</strong></summary>

- **Client pages**
  - `Client/Pages/Notificaciones/AdminIndex.razor`
  - `Client/Pages/Notificaciones/AdminCreate.razor`
  - `Client/Pages/Notificaciones/AdminEdit.razor`
- **Server API**
  - `Server/Controllers/NotificacionesAdminController.cs`
- **Service**
  - `Server/Services/NotificacionesService.cs` (métodos `Admin*`)

</details>

<details>
<summary><strong>2. Endpoints principales</strong></summary>

```http
GET    /api/notificaciones/admin
GET    /api/notificaciones/admin/{id}
POST   /api/notificaciones/admin
PUT    /api/notificaciones/admin/{id}
DELETE /api/notificaciones/admin/{id}
```

</details>

<details>
<summary><strong>3. Formato de datos</strong></summary>

```json
{
  "titulo": "Texto",
  "resumen": "Opcional",
  "detalleHtml": "<b>...</b>",
  "esParaTodos": false,
  "usuarioIds": [1, 2, 3],
  "fechaExpiracion": "2026-12-31T00:00:00"
}
```

</details>

---

## 📊 Resumen Técnico

| Componente | Ubicación | Propósito |
|------------|-----------|-----------|
| Admin UI | Client/Pages/Notificaciones/Admin*.razor | CRUD manual con destinatarios |
| Admin API | Server/Controllers/NotificacionesAdminController.cs | CRUD admin |
| Service | Server/Services/NotificacionesService.cs | Persistencia + destinatarios |
