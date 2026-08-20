---
description: CRUD de reglas SQL para notificaciones automáticas
---

# Notificaciones por Reglas SQL

Permite administrar reglas que ejecutan consultas SQL (SELECT) en intervalos y generan notificaciones automáticas cuando cambia el resultado.

---

## > Instructivo para el Usuario

<details>
<summary><strong>1. Crear una regla</strong></summary>

1. Ir a `Configuración > Usuarios > Reglas SQL (Notif.)`.
2. Click en `Nueva regla`.
3. Completar:
   - **Título** y **SQL (solo SELECT)**
   - **Intervalo** (minutos)
   - **Max filas detalle**
4. Guardar.

</details>

<details>
<summary><strong>2. Activar / Desactivar</strong></summary>

- En el listado, usar `Activar/Desactivar`.
- Una regla inactiva no se ejecuta.

</details>

<details>
<summary><strong>3. Solución de problemas comunes</strong></summary>

### Problema: la regla no guarda
**Solución:**
- El SQL debe empezar con `SELECT`.

</details>

---

## > Instructivo para el Desarrollador

<details>
<summary><strong>1. Componentes principales</strong></summary>

- **Client pages**
  - `Client/Pages/Notificaciones/ReglasSqlIndex.razor`
  - `Client/Pages/Notificaciones/ReglasSqlCreate.razor`
  - `Client/Pages/Notificaciones/ReglasSqlEdit.razor`
- **Server API**
  - `Server/Controllers/NotificacionReglasSqlController.cs`
- **Service**
  - `Server/Services/NotificacionReglasSqlService.cs`
- **Hosted service**
  - `Server/Services/NotificacionReglasSqlHostedService.cs`

</details>

<details>
<summary><strong>2. Endpoints principales</strong></summary>

```http
GET    /api/notificaciones/reglas-sql
GET    /api/notificaciones/reglas-sql/{id}
POST   /api/notificaciones/reglas-sql
PUT    /api/notificaciones/reglas-sql/{id}
DELETE /api/notificaciones/reglas-sql/{id}
```

</details>

<details>
<summary><strong>3. Flujo de funcionamiento</strong></summary>

```
HostedService (timer)
    ↓
Lee reglas activas
    ↓
Ejecuta SELECT
    ↓
Compara hash resultado
    ↓
Genera Notificacion + DetalleHtml
```

</details>

---

## 📊 Resumen Técnico

| Componente | Ubicación | Propósito |
|------------|-----------|-----------|
| CRUD reglas | Client/Pages/Notificaciones/ReglasSql*.razor | UI de administración |
| API reglas | Server/Controllers/NotificacionReglasSqlController.cs | Endpoints CRUD |
| Runner | Server/Services/NotificacionReglasSqlHostedService.cs | Ejecutor periódico |
