# Guía de Configuración de Cascade: Rules, Skills, Workflows y Memories

## 📋 **1. RULES (Reglas)**

Las reglas son instrucciones globales que siempre aplican a tu interacción con Cascade.

**Cómo configurarlas:**
- Se configuran en la configuración de Cascade
- Son persistentes y se aplican a todas las conversaciones
- Definen tu estilo de trabajo preferido

**Ejemplos de uso:**
```
- "Siempre usa español en los comentarios del código"
- "Prefiero código TypeScript sobre JavaScript"
- "Usa single quotes en lugar de double quotes"
- "Siempre agrega logging detallado"
- "Sigue el patrón Repository en el backend"
```

**Cuándo usarlas:**
- Preferencias de estilo de código
- Convenciones de tu equipo
- Patrones arquitectónicos que siempre sigues

---

## 🎯 **2. SKILLS (Habilidades)**

Las skills son capacidades especializadas que Cascade puede usar cuando las necesita.

**Características:**
- Se activan automáticamente según el contexto
- Son herramientas especializadas
- No necesitas invocarlas manualmente

**Ejemplos incluidos en Cascade:**
- Code search (búsqueda en código)
- File operations (lectura/escritura)
- Terminal commands
- Web search
- Debugging tools

**No necesitas configurarlas**, Cascade las usa automáticamente cuando son relevantes.

---

## 🔄 **3. WORKFLOWS (Flujos de trabajo)**

Los workflows son procedimientos paso a paso que defines para tareas repetitivas.

**Ubicación:**
```
.windsurf/workflows/nombre-workflow.md
```

**Formato:**
```markdown
---
description: Cómo hacer deploy de la aplicación
---

# Pasos para Deploy

1. Compilar el proyecto
   ```
   dotnet build --configuration Release
   ```

2. Ejecutar tests
   ```
   dotnet test
   ```

3. Publicar
   ```
   dotnet publish -c Release -o ./publish
   ```

// turbo
4. Subir a servidor
   ```
   scp -r ./publish usuario@servidor:/var/www/
   ```
```

**Características especiales:**
- **`// turbo`**: Permite auto-ejecutar el comando siguiente sin confirmación
- **Slash commands**: Puedes invocarlos con `/nombre-workflow`

**Cuándo crear workflows:**
- Tareas que repites frecuentemente
- Procesos con múltiples pasos
- Procedimientos de deploy
- Configuraciones iniciales de proyectos
- Scripts de mantenimiento

**Ejemplo práctico para tu proyecto:**
```markdown
---
description: Agregar validación de derechos a una página
---

# Agregar Sistema de Derechos

1. Inyectar el servicio en la página
   ```csharp
   @inject DerechosUsuarioClientService derechosService
   @inject NavigationManager NavigationManager
   ```

2. Declarar variables de control
   ```csharp
   private bool tieneAccesoEscritura = false;
   private bool verificandoPermisos = true;
   private const string COD_APLI = "TU_CODIGO";
   private const string DES_APLI = "DESCRIPCION";
   ```

3. Verificar en OnInitializedAsync
   ```csharp
   var nivelAcceso = await derechosService.VerificarDerechoAsync(
       COD_APLI, DES_APLI, mostrarMensaje: true
   );
   tieneAccesoEscritura = nivelAcceso >= 2;
   ```

4. Deshabilitar botones según permisos
   ```html
   <button disabled="@(!tieneAccesoEscritura)">
   ```
```

---

## 💾 **4. MEMORIES (Memorias)**

Las memories son información contextual que Cascade recuerda entre conversaciones.

**Tipos de memories:**

### **A. User-triggered (Explícitas)**
Cuando le dices explícitamente a Cascade que recuerde algo:
```
"Recuerda que en este proyecto usamos el patrón Repository"
"Guarda que el puerto del servidor es 5001"
```

### **B. System-retrieved (Automáticas)**
Cascade automáticamente guarda información importante:
- Decisiones arquitectónicas
- Patrones de código que usas
- Configuraciones importantes
- Soluciones a problemas complejos

**Cuándo se crean automáticamente:**
- Después de resolver un problema complejo
- Al implementar un patrón nuevo
- Al configurar algo importante
- Al documentar decisiones técnicas

**Ejemplo de lo que Cascade recuerda automáticamente:**
```
- "El sistema usa IdFlexoft para match con DER_ACCE.NUM_USER"
- "Los administradores tienen acceso nivel 2 automáticamente"
- "El proyecto compila con dotnet build en la raíz"
```

---

## 🎨 **Cómo Usar Todo Junto - Ejemplo Práctico**

### **Escenario: Desarrollo en tu proyecto IMPLANPROD**

**1. RULES (Una vez, al inicio):**
```
- Siempre usar español en comentarios
- Seguir el patrón Repository existente
- Agregar logging con _logger.LogInformation
- Compilar con dotnet build antes de confirmar cambios
```

**2. WORKFLOWS (Crear según necesites):**

**`.windsurf/workflows/agregar-derechos.md`**
```markdown
---
description: Agregar validación de derechos a una página
---
[Pasos detallados como mostré arriba]
```

**`.windsurf/workflows/crear-entidad.md`**
```markdown
---
description: Crear nueva entidad con auditoría
---
1. Crear clase en Shared/Entities
2. Agregar DbSet en DataContext
3. Crear migración
4. Crear controller
5. Crear servicio
```

**3. MEMORIES (Automáticas):**
Cascade ya está recordando cosas como:
- Sistema de derechos con DER_ACCE
- Perfil ADMINISTRADOR tiene acceso completo
- Estructura del proyecto
- Patrones que usas

**4. SKILLS (Automáticas):**
Cascade usa automáticamente:
- `code_search` cuando necesita buscar código
- `read_file` cuando necesita ver archivos
- `edit` cuando necesita modificar código
- `run_command` cuando necesita compilar

---

## 💡 **Recomendaciones Prácticas**

### **Empieza Simple:**
1. **Define 3-5 rules básicas** sobre tu estilo de código
2. **Crea 1-2 workflows** para tus tareas más repetitivas
3. **Deja que memories se generen automáticamente**
4. **No te preocupes por skills**, Cascade las usa solo

### **Workflows Útiles para tu Proyecto:**
- `/agregar-derechos` - Agregar validación de derechos
- `/nueva-pagina` - Crear nueva página Blazor
- `/nueva-entidad` - Crear entidad con auditoría
- `/deploy` - Proceso de deploy
- `/test` - Ejecutar tests

### **Cuándo Actualizar:**
- **Rules**: Raramente, solo cuando cambias preferencias globales
- **Workflows**: Cuando identificas una tarea repetitiva
- **Memories**: Automático, pero puedes pedir guardar algo específico
- **Skills**: No necesitas, son automáticas

---

## 🚀 **Próximos Pasos Sugeridos**

1. **Define tus rules básicas** en la configuración de Cascade
2. **Crea un workflow** para "Agregar validación de derechos" (lo usarás mucho)
3. **Deja que Cascade aprenda** tu estilo mientras trabajas
4. **Usa slash commands** cuando necesites workflows: `/agregar-derechos`

---

## 📚 **Recursos Adicionales**

### **Ubicaciones de Archivos:**
- Rules: Configuración de Cascade (Settings)
- Workflows: `.windsurf/workflows/*.md`
- Memories: Gestionadas automáticamente por Cascade
- Skills: Integradas en Cascade

### **Comandos Útiles:**
- Ver workflows disponibles: Escribe `/` en el chat
- Crear memory explícita: "Recuerda que..."
- Invocar workflow: `/nombre-workflow`

---

**Fecha de creación:** 16 de enero de 2026
**Proyecto:** IMPLANPROD
**Autor:** Documentación generada por Cascade
