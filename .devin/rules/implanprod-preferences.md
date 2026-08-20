# Preferencias Globales de IMPLANPROD

Reglas de comportamiento y estilo que aplican en todas las interacciones con este proyecto.

## Idioma

- **Responde siempre en español.**
- Los comentarios del código deben redactarse en español claro.

## Calidad de los comentarios

- Comentar el código agregado o modificado de forma clara, simple y didáctica, para que pueda ser interpretado por programadores junior.
- Explicar intención, flujo y decisiones relevantes.
- Evitar comentarios crípticos o excesivamente técnicos cuando no sean necesarios.
- Al modificar código que ya tiene comentarios, revisarlos y actualizarlos para que reflejen el comportamiento real.

## Procesos de prueba

- **NO** dejar procesos de PowerShell/dotnet run en segundo plano bloqueando la ejecución.
- Preferir `dotnet build` y/o pedir confirmación antes de levantar un servidor.

## Inputs numéricos en Blazor

- Cuando se requiera un input numérico sin flechas de incremento/decremento, usar el componente `InputNumeroLimpio.razor`.
- **NO** usar `<input type="number">` ni `<InputNumber>` estándar si el diseño requiere ocultar las flechas.
- **NO** duplicar código CSS; usar la clase `.sin-flechas` existente en `Client/wwwroot/css/app.css`.
- Ubicación del componente: `Client/Shared/Components/InputNumeroLimpio.razor`.

## Componentes compartidos

- Al usar el componente `<InstructivoDesarrollador>` en un archivo `.razor`, declarar:
  ```razor
  @using IMPLANPROD.Client.Shared.Components
  ```
