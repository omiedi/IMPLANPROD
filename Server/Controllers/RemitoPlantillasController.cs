using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs.Impresiones;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/impresiones/remito-plantillas")]
    [Authorize]
    public class RemitoPlantillasController : ControllerBase
    {
        private readonly RemitoPlantillaService _service;
        private readonly PdfService _pdfService;
        private readonly ILogger<RemitoPlantillasController> _logger;

        public RemitoPlantillasController(
            RemitoPlantillaService service,
            PdfService pdfService,
            ILogger<RemitoPlantillasController> logger)
        {
            _service = service;
            _pdfService = pdfService;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> GetAllAsync(CancellationToken cancellationToken)
        {
            try
            {
                return Ok(await _service.GetAllAsync(cancellationToken));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al listar plantillas de remito");
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetByIdAsync([FromRoute] int id, CancellationToken cancellationToken)
        {
            try
            {
                var dto = await _service.GetByIdAsync(id, cancellationToken);
                return dto == null ? NotFound() : Ok(dto);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener plantilla de remito {Id}", id);
                return BadRequest(ex.Message);
            }
        }

        [HttpPost]
        public async Task<IActionResult> SaveAsync([FromBody] RemitoPlantillaUpsertDto dto, CancellationToken cancellationToken)
        {
            try
            {
                var id = await _service.SaveAsync(dto, cancellationToken);
                return Ok(id);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al guardar plantilla de remito");
                return BadRequest(ex.InnerException?.Message ?? ex.Message);
            }
        }

        [HttpPost("preview")]
        public async Task<IActionResult> PreviewAsync([FromBody] RemitoPlantillaPreviewRequestDto dto, CancellationToken cancellationToken)
        {
            try
            {
                var preview = await _service.PreviewAsync(dto, cancellationToken);
                await _service.EnrichPreviewWithObjectSqlAsync(dto.LienzoJson, preview, dto.Parameters, cancellationToken);
                return Ok(preview);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en preview de plantilla de remito");
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("{id:int}/pdf")]
        public async Task<IActionResult> GeneratePdfAsync(
            [FromRoute] int id,
            CancellationToken cancellationToken)
        {
            try
            {
                return await BuildPdfResponseAsync(id, null, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar PDF de remito para plantilla {Id}", id);
                return BadRequest(ex.InnerException?.Message ?? ex.Message);
            }
        }

        [HttpGet("{id:int}/pdf/remito/{nroRemito}")]
        public async Task<IActionResult> GeneratePdfByRemitoAsync(
            [FromRoute] int id,
            [FromRoute] string nroRemito,
            CancellationToken cancellationToken)
        {
            try
            {
                return await BuildPdfResponseAsync(id, nroRemito, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar PDF de remito para plantilla {Id} y remito {NroRemito}", id, nroRemito);
                return BadRequest(ex.InnerException?.Message ?? ex.Message);
            }
        }

        [HttpGet("pdf/milimetrada")]
        public async Task<IActionResult> GenerateMillimeterPagePdfAsync(
            [FromQuery] string? orientation,
            CancellationToken cancellationToken)
        {
            try
            {
                var pdfBytes = await _pdfService.GenerateMillimeterPageAsync(orientation, cancellationToken);
                var orientationSuffix = string.Equals(orientation, "Landscape", StringComparison.OrdinalIgnoreCase)
                    ? "Landscape"
                    : "Portrait";

                return File(pdfBytes, "application/pdf", $"PaginaMilimetrada_{orientationSuffix}.pdf");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar PDF de página milimetrada");
                return BadRequest(ex.InnerException?.Message ?? ex.Message);
            }
        }

        private async Task<IActionResult> BuildPdfResponseAsync(int id, string? fallbackDoprilar, CancellationToken cancellationToken)
        {
            var parameters = BuildParametersFromQuery();
            if (!parameters.Any(x => string.Equals(x.Name, "Doprilar", StringComparison.OrdinalIgnoreCase))
                && !string.IsNullOrWhiteSpace(fallbackDoprilar))
            {
                parameters.Add(new RemitoPlantillaParameterDto { Name = "Doprilar", Value = fallbackDoprilar });
            }

            var doprilar = parameters
                .FirstOrDefault(x => string.Equals(x.Name, "Doprilar", StringComparison.OrdinalIgnoreCase))
                ?.Value;

            if (string.IsNullOrWhiteSpace(doprilar))
            {
                return BadRequest("El número de remito es obligatorio.");
            }

            var plantilla = await _service.GetByIdAsync(id, cancellationToken);
            if (plantilla == null)
            {
                return NotFound();
            }

            var previewRequest = new RemitoPlantillaPreviewRequestDto
            {
                LienzoJson = plantilla.LienzoJson,
                Datasets = plantilla.Datasets,
                Parameters = parameters
            };

            var preview = await _service.PreviewAsync(previewRequest, cancellationToken);
            await _service.EnrichPreviewWithObjectSqlAsync(plantilla.LienzoJson, preview, previewRequest.Parameters, cancellationToken);
            var pdfBytes = await _pdfService.GenerateRemitoFromTemplateAsync(plantilla, preview, cancellationToken);

            var safeNombre = string.IsNullOrWhiteSpace(plantilla.Nombre) ? "Remito" : plantilla.Nombre.Trim();
            var fileName = $"{safeNombre}_{doprilar}.pdf";

            return File(pdfBytes, "application/pdf", fileName);
        }

        private List<RemitoPlantillaParameterDto> BuildParametersFromQuery()
        {
            var parameters = new List<RemitoPlantillaParameterDto>();
            foreach (var queryItem in Request.Query)
            {
                if (string.IsNullOrWhiteSpace(queryItem.Key))
                {
                    continue;
                }

                parameters.Add(new RemitoPlantillaParameterDto
                {
                    Name = queryItem.Key,
                    Value = queryItem.Value.FirstOrDefault()
                });
            }

            return parameters;
        }
    }
}
