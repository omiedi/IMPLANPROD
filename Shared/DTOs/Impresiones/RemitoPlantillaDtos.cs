using System.Text.Json.Serialization;

namespace IMPLANPROD.Shared.DTOs.Impresiones
{
    public class RemitoPlantillaItemDto
    {
        public int Id { get; set; }
        public string Nombre { get; set; } = string.Empty;
        public DateTime FechaModificacion { get; set; }
    }

    public class RemitoPlantillaUpsertDto
    {
        public int Id { get; set; }
        public string Nombre { get; set; } = string.Empty;
        public string? LienzoJson { get; set; }
        public List<RemitoPlantillaDatasetDto> Datasets { get; set; } = new();
    }

    public class RemitoPlantillaDetailDto : RemitoPlantillaUpsertDto
    {
        public DateTime FechaCreacion { get; set; }
        public DateTime FechaModificacion { get; set; }
    }

    public class RemitoPlantillaDatasetDto
    {
        public string Key { get; set; } = string.Empty;
        public string Sql { get; set; } = string.Empty;
        public int MaxRows { get; set; } = 500;
    }

    public class RemitoPlantillaParameterDto
    {
        public string Name { get; set; } = string.Empty;
        public string? Value { get; set; }
    }

    public class RemitoPlantillaPreviewRequestDto
    {
        public string? LienzoJson { get; set; }
        public List<RemitoPlantillaDatasetDto> Datasets { get; set; } = new();
        public List<RemitoPlantillaParameterDto> Parameters { get; set; } = new();
    }

    public class RemitoPlantillaPreviewResultDto
    {
        public Dictionary<string, List<Dictionary<string, object?>>> Results { get; set; } = new();
    }

    public class RemitoPlantillaConfigStorageDto
    {
        public string? LienzoJson { get; set; }
        public List<RemitoPlantillaDatasetDto> Datasets { get; set; } = new();
    }
}
