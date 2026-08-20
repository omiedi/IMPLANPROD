using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    public class BootstrapSuperAdminStatusDto
    {
        public int TotalSuperAdmins { get; set; }
        public bool RequiereBootstrap => TotalSuperAdmins <= 0;
    }

    public class SeguridadModuloDto
    {
        public int Id { get; set; }
        public string Codigo { get; set; } = string.Empty;
        public string Nombre { get; set; } = string.Empty;
        public int Orden { get; set; }
        public List<SeguridadPaginaDto> Paginas { get; set; } = new();
    }

    public class SeguridadPaginaDto
    {
        public int Id { get; set; }
        public int ModuloId { get; set; }
        public string Nombre { get; set; } = string.Empty;
        public string Ruta { get; set; } = string.Empty;
        public bool EsConfiguracion { get; set; }
        public int Orden { get; set; }
    }

    public class AsignarPaginasUsuarioRequestDto
    {
        [Required]
        public List<int> PaginaIds { get; set; } = new();
    }
}
