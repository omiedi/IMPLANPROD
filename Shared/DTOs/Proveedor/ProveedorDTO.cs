using System;

namespace IMPLANPROD.Shared.DTOs
{
    public class ProveedorDTO
    {
        public int Id { get; set; }
        public int? NumeCli { get; set; }
        public string RazonSocial { get; set; } = string.Empty;
        public string Cuit { get; set; } = string.Empty;
        public string Telefono { get; set; } = string.Empty;
        public string Domicilio { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
    }
}
