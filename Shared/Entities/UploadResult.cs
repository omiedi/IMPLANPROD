namespace IMPLANPROD.Shared.Entities
{
    public class UploadResult
    {
        public int Id { get; set; }
        public string? Nombre { get; set; }
        public string? StoredFileName { get; set; }
        public string? ContentType { get; set; }
    }
}
//crear la carpeta uploads en la ruta de server
//CREATE TABLE UploadResult (
//    Id INT PRIMARY KEY IDENTITY(1,1),
//    Nombre NVARCHAR(MAX) NULL,
//    StoredFileName NVARCHAR(MAX) NULL,
//    ContentType NVARCHAR(MAX) NULL
//);

