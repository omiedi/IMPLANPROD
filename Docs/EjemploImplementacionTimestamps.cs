using IMPLANPROD.Shared.Entities;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// EJEMPLO: Nueva entidad con sistema de timestamps automáticos
    /// Esta entidad demuestra cómo implementar ITimestampEntity desde cero
    /// </summary>
    [Table("EJEMPLO_TIMESTAMPS")]
    public class EjemploTimestamps : ITimestampEntity
    {
        /// <summary>
        /// Clave primaria
        /// </summary>
        [Key]
        [Column("ID")]
        public int Id { get; set; }

        /// <summary>
        /// Nombre del ejemplo
        /// </summary>
        [Required(ErrorMessage = "El nombre es obligatorio")]
        [StringLength(100, ErrorMessage = "El nombre no puede exceder los 100 caracteres")]
        [Column("NOMBRE")]
        public string Nombre { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del ejemplo
        /// </summary>
        [StringLength(500, ErrorMessage = "La descripción no puede exceder los 500 caracteres")]
        [Column("DESCRIPCION")]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Estado activo/inactivo
        /// </summary>
        [Column("ACTIVO")]
        public bool Activo { get; set; } = true;

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CODI_EMPR")]
        public int CodiEmpr { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================
        
        /// <summary>
        /// Fecha de creación del registro
        /// Se establece automáticamente cuando se crea el registro
        /// </summary>
        [Column("ADD_RECORD")]
        public DateTime AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente cada vez que se modifica el registro
        /// </summary>
        [Column("LAST_UPDATE")]
        public DateTime LastUpdate { get; set; }
    }
}

/*
-- SCRIPT SQL PARA CREAR LA TABLA CON TIMESTAMPS:

CREATE TABLE [dbo].[EJEMPLO_TIMESTAMPS] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [NOMBRE] nvarchar(100) NOT NULL,
    [DESCRIPCION] nvarchar(500) NULL,
    [ACTIVO] bit NOT NULL DEFAULT 1,
    [CODI_EMPR] int NOT NULL,
    [ADD_RECORD] datetime2 NOT NULL DEFAULT GETDATE(),
    [LAST_UPDATE] datetime2 NOT NULL DEFAULT GETDATE(),
    
    CONSTRAINT [PK_EJEMPLO_TIMESTAMPS] PRIMARY KEY CLUSTERED ([ID] ASC)
);

-- Índices para optimizar consultas por fechas
CREATE INDEX [IX_EJEMPLO_TIMESTAMPS_AddRecord] ON [dbo].[EJEMPLO_TIMESTAMPS] ([ADD_RECORD]);
CREATE INDEX [IX_EJEMPLO_TIMESTAMPS_LastUpdate] ON [dbo].[EJEMPLO_TIMESTAMPS] ([LAST_UPDATE]);
CREATE INDEX [IX_EJEMPLO_TIMESTAMPS_Activo] ON [dbo].[EJEMPLO_TIMESTAMPS] ([ACTIVO]);

-- Comentarios en la tabla
EXEC sys.sp_addextendedproperty 
    @name=N'MS_Description', 
    @value=N'Tabla de ejemplo que demuestra el uso del sistema de timestamps automáticos', 
    @level0type=N'SCHEMA', @level0name=N'dbo', 
    @level1type=N'TABLE', @level1name=N'EJEMPLO_TIMESTAMPS';

EXEC sys.sp_addextendedproperty 
    @name=N'MS_Description', 
    @value=N'Fecha de creación del registro (automática)', 
    @level0type=N'SCHEMA', @level0name=N'dbo', 
    @level1type=N'TABLE', @level1name=N'EJEMPLO_TIMESTAMPS', 
    @level2type=N'COLUMN', @level2name=N'ADD_RECORD';

EXEC sys.sp_addextendedproperty 
    @name=N'MS_Description', 
    @value=N'Fecha de última actualización del registro (automática)', 
    @level0type=N'SCHEMA', @level0name=N'dbo', 
    @level1type=N'TABLE', @level1name=N'EJEMPLO_TIMESTAMPS', 
    @level2type=N'COLUMN', @level2name=N'LAST_UPDATE';
*/

/*
EJEMPLO DE USO EN CONTROLADOR:

[ApiController]
[Route("api/[controller]")]
public class EjemploTimestampsController : ControllerBase
{
    private readonly DataContext _context;

    public EjemploTimestampsController(DataContext context)
    {
        _context = context;
    }

    [HttpPost]
    public async Task<ActionResult<EjemploTimestamps>> Create(EjemploTimestamps ejemplo)
    {
        // NO es necesario establecer AddRecord o LastUpdate manualmente
        // El TimestampInterceptor lo hace automáticamente
        
        _context.Set<EjemploTimestamps>().Add(ejemplo);
        await _context.SaveChangesAsync();
        
        // En este punto:
        // - ejemplo.AddRecord tendrá la fecha/hora actual
        // - ejemplo.LastUpdate tendrá la fecha/hora actual
        
        return CreatedAtAction(nameof(GetById), new { id = ejemplo.Id }, ejemplo);
    }

    [HttpPut("{id}")]
    public async Task<IActionResult> Update(int id, EjemploTimestamps ejemplo)
    {
        if (id != ejemplo.Id)
            return BadRequest();

        _context.Entry(ejemplo).State = EntityState.Modified;
        
        // NO es necesario establecer LastUpdate manualmente
        // El TimestampInterceptor lo hace automáticamente
        
        await _context.SaveChangesAsync();
        
        // En este punto:
        // - ejemplo.AddRecord mantiene su valor original
        // - ejemplo.LastUpdate se actualiza con la fecha/hora actual
        
        return NoContent();
    }
}
*/
