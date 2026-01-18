using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Vml;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Representa una fórmula en el sistema de producción
    /// </summary>
    public class Formulas : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        public int? CodIConj { get; set; }

        public short? CodiEmpr { get; set; }

        [StringLength(24)]
        public string? CodXConj { get; set; }

        public int? CodiElem { get; set; }

        [StringLength(24)]
        public string? CodXElem { get; set; }

        public short? NuorIt { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? UsoBruto { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Unis_Form { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Lar_Corte { get; set; }

        public int? Pladispo { get; set; }

        public string? RefeComp { get; set; }

        public short? Marborra { get; set; }

        [StringLength(8)]
        public string? UMED_FORM { get; set; }

        [StringLength(48)]
        public string? COMPONENTE { get; set; }

        public short? MODUXBB { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? CORTE { get; set; }

        [StringLength(512)]
        public string? VALOBADA_1 { get; set; }

        // Propiedades de navegación

        //[ForeignKey("CodIConj")]
        //public Master? codint { get; set; }

        // Propiedad de navegación 
        
        [NotMapped] // Esta propiedad no se mapeará a la base de datos
        public string? MasterDescripcion { get; set; } // Propiedad virtual que no se almacena en la tabla

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

        //EJEMPLO DE Propiedad de navegación

        // Propiedad de navegación
        // ** EN LA ENTIDAD FORMULAS NAVEGACION CON EL MASTER 
        //[ForeignKey("CodiElem")]
        //public Master? MasterDescripcion { get; set; }
        //
        // ** EN EL CONTROLADOR AGREGAMOS EL INCLUDE PARA QUE RELACIONE LA DESCRIPCION DEL MASTER

        //[HttpGet("conjunto/{codigoConjunto:int}")]
        //public async Task<IActionResult> GetComponentesConjuntoAsync(int codigoConjunto)
        //{
        //    try
        //    {
        //      var componentes = await _context.Formulas
        //     .Where(f => f.CodIConj == codigoConjunto)
        //     .Include(f => f.MasterDescripcion)
        //     .OrderBy(f => f.NuorIt)
        //     .ToListAsync();

        //        return Ok(componentes);
        //    }
        //    catch (Exception ex)
        //    {

        //        throw;
        //    }

        //}

        //
        // ** EN EL HTML MOSTRAMOS EL CODXELEM Y LA DESCRIPCION QUE LA TRAE DEL MASTERDESCRIPCION
        // ** @componente SE CARGA CON UN GET DE ENTIDAD FORMULAS COMO TODA LISTA PARA MOSTRAR EN EL HTML 
        //<td>@componente.CodXElem</td>
        //<td>@componente.MasterDescripcion?.Descripcion</td> <!-- Ajuste aquí -->


    }

    //cambiar manualmente el nombre del campo formulas_id x ID
    //ALTER TABLE formulas ALTER COLUMN CodIConj INT
    //ALTER TABLE formulas ALTER COLUMN CodiElem INT
    //ALTER TABLE formulas ALTER COLUMN[Pladispo] INT

    //ALTER TABLE formulas
    //ALTER COLUMN UsoBruto DECIMAL(18,4);
    //        ALTER TABLE formulas
    //        ALTER COLUMN[Unis_Form] DECIMAL(18,4);

    //        ALTER TABLE formulas
    //        ALTER COLUMN[Lar_Corte] DECIMAL(18,4);
    //        ALTER TABLE formulas
    //        ALTER COLUMN[CORTE] DECIMAL(18,4);
}
