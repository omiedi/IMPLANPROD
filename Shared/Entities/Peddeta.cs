using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa el detalle de un pedido en el sistema.
    /// Mapea a la tabla PEDDETA en la base de datos.
    /// </summary>
    [Table("PEDDETA")]
    public class Peddeta : IAuditableEntity
    {
        /// <summary>
        /// Identificador único del detalle de pedido (campo IDENTITY)
        /// </summary>
        [Key]
        [Column("PEDDETA_ID")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int PeddetaId { get; set; }

        /// <summary>
        /// Precio de lista
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PreList { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Cantidad asignada
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CanAsig { get; set; }

        /// <summary>
        /// Cantidad pedida
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CanPed { get; set; }

        /// <summary>
        /// Precio unitario
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PreUnid { get; set; }

        /// <summary>
        /// Importe total
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? ImpoTot { get; set; }

        /// <summary>
        /// Importe anticipado
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? ImpoAnti { get; set; }

        /// <summary>
        /// Cantidad entregada
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantEnt { get; set; }

        /// <summary>
        /// Cantidad de saldo
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CanSaldo { get; set; }

        /// <summary>
        /// Fecha de última entrega
        /// </summary>
        public DateTime? FeUltEnt { get; set; }

        /// <summary>
        /// Fecha de aprobación
        /// </summary>
        public DateTime? FApro_Nope { get; set; }

        /// <summary>
        /// Acumulado preparado
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? AcuPrepa { get; set; }

        /// <summary>
        /// Acumulado despachado
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? AcuDespa { get; set; }

        /// <summary>
        /// Número de remito
        /// </summary>
        [StringLength(24)]
        public string? NroRemi { get; set; }

        /// <summary>
        /// Tipo de pedido item
        /// </summary>
        [StringLength(16)]
        public string? TPedItem { get; set; }

        /// <summary>
        /// Número de pedido
        /// </summary>
        public int? NroPed { get; set; }

        /// <summary>
        /// Tipo de pedido
        /// </summary>
        [StringLength(8)]
        public string? TipoPed { get; set; }

        /// <summary>
        /// Fecha de emisión
        /// </summary>
        public DateTime? FechEmis { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        public int? NroClie { get; set; }

        /// <summary>
        /// Razón social del cliente
        /// </summary>
        [StringLength(64)]
        public string? RaSoClie { get; set; }

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        [StringLength(16)]
        public string? NroOcom { get; set; }

        /// <summary>
        /// Fecha solicitada de entrega
        /// </summary>
        public DateTime? FeSolEnt { get; set; }

        /// <summary>
        /// Fin de vigencia
        /// </summary>
        public DateTime? FinVigen { get; set; }

        /// <summary>
        /// Estado del pedido
        /// </summary>
        public short? Status { get; set; }

        /// <summary>
        /// Estado de despacho
        /// </summary>
        public short? StaDespa { get; set; }

        /// <summary>
        /// Lista de precios
        /// </summary>
        public short? ListPre { get; set; }

        /// <summary>
        /// Condición de pago
        /// </summary>
        public short? CondPag { get; set; }

        /// <summary>
        /// Vendedor
        /// </summary>
        public short? Vended { get; set; }

        /// <summary>
        /// Porcentaje de comisión del vendedor
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PorComVend { get; set; }

        /// <summary>
        /// Número de orden item
        /// </summary>
        [Column("NuOrItem")]
        public short? NuOrItem { get; set; }

        /// <summary>
        /// Código interno
        /// </summary>
        public int? CodiInt { get; set; }

        /// <summary>
        /// Código externo
        /// </summary>
        [StringLength(24)]
        public string? CodiExt { get; set; }

        /// <summary>
        /// Descripción del producto
        /// </summary>
        [StringLength(384)]
        public string? Descrip { get; set; }

        /// <summary>
        /// Sub código
        /// </summary>
        [StringLength(8)]
        public string? SubCod { get; set; }

        /// <summary>
        /// Tipo de color
        /// </summary>
        [StringLength(16)]
        public string? TColor { get; set; }

        /// <summary>
        /// Tipo de talle
        /// </summary>
        [StringLength(16)]
        public string? TTalle { get; set; }

        /// <summary>
        /// Tipo de largo
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TLargo { get; set; }

        /// <summary>
        /// Tipo de ancho
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TAncho { get; set; }

        /// <summary>
        /// Tipo de espesor
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TEspesor { get; set; }

        /// <summary>
        /// Código base
        /// </summary>
        [StringLength(16)]
        public string? CBase { get; set; }

        /// <summary>
        /// Vendedor representante
        /// </summary>
        [StringLength(8)]
        public string? VenRep { get; set; }

        /// <summary>
        /// Moneda de emisión
        /// </summary>
        public short? MoneEmis { get; set; }

        /// <summary>
        /// Tipo de cambio
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TipoCam { get; set; }

        /// <summary>
        /// Descuento de lista
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? DescLista { get; set; }

        /// <summary>
        /// Descuento combinado
        /// </summary>
        [StringLength(64)]
        public string? DescCombi { get; set; }

        /// <summary>
        /// Descuento de rubro
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? DescRubro { get; set; }

        /// <summary>
        /// Descuento total
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? DescTotal { get; set; }

        /// <summary>
        /// Número de serie
        /// </summary>
        [StringLength(32)]
        public string? NumeroSerie { get; set; }

        /// <summary>
        /// Valor observado 1
        /// </summary>
        [StringLength(256)]
        public string? VALOBADA_1 { get; set; }

        /// <summary>
        /// Identificación de lote
        /// </summary>
        [StringLength(32)]
        public string? IdenLote { get; set; }

        /// <summary>
        /// Lista de item
        /// </summary>
        public short? LISTA_DE_ITEM { get; set; }

        /// <summary>
        /// Valor observado 2
        /// </summary>
        [StringLength(256)]
        public string? VALOBADA_2 { get; set; }

        /// <summary>
        /// Valor observado 3
        /// </summary>
        [StringLength(256)]
        public string? VALOBADA_3 { get; set; }

        /// <summary>
        /// Número de proveedor
        /// </summary>
        public int? Num_Prov { get; set; }

        /// <summary>
        /// Observaciones del item
        /// </summary>
        [StringLength(512)]
        public string? ObservaItem { get; set; }

        /// <summary>
        /// Fecha de preparación
        /// </summary>
        public DateTime? FechPrepa { get; set; }

        /// <summary>
        /// Fecha informado
        /// </summary>
        public DateTime? FechInformado { get; set; }

        /// <summary>
        /// Fecha de retiro
        /// </summary>
        public DateTime? FechRetiro { get; set; }

        /// <summary>
        /// Sin cargo
        /// </summary>
        public short? SIN_CARGO { get; set; }

        /// <summary>
        /// Armador
        /// </summary>
        public short? Armador { get; set; }

        /// <summary>
        /// Número de armado
        /// </summary>
        public int? NumeroArmado { get; set; }

        /// <summary>
        /// Anotaciones del item
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? AnotaItem { get; set; }

        /// <summary>
        /// Fecha de último cambio
        /// </summary>
        public DateTime? FechUltCambio { get; set; }

        /// <summary>
        /// Marca de borrado
        /// </summary>
        public int? MarBorra { get; set; }

        /// <summary>
        /// Número de cotización
        /// </summary>
        [StringLength(16)]
        public string? NROCOTI { get; set; }

        /// <summary>
        /// Orden de fabricación
        /// </summary>
        [StringLength(18)]
        public string? OrdenFabricacion { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int? CodiEmprNet { get; set; } = 0;

        /// <summary>
        /// Identificador único global de la fila
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();

        [Display(Name = "Fecha de Creación")]
        public DateTime? AddRecord { get; set; }

        [Display(Name = "Última Actualización")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// Se actualiza automáticamente con el ID del usuario actual
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
-- Instrucciones SQL para modificar la tabla PEDDETA según los cambios realizados en la entidad:
ALTER TABLE peddeta 
ADD NumUsuar INT NULL;

-- 1. Cambiar CodiEmpr de smallint a int
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [CodiEmpr] int;

-- 2. Cambiar CodiInt de smallint a int
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [CodiInt] int;

-- 3. Cambiar todos los campos float a decimal(18,4)
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [PreList] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [CanAsig] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [CanPed] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [PreUnid] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [ImpoTot] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [ImpoAnti] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [CantEnt] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [CanSaldo] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [AcuPrepa] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [AcuDespa] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [PorComVend] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [TLargo] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [TAncho] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [TEspesor] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [TipoCam] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [DescLista] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [DescRubro] decimal(18,4);
ALTER TABLE [dbo].[PEDDETA] ALTER COLUMN [DescTotal] decimal(18,4);

-- 3. Verificar y mantener campo rowguid
-- NOTA: El campo rowguid debe mantenerse en la tabla
IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'PEDDETA' AND COLUMN_NAME = 'rowguid')
BEGIN
    ALTER TABLE [dbo].[PEDDETA] ADD [rowguid] UNIQUEIDENTIFIER ROWGUIDCOL NOT NULL DEFAULT NEWID();
END

-- NOTA: El campo PEDDETA_ID ya es IDENTITY en la tabla original, por lo que no requiere modificación.
*/

//ALTER TABLE PEDDETA
//ADD     AddRecord DATETIME2 NOT NULL DEFAULT GETDATE(),
//   LastUpdate DATETIME2 NOT NULL DEFAULT GETDATE(),
//   CodiEmprNet INT NOT NULL DEFAULT 0;


//  *****   COMPARACION DE DATOS ENTRE DOS TABLAS IGUALES Y DOS BASES DIFERENTES *****
//  *****   PARA COMPARAR SI SE GRABAN BIEN LOS MISMOS CAMPOS EN UNA TRANSACCION EN ARCON COMO EN ARCON2 *****

/*
 * CREATE OR ALTER PROCEDURE sp_Comparar_Tablas_Detalle
(
    @BaseA      SYSNAME,
    @BaseB      SYSNAME,
    @Tabla      SYSNAME,
    @Condicion  NVARCHAR(MAX),
    @Clave1     SYSNAME,
    @Clave2     SYSNAME
)
AS
BEGIN
    SET NOCOUNT ON;

DECLARE @SQL NVARCHAR(MAX) = '';

 Obtener columnas comunes con tipo 
DECLARE @SQLCols NVARCHAR(MAX) = '
        SELECT c1.name, t1.name AS tipo
        FROM ' + QUOTENAME(@BaseA) + '.sys.columns c1
        JOIN ' + QUOTENAME(@BaseA) + '.sys.types t1 ON c1.user_type_id = t1.user_type_id
        JOIN ' + QUOTENAME(@BaseB) + '.sys.columns c2 ON c1.name = c2.name
        WHERE c1.object_id = OBJECT_ID(''' + @BaseA + '.dbo.' + @Tabla + ''')
          AND c2.object_id = OBJECT_ID(''' + @BaseB + '.dbo.' + @Tabla + ''')
          AND c1.name NOT IN (''id'')';

    CREATE TABLE #Cols (name SYSNAME, tipo SYSNAME);
    INSERT INTO #Cols EXEC sp_executesql @SQLCols;

    SELECT @SQL = @SQL + '
    SELECT
        COALESCE(A.' + @Clave1 + ', B.' + @Clave1 + ') AS ' + @Clave1 + ',
    COALESCE(A.' + @Clave2 + ', B.' + @Clave2 + ') AS ' + @Clave2 + ',
        ''' + name + ''' AS campo,
    CAST(A.' + name + ' AS NVARCHAR(MAX)) AS valor_basea,
    CAST(B.' + name + ' AS NVARCHAR(MAX)) AS valor_baseb
    FROM ' + QUOTENAME(@BaseA) + '.dbo.' + QUOTENAME(@Tabla) + ' A
    FULL JOIN ' + QUOTENAME(@BaseB) + '.dbo.' + QUOTENAME(@Tabla) + ' B
        ON A.' + @Clave1 + ' = B.' + @Clave1 + '
       AND A.' + @Clave2 + ' = B.' + @Clave2 + '
    WHERE (' + @Condicion + ')
      AND ' +
      CASE 
        WHEN tipo IN ('int','decimal','numeric','float','real','money','smallmoney')
        THEN '
        ISNULL(TRY_CAST(A.' + name + ' AS DECIMAL(18,6)), 0)
        <>
        ISNULL(TRY_CAST(B.' + name + ' AS DECIMAL(18,6)), 0)'
        ELSE '
        ISNULL(CAST(A.' + name + ' AS NVARCHAR(MAX)), '''')
        <>
        ISNULL(CAST(B.' + name + ' AS NVARCHAR(MAX)), '''')'
      END + '
    UNION ALL'
    FROM #Cols;

    DROP TABLE #Cols;

    SET @SQL = LEFT(@SQL, LEN(@SQL) - 9);

EXEC sp_executesql @SQL;
END
GO*/


// *****  PARA EJECUTAR EL SP DE COMPARACION ******
/*
 * 	EXEC sp_Comparar_Tablas_Detalle
    @BaseA      = 'arcon',
    @BaseB      = 'arcon2',
    @Tabla      = 'peddeta',
    @Condicion  = 'A.nroped = 9800 OR B.nroped = 9600',
    @Clave1     = 'peddeta_id',
    @Clave2     = 'codiint';
*/