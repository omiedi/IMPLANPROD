---
description: CLAVE: AGREGAR IMPRESION
---




Cuando solicito agregar impresion a una tabla tipo html, debes tomar el ejemplo de la page PedidosVentaForm.razor agregando el boton de impresion y el servicio de impresion. Imprimir/Excel > presentar el modal <ImpresionLista TItem="Pedenca"
                Datos="pedidos"
                Titulo="@TituloReporte"
                UrlExportarExcel="api/excel/exportar-pedidos"
                Encabezados="encabezadosPedidos"
                ModoDirecto="false"
                MostrarModal="mostrarModalImpresion"
                OnCerrar="ResetearImpresion" />  utilizar el componente C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Client\Pages\Reportes\ImpresionLista.razor  usando esta js C:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Client\wwwroot\js\utilities.js  
                la impresion siempre debe mostrar las columnas de pantalla y respetar los filtros
  tambien se debe crear en excelContoller     el post como este ejemplo /// <summary>
/// Exporta pedidos de venta a Excel
/// </summary>
/// <param name="datos">Lista de pedidos a exportar</param>
/// <returns>Archivo Excel con los pedidos</returns>
[HttpPost("exportar-pedidos")]
public IActionResult ExportarPedidos([FromBody] List<Pedenca> datos)
{
    try
    {
        var encabezados = new Dictionary<string, string>
        {
            { "NroPed", "N° Pedido" },
            { "FechEmis", "Fecha Emisión" },
            { "RasoClie", "Cliente" },
            { "NroClie", "N° Cliente" },
            { "NroOcom", "N° OC" },
            { "Status", "Estado" },
            { "FeSolEnt", "Fecha Entrega" },
            { "VenRep", "Vendedor" },
            { "Vended", "Cod. Vendedor" },
            { "Referencia", "Referencia" },
            { "TranspEnt", "Transporte" },
            { "NTranspo", "N° Transporte" },
            { "ListPre", "Lista Precio" },
            { "CondPag", "Cond. Pago" },
            { "PorComVend", "% Comisión" }
        };

        var orden = new List<string>
        {
            "NroPed", "FechEmis", "RasoClie", "NroClie", "NroOcom", 
            "Status", "FeSolEnt", "VenRep", "Vended", "Referencia",
            "TranspEnt", "NTranspo", "ListPre", "CondPag", "PorComVend"
        };

        var archivo = _excelService.GenerarExcel(datos, "Pedidos de Venta", encabezados, orden);

        return File(archivo,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"Pedidos_Venta_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "Error al generar Excel de Pedidos de Venta");
        return StatusCode(500, "Error al generar el archivo excel");
    }
}
         