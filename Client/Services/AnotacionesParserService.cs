using IMPLANPROD.Shared.DTOs;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio para parsear las anotaciones de pedidos desde el formato VB6 (ntext)
    /// Mantiene compatibilidad con la estructura original de posiciones fijas
    /// </summary>
    public class AnotacionesParserService
    {
        /// <summary>
        /// Parsea el texto de anotaciones desde el formato VB6 a una estructura organizada
        /// </summary>
        /// <param name="textoAnotaciones">Texto completo de anotaciones desde pedenca.anotaped</param>
        /// <returns>Lista de anotaciones parseadas</returns>
        public List<AnotacionPedidoDTO> ParsearAnotaciones(string textoAnotaciones)
        {
            var anotaciones = new List<AnotacionPedidoDTO>();
            
            if (string.IsNullOrWhiteSpace(textoAnotaciones))
                return anotaciones;

            // Dividir por separadores de anotaciones (líneas de iguales)
            var bloques = textoAnotaciones.Split(new[] { "========================================================================================" }, 
                StringSplitOptions.RemoveEmptyEntries);

            foreach (var bloque in bloques)
            {
                if (string.IsNullOrWhiteSpace(bloque)) continue;

                var anotacion = ParsearBloqueAnotacion(bloque.Trim());
                if (anotacion != null)
                {
                    anotaciones.Add(anotacion);
                }
            }

            return anotaciones;
        }

        /// <summary>
        /// Parsea un bloque individual de anotación
        /// </summary>
        private AnotacionPedidoDTO? ParsearBloqueAnotacion(string bloque)
        {
            try
            {
                var lineas = bloque.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (lineas.Length == 0) return null;

                var anotacion = new AnotacionPedidoDTO
                {
                    TextoOriginal = bloque
                };

                // Parsear la primera línea que contiene fecha, descripción y usuario
                var primeraLinea = lineas[0];
                if (primeraLinea.Length >= 8)
                {
                    // Extraer fecha (posiciones 1-8)
                    anotacion.Fecha = primeraLinea.Substring(0, Math.Min(8, primeraLinea.Length)).Trim();
                }

                if (primeraLinea.Length >= 69)
                {
                    // Extraer usuario (posición 69, 24 caracteres)
                    var inicioUsuario = Math.Min(68, primeraLinea.Length - 1); // Posición 69 = índice 68
                    var longitudUsuario = Math.Min(24, primeraLinea.Length - inicioUsuario);
                    anotacion.Usuario = primeraLinea.Substring(inicioUsuario, longitudUsuario).Trim();
                }

                if (primeraLinea.Length >= 11)
                {
                    // Extraer descripción (posiciones 11-56)
                    var inicioDescripcion = 10; // Posición 11 = índice 10
                    var finDescripcion = Math.Min(66, primeraLinea.Length); // Posición 56 = hasta índice 66
                    if (primeraLinea.Length >= 69)
                    {
                        finDescripcion = Math.Min(68, primeraLinea.Length); // Hasta antes del usuario
                    }
                    
                    if (finDescripcion > inicioDescripcion)
                    {
                        anotacion.Descripcion = primeraLinea.Substring(inicioDescripcion, finDescripcion - inicioDescripcion).Trim();
                    }
                }

                // Parsear productos (líneas que empiezan con "- ")
                anotacion.Productos = ParsearProductos(lineas);

                // Parsear observaciones
                anotacion.Observaciones = ParsearObservaciones(lineas);

                return anotacion;
            }
            catch (Exception)
            {
                // Si hay error en el parsing, devolver anotación básica con texto original
                return new AnotacionPedidoDTO
                {
                    TextoOriginal = bloque,
                    Descripcion = "Error al parsear anotación"
                };
            }
        }

        /// <summary>
        /// Parsea los productos de las líneas de anotación.
        /// Detecta líneas de producto por los prefijos "- " (modificado), 
        /// "+ " (agregado) y "X " (eliminado) que usa el servicio de anotaciones.
        /// </summary>
        private List<ProductoAnotacionDTO> ParsearProductos(string[] lineas)
        {
            var productos = new List<ProductoAnotacionDTO>();

            foreach (var linea in lineas)
            {
                var lineaTrim = linea.Trim();
                if (lineaTrim.StartsWith("- ") || lineaTrim.StartsWith("+ ") || lineaTrim.StartsWith("X "))
                {
                    var producto = ParsearLineaProducto(linea);
                    if (producto != null)
                    {
                        productos.Add(producto);
                    }
                }
            }

            return productos;
        }

        /// <summary>
        /// Parsea una línea individual de producto.
        /// Formato VB6 fijo de 76 chars (FORPED07.frm línea 3631-3635):
        ///   Pos 11 (idx 10): prefijo "- "/"X "/"+ " + código (18 chars)
        ///   Pos 29 (idx 28): descripción (38 chars, los últimos 10 los pisa la cantidad)
        ///   Pos 57 (idx 56): cantidad (10 chars)
        /// Se usa extracción por posiciones fijas en lugar de regex porque el regex
        /// falla cuando la cantidad tiene formato cultural (comas/puntos) o cuando
        /// la descripción contiene números al final.
        /// </summary>
        private ProductoAnotacionDTO? ParsearLineaProducto(string linea)
        {
            try
            {
                // Buscar el marcador "- " (o "+ " / "X ") en la línea.
                // En VB6 está en idx 10, pero puede variar si la línea viene de
                // un bloque envuelto a 96 chars.
                int idxMarcador = -1;
                string prefijo = "";
                foreach (var pref in new[] { "- ", "+ ", "X " })
                {
                    var idx = linea.IndexOf(pref);
                    if (idx >= 0 && (idxMarcador < 0 || idx < idxMarcador))
                    {
                        idxMarcador = idx;
                        prefijo = pref;
                    }
                }

                if (idxMarcador < 0)
                    return null;

                // El código empieza 2 chars después del marcador.
                int idxCodigo = idxMarcador + 2;
                // La descripción empieza 18 chars después del marcador (pos 29 = idxMarcador + 18).
                int idxDescripcion = idxMarcador + 18;
                // La cantidad empieza 46 chars después del marcador (pos 57 = idxMarcador + 46).
                int idxCantidad = idxMarcador + 46;

                // Extraer código (16 chars)
                string codigo = string.Empty;
                if (idxCodigo < linea.Length)
                {
                    var finCodigo = Math.Min(idxCodigo + 16, linea.Length);
                    codigo = linea.Substring(idxCodigo, finCodigo - idxCodigo).Trim();
                }

                // Extraer descripción (28 chars, antes de la cantidad)
                string descripcion = string.Empty;
                if (idxDescripcion < linea.Length)
                {
                    var finDescripcion = Math.Min(idxDescripcion + 28, linea.Length);
                    descripcion = linea.Substring(idxDescripcion, finDescripcion - idxDescripcion).Trim();
                }

                // Extraer cantidad (10 chars)
                string cantidadTexto = string.Empty;
                if (idxCantidad < linea.Length)
                {
                    var finCantidad = Math.Min(idxCantidad + 10, linea.Length);
                    cantidadTexto = linea.Substring(idxCantidad, finCantidad - idxCantidad).Trim();
                }

                if (string.IsNullOrEmpty(codigo))
                    return null;

                // Parsear cantidad: aceptar coma o punto como separador decimal (compatibilidad VB6/C#)
                var cantidadNormalizada = cantidadTexto.Replace(",", ".");
                decimal cantidad = 0m;
                if (!string.IsNullOrEmpty(cantidadNormalizada) &&
                    !decimal.TryParse(cantidadNormalizada, NumberStyles.Any, CultureInfo.InvariantCulture, out cantidad))
                {
                    // Si no es numérico (ej: "ELIMINADO"), dejar 0
                    cantidad = 0m;
                }

                return new ProductoAnotacionDTO
                {
                    Codigo = codigo,
                    Descripcion = descripcion,
                    Cantidad = cantidad
                };
            }
            catch
            {
                // Ignorar errores de parsing de productos individuales
            }

            return null;
        }

        /// <summary>
        /// Parsea las observaciones de las líneas de anotación
        /// </summary>
        private string ParsearObservaciones(string[] lineas)
        {
            var observaciones = new List<string>();
            bool enObservaciones = false;

            foreach (var linea in lineas)
            {
                if (linea.Trim().StartsWith("Observaciones:"))
                {
                    enObservaciones = true;
                    var obs = linea.Trim().Substring("Observaciones:".Length).Trim();
                    if (!string.IsNullOrEmpty(obs))
                    {
                        observaciones.Add(obs);
                    }
                }
                else if (enObservaciones && !linea.Trim().StartsWith("---"))
                {
                    var obs = linea.Trim();
                    if (!string.IsNullOrEmpty(obs))
                    {
                        observaciones.Add(obs);
                    }
                }
                else if (linea.Trim().StartsWith("---"))
                {
                    enObservaciones = false;
                }
            }

            return string.Join(" ", observaciones);
        }
    }
}
