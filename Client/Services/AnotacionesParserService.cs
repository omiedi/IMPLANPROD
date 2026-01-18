using IMPLANPROD.Shared.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

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
        /// Parsea los productos de las líneas de anotación
        /// </summary>
        private List<ProductoAnotacionDTO> ParsearProductos(string[] lineas)
        {
            var productos = new List<ProductoAnotacionDTO>();

            foreach (var linea in lineas)
            {
                if (linea.Trim().StartsWith("- "))
                {
                    Console.WriteLine($"Parseando línea de producto: '{linea}'");
                    var producto = ParsearLineaProducto(linea);
                    if (producto != null)
                    {
                        Console.WriteLine($"Producto parseado: {producto.Codigo} - {producto.Descripcion} - {producto.Cantidad}");
                        productos.Add(producto);
                    }
                    else
                    {
                        Console.WriteLine($"No se pudo parsear la línea: '{linea}'");
                    }
                }
            }

            Console.WriteLine($"Total productos parseados: {productos.Count}");
            return productos;
        }

        /// <summary>
        /// Parsea una línea individual de producto
        /// Formato esperado: "- CODIGO       DESCRIPCION                         CANTIDAD"
        /// </summary>
        private ProductoAnotacionDTO? ParsearLineaProducto(string linea)
        {
            try
            {
                // Remover el "- " inicial
                var contenido = linea.Substring(2).Trim();

                // Usar regex para extraer código, descripción y cantidad
                // Patrón mejorado: código (espacios) descripción (espacios) cantidad (con comas y puntos)
                var patron = @"^(\S+)\s+(.+?)\s+([\d,\.]+)$";
                var match = Regex.Match(contenido, patron);

                if (match.Success)
                {
                    var cantidadTexto = match.Groups[3].Value.Replace(",", ".");
                    return new ProductoAnotacionDTO
                    {
                        Codigo = match.Groups[1].Value.Trim(),
                        Descripcion = match.Groups[2].Value.Trim(),
                        Cantidad = decimal.TryParse(cantidadTexto, out var cantidad) ? cantidad : 0
                    };
                }

                // Si el regex no funciona, intentar parsing manual
                var partes = contenido.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (partes.Length >= 3)
                {
                    var codigo = partes[0];
                    var cantidadTexto = partes[partes.Length - 1].Replace(",", ".");
                    var cantidad = decimal.TryParse(cantidadTexto, out var cant) ? cant : 0;
                    var descripcion = string.Join(" ", partes.Skip(1).Take(partes.Length - 2));

                    return new ProductoAnotacionDTO
                    {
                        Codigo = codigo,
                        Descripcion = descripcion,
                        Cantidad = cantidad
                    };
                }
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
