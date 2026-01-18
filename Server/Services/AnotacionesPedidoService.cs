using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;
using System.Text;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para manejar las anotaciones automáticas de pedidos
    /// Registra todas las actividades y cambios en el campo AnotaPed
    /// </summary>
    public class AnotacionesPedidoService
    {
        private readonly DataContext _context;
        private readonly ILogger<AnotacionesPedidoService> _logger;

        public AnotacionesPedidoService(DataContext context, ILogger<AnotacionesPedidoService> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Agrega una nueva anotación al pedido
        /// Equivalente a AGREGA_ANOTAPED_NEW en VB6
        /// </summary>
        /// <param name="nroPedido">Número del pedido</param>
        /// <param name="texto">Texto de la anotación</param>
        /// <param name="usuario">Usuario que realiza la acción (opcional)</param>
        /// <param name="fecha">Fecha específica (opcional, por defecto usa fecha actual)</param>
        public async Task AgregarAnotacionAsync(int nroPedido, string texto, string? usuario = null, DateTime? fecha = null)
        {
            try
            {
                if (nroPedido <= 0 || string.IsNullOrWhiteSpace(texto))
                    return;

                // Usar fecha actual si no se especifica
                var fechaAnotacion = fecha ?? DateTime.Now;
                var fechaTexto = fechaAnotacion.ToString("dd/MM/yy");
                var usuarioTexto = usuario ?? "SISTEMA";

                // Fraccionar el texto en líneas de máximo 56 caracteres
                var lineasTexto = FraccionarTexto(texto, 56);
                
                var anotacionCompleta = new StringBuilder();

                // Procesar cada línea fraccionada
                for (int i = 0; i < lineasTexto.Count; i++)
                {
                    var lineaFormateada = new string(' ', 96); // Equivalente a Space$(96)
                    var caracteres = lineaFormateada.ToCharArray();

                    // Si es la primera línea, agregar fecha y usuario
                    if (i == 0)
                    {
                        // Fecha en posición 1-8
                        ReemplazarEnPosicion(caracteres, fechaTexto, 0, 8);
                        // Usuario en posición 69-92 (24 caracteres)
                        ReemplazarEnPosicion(caracteres, usuarioTexto, 68, 24);
                    }

                    // Texto principal en posición 11-66 (56 caracteres)
                    ReemplazarEnPosicion(caracteres, lineasTexto[i], 10, 56);

                    anotacionCompleta.AppendLine(new string(caracteres));
                }

                // Agregar separador
                anotacionCompleta.AppendLine(new string('=', 88));

                // Obtener el pedido y actualizar las anotaciones
                var pedido = await _context.Pedencas
                    .FirstOrDefaultAsync(p => p.NroPed == nroPedido);

                if (pedido != null)
                {
                    var anotacionExistente = pedido.AnotaPed ?? string.Empty;
                    pedido.AnotaPed = anotacionExistente + " " + anotacionCompleta.ToString();
                    
                    await _context.SaveChangesAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al agregar anotación al pedido {NroPedido}", nroPedido);
            }
        }

        /// <summary>
        /// Registra los cambios en el detalle del pedido
        /// Equivalente al procesamiento del FOR en VB6
        /// </summary>
        /// <param name="nroPedido">Número del pedido</param>
        /// <param name="detalles">Lista de detalles del pedido</param>
        /// <param name="accion">Acción realizada (CREADO, EDITADO, etc.)</param>
        public async Task RegistrarCambiosDetalleAsync(int nroPedido, List<Peddeta> detalles, string accion = "EDITADO")
        {
            try
            {
                var textoDetalle = new StringBuilder();
                textoDetalle.AppendLine($"=== PEDIDO {accion} ===");

                foreach (var detalle in detalles)
                {
                    var lineaDetalle = new string(' ', 76); // Equivalente a Space$(76)
                    var caracteres = lineaDetalle.ToCharArray();

                    // Código externo en posición 11-28 (18 caracteres)
                    var codigo = detalle.CodiExt ?? detalle.CodiInt?.ToString() ?? "";
                    ReemplazarEnPosicion(caracteres, "- " + codigo, 10, 18);

                    // Descripción en posición 29-66 (38 caracteres)
                    var descripcion = detalle.Descrip ?? "";
                    ReemplazarEnPosicion(caracteres, descripcion, 28, 38);

                    // Cantidad en posición 57-66 (10 caracteres)
                    var cantidad = (detalle.CanPed ?? 0).ToString("N2");
                    ReemplazarEnPosicion(caracteres, cantidad, 56, 10);

                    textoDetalle.AppendLine(new string(caracteres));
                }

                await AgregarAnotacionAsync(nroPedido, textoDetalle.ToString());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar cambios de detalle del pedido {NroPedido}", nroPedido);
            }
        }

        /// <summary>
        /// Registra solo los cambios reales comparando detalle anterior vs nuevo
        /// </summary>
        /// <param name="nroPedido">Número del pedido</param>
        /// <param name="detalleAnterior">Detalle antes de los cambios</param>
        /// <param name="detalleNuevo">Detalle después de los cambios</param>
        public async Task RegistrarCambiosRealesAsync(int nroPedido, List<Peddeta> detalleAnterior, List<Peddeta> detalleNuevo)
        {
            try
            {
                _logger.LogInformation("Iniciando comparación de cambios para pedido {NroPedido}. Items anteriores: {Anteriores}, Items nuevos: {Nuevos}", 
                    nroPedido, detalleAnterior.Count, detalleNuevo.Count);

                var cambiosDetectados = new List<string>();
                
                // Crear diccionarios para comparación por NuOrItem
                var anteriorDict = detalleAnterior.ToDictionary(d => d.NuOrItem ?? 0);
                var nuevoDict = detalleNuevo.ToDictionary(d => d.NuOrItem ?? 0);

                // Detectar items modificados
                foreach (var itemNuevo in detalleNuevo)
                {
                    var nuOrItem = itemNuevo.NuOrItem ?? 0;
                    
                    if (anteriorDict.ContainsKey(nuOrItem))
                    {
                        var itemAnterior = anteriorDict[nuOrItem];
                        
                        // Comparar campos relevantes
                        var cantidadCambio = itemAnterior.CanPed != itemNuevo.CanPed;
                        var precioCambio = itemAnterior.PreUnid != itemNuevo.PreUnid;
                        var descripcionCambio = itemAnterior.Descrip != itemNuevo.Descrip;

                        _logger.LogInformation("Item {NuOrItem} - Cantidad: {CantAnterior} -> {CantNueva} (cambió: {CantCambio}), " +
                            "Precio: {PrecioAnterior} -> {PrecioNuevo} (cambió: {PrecioCambio}), " +
                            "Descripción cambió: {DescCambio}", 
                            nuOrItem, itemAnterior.CanPed, itemNuevo.CanPed, cantidadCambio,
                            itemAnterior.PreUnid, itemNuevo.PreUnid, precioCambio, descripcionCambio);

                        if (cantidadCambio || precioCambio || descripcionCambio)
                        {
                            var lineaDetalle = new string(' ', 76);
                            var caracteres = lineaDetalle.ToCharArray();

                            var codigo = itemNuevo.CodiExt ?? itemNuevo.CodiInt?.ToString() ?? "";
                            ReemplazarEnPosicion(caracteres, "- " + codigo, 10, 18);

                            var descripcion = itemNuevo.Descrip ?? "";
                            ReemplazarEnPosicion(caracteres, descripcion, 28, 38);

                            var cantidad = (itemNuevo.CanPed ?? 0).ToString("N2");
                            ReemplazarEnPosicion(caracteres, cantidad, 56, 10);

                            cambiosDetectados.Add(new string(caracteres));
                            _logger.LogInformation("Cambio detectado para item {NuOrItem}: {Codigo}", nuOrItem, codigo);
                        }
                    }
                    else
                    {
                        // Item nuevo agregado
                        var lineaDetalle = new string(' ', 76);
                        var caracteres = lineaDetalle.ToCharArray();

                        var codigo = itemNuevo.CodiExt ?? itemNuevo.CodiInt?.ToString() ?? "";
                        ReemplazarEnPosicion(caracteres, "+ " + codigo, 10, 18);

                        var descripcion = itemNuevo.Descrip ?? "";
                        ReemplazarEnPosicion(caracteres, descripcion, 28, 38);

                        var cantidad = (itemNuevo.CanPed ?? 0).ToString("N2");
                        ReemplazarEnPosicion(caracteres, cantidad, 56, 10);

                        cambiosDetectados.Add(new string(caracteres));
                    }
                }

                // Detectar items eliminados
                foreach (var itemAnterior in detalleAnterior)
                {
                    var nuOrItem = itemAnterior.NuOrItem ?? 0;
                    
                    if (!nuevoDict.ContainsKey(nuOrItem))
                    {
                        var lineaDetalle = new string(' ', 76);
                        var caracteres = lineaDetalle.ToCharArray();

                        var codigo = itemAnterior.CodiExt ?? itemAnterior.CodiInt?.ToString() ?? "";
                        ReemplazarEnPosicion(caracteres, "X " + codigo, 10, 18);

                        var descripcion = itemAnterior.Descrip ?? "";
                        ReemplazarEnPosicion(caracteres, descripcion, 28, 38);

                        ReemplazarEnPosicion(caracteres, "ELIMINADO", 56, 10);

                        cambiosDetectados.Add(new string(caracteres));
                    }
                }

                // Solo registrar si hay cambios reales
                _logger.LogInformation("Cambios detectados: {CantidadCambios}", cambiosDetectados.Count);
                
                if (cambiosDetectados.Any())
                {
                    var textoDetalle = new StringBuilder();
                    textoDetalle.AppendLine("=== PEDIDO EDITADO ===");
                    foreach (var cambio in cambiosDetectados)
                    {
                        textoDetalle.AppendLine(cambio);
                    }

                    _logger.LogInformation("Registrando anotación para pedido {NroPedido} con {CantidadCambios} cambios", 
                        nroPedido, cambiosDetectados.Count);
                    await AgregarAnotacionAsync(nroPedido, textoDetalle.ToString());
                }
                else
                {
                    _logger.LogInformation("No se detectaron cambios reales para el pedido {NroPedido}", nroPedido);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar cambios reales del pedido {NroPedido}", nroPedido);
            }
        }

        /// <summary>
        /// Registra una entrega de cantidades
        /// </summary>
        public async Task RegistrarEntregaAsync(int nroPedido, List<Peddeta> itemsEntregados)
        {
            try
            {
                var textoEntrega = new StringBuilder();
                textoEntrega.AppendLine("=== ENTREGA REGISTRADA ===");

                foreach (var item in itemsEntregados.Where(i => (i.CantEnt ?? 0) > 0))
                {
                    textoEntrega.AppendLine($"- {item.CodiExt ?? item.CodiInt?.ToString()}: {item.Descrip}");
                    textoEntrega.AppendLine($"  Cantidad entregada: {item.CantEnt:N2}");
                    textoEntrega.AppendLine($"  Saldo pendiente: {item.CanSaldo:N2}");
                }

                await AgregarAnotacionAsync(nroPedido, textoEntrega.ToString());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar entrega del pedido {NroPedido}", nroPedido);
            }
        }

        /// <summary>
        /// Fracciona un texto en líneas de longitud máxima especificada
        /// Equivalente a FRATEXTO en VB6
        /// </summary>
        private List<string> FraccionarTexto(string texto, int longitudMaxima)
        {
            var lineas = new List<string>();
            
            if (string.IsNullOrEmpty(texto))
                return lineas;

            var palabras = texto.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var lineaActual = new StringBuilder();

            foreach (var palabra in palabras)
            {
                if (lineaActual.Length + palabra.Length + 1 <= longitudMaxima)
                {
                    if (lineaActual.Length > 0)
                        lineaActual.Append(" ");
                    lineaActual.Append(palabra);
                }
                else
                {
                    if (lineaActual.Length > 0)
                    {
                        lineas.Add(lineaActual.ToString());
                        lineaActual.Clear();
                    }
                    lineaActual.Append(palabra);
                }
            }

            if (lineaActual.Length > 0)
                lineas.Add(lineaActual.ToString());

            return lineas;
        }

        /// <summary>
        /// Reemplaza texto en una posición específica del array de caracteres
        /// Equivalente a REPLA en VB6
        /// </summary>
        private void ReemplazarEnPosicion(char[] destino, string texto, int posicionInicio, int longitudMaxima)
        {
            if (string.IsNullOrEmpty(texto) || posicionInicio < 0 || posicionInicio >= destino.Length)
                return;

            var textoAInsertar = texto.Length > longitudMaxima ? texto.Substring(0, longitudMaxima) : texto;
            
            for (int i = 0; i < textoAInsertar.Length && (posicionInicio + i) < destino.Length; i++)
            {
                destino[posicionInicio + i] = textoAInsertar[i];
            }
        }
    }
}
