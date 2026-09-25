using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
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
        /// Persiste los items del pedido en la tabla ANOTAPED de forma estructurada.
        /// Reemplaza el guardado en el campo ntext PEDENCA.AnotaPed por registros individuales.
        /// Todos los items de un mismo update comparten el mismo IdAnotacion (incremental por pedido).
        /// </summary>
        /// <param name="nroPedido">Número de pedido (PEDENCA.NroPed)</param>
        /// <param name="detalles">Lista de items del pedido (PEDDETA)</param>
        /// <param name="observaGral">Observaciones generales del pedido (PEDENCA.Observa), truncado a 1024</param>
        /// <param name="nombreUsuario">Nombre del usuario activo</param>
        /// <param name="numUsuar">NumUsuar del usuario (usuarios.idflexoft)</param>
        /// <param name="accion">Descripción de la acción (CREADO, EDITADO, etc.)</param>
        public async Task GuardarEnAnotapedAsync(
            int nroPedido,
            List<Peddeta> detalles,
            string? observaGral,
            string? nombreUsuario,
            int? numUsuar,
            string accion = "CREADO")
        {
            try
            {
                if (nroPedido <= 0 || detalles == null || !detalles.Any())
                    return;

                // Calcular el siguiente IdAnotacion para este pedido (incremental por pedido)
                var maxIdAnotacion = await _context.Anotapeds
                    .Where(a => a.NroPed == nroPedido)
                    .MaxAsync(a => (int?)a.IdAnotacion) ?? 0;
                var nuevoIdAnotacion = maxIdAnotacion + 1;

                // Truncar ObservaGral a 1024 caracteres
                var observaGralTruncado = string.IsNullOrEmpty(observaGral)
                    ? null
                    : observaGral.Length > 1024 ? observaGral.Substring(0, 1024) : observaGral;

                var ahora = DateTime.Now;
                var usuarioTexto = string.IsNullOrEmpty(nombreUsuario) ? "SISTEMA" : nombreUsuario;

                var registros = new List<Anotaped>();
                foreach (var det in detalles)
                {
                    var anotaped = new Anotaped
                    {
                        IdAnotacion = nuevoIdAnotacion,
                        NroPed = nroPedido,
                        CodiInt = det.CodiInt ?? 0,
                        CodiExt = det.CodiExt,
                        Descrip = det.Descrip,
                        CanPed = det.CanPed,
                        CantEnt = det.CantEnt,
                        ObservaItem = det.ObservaItem,
                        ObservaGral = observaGralTruncado,
                        FechaCambio = ahora,
                        Usuario = usuarioTexto,
                        CodiEmpr = 0,
                        AddRecord = ahora,
                        LastUpdate = ahora,
                        CodiEmprNet = 0,
                        NumUsuar = numUsuar ?? 0
                    };
                    registros.Add(anotaped);
                }

                _context.Anotapeds.AddRange(registros);
                await _context.SaveChangesAsync();

                _logger.LogInformation(
                    "Guardados {Cantidad} registros en ANOTAPED para pedido {NroPed}, IdAnotacion={IdAnotacion}, acción={Accion}",
                    registros.Count, nroPedido, nuevoIdAnotacion, accion);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al guardar en ANOTAPED para pedido {NroPedido}", nroPedido);
            }
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
        /// Equivalente al procesamiento del FOR en VB6 (FORPED07.frm línea 3631-3635).
        /// Las líneas de producto se guardan directamente en formato VB6 de 76 chars
        /// (Space$(76) + REPLA en posiciones 11, 29 y 57) SIN pasar por AgregarAnotacionAsync,
        /// porque esa función re-envuelve el texto a 56 chars y destruye el formato fijo.
        /// </summary>
        /// <param name="nroPedido">Número del pedido</param>
        /// <param name="detalles">Lista de detalles del pedido</param>
        /// <param name="accion">Acción realizada (CREADO, EDITADO, etc.)</param>
        public async Task RegistrarCambiosDetalleAsync(int nroPedido, List<Peddeta> detalles, string accion = "EDITADO")
        {
            try
            {
                if (nroPedido <= 0 || detalles == null || !detalles.Any())
                    return;

                var fechaTexto = DateTime.Now.ToString("dd/MM/yy");
                var usuarioTexto = "SISTEMA";

                var anotacionCompleta = new StringBuilder();

                // Línea de encabezado: 96 chars con fecha + descripción + usuario (igual a VB6)
                var lineaHeader = new string(' ', 96).ToCharArray();
                ReemplazarEnPosicion(lineaHeader, fechaTexto, 0, 8);
                ReemplazarEnPosicion(lineaHeader, accion, 10, 56);
                ReemplazarEnPosicion(lineaHeader, usuarioTexto, 68, 24);
                anotacionCompleta.AppendLine(new string(lineaHeader));

                // Líneas de producto: 76 chars en formato VB6 fijo
                // Pos 11 (idx 10): "- " + código (18 chars)
                // Pos 29 (idx 28): descripción (38 chars, los últimos 10 los pisa la cantidad)
                // Pos 57 (idx 56): cantidad (10 chars)
                foreach (var detalle in detalles)
                {
                    var lineaDetalle = new string(' ', 76).ToCharArray();

                    var codigo = detalle.CodiExt ?? detalle.CodiInt?.ToString() ?? "";
                    ReemplazarEnPosicion(lineaDetalle, "- " + codigo, 10, 18);

                    var descripcion = detalle.Descrip ?? "";
                    ReemplazarEnPosicion(lineaDetalle, descripcion, 28, 38);

                    var cantidad = (detalle.CanPed ?? 0).ToString("F2", CultureInfo.InvariantCulture);
                    ReemplazarEnPosicion(lineaDetalle, cantidad, 56, 10);

                    anotacionCompleta.AppendLine(new string(lineaDetalle));
                }

                // Separador: 88 "=" (igual a VB6)
                anotacionCompleta.AppendLine(new string('=', 88));

                // Guardar directamente en AnotaPed, sin pasar por AgregarAnotacionAsync
                var pedido = await _context.Pedencas.FirstOrDefaultAsync(p => p.NroPed == nroPedido);
                if (pedido != null)
                {
                    var anotacionExistente = pedido.AnotaPed ?? string.Empty;
                    pedido.AnotaPed = anotacionExistente + " " + anotacionCompleta.ToString();
                    await _context.SaveChangesAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar cambios de detalle del pedido {NroPedido}", nroPedido);
            }
        }

        /// <summary>
        /// Registra solo los cambios reales comparando detalle anterior vs nuevo.
        /// Las líneas de producto se guardan directamente en formato VB6 de 76 chars
        /// SIN pasar por AgregarAnotacionAsync para evitar el re-envoltorio a 56 chars.
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

                var cambiosDetectados = new List<(string Prefijo, string Codigo, string Descripcion, string Cantidad)>();
                
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
                            var codigo = itemNuevo.CodiExt ?? itemNuevo.CodiInt?.ToString() ?? "";
                            var descripcion = itemNuevo.Descrip ?? "";
                            var cantidad = (itemNuevo.CanPed ?? 0).ToString("F2", CultureInfo.InvariantCulture);
                            cambiosDetectados.Add(("- ", codigo, descripcion, cantidad));
                            _logger.LogInformation("Cambio detectado para item {NuOrItem}: {Codigo}", nuOrItem, codigo);
                        }
                    }
                    else
                    {
                        // Item nuevo agregado
                        var codigo = itemNuevo.CodiExt ?? itemNuevo.CodiInt?.ToString() ?? "";
                        var descripcion = itemNuevo.Descrip ?? "";
                        var cantidad = (itemNuevo.CanPed ?? 0).ToString("F2", CultureInfo.InvariantCulture);
                        cambiosDetectados.Add(("+ ", codigo, descripcion, cantidad));
                    }
                }

                // Detectar items eliminados
                foreach (var itemAnterior in detalleAnterior)
                {
                    var nuOrItem = itemAnterior.NuOrItem ?? 0;
                    
                    if (!nuevoDict.ContainsKey(nuOrItem))
                    {
                        var codigo = itemAnterior.CodiExt ?? itemAnterior.CodiInt?.ToString() ?? "";
                        var descripcion = itemAnterior.Descrip ?? "";
                        cambiosDetectados.Add(("X ", codigo, descripcion, "ELIMINADO"));
                    }
                }

                // Solo registrar si hay cambios reales
                _logger.LogInformation("Cambios detectados: {CantidadCambios}", cambiosDetectados.Count);
                
                if (cambiosDetectados.Any())
                {
                    var fechaTexto = DateTime.Now.ToString("dd/MM/yy");
                    var usuarioTexto = "SISTEMA";

                    var anotacionCompleta = new StringBuilder();

                    // Línea de encabezado: 96 chars con fecha + "PEDIDO EDITADO" + usuario
                    var lineaHeader = new string(' ', 96).ToCharArray();
                    ReemplazarEnPosicion(lineaHeader, fechaTexto, 0, 8);
                    ReemplazarEnPosicion(lineaHeader, "PEDIDO EDITADO", 10, 56);
                    ReemplazarEnPosicion(lineaHeader, usuarioTexto, 68, 24);
                    anotacionCompleta.AppendLine(new string(lineaHeader));

                    // Líneas de producto: 76 chars en formato VB6 fijo
                    foreach (var cambio in cambiosDetectados)
                    {
                        var lineaDetalle = new string(' ', 76).ToCharArray();
                        ReemplazarEnPosicion(lineaDetalle, cambio.Prefijo + cambio.Codigo, 10, 18);
                        ReemplazarEnPosicion(lineaDetalle, cambio.Descripcion, 28, 38);
                        ReemplazarEnPosicion(lineaDetalle, cambio.Cantidad, 56, 10);
                        anotacionCompleta.AppendLine(new string(lineaDetalle));
                    }

                    // Separador: 88 "="
                    anotacionCompleta.AppendLine(new string('=', 88));

                    _logger.LogInformation("Registrando anotación para pedido {NroPedido} con {CantidadCambios} cambios", 
                        nroPedido, cambiosDetectados.Count);

                    // Guardar directamente en AnotaPed, sin pasar por AgregarAnotacionAsync
                    var pedido = await _context.Pedencas.FirstOrDefaultAsync(p => p.NroPed == nroPedido);
                    if (pedido != null)
                    {
                        var anotacionExistente = pedido.AnotaPed ?? string.Empty;
                        pedido.AnotaPed = anotacionExistente + " " + anotacionCompleta.ToString();
                        await _context.SaveChangesAsync();
                    }
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
        /// Registra la anotación automática cuando se marcan ítems como cumplidos
        /// desde la utilidad "Anular Items de Pedido" de /pedidos-venta.
        /// Escribe en los dos soportes de anotaciones del pedido:
        ///  1) PEDENCA.AnotaPed (vista texto): bloque con título "ITEMS MARCADOS COMO
        ///     CUMPLIDOS" y una línea por ítem en el formato fijo de producto
        ///     (76 chars: "- " + código en pos 11, descripción en pos 29 y cantidad
        ///     cumplida = saldo pendiente en pos 57), que la vista texto parsea y
        ///     muestra como tabla de productos (detalle resaltado en negrita).
        ///  2) ANOTAPED (vista tabla): un registro por ítem compartiendo un nuevo
        ///     IdAnotacion, con ObservaGral = "Items Marcados Como Cumplidos" como
        ///     título del bloque.
        /// </summary>
        /// <param name="nroPedido">Número de pedido (PEDENCA.NroPed)</param>
        /// <param name="itemsCumplidos">Ítems marcados como cumplidos (con sus cantidades)</param>
        /// <param name="nombreUsuario">Nombre del usuario activo</param>
        /// <param name="numUsuar">NumUsuar del usuario (usuarios.idflexoft)</param>
        public async Task RegistrarItemsCumplidosAsync(int nroPedido, List<Peddeta> itemsCumplidos, string? nombreUsuario, int? numUsuar)
        {
            try
            {
                if (nroPedido <= 0 || itemsCumplidos == null || !itemsCumplidos.Any())
                    return;

                var ahora = DateTime.Now;
                var usuarioTexto = string.IsNullOrEmpty(nombreUsuario) ? "SISTEMA" : nombreUsuario;
                const string tituloAnotacion = "Items Marcados Como Cumplidos";

                // ── 1) Vista texto: bloque en PEDENCA.AnotaPed ──────────────────────
                var anotacionCompleta = new StringBuilder();

                // Línea de encabezado de 96 chars: fecha (1-8) + título (11-66) + usuario (69-92)
                var lineaHeader = new string(' ', 96).ToCharArray();
                ReemplazarEnPosicion(lineaHeader, ahora.ToString("dd/MM/yy"), 0, 8);
                ReemplazarEnPosicion(lineaHeader, tituloAnotacion.ToUpper(), 10, 56);
                ReemplazarEnPosicion(lineaHeader, usuarioTexto, 68, 24);
                anotacionCompleta.AppendLine(new string(lineaHeader));

                // Líneas de ítem: formato fijo de 76 chars igual al resto de las
                // anotaciones de producto, con prefijo "- " para que el parser de la
                // vista texto las reconozca y las muestre en la tabla de productos
                // (código y cantidad en negrita, resaltando el detalle).
                //   pos 11 (idx 10): "- " + código (18 chars)
                //   pos 29 (idx 28): descripción (38 chars)
                //   pos 57 (idx 56): cantidad cumplida = saldo pendiente (10 chars)
                foreach (var item in itemsCumplidos)
                {
                    var lineaItem = new string(' ', 76).ToCharArray();
                    var codigo = item.CodiExt ?? item.CodiInt?.ToString() ?? "";
                    var saldo = (item.CanPed ?? 0m) - (item.CantEnt ?? 0m);

                    ReemplazarEnPosicion(lineaItem, "- " + codigo, 10, 18);
                    ReemplazarEnPosicion(lineaItem, item.Descrip ?? "", 28, 38);
                    ReemplazarEnPosicion(lineaItem, saldo.ToString("F1", CultureInfo.InvariantCulture).PadLeft(10), 56, 10);

                    anotacionCompleta.AppendLine(new string(lineaItem));
                }

                // Separador de bloque: 88 "=" (igual que el resto de las anotaciones)
                anotacionCompleta.AppendLine(new string('=', 88));

                var pedido = await _context.Pedencas.FirstOrDefaultAsync(p => p.NroPed == nroPedido);
                if (pedido != null)
                {
                    pedido.AnotaPed = (pedido.AnotaPed ?? string.Empty) + " " + anotacionCompleta.ToString();
                }

                // ── 2) Vista tabla: registros ANOTAPED ──────────────────────────────
                // Todos los ítems comparten el mismo IdAnotacion (incremental por pedido)
                // y ObservaGral lleva el título, que la vista muestra como
                // "Observación General" del grupo.
                var maxIdAnotacion = await _context.Anotapeds
                    .Where(a => a.NroPed == nroPedido)
                    .MaxAsync(a => (int?)a.IdAnotacion) ?? 0;
                var nuevoIdAnotacion = maxIdAnotacion + 1;

                var registros = itemsCumplidos.Select(item => new Anotaped
                {
                    IdAnotacion = nuevoIdAnotacion,
                    NroPed = nroPedido,
                    CodiInt = item.CodiInt ?? 0,
                    CodiExt = item.CodiExt,
                    Descrip = item.Descrip,
                    CanPed = item.CanPed,
                    CantEnt = item.CantEnt,
                    ObservaItem = item.ObservaItem,
                    ObservaGral = tituloAnotacion,
                    FechaCambio = ahora,
                    Usuario = usuarioTexto,
                    CodiEmpr = 0,
                    AddRecord = ahora,
                    LastUpdate = ahora,
                    CodiEmprNet = 0,
                    NumUsuar = numUsuar ?? 0
                }).ToList();

                _context.Anotapeds.AddRange(registros);
                await _context.SaveChangesAsync();

                _logger.LogInformation(
                    "Anotación '{Titulo}' registrada para pedido {NroPed}: {Cantidad} ítems (IdAnotacion={IdAnotacion})",
                    tituloAnotacion, nroPedido, registros.Count, nuevoIdAnotacion);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar anotación de ítems cumplidos del pedido {NroPedido}", nroPedido);
            }
        }

        /// <summary>
        /// Registra la anotación automática "Items Remitidos" cuando se genera un
        /// remito de cliente que está asociado a un pedido (MOVISTO.NuPedCli &gt; 0).
        /// Mantiene el mismo criterio implementado para los ítems marcados como
        /// cumplidos: escribe en los dos soportes de anotaciones del pedido:
        ///  1) PEDENCA.AnotaPed (vista texto): bloque con título "ITEMS REMITIDOS -
        ///     REMITO {remito}" y una línea por ítem en el formato fijo de producto
        ///     (76 chars: "- " + código en pos 11, descripción en pos 29 y cantidad
        ///     remitida en pos 57), que la vista texto parsea y muestra como tabla
        ///     de productos (detalle resaltado).
        ///  2) ANOTAPED (vista tabla): un registro por ítem compartiendo un nuevo
        ///     IdAnotacion, con ObservaGral = "Items Remitidos - Remito {N°}"
        ///     (título del bloque, mostrado en negrita) y ObservaItem con la
        ///     referencia del remito.
        /// Solo se anotan los movimientos con NuPedCli &gt; 0: los ítems del remito
        /// que no están asociados a un pedido no generan anotación.
        /// </summary>
        /// <param name="nroPedido">Número de pedido (PEDENCA.NroPed = MOVISTO.NuPedCli)</param>
        /// <param name="movimientos">Movimientos MOVISTO del remito asociados a este pedido</param>
        /// <param name="remitoLargo">Documento del remito (MOVISTO.DoPriLar)</param>
        /// <param name="nombreUsuario">Nombre del usuario activo</param>
        /// <param name="numUsuar">NumUsuar del usuario (usuarios.idflexoft)</param>
        public async Task RegistrarItemsRemitidosAsync(int nroPedido, List<Movisto> movimientos, string remitoLargo, string? nombreUsuario, int? numUsuar)
        {
            try
            {
                if (nroPedido <= 0 || movimientos == null || !movimientos.Any())
                    return;

                var ahora = DateTime.Now;
                var usuarioTexto = string.IsNullOrEmpty(nombreUsuario) ? "SISTEMA" : nombreUsuario;
                const string tituloAnotacion = "Items Remitidos";

                // Consolidar los movimientos por ítem del pedido (Cod_Inte + NumeOrdPed):
                // un producto trazable puede generar varios movimientos (uno por lote)
                // y en la anotación se muestra una sola línea con la cantidad total remitida.
                var items = movimientos
                    .GroupBy(m => new { m.Cod_Inte, m.NumeOrdPed })
                    .Select(g => new
                    {
                        CodiInt = g.Key.Cod_Inte ?? 0,
                        CodiExt = g.First().Cod_Exte,
                        Descrip = g.First().Descri_Lar ?? g.First().Descrip,
                        CantidadRemitida = g.Sum(m => m.CantSalid ?? 0m),
                        NuOrItem = g.Key.NumeOrdPed ?? (short)0
                    })
                    .OrderBy(x => x.NuOrItem)
                    .ToList();

                // ── 1) Vista texto: bloque en PEDENCA.AnotaPed ──────────────────────
                var anotacionCompleta = new StringBuilder();

                // Línea de encabezado de 96 chars: fecha (1-8) + título (11-66) + usuario (69-92)
                var lineaHeader = new string(' ', 96).ToCharArray();
                ReemplazarEnPosicion(lineaHeader, ahora.ToString("dd/MM/yy"), 0, 8);
                ReemplazarEnPosicion(lineaHeader, $"{tituloAnotacion.ToUpper()} - REMITO {remitoLargo}", 10, 56);
                ReemplazarEnPosicion(lineaHeader, usuarioTexto, 68, 24);
                anotacionCompleta.AppendLine(new string(lineaHeader));

                // Líneas de ítem: formato fijo de 76 chars igual al resto de las
                // anotaciones de producto, con prefijo "- " para que el parser de la
                // vista texto las reconozca y las muestre en la tabla de productos.
                //   pos 11 (idx 10): "- " + código (18 chars)
                //   pos 29 (idx 28): descripción (38 chars)
                //   pos 57 (idx 56): cantidad remitida (10 chars)
                foreach (var item in items)
                {
                    var lineaItem = new string(' ', 76).ToCharArray();
                    var codigo = !string.IsNullOrEmpty(item.CodiExt) ? item.CodiExt : item.CodiInt.ToString();
                    ReemplazarEnPosicion(lineaItem, "- " + codigo, 10, 18);
                    ReemplazarEnPosicion(lineaItem, item.Descrip ?? "", 28, 38);
                    ReemplazarEnPosicion(lineaItem, item.CantidadRemitida.ToString("F1", CultureInfo.InvariantCulture).PadLeft(10), 56, 10);
                    anotacionCompleta.AppendLine(new string(lineaItem));
                }

                // Separador de bloque: 88 "=" (igual que el resto de las anotaciones)
                anotacionCompleta.AppendLine(new string('=', 88));

                var pedido = await _context.Pedencas.FirstOrDefaultAsync(p => p.NroPed == nroPedido);
                if (pedido != null)
                {
                    pedido.AnotaPed = (pedido.AnotaPed ?? string.Empty) + " " + anotacionCompleta.ToString();
                }

                // ── 2) Vista tabla: registros ANOTAPED ──────────────────────────────
                // Todos los ítems comparten el mismo IdAnotacion (incremental por pedido).
                var maxIdAnotacion = await _context.Anotapeds
                    .Where(a => a.NroPed == nroPedido)
                    .MaxAsync(a => (int?)a.IdAnotacion) ?? 0;
                var nuevoIdAnotacion = maxIdAnotacion + 1;

                var registros = items.Select(item => new Anotaped
                {
                    IdAnotacion = nuevoIdAnotacion,
                    NroPed = nroPedido,
                    CodiInt = item.CodiInt,
                    CodiExt = item.CodiExt,
                    Descrip = item.Descrip,
                    CanPed = item.CantidadRemitida,
                    CantEnt = item.CantidadRemitida,
                    ObservaItem = $"Remitido por remito {remitoLargo}",
                    // El título incluye el N° de remito (MOVISTO.DoPriLar) para
                    // que quede visible en el footer de la Vista Tabla
                    ObservaGral = $"{tituloAnotacion} - Remito {remitoLargo}",
                    FechaCambio = ahora,
                    Usuario = usuarioTexto,
                    CodiEmpr = 0,
                    AddRecord = ahora,
                    LastUpdate = ahora,
                    CodiEmprNet = 0,
                    NumUsuar = numUsuar ?? 0
                }).ToList();

                _context.Anotapeds.AddRange(registros);
                await _context.SaveChangesAsync();

                _logger.LogInformation(
                    "Anotación '{Titulo}' registrada para pedido {NroPed}: {Cantidad} ítems remitidos por {Remito} (IdAnotacion={IdAnotacion})",
                    tituloAnotacion, nroPedido, registros.Count, remitoLargo, nuevoIdAnotacion);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar anotación de ítems remitidos del pedido {NroPedido}", nroPedido);
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
