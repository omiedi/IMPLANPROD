using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para manejar las consignaciones (MACONSIG) según la lógica VB6
    /// </summary>
    public class ConsignacionService
    {
        private readonly DataContext _context;

        public ConsignacionService(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Procesa las consignaciones para una orden de compra según la lógica VB6
        /// Por cada item de OCOMDETA que tenga estructura (sea padre en Formulas),
        /// agrega sus componentes hijos a MACONSIG
        /// </summary>
        /// <param name="numeroOrden">Número de orden de compra</param>
        /// <param name="numeroProveedor">Número de proveedor</param>
        /// <param name="razonSocialProveedor">Razón social del proveedor</param>
        /// <param name="fechaEmision">Fecha de emisión de la orden</param>
        /// <param name="monedaEmision">Moneda de emisión</param>
        /// <returns>True si se procesaron consignaciones, False si no había items con estructura</returns>
        public async Task<bool> ProcesarConsignacionesOrdenCompraAsync(
            int numeroOrden, 
            int numeroProveedor, 
            string razonSocialProveedor, 
            DateTime fechaEmision,
            short monedaEmision)
        {
            Console.WriteLine($"[ConsignacionService] 🔄 Iniciando procesamiento de consignaciones para O.C. {numeroOrden}");

            try
            {
                Console.WriteLine($"[ConsignacionService] 🔍 DEBUGGING - Buscando items en OCOMDETA para O.C. {numeroOrden}");
                
                // Primero verificar si existen items en general (sin filtro de estado)
                var todosLosItems = await _context.Ocomdetas
                    .Where(d => d.NumeOcom == numeroOrden)
                    .ToListAsync();
                
                Console.WriteLine($"[ConsignacionService] 📊 DEBUGGING - Total items encontrados (todos los estados): {todosLosItems.Count}");
                
                if (todosLosItems.Any())
                {
                    foreach (var item in todosLosItems)
                    {
                        Console.WriteLine($"[ConsignacionService] 📋 DEBUGGING - Item: NuOrdItem={item.NuOrdItem}, CodInte={item.CodInte}, StatOcom={item.StatOcom}");
                    }
                }

                // 1. Obtener todos los items activos de la orden de compra
                var itemsOrdenCompra = await _context.Ocomdetas
                    .Where(d => d.NumeOcom == numeroOrden && d.StatOcom == 0)
                    .ToListAsync();

                Console.WriteLine($"[ConsignacionService] 📊 DEBUGGING - Items activos (StatOcom=0): {itemsOrdenCompra.Count}");

                if (!itemsOrdenCompra.Any())
                {
                    Console.WriteLine($"[ConsignacionService] ⚠️ No se encontraron items activos en O.C. {numeroOrden}");
                    Console.WriteLine($"[ConsignacionService] 💡 POSIBLE CAUSA: Los items pueden tener StatOcom != 0 después de la edición");
                    return false;
                }

                Console.WriteLine($"[ConsignacionService] 📋 Encontrados {itemsOrdenCompra.Count} items en O.C. {numeroOrden}");

                bool seProcesoAlgunaConsignacion = false;
                short nuOrdItem = 1; // Contador para NUORDITEM

                // 2. Por cada item de la orden, verificar si tiene estructura (es padre)
                foreach (var item in itemsOrdenCompra)
                {
                    if (!item.CodInte.HasValue)
                    {
                        Console.WriteLine($"[ConsignacionService] ⚠️ Item sin código interno, saltando...");
                        continue;
                    }

                    // 3. Buscar si este código es padre en la tabla Formulas (solo primer nivel)
                    Console.WriteLine($"[ConsignacionService] 🔍 Buscando estructura para código interno: {item.CodInte.Value}");
                    
                    // Primero verificar si existen registros en Formulas para este código
                    var totalFormulas = await _context.Formulas
                        .Where(f => f.CodIConj == item.CodInte.Value)
                        .CountAsync();
                    
                    Console.WriteLine($"[ConsignacionService] 📊 Total registros en Formulas para código {item.CodInte.Value}: {totalFormulas}");
                    
                    var componentesHijos = await _context.Formulas
                        .Where(f => f.CodIConj == item.CodInte.Value && f.Marborra != 1) // No borrados
                        .ToListAsync();

                    Console.WriteLine($"[ConsignacionService] 📋 Componentes hijos no borrados encontrados: {componentesHijos.Count}");

                    if (!componentesHijos.Any())
                    {
                        Console.WriteLine($"[ConsignacionService] ℹ️ Código {item.CodInte} no tiene estructura (no es padre)");
                        continue;
                    }

                    Console.WriteLine($"[ConsignacionService] 🔍 Código {item.CodInte} tiene {componentesHijos.Count} componentes hijos");

                    // 4. Por cada componente hijo, procesar según lógica de edición
                    foreach (var componente in componentesHijos)
                    {
                        // PASO 2: Igualar Cant_Asig = Cant_Entreg en registros existentes con entregas
                        if (componente.CodiElem.HasValue)
                        {
                            await IgualarCantidadAsignadaAEntregada(numeroOrden, componente.CodiElem.Value);
                        }

                        // PASO 3: Crear nuevo registro por remanente
                        await CrearRegistroConsignacionAsync(
                            numeroOrden,
                            item.NuOrdItem ?? 0,
                            numeroProveedor,
                            razonSocialProveedor,
                            fechaEmision,
                            monedaEmision,
                            componente,
                            item.CantOcom ?? 0,
                            nuOrdItem++);

                        seProcesoAlgunaConsignacion = true;
                    }
                }

                if (seProcesoAlgunaConsignacion)
                {
                    Console.WriteLine($"[ConsignacionService] ✅ Consignaciones procesadas exitosamente para O.C. {numeroOrden}");
                }
                else
                {
                    Console.WriteLine($"[ConsignacionService] ℹ️ No se encontraron items con estructura en O.C. {numeroOrden}");
                }

                return seProcesoAlgunaConsignacion;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ConsignacionService] ❌ Error procesando consignaciones: {ex.Message}");
                Console.WriteLine($"[ConsignacionService] StackTrace: {ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Crea un registro individual de consignación en MACONSIG
        /// Replica la lógica VB6 del código proporcionado
        /// </summary>
        private async Task CrearRegistroConsignacionAsync(
            int numeroOrden,
            int numeroItemOrden,
            int numeroProveedor,
            string razonSocialProveedor,
            DateTime fechaEmision,
            short monedaEmision,
            Formulas componente,
            decimal cantidadPadre,
            short nuOrdItem)
        {
            try
            {
                Console.WriteLine($"[ConsignacionService] 🔧 Creando consignación para componente: {componente.CodXElem} (Código interno: {componente.CodiElem})");
                // 1. Generar número de orden formateado según VB6
                // NOCOM$ = TRIM$(Str$(NRO&))
                // While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
                // ITNU$ = TRIM$(Str$(ORDIT%))
                // While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
                // NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
                
                string numeroOrdenStr = numeroOrden.ToString().PadLeft(6, '0');
                string numeroItemStr = numeroItemOrden.ToString().PadLeft(2, '0');
                string nroOcomFormateado = $"OC-{numeroOrdenStr}-{numeroItemStr}";

                // 2. Calcular cantidad deseada según uso por unidad (CANTIASIGN#)
                // CANTIASIGN# = CAUNIS * USOI(i%) (cantidad deseada)
                decimal cantidadDeseada = cantidadPadre * (componente.UsoBruto ?? 0) * (componente.Unis_Form ?? 1);
                Console.WriteLine($"[ConsignacionService] 📊 Cantidad deseada calculada: {cantidadDeseada:F4}");

                // 3. Verificar cantidad ya asignada (CANYASIG#)
                // CANYASIG# = SUM(CANT_ASIG) donde Nro_ocom = 'OC-...' AND CANT_ASIG > 0 AND COD_INTE = CII%
                string patronOrdenComponente = $"OC-{numeroOrden.ToString().PadLeft(6, '0')}-";
                
                var cantidadYaAsignada = await _context.Maconsigs
                    .Where(m => m.NroOcom != null && 
                               m.NroOcom.StartsWith(patronOrdenComponente) && 
                               m.CantAsig > 0 && 
                               m.CodInte == componente.CodiElem)
                    .SumAsync(m => m.CantAsig ?? 0);

                Console.WriteLine($"[ConsignacionService] 📋 Cantidad ya asignada: {cantidadYaAsignada:F4}");

                // 4. Calcular remanente (CANTIASIGN# - CANYASIG#)
                decimal cantidadRemanente = cantidadDeseada - cantidadYaAsignada;
                Console.WriteLine($"[ConsignacionService] ⚖️ Remanente calculado: {cantidadRemanente:F4} = {cantidadDeseada:F4} - {cantidadYaAsignada:F4}");

                // 5. Validar si debe crear registro
                // En edición: crear registro gemelo si hay remanente > 0 (sin importar si es <= 0.05)
                // En creación: aplicar lógica VB6 original (remanente > 0.05)
                
                // Verificar si hay registros existentes con entregas (indica que es edición)
                var hayRegistrosConEntregas = await _context.Maconsigs
                    .AnyAsync(m => m.NroOcom != null && 
                              m.NroOcom.StartsWith(patronOrdenComponente) && 
                              m.CodInte == componente.CodiElem &&
                              m.CantEntreg > 0);

                bool esEdicion = hayRegistrosConEntregas;
                
                if (!esEdicion && cantidadRemanente <= 0.05m)
                {
                    Console.WriteLine($"[ConsignacionService] ⚠️ CREACIÓN - Remanente {cantidadRemanente:F4} <= 0.05, no se crea consignación");
                    return;
                }
                
                if (esEdicion && cantidadRemanente <= 0)
                {
                    Console.WriteLine($"[ConsignacionService] ⚠️ EDICIÓN - Remanente {cantidadRemanente:F4} <= 0, no se crea consignación");
                    return;
                }

                if (string.IsNullOrEmpty(componente.CodXElem))
                {
                    Console.WriteLine($"[ConsignacionService] ⚠️ Código externo vacío para componente {componente.CodiElem}, no se crea consignación");
                    return;
                }

                string tipoOperacion = esEdicion ? "EDICIÓN" : "CREACIÓN";
                Console.WriteLine($"[ConsignacionService] ✅ {tipoOperacion} - Validaciones pasadas, creando consignación por remanente: {cantidadRemanente:F4}");

                Console.WriteLine($"[ConsignacionService] ✅ Validaciones pasadas, creando consignación por remanente: {cantidadRemanente:F4}");

                // 6. Obtener información del componente desde Master
                var masterComponente = await _context.masters
                    .FirstOrDefaultAsync(m => m.Codint == componente.CodiElem);

                if (masterComponente == null)
                {
                    Console.WriteLine($"[ConsignacionService] ⚠️ No se encontró Master para código {componente.CodiElem}");
                    return;
                }

                // 7. Calcular valor total según VB6 usando el remanente
                // Valo_Tot = REMANENTE * costo unitario * master.monecos
                decimal costoUnitario = masterComponente.Cosuni ?? 0;
                decimal factorMoneda = masterComponente.Monecos ?? 1;
                decimal valorTotal = cantidadRemanente * costoUnitario * factorMoneda;

                // 5. Crear registro MACONSIG según estructura VB6
                var consignacion = new Maconsig
                {
                    // Campos básicos de la orden
                    CodiEmpr = 0, // !CODIEMPR = 0
                    FemiOcom = fechaEmision, // !FEMI_OCOM = hoy
                    NproOcom = numeroProveedor, // !Npro_Ocom = Nroproveedor
                    RaSoProv = razonSocialProveedor, // !Raso_Prov = razon social proveedor
                    NroOcom = nroOcomFormateado, // !nro_ocom = NOCOM$

                    // Información del componente
                    CodInte = componente.CodiElem, // !cod_inte = CII%
                    CodExte = componente.CodXElem, // !Cod_Exte = CODEXT$(CII%)
                    Descritem = masterComponente.Descripcion, // !DESCRITEM = DESCRIT0$(CII%)
                    Unidad = masterComponente.Umedida, // !Unidad = unidad de medida

                    // Cantidades y valores
                    CantAsig = cantidadRemanente, // !cant_asig = REMANENTE (cantidad deseada - ya asignada)
                    ValoUnit = costoUnitario, // !Valo_Unit = costo unitario (master.cosuni)
                    Moneda = monedaEmision, // !Moneda = moneda de emision
                    ValoTot = valorTotal, // !Valo_Tot = CANTI# * costo unitario * master.monecos

                    // Campos inicializados según VB6
                    SeccEjec = 0, // !Secc_Ejec = 0
                    DenoSecc = "", // !Deno_Secc = ""
                    NroOrfa = "", // !NRO_ORFA = ""
                    LoteOcom = "", // !LOTE_OCOM = ""
                    CantEntreg = 0, // !cant_entreg = 0
                    FechaEntreg = null, // !FECHA_ENTREG = CVDAT(0) - null en C#
                    NRemiConsig = "", // !NREMI_CONSIG = ""
                    CantRendid = 0, // !CANT_RENDID = 0
                    CantDevu = 0, // !cant_devu = 0
                    SaldoPend = 0, // !Saldo_Pend = 0
                    Observa = "", // !Observa = ""
                    NuOrdItem = nuOrdItem // !NUORDITEM = NUORD%
                };

                // 6. Agregar a contexto
                _context.Maconsigs.Add(consignacion);

                Console.WriteLine($"[ConsignacionService] ➕ Creada consignación: {nroOcomFormateado} - {componente.CodXElem} - Remanente: {cantidadRemanente:F4}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ConsignacionService] ❌ Error creando consignación individual: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// PASO 2: Iguala Cant_Asig = Cant_Entreg en registros existentes con entregas
        /// Y crea registros gemelos por la diferencia
        /// </summary>
        /// <param name="numeroOrden">Número de orden de compra</param>
        /// <param name="codigoComponente">Código interno del componente</param>
        private async Task IgualarCantidadAsignadaAEntregada(int numeroOrden, int codigoComponente)
        {
            try
            {
                // Buscar registros existentes de este componente con entregas
                string patronOrden = $"OC-{numeroOrden.ToString().PadLeft(6, '0')}-";
                
                var registrosConEntregas = await _context.Maconsigs
                    .Where(m => m.NroOcom != null && 
                               m.NroOcom.StartsWith(patronOrden) && 
                               m.CodInte == codigoComponente &&
                               m.CantEntreg > 0)
                    .ToListAsync();

                Console.WriteLine($"[ConsignacionService] 🔄 PASO 2 - Componente {codigoComponente}: {registrosConEntregas.Count} registros con entregas");

                foreach (var registro in registrosConEntregas)
                {
                    decimal cantEntregOriginal = registro.CantEntreg ?? 0;
                    decimal cantAsigOriginal = registro.CantAsig ?? 0;
                    decimal diferencia = cantAsigOriginal - cantEntregOriginal;
                    
                    // Igualar Cant_Asig = Cant_Entreg
                    registro.CantAsig = registro.CantEntreg;
                    
                    Console.WriteLine($"[ConsignacionService] 📝 Registro {registro.NroOcom}: CantAsig {cantAsigOriginal:F4} -> {registro.CantAsig:F4} (= CantEntreg {cantEntregOriginal:F4})");
                    
                    // Crear registro gemelo por la diferencia si es > 0
                    if (diferencia > 0)
                    {
                        await CrearRegistroGemelo(registro, diferencia);
                        Console.WriteLine($"[ConsignacionService] 👥 Creado registro gemelo por diferencia: {diferencia:F4}");
                    }
                }

                if (registrosConEntregas.Any())
                {
                    Console.WriteLine($"[ConsignacionService] ✅ PASO 2 completado - {registrosConEntregas.Count} registros ajustados para componente {codigoComponente}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ConsignacionService] ❌ Error en PASO 2 para componente {codigoComponente}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Crea un registro gemelo basado en un registro existente, con la diferencia como Cant_Asig
        /// </summary>
        private Task CrearRegistroGemelo(Maconsig registroOriginal, decimal diferencia)
        {
            var registroGemelo = new Maconsig
            {
                // Copiar todos los campos del registro original
                CodiEmpr = registroOriginal.CodiEmpr,
                FemiOcom = registroOriginal.FemiOcom,
                NproOcom = registroOriginal.NproOcom,
                RaSoProv = registroOriginal.RaSoProv,
                NroOcom = registroOriginal.NroOcom,
                CodInte = registroOriginal.CodInte,
                CodExte = registroOriginal.CodExte,
                Descritem = registroOriginal.Descritem,
                Unidad = registroOriginal.Unidad,
                ValoUnit = registroOriginal.ValoUnit,
                Moneda = registroOriginal.Moneda,
                SeccEjec = registroOriginal.SeccEjec,
                DenoSecc = registroOriginal.DenoSecc,
                NroOrfa = registroOriginal.NroOrfa,
                LoteOcom = registroOriginal.LoteOcom,
                NRemiConsig = registroOriginal.NRemiConsig,
                Observa = registroOriginal.Observa,
                NuOrdItem = registroOriginal.NuOrdItem,
                
                // Campos específicos del gemelo
                CantAsig = diferencia,           // La diferencia
                CantEntreg = 0,                  // Sin entregas
                FechaEntreg = null,              // Sin fecha de entrega
                CantRendid = 0,                  // Sin rendimiento
                CantDevu = 0,                    // Sin devoluciones
                SaldoPend = 0,                   // Sin saldo pendiente
                ValoTot = diferencia * (registroOriginal.ValoUnit ?? 0) // Recalcular valor total
            };

            _context.Maconsigs.Add(registroGemelo);
            return Task.CompletedTask;
        }

        /// <summary>
        /// Elimina las consignaciones existentes de una orden de compra
        /// Se usa cuando se edita una orden para limpiar consignaciones anteriores
        /// </summary>
        /// <param name="numeroOrden">Número de orden de compra</param>
        public async Task EliminarConsignacionesOrdenAsync(int numeroOrden)
        {
            try
            {
                // Buscar consignaciones existentes que coincidan con el patrón de la orden
                string patronOrden = $"OC-{numeroOrden.ToString().PadLeft(6, '0')}-";
                
                var consignacionesExistentes = await _context.Maconsigs
                    .Where(c => c.NroOcom != null && c.NroOcom.StartsWith(patronOrden))
                    .ToListAsync();

                if (consignacionesExistentes.Any())
                {
                    _context.Maconsigs.RemoveRange(consignacionesExistentes);
                    Console.WriteLine($"[ConsignacionService] 🗑️ Eliminadas {consignacionesExistentes.Count} consignaciones existentes de O.C. {numeroOrden}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ConsignacionService] ❌ Error eliminando consignaciones: {ex.Message}");
                throw;
            }
        }
    }
}
