using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para actualizar costos en la tabla MASTER según configuración ACOSREP
    /// Replica la lógica VB6 de actualización de costos al guardar órdenes de compra
    /// </summary>
    public class ActualizacionCostoMasterService
    {
        private readonly DataContext _context;

        public ActualizacionCostoMasterService(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Actualiza los costos en MASTER para los items de una orden de compra
        /// según la configuración ACOSREP y el proveedor habitual
        /// </summary>
        /// <param name="numeroOrden">Número de la orden de compra</param>
        /// <param name="numeroProveedor">Número del proveedor de la orden</param>
        /// <returns>Cantidad de productos actualizados</returns>
        public async Task<int> ActualizarCostosMasterAsync(int numeroOrden, int numeroProveedor)
        {
            try
            {
                Console.WriteLine($"[ActualizacionCostoMaster] 🔍 Iniciando actualización de costos para O.C. {numeroOrden}, Proveedor {numeroProveedor}");

                // 1. OBTENER CONFIGURACIÓN ACOSREP
                // ACOSREP$ = UCase$(Control$("OCOMPRA", "ACOSREP"))
                var configuracionAcosrep = await _context.Flexcrls
                    .FirstOrDefaultAsync(f => f.PROBUS == "OCOMPRA" && f.CLABUS == "ACOSREP");

                string acosrep = configuracionAcosrep?.VALCONTROL?.ToUpper() ?? "SIEMPRE";
                Console.WriteLine($"[ActualizacionCostoMaster] ⚙️ Configuración ACOSREP: '{acosrep}'");

                // Si está configurado como NUNCA, no actualizar nada
                if (acosrep == "NUNCA")
                {
                    Console.WriteLine($"[ActualizacionCostoMaster] ⚠️ ACOSREP = NUNCA, no se actualizarán costos");
                    return 0;
                }

                // 2. OBTENER DETALLE DE LA ORDEN DE COMPRA (solo items activos)
                var detalleOrden = await _context.Ocomdetas
                    .Where(d => d.NumeOcom == numeroOrden && d.StatOcom == 0)
                    .ToListAsync();

                Console.WriteLine($"[ActualizacionCostoMaster] 📋 Items activos encontrados: {detalleOrden.Count}");

                if (!detalleOrden.Any())
                {
                    Console.WriteLine($"[ActualizacionCostoMaster] ⚠️ No hay items activos para procesar");
                    return 0;
                }

                int productosActualizados = 0;

                // 3. PROCESAR CADA ITEM DEL DETALLE
                foreach (var item in detalleOrden)
                {
                    // Validar que el item tenga código interno
                    if (!item.CodInte.HasValue || item.CodInte.Value <= 0)
                    {
                        Console.WriteLine($"[ActualizacionCostoMaster] ⏭️ Item sin código interno, saltando...");
                        continue;
                    }

                    // Buscar el producto en MASTER
                    var producto = await _context.masters
                        .FirstOrDefaultAsync(m => m.Codint == item.CodInte.Value);

                    if (producto == null)
                    {
                        Console.WriteLine($"[ActualizacionCostoMaster] ⚠️ Producto con CodInt {item.CodInte.Value} no encontrado en MASTER");
                        continue;
                    }

                    // LÓGICA VB6 CON 3 OPCIONES:
                    // PRHBL% = PROVHABI%(CI0%) - Proveedor habitual del producto
                    // If PRHBL% < 1 Or PRHBL% = NPRO% Or ACOSREP$ = "SIEMPRE" Then
                    short proveedorHabitual = producto.Provhabi ?? 0;
                    bool debeActualizar = false;

                    // OPCIÓN 1: SIEMPRE - Actualizar siempre sin importar el proveedor
                    if (acosrep == "SIEMPRE")
                    {
                        debeActualizar = true;
                        Console.WriteLine($"[ActualizacionCostoMaster] ✅ Producto {producto.Codigo} se actualizará por ACOSREP=SIEMPRE");
                    }
                    // OPCIÓN 2: PROHABIT - Solo actualizar si es proveedor habitual o no tiene proveedor habitual
                    else if (acosrep == "PROHABIT")
                    {
                        if (proveedorHabitual < 1)
                        {
                            // No tiene proveedor habitual definido
                            debeActualizar = true;
                            Console.WriteLine($"[ActualizacionCostoMaster] ✅ Producto {producto.Codigo} sin proveedor habitual, se actualizará (PROHABIT)");
                        }
                        else if (proveedorHabitual == numeroProveedor)
                        {
                            // El proveedor habitual es el mismo de la orden
                            debeActualizar = true;
                            Console.WriteLine($"[ActualizacionCostoMaster] ✅ Producto {producto.Codigo} con proveedor habitual {proveedorHabitual} = proveedor actual, se actualizará (PROHABIT)");
                        }
                        else
                        {
                            // Tiene proveedor habitual diferente, no actualizar
                            Console.WriteLine($"[ActualizacionCostoMaster] ⏭️ Producto {producto.Codigo} tiene proveedor habitual diferente ({proveedorHabitual} != {numeroProveedor}), no se actualiza (PROHABIT)");
                        }
                    }
                    // OPCIÓN 3: NUNCA - No actualizar (ya se maneja al inicio del método)

                    if (!debeActualizar)
                    {
                        Console.WriteLine($"[ActualizacionCostoMaster] ⏭️ Producto {producto.Codigo} no cumple criterios de actualización (ProvHabi={proveedorHabitual}, ProvActual={numeroProveedor}, ACOSREP={acosrep})");
                        continue;
                    }

                    // OBTENER COSTO DE COMPRA ACTIVO
                    // COSTO_COMPRA_ACTIVA es el precio unitario neto del item
                    decimal costoCompraActivo = item.PrUnNeto ?? item.PreUnit ?? 0m;

                    // Validar que el costo sea significativo
                    // If COSTO_COMPRA_ACTIVA >= 0.000005 Then
                    if (costoCompraActivo >= 0.000005m)
                    {
                        // Guardar costo anterior para logging
                        decimal costoAnterior = producto.Cosuni ?? 0m;

                        // ACTUALIZAR CAMPOS EN MASTER
                        // !COSUNI = COSTO_COMPRA_ACTIVA
                        producto.Cosuni = costoCompraActivo;

                        // !FECOSTO = HOY
                        producto.Fecosto = DateTime.Now;

                        // !Monecos = MONECO% (moneda de costo de la orden)
                        producto.Monecos = item.MoneCos;

                        // If PRHBL% < 1 Then !PROVHABI = NPRO%
                        if (proveedorHabitual < 1)
                        {
                            producto.Provhabi = (short)numeroProveedor;
                            Console.WriteLine($"[ActualizacionCostoMaster] 📝 Asignando proveedor habitual {numeroProveedor} al producto {producto.Codigo}");
                        }

                        productosActualizados++;

                        Console.WriteLine($"[ActualizacionCostoMaster] 💰 ACTUALIZADO Producto {producto.Codigo}:");
                        Console.WriteLine($"  - COSUNI: {costoAnterior:C4} → {producto.Cosuni:C4}");
                        Console.WriteLine($"  - FECOSTO: {producto.Fecosto:yyyy-MM-dd}");
                        Console.WriteLine($"  - Monecos: {producto.Monecos}");
                        Console.WriteLine($"  - Provhabi: {producto.Provhabi}");
                    }
                    else
                    {
                        Console.WriteLine($"[ActualizacionCostoMaster] ⚠️ Producto {producto.Codigo} con costo insignificante ({costoCompraActivo}), no se actualiza");
                    }
                }

                // 4. GUARDAR CAMBIOS EN LA BASE DE DATOS
                if (productosActualizados > 0)
                {
                    await _context.SaveChangesAsync();
                    Console.WriteLine($"[ActualizacionCostoMaster] ✅ Actualizados {productosActualizados} productos en MASTER");
                }
                else
                {
                    Console.WriteLine($"[ActualizacionCostoMaster] ℹ️ No se actualizó ningún producto");
                }

                return productosActualizados;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ActualizacionCostoMaster] ❌ ERROR: {ex.Message}");
                Console.WriteLine($"[ActualizacionCostoMaster] StackTrace: {ex.StackTrace}");
                throw;
            }
        }
    }
}
