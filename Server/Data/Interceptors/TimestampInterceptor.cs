using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data.Interfaces;

namespace IMPLANPROD.Server.Data.Interceptors
{
    /// <summary>
    /// Interceptor que actualiza automáticamente los campos de timestamp (AddRecord y LastUpdate)
    /// para todas las entidades que implementen la interfaz ITimestampEntity
    /// </summary>
    public class TimestampInterceptor : SaveChangesInterceptor
    {
        /// <summary>
        /// Se ejecuta antes de guardar cambios en la base de datos
        /// Actualiza automáticamente los campos de fecha de creación y última actualización
        /// </summary>
        /// <param name="eventData">Datos del evento de guardado</param>
        /// <param name="result">Resultado de la interceptación</param>
        /// <returns>Resultado de la interceptación</returns>
        public override InterceptionResult<int> SavingChanges(DbContextEventData eventData, InterceptionResult<int> result)
        {
            var context = eventData.Context;

            // Si no hay contexto, continuar con el comportamiento base
            if (context == null) return base.SavingChanges(eventData, result);

            // Obtener todas las entidades que implementan ITimestampEntity y están siendo agregadas o modificadas
            var entries = context.ChangeTracker.Entries()
                .Where(e => e.Entity is ITimestampEntity &&
                            (e.State == EntityState.Added || e.State == EntityState.Modified));

            // Actualizar los timestamps para cada entidad
            foreach (var entry in entries)
            {
                var timestampEntity = (ITimestampEntity)entry.Entity;
                
                // Si es una nueva entidad, establecer la fecha de creación
                if (entry.State == EntityState.Added)
                {
                    timestampEntity.AddRecord = DateTime.Now;
                }
                
                // Siempre actualizar la fecha de última modificación
                timestampEntity.LastUpdate = DateTime.Now;
            }

            return base.SavingChanges(eventData, result);
        }

        /// <summary>
        /// Versión asíncrona del método SavingChanges
        /// Se ejecuta antes de guardar cambios de forma asíncrona en la base de datos
        /// </summary>
        /// <param name="eventData">Datos del evento de guardado</param>
        /// <param name="result">Resultado de la interceptación</param>
        /// <param name="cancellationToken">Token de cancelación</param>
        /// <returns>Resultado de la interceptación asíncrona</returns>
        public override ValueTask<InterceptionResult<int>> SavingChangesAsync(DbContextEventData eventData, InterceptionResult<int> result, CancellationToken cancellationToken = default)
        {
            var context = eventData.Context;

            // Si no hay contexto, continuar con el comportamiento base
            if (context == null) return base.SavingChangesAsync(eventData, result, cancellationToken);

            // Obtener todas las entidades que implementan ITimestampEntity y están siendo agregadas o modificadas
            var entries = context.ChangeTracker.Entries()
                .Where(e => e.Entity is ITimestampEntity &&
                            (e.State == EntityState.Added || e.State == EntityState.Modified));

            // Actualizar los timestamps para cada entidad
            foreach (var entry in entries)
            {
                var timestampEntity = (ITimestampEntity)entry.Entity;
                
                // Si es una nueva entidad, establecer la fecha de creación
                if (entry.State == EntityState.Added)
                {
                    timestampEntity.AddRecord = DateTime.Now;
                }
                
                // Siempre actualizar la fecha de última modificación
                timestampEntity.LastUpdate = DateTime.Now;
            }

            return base.SavingChangesAsync(eventData, result, cancellationToken);
        }
    }
}
