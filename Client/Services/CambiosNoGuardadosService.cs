using System;
using System.Threading.Tasks;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio para manejar la detección de cambios no guardados en formularios
    /// Proporciona funcionalidad para prevenir navegación accidental sin guardar cambios
    /// </summary>
    public class CambiosNoGuardadosService
    {
        /// <summary>
        /// Indica si hay cambios no guardados en el formulario actual
        /// </summary>
        public bool TieneCambiosNoGuardados { get; set; } = false;

        /// <summary>
        /// Mensaje personalizado para mostrar cuando hay cambios no guardados
        /// </summary>
        public string MensajePersonalizado { get; set; } = "Hay cambios sin guardar. ¿Desea salir sin guardar los cambios?";

        /// <summary>
        /// Evento que se dispara cuando se detectan cambios
        /// </summary>
        public event Action<bool>? CambiosDetectados;

        /// <summary>
        /// Marca que hay cambios no guardados
        /// </summary>
        public void MarcarCambios()
        {
            if (!TieneCambiosNoGuardados)
            {
                TieneCambiosNoGuardados = true;
                CambiosDetectados?.Invoke(true);
            }
        }

        /// <summary>
        /// Marca que los cambios han sido guardados
        /// </summary>
        public void MarcarGuardado()
        {
            if (TieneCambiosNoGuardados)
            {
                TieneCambiosNoGuardados = false;
                CambiosDetectados?.Invoke(false);
            }
        }

        /// <summary>
        /// Limpia el estado de cambios (útil al navegar a una nueva página)
        /// </summary>
        public void LimpiarEstado()
        {
            TieneCambiosNoGuardados = false;
            MensajePersonalizado = "Hay cambios sin guardar. ¿Desea salir sin guardar los cambios?";
        }

        /// <summary>
        /// Establece un mensaje personalizado para la confirmación
        /// </summary>
        public void EstablecerMensaje(string mensaje)
        {
            MensajePersonalizado = mensaje;
        }
    }
}
