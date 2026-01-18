namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Interfaz para validaciones específicas por empresa y dependencia
    /// </summary>
    public interface IEmpresaDependenciaService
    {
        /// <summary>
        /// Obtiene el nombre de la empresa del usuario actual
        /// </summary>
        string? ObtenerNombreEmpresaActual();

        /// <summary>
        /// Obtiene el nombre de la dependencia del usuario actual
        /// </summary>
        string? ObtenerNombreDependenciaActual();

        /// <summary>
        /// Verifica si el usuario pertenece a una empresa específica
        /// </summary>
        bool EsEmpresa(string nombreEmpresa);

        /// <summary>
        /// Verifica si el usuario pertenece a una dependencia específica
        /// </summary>
        bool EsDependencia(string nombreDependencia);

        /// <summary>
        /// Verifica si el usuario pertenece a una empresa Y dependencia específicas
        /// </summary>
        bool EsEmpresaYDependencia(string nombreEmpresa, string nombreDependencia);

        /// <summary>
        /// Obtiene el ID de la empresa del usuario actual
        /// </summary>
        int? ObtenerEmpresaIdActual();

        /// <summary>
        /// Obtiene el ID de la dependencia del usuario actual
        /// </summary>
        int? ObtenerDependenciaIdActual();
    }
}
