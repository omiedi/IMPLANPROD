using System.Security.Claims;
using System;
using IMPLANPROD.Server.Data;
using DocumentFormat.OpenXml.Drawing;
using IMPLANPROD.Shared.Entities;
using System.Net;
using System.Runtime.ConstrainedExecution;
using System.Runtime.Intrinsics.X86;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Services
{

    public class Codiempresa : ICodiempresa
    {
        private readonly DataContext _context;


            public Codiempresa(DataContext context)
            {
                _context = context;
            }

        public short? IObtenerCodiempresa(ClaimsPrincipal user)
        {
            if (user?.Identity?.IsAuthenticated ?? false)
            {
                var username = user.Identity?.Name;
                var email = user.FindFirst(ClaimTypes.Email)?.Value;

                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
                {
                    var usuario = _context.Usuarios.FirstOrDefault(u =>
                        u.NombreUsuario == username && u.Email == email);
                    int? empresaIdn = usuario?.EmpresaId;
                    int? dependenciaIdn = usuario?.DependenciaId;

                    if (usuario != null && empresaIdn.HasValue && dependenciaIdn.HasValue)
                    {
                        var empresaId = empresaIdn.Value;
                        var dependenciaId = dependenciaIdn.Value;

                        if (empresaId >= 0 && empresaId <= 99 &&
                            dependenciaId >= 0 && dependenciaId <= 327)
                        {
                            short codiempr = (short)(dependenciaId * 100 + empresaId);
                            return codiempr;
                        }

                        // Opcional: throw new InvalidOperationException("Valores fuera de rango");
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Obtiene el nombre de la empresa (razón social) a partir del ID del usuario.
        /// Busca primero en la tabla Usuarios para obtener el EmpresaId,
        /// luego consulta la tabla Empresas para obtener la razón social.
        /// </summary>
        /// <param name="usuarioId">ID del usuario</param>
        /// <returns>Nombre de la empresa (razón social) o null si no se encuentra</returns>
        public string? ObtenerNombreEmpresa(int usuarioId)
        {
            try
            {
                // Buscar el usuario por su ID para obtener el EmpresaId
                var usuario = _context.Usuarios.AsNoTracking().FirstOrDefault(u => u.Id == usuarioId);
                
                if (usuario?.EmpresaId == null)
                {
                    return null; // Usuario no encontrado o sin empresa asignada
                }

                // Buscar la empresa por su ID para obtener la razón social
                var empresa = _context.Empresas.AsNoTracking().FirstOrDefault(e => e.Id == usuario.EmpresaId);
                
                return empresa?.razon; // Retorna la razón social o null si no se encuentra
            }
            catch (Exception)
            {
                // En caso de error, retornar null
                return null;
            }
        }

        /// <summary>
        /// Obtiene el nombre de la dependencia a partir del ID del usuario.
        /// Busca primero en la tabla Usuarios para obtener el DependenciaId,
        /// luego consulta la tabla Dependencias para obtener el nombre.
        /// </summary>
        /// <param name="usuarioId">ID del usuario</param>
        /// <returns>Nombre de la dependencia o null si no se encuentra</returns>
        public string? ObtenerNombreDependencia(int usuarioId)
        {
            try
            {
                // Buscar el usuario por su ID para obtener el DependenciaId
                var usuario = _context.Usuarios.AsNoTracking().FirstOrDefault(u => u.Id == usuarioId);
                
                if (usuario?.DependenciaId == null)
                {
                    return null; // Usuario no encontrado o sin dependencia asignada
                }

                // Buscar la dependencia por su ID para obtener el nombre
                var dependencia = _context.Dependencias.AsNoTracking().FirstOrDefault(d => d.Id == usuario.DependenciaId);
                
                return dependencia?.dependencia; // Retorna el nombre de la dependencia o null si no se encuentra
            }
            catch (Exception)
            {
                // En caso de error, retornar null
                return null;
            }
        }
     
    }
//    - dependenciaId puede ser hasta 327 → ocupa el bloque de las centenas.
//- empresaId puede ser hasta 99 → ocupa las dos cifras finales

//    ¿Qué hace exactamente % 100?
//- El operador % devuelve el residuo de una división.
//- Entonces codiempr % 100 significa: ¿qué sobra al dividir codiempr entre 100?
//- Como empresaId fue el valor que “se agregó al final” en la codificación (de menor magnitud), el resultado del módulo 100 es justamente ese valor.


}

