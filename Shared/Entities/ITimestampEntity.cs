using System;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Interface para entidades que requieren campos de timestamp (fecha de creación y última actualización)
    /// </summary>
    public interface ITimestampEntity
    {
        DateTime? AddRecord { get; set; }
        DateTime? LastUpdate { get; set; }
    }
}
