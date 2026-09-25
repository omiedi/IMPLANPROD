using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para generar un alta y consumo de Producción Propia.
    /// A diferencia de la orden de fabricación, este proceso genera directamente
    /// los movimientos de stock (ingreso del producto terminado + consumo de componentes)
    /// en la tabla MOVISTO, sin pasar por el circuito de órdenes de fabricación.
    ///
    /// Reglas de negocio (equivalente VB6):
    /// - NRO& = MAX(TRY_CAST(SUBSTRING(DOPRILAR,4,LEN(DOPRILAR)) AS INT)) + 1
    ///         WHERE DOPRILAR LIKE 'OF-%' AND DOPRILAR NOT LIKE '%-1'
    /// - DOPRI$ = "OF-" + NRO&
    /// - DOSEC$ = "PF-" + NRO&
    /// - CMOV$  = "PF" (producto fabricado)
    /// - REFE$  = Referencia ingresada (o "Informe de Produccion" si está vacío)
    /// - DEPX%  = Depósito de entrada (donde ingresa el producto terminado)
    /// - CANTI  = Cantidad a fabricar (ingresa como CantIngre > 0)
    ///
    /// Para cada componente de la fórmula:
    /// - CMOV$  = "CF" (consumo de fabricación)
    /// - REFE$  = "CONSUMO S/DESP. " + CODCON$ (código externo del producto terminado)
    /// - DEPX%  = Depósito de consumo informado en pantalla
    /// - NUORIT% = NUORIT% + 1 (incremental por componente)
    /// - Cantidad negativa → CantSalid > 0
    /// </summary>
    public class GenerarPropiaProduccionDTO
    {
        /// <summary>
        /// Fecha del movimiento (FECHA% en VB6).
        /// </summary>
        public DateTime Fecha { get; set; }

        /// <summary>
        /// Código interno del producto terminado a fabricar (CICONJ% en VB6).
        /// </summary>
        public int Cod_Inte { get; set; }

        /// <summary>
        /// Cantidad a fabricar (CANTI en VB6). Debe ser mayor a 0.
        /// Es nullable para compatibilidad con InputNumeroLimpio.
        /// </summary>
        public decimal? Cantidad { get; set; }

        /// <summary>
        /// Referencia del movimiento (REFE$ en VB6).
        /// Si está vacío se asigna "Informe de Produccion".
        /// </summary>
        public string? Referencia { get; set; }

        /// <summary>
        /// Depósito de entrada donde ingresa el producto terminado (DEPX% en VB6).
        /// </summary>
        public int DepositoEntrada { get; set; }

        /// <summary>
        /// Depósito de consumo de donde salen los componentes (Depósito de consumo).
        /// </summary>
        public int DepositoConsumo { get; set; }

        /// <summary>
        /// Componentes a consumir, con la cantidad a utilizar editada por el usuario.
        /// Se obtienen de la fórmula del producto terminado, pero el usuario puede
        /// modificar las cantidades antes de generar el movimiento.
        /// </summary>
        public List<FormulaComponenteDTO> Componentes { get; set; } = new();
    }
}
