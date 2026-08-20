/**
 * Desplaza la vista hasta un elemento identificado por su id, dejándolo visible y centrado.
 * Funciona tanto con el scroll de la ventana como con el de un contenedor scrolleable.
 * @param {string} elementId - ID del elemento al que se quiere hacer scroll
 * @returns {boolean} true si el elemento existe y se hizo el scroll
 */
export function scrollToElementById(elementId) {
    const elemento = document.getElementById(elementId);
    if (!elemento) {
        return false;
    }

    elemento.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
    return true;
}
