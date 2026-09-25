
window.imprimirContenido = function (id) {
    const contenido = document.getElementById(id).innerHTML;

    const popup = window.open('', '', 'width=900,height=700');
    popup.document.write(`
        <html>
            <head>
                <title>Impresión</title>
                <style>
                    body { font-family: sans-serif; padding: 20px; }
                    table { border-collapse: collapse; width: 100%; }
                    th, td { border: 1px solid #ddd; padding: 8px; }
                </style>
            </head>
            <body>${contenido}</body>
        </html>`);
    popup.document.close();
    popup.print();
};

/**
 * Imprime un elemento específico del DOM
 * @param {string} elementId - ID del elemento a imprimir
 */
window.imprimirElemento = function (elementId) {
    const elemento = document.getElementById(elementId);
    if (!elemento) {
        console.error('No se encontró el elemento con ID:', elementId);
        return;
    }

    const contenido = elemento.innerHTML;
    const ventanaImpresion = window.open('', '_blank', 'width=900,height=700');
    
    if (!ventanaImpresion) {
        console.error('No se pudo abrir la ventana de impresión. Verifique que los pop-ups estén permitidos.');
        alert('No se pudo abrir la ventana de impresión. Por favor, permita las ventanas emergentes para este sitio.');
        return;
    }
    
    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Impresión</title>
                <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
                <style>
                    body { 
                        font-family: Arial, sans-serif; 
                        padding: 20px; 
                        font-size: 12px;
                    }
                    table { 
                        border-collapse: collapse; 
                        width: 100%; 
                        margin-top: 20px;
                    }
                    th, td { 
                        border: 1px solid #000; 
                        padding: 6px; 
                        text-align: left;
                    }
                    th {
                        background-color: #f0f0f0;
                        font-weight: bold;
                    }
                    img {
                        max-width: 200px;
                        height: auto;
                    }
                    h3, h5 {
                        margin: 10px 0;
                    }
                    .text-center {
                        text-align: center;
                    }
                    @media print {
                        body { margin: 0; padding: 10px; }
                        .no-print { display: none; }
                    }
                </style>
            </head>
            <body>
                ${contenido}
            </body>
        </html>
    `);
    
    ventanaImpresion.document.close();
    
    // Esperar a que todo el contenido (incluyendo imágenes) se cargue
    setTimeout(function() {
        try {
            ventanaImpresion.focus();
            ventanaImpresion.print();
        } catch (error) {
            console.error('Error al intentar imprimir:', error);
        }
    }, 500);
};

/**
 * Imprime el Vale de Transferencias de Stock en una ventana aparte,
 * ocupando una sola hoja A4, centrado por los márgenes de @page.
 * Se usa una ventana nueva (en lugar del truco visibility:hidden sobre
 * el propio modal) porque los contenedores de Bootstrap (.modal, .modal-dialog,
 * .modal-content) generan contextos de posicionamiento/scroll que rompían el
 * cálculo de una sola página al imprimir directamente desde el modal.
 * @param {string} elementId - ID del elemento que contiene el vale (areaImpresionVale)
 */
window.imprimirValeTransferencia = function (elementId) {
    const elemento = document.getElementById(elementId);
    if (!elemento) {
        console.error('No se encontró el elemento con ID:', elementId);
        return;
    }

    const contenido = elemento.innerHTML;
    const ventanaImpresion = window.open('', '_blank', 'width=900,height=700');

    if (!ventanaImpresion) {
        alert('No se pudo abrir la ventana de impresión. Por favor, permita las ventanas emergentes para este sitio.');
        return;
    }

    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Vale de Transferencias de Stock</title>
                <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
                <style>
                    @page {
                        size: A4;
                        margin: 10mm;
                    }
                    html, body {
                        margin: 0;
                        padding: 0;
                    }
                    body {
                        font-family: 'Times New Roman', Times, serif;
                        font-size: 11pt;
                    }
                    .vale-transferencia {
                        border: 1px solid #000;
                        min-height: 277mm;
                        box-sizing: border-box;
                        padding: 5mm;
                        display: flex;
                        flex-direction: column;
                    }
                    .vale-detalle {
                        flex: 1 1 auto;
                    }
                    .table {
                        font-size: 12pt;
                    }
                </style>
            </head>
            <body>
                ${contenido}
            </body>
        </html>
    `);

    ventanaImpresion.document.close();

    setTimeout(function () {
        try {
            ventanaImpresion.focus();
            ventanaImpresion.print();
        } catch (error) {
            console.error('Error al intentar imprimir el vale:', error);
        }
    }, 500);
};

window.habilitarBotonMasterIndex = function () {
    const botonMaster = document.getElementById('btnImprimirMaster');
    if (botonMaster) {
        botonMaster.disabled = false;
    }
};
//function habilitarBotonMasterIndex() {
//    const botonMaster = document.getElementById('btnImprimirMaster');
//    if (botonMaster) {
//        botonMaster.disabled = false;
//    }
//}
window.descargarArchivo = function (nombreArchivo, base64String) {
    try {
        console.log('[descargarArchivo] Iniciando descarga:', nombreArchivo);
        
        // Validar parámetros
        if (!nombreArchivo || !base64String) {
            console.error('[descargarArchivo] Parámetros inválidos:', { nombreArchivo, base64String: base64String ? 'presente' : 'ausente' });
            throw new Error('Nombre de archivo o contenido Base64 no proporcionado');
        }
        
        // Limpiar el string Base64 (remover espacios y saltos de línea)
        const base64Limpio = base64String.replace(/\s/g, '');
        
        // Validar que sea Base64 válido
        if (!/^[A-Za-z0-9+/]*={0,2}$/.test(base64Limpio)) {
            console.error('[descargarArchivo] String Base64 inválido');
            throw new Error('El contenido no es un string Base64 válido');
        }
        
        // Convertir base64 a bytes
        const byteCharacters = atob(base64Limpio);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        
        console.log('[descargarArchivo] Archivo convertido:', byteArray.length, 'bytes');
        
        // Crear blob y descargar
        const blob = new Blob([byteArray], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = nombreArchivo;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        console.log('[descargarArchivo] ✅ Descarga completada');
    } catch (error) {
        console.error('[descargarArchivo] ❌ Error:', error);
        alert('Error al descargar el archivo: ' + error.message);
        throw error;
    }
};

window.saveAsFile = function (fileName, base64String) {
    try {
        if (!fileName || !base64String) {
            throw new Error('Nombre de archivo o contenido Base64 no proporcionado');
        }

        const base64Limpio = base64String.replace(/\s/g, '');
        const byteCharacters = atob(base64Limpio);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);

        const lower = (fileName || '').toLowerCase();
        const mimeType = lower.endsWith('.xlsx')
            ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            : 'application/octet-stream';

        const blob = new Blob([byteArray], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        return true;
    } catch (error) {
        console.error('[saveAsFile] ❌ Error:', error);
        alert('Error al descargar el archivo: ' + error.message);
        return false;
    }
};

window.openPdfInNewTab = function (base64String) {
    try {
        if (!base64String) {
            throw new Error('Contenido Base64 no proporcionado');
        }

        const base64Limpio = base64String.replace(/\s/g, '');
        const byteCharacters = atob(base64Limpio);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }

        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);

        const newWindow = window.open(url, '_blank');
        if (!newWindow) {
            URL.revokeObjectURL(url);
            throw new Error('No se pudo abrir la nueva pestaña. Verifique bloqueador de pop-ups.');
        }

        setTimeout(() => URL.revokeObjectURL(url), 30000);
        return true;
    } catch (error) {
        console.error('[openPdfInNewTab] ❌ Error:', error);
        alert('Error al abrir el PDF: ' + error.message);
        return false;
    }
};

//function descargarArchivo(bytes, fileName, mimeType) {
//    const blob = new Blob([bytes], { type: mimeType });
//    const url = URL.createObjectURL(blob);
//    const link = document.createElement("a");
//    link.href = url;
//    link.download = fileName;
//    link.click();
//    URL.revokeObjectURL(url);
//}

/**
 * Enfoca un elemento del DOM
 * @param {HTMLElement} element - Elemento a enfocar
 */
window.focusElement = function (element) {
    if (element) {
        setTimeout(() => {
            element.focus();
        }, 100);
    }
};

/**
 * Enfoca un input y coloca el cursor al final del texto.
 * @param {HTMLInputElement} element - Input a enfocar
 */
window.focusCodigoInput = function (element) {
    if (element) {
        setTimeout(() => {
            element.focus();
            // Colocar el cursor al final del contenido
            var len = element.value.length;
            if (element.setSelectionRange) {
                element.setSelectionRange(len, len);
            } else if (element.createTextRange) {
                var range = element.createTextRange();
                range.collapse(true);
                range.moveEnd('character', len);
                range.moveStart('character', len);
                range.select();
            }
        }, 100);
    }
};

/**
 * Agrega un event listener para la tecla Escape en un modal
 * @param {string} modalId - ID del modal
 * @param {object} dotNetRef - Referencia al objeto .NET
 */
window.addEscapeKeyListener = function (modalId, dotNetRef) {
    const modal = document.getElementById(modalId);
    if (modal) {
        const handleEscape = function (event) {
            if (event.key === 'Escape') {
                event.preventDefault();
                event.stopPropagation();
                dotNetRef.invokeMethodAsync('CerrarConEscape');
            }
        };
        
        // Guardar la referencia del handler para poder removerlo después
        modal._escapeHandler = handleEscape;
        modal._dotNetRef = dotNetRef;
        
        // Agregar el event listener al documento
        document.addEventListener('keydown', handleEscape);
        
        // Enfocar el modal para asegurar que capture los eventos de teclado
        modal.focus();
    }
};

/**
 * Remueve el event listener de la tecla Escape de un modal
 * @param {string} modalId - ID del modal
 */
window.removeEscapeKeyListener = function (modalId) {
    const modal = document.getElementById(modalId);
    if (modal && modal._escapeHandler) {
        // Remover el event listener
        document.removeEventListener('keydown', modal._escapeHandler);

        // Limpiar las referencias
        delete modal._escapeHandler;
        if (modal._dotNetRef) {
            modal._dotNetRef.dispose();
            delete modal._dotNetRef;
        }
    }
};

/**
 * Imprime el vale interno en formato A4
 * @param {string} elementId - ID del elemento a imprimir
 */
window.imprimirVale = function (elementId) {
    const elemento = document.getElementById(elementId);
    if (!elemento) {
        console.error('No se encontró el elemento con ID:', elementId);
        return;
    }

    const contenido = elemento.innerHTML;
    const ventanaImpresion = window.open('', '_blank', 'width=900,height=700');

    if (!ventanaImpresion) {
        console.error('No se pudo abrir la ventana de impresión. Verifique que los pop-ups estén permitidos.');
        alert('No se pudo abrir la ventana de impresión. Por favor, permita las ventanas emergentes para este sitio.');
        return;
    }

    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Vale Interno</title>
                <style>
                    @page {
                        margin: 0;
                        size: A4;
                    }
                    body {
                        font-family: 'Courier New', monospace;
                        font-size: 12px;
                        margin: 20px;
                        padding: 0;
                    }
                    .d-flex {
                        display: flex;
                    }
                    .justify-content-between {
                        justify-content: space-between;
                    }
                    .align-items-start {
                        align-items: flex-start;
                    }
                    .mb-2 {
                        margin-bottom: 0.5rem;
                    }
                    .text-end {
                        text-align: right;
                    }
                    .row {
                        display: flex;
                        flex-direction: row;
                    }
                    .mb-2 {
                        margin-bottom: 0.5rem;
                    }
                    .col-3 {
                        flex: 0 0 25%;
                        max-width: 25%;
                    }
                    .col-6 {
                        flex: 0 0 50%;
                        max-width: 50%;
                    }
                    .text-end {
                        text-align: right;
                    }
                    .mt-3 {
                        margin-top: 1rem;
                    }
                    strong {
                        font-weight: bold;
                    }
                    @media print {
                        body { margin: 20px; }
                    }
                </style>
            </head>
            <body>
                ${contenido}
            </body>
        </html>
    `);

    ventanaImpresion.document.close();

    setTimeout(function() {
        try {
            ventanaImpresion.focus();
            ventanaImpresion.print();
        } catch (error) {
            console.error('Error al intentar imprimir:', error);
        }
    }, 500);
};

/**
 * Genera un PDF del vale interno en formato A4
 * @param {string} elementId - ID del elemento a convertir en PDF
 * @param {string} nombreArchivo - Nombre del archivo PDF
 */
window.generarPdfVale = function (elementId, nombreArchivo) {
    const elemento = document.getElementById(elementId);
    if (!elemento) {
        console.error('No se encontró el elemento con ID:', elementId);
        return;
    }

    if (typeof html2pdf === 'undefined') {
        console.error('La librería html2pdf no está disponible.');
        alert('No se pudo generar el PDF. La librería no está disponible.');
        return;
    }

    const opciones = {
        margin: 15,
        filename: nombreArchivo || 'ValeInterno.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    html2pdf().set(opciones).from(elemento).save();
};

/**
 * Imprime una Orden de Fabricación en formato A4 (hoja vertical).
 * Recibe un DTO del servidor con la orden, el usuario, la revisión del master
 * y el detalle de operaciones. Genera una ventana emergente con una tabla
 * fija según el diseño de la empresa.
 *
 * NOTA PARA EL DESARROLLADOR JUNIOR:
 * - Esta función no depende de un <div> oculto en el DOM. Recibe los datos en JSON
 *   y construye el HTML dentro de la ventana de impresión.
 * - El estilo @page fuerza el tamaño A4 y unos márgenes razonables.
 * - Las fechas ya vienen formateadas como texto desde el endpoint del servidor para
 *   evitar problemas de zona horaria al formatear en JS.
 *
 * @param {Object} datos - DTO OrdenFabricacionImpresionDTO serializado.
 * @param {number} datos.numeOrFa - Número interno de la orden.
 * @param {string} datos.idSubOr - Número visible (ej: OF-000001-1).
 * @param {string} datos.fechGenFormateada - Fecha de generación (dd/MM/yyyy).
 * @param {string} datos.cod_Exte - Código externo del producto.
 * @param {string} datos.descrip - Denominación del producto.
 * @param {string} datos.canProyFormateada - Cantidad proyectada (lote).
 * @param {string} datos.nombreUsuario - Nombre del usuario que lanzó la orden.
 * @param {string} datos.ferevisFormateada - Fecha de revisión del master.
 * @param {Array} datos.operaciones - Lista de operaciones { numeOper, descOper }.
 */
window.imprimirOrdenFabricacion = function (datos) {
    if (!datos) {
        console.error('[imprimirOrdenFabricacion] No se recibieron datos para imprimir.');
        alert('No se recibieron datos para imprimir la Orden de Fabricación.');
        return;
    }

    // Normalizar valores vacíos para que el formulario no muestre "undefined" o "null".
    const def = function (valor, porDefecto) {
        return (valor === null || valor === undefined || valor === '') ? (porDefecto ?? '') : String(valor);
    };

    // Construir filas de operaciones a partir del array recibido.
    let filasOperaciones = '';
    if (datos.operaciones && datos.operaciones.length > 0) {
        datos.operaciones.forEach(function (op) {
            filasOperaciones += `
                <tr>
                    <td style="border:1px solid #000; padding:4px 8px; text-align:center; width:15%;">${def(op.numeOper)}</td>
                    <td style="border:1px solid #000; padding:4px 8px; width:85%;">${def(op.descOper)}</td>
                </tr>`;
        });
    } else {
        filasOperaciones = `
            <tr>
                <td colspan="2" style="border:1px solid #000; padding:4px 8px; text-align:center;">Sin operaciones registradas para este producto.</td>
            </tr>`;
    }

    const ventanaImpresion = window.open('', '_blank', 'width=900,height=700');
    if (!ventanaImpresion) {
        console.error('[imprimirOrdenFabricacion] No se pudo abrir la ventana de impresión.');
        alert('No se pudo abrir la ventana de impresión. Por favor, permita las ventanas emergentes.');
        return;
    }

    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Orden de Fabricación ${def(datos.idSubOr)}</title>
                <style>
                    @page {
                        size: A4;
                        margin: 10mm;
                    }
                    html, body {
                        margin: 0;
                        padding: 0;
                    }
                    body {
                        font-family: Arial, Helvetica, sans-serif;
                        font-size: 11pt;
                        color: #000;
                    }
                    .hoja-a4 {
                        width: 190mm;
                        min-height: 277mm;
                        box-sizing: border-box;
                        padding: 5mm;
                    }
                    .titulo {
                        text-align: center;
                        font-weight: bold;
                        font-size: 14pt;
                        margin-bottom: 8px;
                        text-transform: uppercase;
                    }
                    table {
                        border-collapse: collapse;
                        width: 100%;
                        margin-bottom: 10px;
                    }
                    th, td {
                        border: 1px solid #000;
                        padding: 5px 8px;
                        vertical-align: middle;
                    }
                    .etiqueta {
                        font-weight: bold;
                        background-color: #f0f0f0;
                    }
                    .subtitulo {
                        font-weight: bold;
                        margin-top: 12px;
                        margin-bottom: 4px;
                    }
                    .texto-fijo {
                        font-style: italic;
                    }
                </style>
            </head>
            <body>
                <div class="hoja-a4">
                    <div class="titulo">Orden de Fabricación</div>

                    <!-- 1° fila: Revisión, N° y Fecha -->
                    <table>
                        <tr>
                            <td class="etiqueta" style="width:15%">Revisión</td>
                            <td style="width:10%">${def(datos.revision)}</td>
                            <td class="etiqueta" style="width:15%">N°</td>
                            <td style="width:30%">${def(datos.idSubOr)}</td>
                            <td class="etiqueta" style="width:15%">Fecha</td>
                            <td style="width:15%">${def(datos.fechGenFormateada)}</td>
                        </tr>
                    </table>

                    <!-- 2° fila: Código, Denominación, Lote y texto fijo -->
                    <table>
                        <tr>
                            <td class="etiqueta" style="width:12%">Código</td>
                            <td style="width:23%">${def(datos.cod_Exte)}</td>
                            <td class="etiqueta" style="width:15%">Denominación</td>
                            <td style="width:25%">${def(datos.descrip)}</td>
                            <td class="etiqueta" style="width:10%">Lote</td>
                            <td style="width:15%">${def(datos.canProyFormateada)}</td>
                        </tr>
                        <tr>
                            <td class="etiqueta">Denominación-dimensiones</td>
                            <td colspan="5" class="texto-fijo">Según Vale de Materiales</td>
                        </tr>
                    </table>

                    <!-- 3° fila: Lanzó, Preparó/Aprobó y Fecha de revisión -->
                    <table>
                        <tr>
                            <td class="etiqueta" style="width:15%">Lanzó</td>
                            <td style="width:35%">${def(datos.nombreUsuario)}</td>
                            <td class="etiqueta" style="width:18%">Preparó y Aprobó</td>
                            <td style="width:17%">Jorge Gonzalez</td>
                            <td class="etiqueta" style="width:10%">Fecha</td>
                            <td style="width:15%">${def(datos.ferevisFormateada)}</td>
                        </tr>
                    </table>

                    <!-- Detalle de operaciones: SECOPER del producto -->
                    <div class="subtitulo">Detalle de operaciones</div>
                    <table>
                        <thead>
                            <tr>
                                <th class="etiqueta" style="width:15%">N° Op</th>
                                <th class="etiqueta">Descripción</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${filasOperaciones}
                        </tbody>
                    </table>
                </div>
            </body>
        </html>
    `);

    ventanaImpresion.document.close();

    setTimeout(function () {
        try {
            ventanaImpresion.focus();
            ventanaImpresion.print();
        } catch (error) {
            console.error('[imprimirOrdenFabricacion] Error al imprimir:', error);
        }
    }, 500);
};

/**
 * NOTA DE MANTENIMIENTO / JUNIOR:
 * El navegador bloquea window.open() si no se ejecuta directamente como respuesta a una
 * interacción del usuario. La función imprimirOrdenFabricacion() de arriba intenta abrir
 * una ventana + construir el HTML en el mismo paso, pero cuando la página Blazor hace
 * await de una llamada a la API entre el click y la invocación JS, el "user activation"
 * expira y el pop-up se bloquea (no aparece ni vista previa).
 *
 * Solución: separar en dos pasos:
 *   1) prepararVentanaImpresionOrden() -> abre la ventana en blanco inmediatamente con un
 *      mensaje de "Cargando..." y la guarda en window.ventanaOrdenFabricacion.
 *   2) imprimirOrdenFabricacionEnVentanaPreparada(datos) -> escribe el HTML en la ventana
 *      que ya está abierta y llama a print().
 *
 * El cliente debe llamar a (1) justo después del click, antes de hacer await de la API,
 * y luego a (2) una vez que reciba los datos.
 */

/**
 * Abre una ventana de impresión A4 en blanco, inmediatamente, para reservar el pop-up
 * mientras el cliente consulta la API. Devuelve true si se abrió, false si fue bloqueada.
 * @returns {boolean}
 */
window.prepararVentanaImpresionOrden = function () {
    const ventanaImpresion = window.open('', '_blank', 'width=900,height=700');

    if (!ventanaImpresion) {
        console.error('[prepararVentanaImpresionOrden] El navegador bloqueó la ventana de impresión.');
        alert('El navegador bloqueó la ventana de impresión.\n\n' +
              'Por favor permita las ventanas emergentes para este sitio.');
        return false;
    }

    // Guardar la referencia para que la segunda función pueda escribir sobre ella.
    window.ventanaOrdenFabricacion = ventanaImpresion;

    // Mostrar un mensaje de carga mientras Blazor trae los datos del servidor.
    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Cargando Orden de Fabricación...</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 100vh;
                        margin: 0;
                    }
                    .cargando {
                        font-size: 16pt;
                        color: #555;
                    }
                </style>
            </head>
            <body>
                <div class="cargando">Cargando datos para imprimir...</div>
            </body>
        </html>
    `);
    ventanaImpresion.document.close();
    return true;
};

/**
 * Escribe el contenido de la Orden de Fabricación en la ventana que ya abrió
 * prepararVentanaImpresionOrden() y dispara la impresión.
 * @param {Object} datos - DTO OrdenFabricacionImpresionDTO serializado.
 */
window.imprimirOrdenFabricacionEnVentanaPreparada = function (datos) {
    const ventanaImpresion = window.ventanaOrdenFabricacion;

    if (!ventanaImpresion || ventanaImpresion.closed) {
        console.error('[imprimirOrdenFabricacionEnVentanaPreparada] La ventana de impresión no está disponible.');
        alert('La ventana de impresión se cerró o no se abrió. Por favor intente de nuevo.');
        return;
    }

    // Normalizar valores vacíos para que el formulario no muestre "undefined" o "null".
    const def = function (valor, porDefecto) {
        return (valor === null || valor === undefined || valor === '') ? (porDefecto ?? '') : String(valor);
    };

    // Armar páginas de la Orden de Fabricación.
    // Cada hoja repite el encabezado y lleva a lo sumo FILAS_POR_PAGINA operaciones.
    // No se generan hojas adicionales con filas vacías.
    const FILAS_POR_PAGINA = 12;
    const todasOperaciones = (datos.operaciones && Array.isArray(datos.operaciones)) ? datos.operaciones : [];
    const paginas = [];
    for (let i = 0; i < todasOperaciones.length; i += FILAS_POR_PAGINA) {
        paginas.push(todasOperaciones.slice(i, i + FILAS_POR_PAGINA));
    }
    if (paginas.length === 0) {
        paginas.push([]);
    }

    function armarFilas(operaciones) {
        if (!operaciones || operaciones.length === 0) {
            return `<tr><td colspan="4" style="text-align:center;">Sin operaciones registradas para este producto.</td></tr>`;
        }
        let html = '';
        operaciones.forEach(function (op) {
            const desc = def(op.descOper);
            const obse = op.obseOper ? String(op.obseOper).replace(/\n/g, '<br>') : '';
            const contenidoDescripcion = obse ? `${desc}<br><br>${obse}` : desc;
            // Columnas fijas de Control y Fecha / Firma: se mantienen con bordes y renglones vacíos.
            html += `
                <tr>
                    <td style="width:10%">${def(op.numeOper)}</td>
                    <td>${contenidoDescripcion}</td>
                    <td class="linea-firma">&nbsp;</td>
                    <td class="linea-firma">&nbsp;</td>
                </tr>`;
        });
        return html;
    }

    function armarHoja(operacionesPagina, esUltima) {
        const claseUltima = esUltima ? 'ultima-hoja' : '';
        const filas = armarFilas(operacionesPagina);
        return `
            <div class="hoja-a4 ${claseUltima}">
                <div class="titulo">Orden de Fabricación</div>

                <!-- 1° fila: Revisión, N° y Fecha -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:15%">Revisión</td>
                        <td style="width:10%">${def(datos.revision)}</td>
                        <td class="etiqueta" style="width:15%">N°</td>
                        <td style="width:30%">${def(datos.idSubOr)}</td>
                        <td class="etiqueta" style="width:15%">Fecha</td>
                        <td style="width:15%">${def(datos.fechGenFormateada)}</td>
                    </tr>
                </table>

                <!-- 2° fila: Código, Denominación, Lote y texto fijo -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:12%">Código</td>
                        <td style="width:23%">${def(datos.cod_Exte)}</td>
                        <td class="etiqueta" style="width:15%">Denominación</td>
                        <td style="width:25%">${def(datos.descrip)}</td>
                        <td class="etiqueta" style="width:10%">Lote</td>
                        <td style="width:15%">${def(datos.canProyFormateada)}</td>
                    </tr>
                    <tr>
                        <td class="etiqueta">Denominación-dimensiones</td>
                        <td colspan="5" class="texto-fijo">Según Vale de Materiales</td>
                    </tr>
                </table>

                <!-- 3° fila: Lanzó, Preparó/Aprobó y Fecha de revisión -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:15%">Lanzó</td>
                        <td style="width:35%">${def(datos.nombreUsuario)}</td>
                        <td class="etiqueta" style="width:18%">Preparó y Aprobó</td>
                        <td style="width:17%">Jorge Gonzalez</td>
                        <td class="etiqueta" style="width:10%">Fecha</td>
                        <td style="width:15%">${def(datos.ferevisFormateada)}</td>
                    </tr>
                </table>

                <div class="subtitulo">Detalle de operaciones</div>
                <table class="detalle-operaciones">
                    <thead>
                        <tr>
                            <th style="width:10%">N° Op</th>
                            <th style="width:50%">Descripción</th>
                            <th style="width:20%">Control</th>
                            <th style="width:20%">Fecha / Firma</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${filas}
                    </tbody>
                </table>
            </div>
        `;
    }

    const hojasHtml = paginas.map(function (pagina, indice) {
        return armarHoja(pagina, indice === paginas.length - 1);
    }).join('');

    ventanaImpresion.document.open();
    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Orden de Fabricación ${def(datos.idSubOr)}</title>
                <style>
                    @page {
                        size: A4;
                        /* Márgenes: 1.5 cm izquierdo, 0.8 cm derecho, 1 cm arriba, 1 cm abajo */
                        margin: 10mm 8mm 10mm 15mm;
                    }
                    * {
                        box-sizing: border-box;
                    }
                    html, body {
                        margin: 0;
                        padding: 0;
                        width: 100%;
                    }
                    body {
                        font-family: Arial, Helvetica, sans-serif;
                        font-size: 11pt;
                        color: #000;
                    }
                    .hoja-a4 {
                        width: 100%;
                        height: 277mm;
                        border: 1px solid #000; /* marco exterior que ocupa toda la hoja */
                        padding: 4mm;
                        page-break-after: always;
                    }
                    .hoja-a4.ultima-hoja {
                        page-break-after: auto;
                    }
                    .titulo {
                        text-align: center;
                        font-weight: bold;
                        font-size: 14pt;
                        margin-bottom: 8px;
                        text-transform: uppercase;
                    }
                    .fila-formulario {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 6px;
                        /* auto evita que los anchos en porcentaje desborden el marco */
                        table-layout: auto;
                    }
                    .fila-formulario th, .fila-formulario td {
                        border: 1px solid #000;
                        padding: 4px 6px;
                        vertical-align: middle;
                        height: 28px;
                        /* permite romper textos largos y evitar que una celda se estire */
                        word-wrap: break-word;
                    }
                    .etiqueta {
                        font-weight: bold;
                        background-color: #f0f0f0;
                    }
                    .subtitulo {
                        font-weight: bold;
                        margin-top: 8px;
                        margin-bottom: 4px;
                    }
                    .detalle-operaciones {
                        width: 100%;
                        border-collapse: collapse;
                        table-layout: fixed;
                    }
                    .detalle-operaciones th, .detalle-operaciones td {
                        border: 1px solid #000;
                        padding: 4px 6px;
                        vertical-align: middle;
                        word-wrap: break-word;
                    }
                    .detalle-operaciones th {
                        background-color: #f0f0f0;
                        font-weight: bold;
                        height: 24px;
                    }
                    .detalle-operaciones td {
                        min-height: 24px;
                        height: auto;
                    }

                    .texto-fijo {
                        font-style: italic;
                    }
                </style>
            </head>
            <body>
                ${hojasHtml}
            </body>
        </html>
    `);
    ventanaImpresion.document.close();

    setTimeout(function () {
        try {
            ventanaImpresion.focus();
            ventanaImpresion.print();
        } catch (error) {
            console.error('[imprimirOrdenFabricacionEnVentanaPreparada] Error al imprimir:', error);
        }
    }, 500);
};

/**
 * Escribe el contenido del Vale de Materiales en la ventana ya abierta por prepararVentanaImpresionOrden()
 * y dispara la impresión.
 * @param {Object} datos - DTO OrdenFabricacionConsumoImpresionDTO serializado.
 */
window.imprimirConsumoMaterialesEnVentanaPreparada = function (datos) {
    const ventanaImpresion = window.ventanaOrdenFabricacion;

    if (!ventanaImpresion || ventanaImpresion.closed) {
        console.error('[imprimirConsumoMaterialesEnVentanaPreparada] La ventana de impresión no está disponible.');
        alert('La ventana de impresión se cerró o no se abrió. Por favor intente de nuevo.');
        return;
    }

    const def = function (valor, porDefecto) {
        return (valor === null || valor === undefined || valor === '') ? (porDefecto ?? '') : String(valor);
    };

    // Armar páginas del Vale de Materiales.
    const FILAS_POR_PAGINA = 12;
    const todosMateriales = (datos.materiales && Array.isArray(datos.materiales)) ? datos.materiales : [];
    const paginas = [];
    for (let i = 0; i < todosMateriales.length; i += FILAS_POR_PAGINA) {
        paginas.push(todosMateriales.slice(i, i + FILAS_POR_PAGINA));
    }
    if (paginas.length === 0) {
        paginas.push([]);
    }

    function armarFilas(materiales) {
        if (!materiales || materiales.length === 0) {
            return `<tr><td colspan="5" style="text-align:center;">Sin materiales reservados para esta orden.</td></tr>`;
        }
        let html = '';
        materiales.forEach(function (m) {
            html += `
                <tr>
                    <td style="width:20%">${def(m.cod_Exte)}</td>
                    <td style="width:35%">${def(m.descrip)}</td>
                    <td style="width:10%">${def(m.umedida)}</td>
                    <td style="width:15%">${def(m.usoUnitFormateada)}</td>
                    <td style="width:20%">${def(m.cantResFormateada)}</td>
                </tr>`;
        });
        return html;
    }

    function armarHoja(materialesPagina, esUltima) {
        const claseUltima = esUltima ? 'ultima-hoja' : '';
        const filas = armarFilas(materialesPagina);
        return `
            <div class="hoja-a4 ${claseUltima}">
                <div class="titulo">Orden de Fabricación</div>

                <!-- 1° fila: Revisión, N° y Fecha -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:15%">Revisión</td>
                        <td style="width:10%">${def(datos.revision)}</td>
                        <td class="etiqueta" style="width:15%">N°</td>
                        <td style="width:30%">${def(datos.idSubOr)}</td>
                        <td class="etiqueta" style="width:15%">Fecha</td>
                        <td style="width:15%">${def(datos.fechGenFormateada)}</td>
                    </tr>
                </table>

                <!-- 2° fila: Código, Denominación y Lote -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:12%">Código</td>
                        <td style="width:23%">${def(datos.cod_Exte)}</td>
                        <td class="etiqueta" style="width:15%">Denominación</td>
                        <td style="width:25%">${def(datos.descrip)}</td>
                        <td class="etiqueta" style="width:10%">Lote</td>
                        <td style="width:15%">${def(datos.canProyFormateada)}</td>
                    </tr>
                </table>

                <!-- 3° fila: título destacado del Vale de Materiales -->
                <table class="fila-formulario">
                    <tr>
                        <td colspan="6" class="titulo-vale">VALE DE MATERIALES</td>
                    </tr>
                </table>

                <!-- 4° fila: Lanzó, Preparó/Aprobó y Fecha de revisión -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:15%">Lanzó</td>
                        <td style="width:35%">${def(datos.nombreUsuario)}</td>
                        <td class="etiqueta" style="width:18%">Preparó y Aprobó</td>
                        <td style="width:17%">Jorge Gonzalez</td>
                        <td class="etiqueta" style="width:10%">Fecha</td>
                        <td style="width:15%">${def(datos.ferevisFormateada)}</td>
                    </tr>
                </table>

                <div class="subtitulo">MATERIALES</div>
                <table class="detalle-operaciones">
                    <thead>
                        <tr>
                            <th style="width:20%">Código</th>
                            <th style="width:35%">Descripción</th>
                            <th style="width:10%">U-M</th>
                            <th style="width:15%">Uso</th>
                            <th style="width:20%">Consumo</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${filas}
                    </tbody>
                </table>
            </div>
        `;
    }

    const hojasHtml = paginas.map(function (pagina, indice) {
        return armarHoja(pagina, indice === paginas.length - 1);
    }).join('');

    ventanaImpresion.document.open();
    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Vale de Materiales ${def(datos.idSubOr)}</title>
                <style>
                    @page {
                        size: A4;
                        margin: 10mm 8mm 10mm 15mm;
                    }
                    * {
                        box-sizing: border-box;
                    }
                    html, body {
                        margin: 0;
                        padding: 0;
                        width: 100%;
                    }
                    body {
                        font-family: Arial, Helvetica, sans-serif;
                        font-size: 11pt;
                        color: #000;
                    }
                    .hoja-a4 {
                        width: 100%;
                        height: 277mm;
                        border: 1px solid #000;
                        padding: 4mm;
                        page-break-after: always;
                    }
                    .hoja-a4.ultima-hoja {
                        page-break-after: auto;
                    }
                    .titulo {
                        text-align: center;
                        font-weight: bold;
                        font-size: 14pt;
                        margin-bottom: 8px;
                        text-transform: uppercase;
                    }
                    .titulo-vale {
                        text-align: center;
                        font-weight: bold;
                        font-size: 18pt;
                        text-transform: uppercase;
                        padding: 8px;
                        background-color: #f0f0f0;
                    }
                    .fila-formulario {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 6px;
                        table-layout: auto;
                    }
                    .fila-formulario th, .fila-formulario td {
                        border: 1px solid #000;
                        padding: 4px 6px;
                        vertical-align: middle;
                        height: 28px;
                        word-wrap: break-word;
                    }
                    .etiqueta {
                        font-weight: bold;
                        background-color: #f0f0f0;
                    }
                    .subtitulo {
                        font-weight: bold;
                        margin-top: 8px;
                        margin-bottom: 4px;
                    }
                    .detalle-operaciones {
                        width: 100%;
                        border-collapse: collapse;
                        table-layout: fixed;
                    }
                    .detalle-operaciones th, .detalle-operaciones td {
                        border: 1px solid #000;
                        padding: 4px 6px;
                        vertical-align: middle;
                        word-wrap: break-word;
                    }
                    .detalle-operaciones th {
                        background-color: #f0f0f0;
                        font-weight: bold;
                        height: 24px;
                    }
                    .detalle-operaciones td {
                        min-height: 24px;
                        height: auto;
                    }
                    .texto-fijo {
                        font-style: italic;
                    }
                </style>
            </head>
            <body>
                ${hojasHtml}
            </body>
        </html>
    `);
    ventanaImpresion.document.close();

    setTimeout(function () {
        try {
            ventanaImpresion.focus();
            ventanaImpresion.print();
        } catch (error) {
            console.error('[imprimirConsumoMaterialesEnVentanaPreparada] Error al imprimir:', error);
        }
    }, 500);
};

/**
 * Escribe múltiples documentos de una Orden de Fabricación en la ventana ya preparada,
 * concatenando cada documento con salto de página. De este modo se evita abrir varias
 * ventanas emergentes cuando el usuario elige más de una opción en el modal.
 * @param {Array} documentos - Cada elemento debe tener { tipo, datos }.
 *                             tipo: 'orden-fabricacion' | 'consumo-materiales'.
 */
window.escribirDocumentosEnVentanaPreparada = function (documentos) {
    const ventanaImpresion = window.ventanaOrdenFabricacion;

    if (!ventanaImpresion || ventanaImpresion.closed) {
        console.error('[escribirDocumentosEnVentanaPreparada] La ventana de impresión no está disponible.');
        alert('La ventana de impresión se cerró o no se abrió. Por favor intente de nuevo.');
        return;
    }

    const def = function (valor, porDefecto) {
        return (valor === null || valor === undefined || valor === '') ? (porDefecto ?? '') : String(valor);
    };

    function armarFilasOrden(operaciones) {
        if (!operaciones || operaciones.length === 0) {
            return `<tr><td colspan="4" style="text-align:center;">Sin operaciones registradas para este producto.</td></tr>`;
        }
        let html = '';
        operaciones.forEach(function (op) {
            const desc = def(op.descOper);
            const obse = op.obseOper ? String(op.obseOper).replace(/\n/g, '<br>') : '';
            const contenidoDescripcion = obse ? `${desc}<br><br>${obse}` : desc;
            // Cada fila de operación tiene altura fija para alojar 6 líneas de texto
            // en las columnas Control y Fecha/Firma, con fuente pequeña.
            // Columna CONTROL: campos Maq., Oper., Cant., Inic., Fin. para llenar a mano.
            // Columna FECHA/FIRMA: línea de fecha (dd / mm / aaaa) y línea de firma.
            html += `
                <tr style="height:6em;">
                    <td style="width:10%; vertical-align:top;">${def(op.numeOper)}</td>
                    <td style="vertical-align:top;">${contenidoDescripcion}</td>
                    <td style="vertical-align:top; font-size:8pt; line-height:1.4;">
                        Maq.:___________<br>
                        Oper.:___________<br>
                        Cant.:___________<br>
                        Inic.:____________<br>
                        Fin:&nbsp;&nbsp;____________
                    </td>
                    <td style="vertical-align:top; font-size:8pt; line-height:1.4; text-align:center;">
                        <span style="display:inline-block; width:30%; text-align:center;">&nbsp;</span><span style="display:inline-block; width:5%; text-align:center;">/</span><span style="display:inline-block; width:30%; text-align:center;">&nbsp;</span><span style="display:inline-block; width:5%; text-align:center;">/</span><span style="display:inline-block; width:30%; text-align:center;">&nbsp;</span><br>
                        <br>
                        <span style="display:block; border-bottom:1px solid #000; width:100%;">&nbsp;</span>
                    </td>
                </tr>`;
        });
        return html;
    }

    function armarHojasOrden(datos) {
        const FILAS_POR_PAGINA = 8;
        const operaciones = (datos.operaciones && Array.isArray(datos.operaciones)) ? datos.operaciones : [];
        const paginas = [];
        for (let i = 0; i < operaciones.length; i += FILAS_POR_PAGINA) {
            paginas.push(operaciones.slice(i, i + FILAS_POR_PAGINA));
        }
        if (paginas.length === 0) {
            paginas.push([]);
        }

        let html = '';
        paginas.forEach(function (pagina, indice) {
            const esUltima = indice === paginas.length - 1;
            const claseUltima = esUltima ? 'ultima-hoja' : '';
            const filas = armarFilasOrden(pagina);
            html += `
                <div class="hoja-a4 ${claseUltima}">
                    <div class="titulo">Orden de Fabricación</div>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:15%">Revisión</td>
                            <td style="width:10%">${def(datos.revision)}</td>
                            <td class="etiqueta" style="width:15%">N°</td>
                            <td style="width:30%">${def(datos.idSubOr)}</td>
                            <td class="etiqueta" style="width:15%">Fecha</td>
                            <td style="width:15%">${def(datos.fechGenFormateada)}</td>
                        </tr>
                    </table>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:12%">Código</td>
                            <td style="width:23%">${def(datos.cod_Exte)}</td>
                            <td class="etiqueta" style="width:15%">Denominación</td>
                            <td style="width:25%">${def(datos.descrip)}</td>
                            <td class="etiqueta" style="width:10%">Lote</td>
                            <td style="width:15%">${def(datos.canProyFormateada)}</td>
                        </tr>
                        <tr>
                            <td class="etiqueta">Denominación-dimensiones</td>
                            <td colspan="5" class="texto-fijo">Según Vale de Materiales</td>
                        </tr>
                    </table>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:15%">Lanzó</td>
                            <td style="width:35%">${def(datos.nombreUsuario)}</td>
                            <td class="etiqueta" style="width:18%">Preparó y Aprobó</td>
                            <td style="width:17%">Jorge Gonzalez</td>
                            <td class="etiqueta" style="width:10%">Fecha</td>
                            <td style="width:15%">${def(datos.ferevisFormateada)}</td>
                        </tr>
                    </table>

                    <div class="subtitulo">Detalle de operaciones</div>
                    <table class="detalle-of">
                        <thead>
                            <tr>
                                <th style="width:10%">N° Op</th>
                                <th style="width:50%">Descripción</th>
                                <th style="width:20%">Control</th>
                                <th style="width:20%">Fecha / Firma</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${filas}
                        </tbody>
                    </table>
                </div>
            `;
        });
        return html;
    }

    function armarFilasConsumo(materiales) {
        if (!materiales || materiales.length === 0) {
            return `<tr><td colspan="5" style="text-align:center;">Sin materiales reservados para esta orden.</td></tr>`;
        }
        let html = '';
        materiales.forEach(function (m) {
            html += `
                <tr>
                    <td style="width:20%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${def(m.cod_Exte)}</td>
                    <td style="width:57%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${def(m.descrip)}</td>
                    <td style="width:4%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${def(m.umedida)}</td>
                    <td style="width:9%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; text-align:right;">${def(m.usoUnitFormateada)}</td>
                    <td style="width:10%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; text-align:right;">${def(m.cantResFormateada)}</td>
                </tr>`;
        });
        return html;
    }

    function armarHojasConsumo(datos) {
        const FILAS_POR_PAGINA = 32;
        const materiales = (datos.materiales && Array.isArray(datos.materiales)) ? datos.materiales : [];
        const paginas = [];
        for (let i = 0; i < materiales.length; i += FILAS_POR_PAGINA) {
            paginas.push(materiales.slice(i, i + FILAS_POR_PAGINA));
        }
        if (paginas.length === 0) {
            paginas.push([]);
        }

        let html = '';
        paginas.forEach(function (pagina, indice) {
            const esUltima = indice === paginas.length - 1;
            const claseUltima = esUltima ? 'ultima-hoja' : '';
            const filas = armarFilasConsumo(pagina);
            html += `
                <div class="hoja-a4 ${claseUltima}">
                    <div class="titulo">Orden de Fabricación</div>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:15%">Revisión</td>
                            <td style="width:10%">${def(datos.revision)}</td>
                            <td class="etiqueta" style="width:15%">N°</td>
                            <td style="width:30%">${def(datos.idSubOr)}</td>
                            <td class="etiqueta" style="width:15%">Fecha</td>
                            <td style="width:15%">${def(datos.fechGenFormateada)}</td>
                        </tr>
                    </table>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:12%">Código</td>
                            <td style="width:23%">${def(datos.cod_Exte)}</td>
                            <td class="etiqueta" style="width:15%">Denominación</td>
                            <td style="width:25%">${def(datos.descrip)}</td>
                            <td class="etiqueta" style="width:10%">Lote</td>
                            <td style="width:15%">${def(datos.canProyFormateada)}</td>
                        </tr>
                    </table>

                    <table class="fila-formulario">
                        <tr>
                            <td colspan="6" class="titulo-vale">VALE DE MATERIALES</td>
                        </tr>
                    </table>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:15%">Lanzó</td>
                            <td style="width:35%">${def(datos.nombreUsuario)}</td>
                            <td class="etiqueta" style="width:18%">Preparó y Aprobó</td>
                            <td style="width:17%">Jorge Gonzalez</td>
                            <td class="etiqueta" style="width:10%">Fecha</td>
                            <td style="width:15%">${def(datos.ferevisFormateada)}</td>
                        </tr>
                    </table>

                    <div class="subtitulo">MATERIALES</div>
                    <table class="detalle-consumo">
                        <thead>
                            <tr>
                                <th style="width:20%">Código</th>
                                <th style="width:57%">Descripción</th>
                                <th style="width:4%">U-M</th>
                                <th style="width:9%">Uso</th>
                                <th style="width:10%">Consumo</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${filas}
                        </tbody>
                    </table>
                </div>
            `;
        });
        return html;
    }

    function armarFilasPlan(items) {
        if (!items || items.length === 0) {
            return `<tr><td colspan="3">&nbsp;</td></tr>`;
        }
        let html = '';
        items.forEach(function (item) {
            // Se imprime la fila tal cual, incluso si sus 3 campos están vacíos
            // (fila usada como separador visual entre bloques del plan).
            html += `
                <tr>
                    <td style="width:28%">${def(item.dimSegunPlano)}</td>
                    <td style="width:22%">${def(item.instrumento)}</td>
                    <td style="width:50%">${def(item.observaciones)}</td>
                </tr>`;
        });
        return html;
    }

    function armarHojasPlan(datos) {
        const FILAS_POR_PAGINA = 28;
        const todosItems = (datos.items && Array.isArray(datos.items)) ? datos.items : [];
        const paginas = [];
        for (let i = 0; i < todosItems.length; i += FILAS_POR_PAGINA) {
            paginas.push(todosItems.slice(i, i + FILAS_POR_PAGINA));
        }
        if (paginas.length === 0) {
            paginas.push([]);
        }

        const totalHojas = paginas.length;
        let html = '';
        paginas.forEach(function (pagina, indice) {
            const numeroHoja = indice + 1;
            const esUltima = indice === totalHojas - 1;
            const claseUltima = esUltima ? 'ultima-hoja' : '';
            const filas = armarFilasPlan(pagina);
            const observaciones = def(datos.observaciones).replace(/\n/g, '<br>');
            html += `
                <div class="hoja-a4 ${claseUltima}">
                    <table class="encabezado-titulo">
                        <tr>
                            <td class="celda-logo">
                                <div class="logo-empresa">metales</div>
                                <div class="logo-marca">ARCON</div>
                            </td>
                            <td class="celda-titulo">
                                <div class="titulo" style="margin-bottom:0;">PLAN DE INSPECCIÓN</div>
                            </td>
                            <td class="celda-fecha-hoja">
                                <table class="tabla-fecha-hoja">
                                    <tr>
                                        <td class="etiqueta">FECHA</td>
                                        <td>${def(datos.fechaImpresionFormateada)}</td>
                                    </tr>
                                    <tr>
                                        <td class="etiqueta">HOJA</td>
                                        <td>${numeroHoja} DE ${totalHojas}</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:10%">Denominación</td>
                            <td style="width:38%">${def(datos.denominacion)}</td>
                            <td class="etiqueta" style="width:9%">Código</td>
                            <td style="width:13%">${def(datos.codigo)}</td>
                            <td class="etiqueta" style="width:5%">Rev.</td>
                            <td style="width:6%">${def(datos.revision)}</td>
                            <td class="etiqueta" style="width:8%">Origen</td>
                            <td style="width:11%">${def(datos.origen)}</td>
                        </tr>
                    </table>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:8%">AQL</td>
                            <td style="width:8%">${def(datos.aql)}</td>
                            <td class="etiqueta" style="width:14%">Nivel (%)</td>
                            <td style="width:10%">${def(datos.nivelPct)}</td>
                            <td class="etiqueta" style="width:18%">Frecuencia de Inspección</td>
                            <td style="width:20%">${def(datos.frecInspeccion)}</td>
                            <td class="etiqueta" style="width:11%">Puntos de Detención Obligatoria</td>
                            <td style="width:11%">${def(datos.puntoDetencion)}</td>
                        </tr>
                    </table>

                    <table class="fila-formulario">
                        <tr>
                            <td class="etiqueta" style="width:12%">Preparó</td>
                            <td style="width:38%">${def(datos.nombrePreparo)}</td>
                            <td class="etiqueta" style="width:12%">Aprobó</td>
                            <td style="width:26%">${def(datos.nombreAprobo)}</td>
                            <td class="etiqueta" style="width:6%">Fecha</td>
                            <td style="width:6%">${def(datos.fechaAproboFormateada)}</td>
                        </tr>
                    </table>

                    <div class="contenido-detalle">
                        <table class="detalle-plan">
                            <thead>
                                <tr>
                                    <th style="width:28%">Dimensión s/Plano</th>
                                    <th style="width:22%">Instrumento</th>
                                    <th style="width:50%">Resultado</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${filas}
                            </tbody>
                        </table>
                    </div>

                    <div class="pie-observaciones">
                        <span class="etiqueta-pie">Observaciones:</span>
                        <span class="texto-pie">${observaciones}</span>
                    </div>
                </div>
            `;
        });
        return html;
    }

    let hojasHtml = '';
    if (!documentos || !Array.isArray(documentos) || documentos.length === 0) {
        console.error('[escribirDocumentosEnVentanaPreparada] No se recibieron documentos para imprimir.');
        return;
    }

    documentos.forEach(function (doc) {
        if (doc.tipo === 'orden-fabricacion') {
            hojasHtml += armarHojasOrden(doc.datos);
        } else if (doc.tipo === 'consumo-materiales') {
            hojasHtml += armarHojasConsumo(doc.datos);
        } else if (doc.tipo === 'plan-inspeccion') {
            hojasHtml += armarHojasPlan(doc.datos);
        } else {
            console.warn('[escribirDocumentosEnVentanaPreparada] Tipo de documento desconocido:', doc.tipo);
        }
    });

    if (!hojasHtml) {
        console.error('[escribirDocumentosEnVentanaPreparada] No se generó contenido para imprimir.');
        return;
    }

    ventanaImpresion.document.open();
    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Documentos de Orden ${def(documentos[0].datos.idSubOr)}</title>
                <style>
                    @page {
                        size: A4;
                        margin: 10mm 8mm 10mm 15mm;
                    }
                    * {
                        box-sizing: border-box;
                    }
                    html, body {
                        margin: 0;
                        padding: 0;
                        width: 100%;
                    }
                    body {
                        font-family: Arial, Helvetica, sans-serif;
                        font-size: 11pt;
                        color: #000;
                    }
                    .hoja-a4 {
                        width: 100%;
                        height: 277mm;
                        border: 1px solid #000;
                        padding: 4mm;
                        page-break-after: always;
                        display: flex;
                        flex-direction: column;
                    }
                    .hoja-a4.ultima-hoja {
                        page-break-after: auto;
                    }
                    .titulo {
                        text-align: center;
                        font-weight: bold;
                        font-size: 14pt;
                        margin-bottom: 8px;
                        text-transform: uppercase;
                    }
                    .encabezado-titulo {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 6px;
                    }
                    .encabezado-titulo td {
                        border: 1px solid #000;
                        vertical-align: middle;
                    }
                    .celda-logo {
                        width: 15%;
                        text-align: center;
                        padding: 2px;
                    }
                    .logo-empresa {
                        font-size: 8pt;
                        color: #555;
                    }
                    .logo-marca {
                        font-weight: bold;
                        font-size: 14pt;
                        letter-spacing: 1px;
                    }
                    .celda-titulo {
                        width: 60%;
                        text-align: center;
                    }
                    .celda-fecha-hoja {
                        width: 25%;
                        padding: 0;
                    }
                    .tabla-fecha-hoja {
                        width: 100%;
                        height: 100%;
                        border-collapse: collapse;
                    }
                    .tabla-fecha-hoja td {
                        border: 1px solid #000;
                        padding: 2px 6px;
                        font-size: 9pt;
                    }
                    .contenido-detalle {
                        flex: 1 1 auto;
                    }
                    .pie-observaciones {
                        flex: 0 0 auto;
                        border: 1px solid #000;
                        border-top: none;
                        padding: 4px 6px;
                        font-size: 9pt;
                        min-height: 14mm;
                    }
                    .etiqueta-pie {
                        font-weight: bold;
                        margin-right: 6px;
                    }
                    .titulo-vale {
                        text-align: center;
                        font-weight: bold;
                        font-size: 18pt;
                        text-transform: uppercase;
                        padding: 8px;
                        background-color: #f0f0f0;
                    }
                    .fila-formulario {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 6px;
                        table-layout: auto;
                    }
                    .fila-formulario th, .fila-formulario td {
                        border: 1px solid #000;
                        padding: 4px 6px;
                        vertical-align: middle;
                        height: 28px;
                        word-wrap: break-word;
                    }
                    .etiqueta {
                        font-weight: bold;
                        background-color: #f0f0f0;
                    }
                    .subtitulo {
                        font-weight: bold;
                        margin-top: 8px;
                        margin-bottom: 4px;
                    }
                    .detalle-of,
                    .detalle-consumo,
                    .detalle-plan {
                        width: 100%;
                        border-collapse: collapse;
                        table-layout: fixed;
                    }
                    .detalle-of th, .detalle-of td,
                    .detalle-consumo th, .detalle-consumo td,
                    .detalle-plan th, .detalle-plan td {
                        border: 1px solid #000;
                        padding: 2px 6px;
                        vertical-align: middle;
                        word-wrap: break-word;
                    }
                    .detalle-of th,
                    .detalle-consumo th,
                    .detalle-plan th {
                        background-color: #f0f0f0;
                        font-weight: bold;
                        height: 20px;
                    }
                    .detalle-of td,
                    .detalle-consumo td,
                    .detalle-plan td {
                        min-height: 20px;
                        height: auto;
                    }
                    .detalle-plan th {
                        text-align: center;
                    }
                    .texto-fijo {
                        font-style: italic;
                    }
                </style>
            </head>
            <body>
                ${hojasHtml}
            </body>
        </html>
    `);
    ventanaImpresion.document.close();

    setTimeout(function () {
        try {
            ventanaImpresion.focus();
            ventanaImpresion.print();
        } catch (error) {
            console.error('[escribirDocumentosEnVentanaPreparada] Error al imprimir:', error);
        }
    }, 500);
};

/**
 * Abre una ventana de impresión A4 en blanco, inmediatamente, para reservar el pop-up
 * mientras el cliente consulta la API del Plan de Inspección. Devuelve true si se abrió,
 * false si fue bloqueada por el navegador.
 * @returns {boolean}
 */
window.prepararVentanaImpresionPlanInspeccion = function () {
    const ventanaImpresion = window.open('', '_blank', 'width=900,height=700');

    if (!ventanaImpresion) {
        console.error('[prepararVentanaImpresionPlanInspeccion] El navegador bloqueó la ventana de impresión.');
        alert('El navegador bloqueó la ventana de impresión.\n\n' +
              'Por favor permita las ventanas emergentes para este sitio.');
        return false;
    }

    // Guardar la referencia para que la segunda función pueda escribir sobre ella.
    window.ventanaPlanInspeccion = ventanaImpresion;

    // Mostrar un mensaje de carga mientras Blazor trae los datos del servidor.
    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Cargando Plan de Inspección...</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 100vh;
                        margin: 0;
                    }
                    .cargando {
                        font-size: 16pt;
                        color: #555;
                    }
                </style>
            </head>
            <body>
                <div class="cargando">Cargando datos para imprimir...</div>
            </body>
        </html>
    `);
    ventanaImpresion.document.close();
    return true;
};

/**
 * Escribe el contenido del Plan de Inspección y Ensayo en la ventana ya abierta por
 * prepararVentanaImpresionPlanInspeccion() y dispara la impresión.
 * Reproduce el formulario "PLAN DE INSPECCION" en papel A4: encabezado del producto
 * (Denominación/Código/Revisión), encabezado del plan (AQL/Nivel/Frecuencia/Puntos de
 * Detención/Preparó/Aprobó) y el detalle PLAINSP (Dimensión s/Plano, Instrumento, Resultado).
 * Si el detalle excede una hoja se generan hojas adicionales repitiendo todo el encabezado,
 * y se preservan las filas vacías del detalle porque el usuario las usa como separador visual.
 * @param {Object} datos - DTO PlanInspeccionImpresionDTO serializado.
 */
window.imprimirPlanInspeccionEnVentanaPreparada = function (datos) {
    const ventanaImpresion = window.ventanaPlanInspeccion;

    if (!ventanaImpresion || ventanaImpresion.closed) {
        console.error('[imprimirPlanInspeccionEnVentanaPreparada] La ventana de impresión no está disponible.');
        alert('La ventana de impresión se cerró o no se abrió. Por favor intente de nuevo.');
        return;
    }

    const def = function (valor, porDefecto) {
        return (valor === null || valor === undefined || valor === '') ? (porDefecto ?? '') : String(valor);
    };

    // Armar páginas del detalle. No se filtran ni se colapsan filas vacías: son separadores
    // visuales intencionales del usuario en la grilla PLAINSP. Tampoco se agregan filas de
    // relleno adicionales; solo se pagina cuando el detalle real excede una hoja.
    const FILAS_POR_PAGINA = 28;
    const todosItems = (datos.items && Array.isArray(datos.items)) ? datos.items : [];
    const paginas = [];
    for (let i = 0; i < todosItems.length; i += FILAS_POR_PAGINA) {
        paginas.push(todosItems.slice(i, i + FILAS_POR_PAGINA));
    }
    if (paginas.length === 0) {
        paginas.push([]);
    }

    function armarFilas(items) {
        if (!items || items.length === 0) {
            return `<tr><td colspan="3">&nbsp;</td></tr>`;
        }
        let html = '';
        items.forEach(function (item) {
            // Se imprime la fila tal cual, incluso si sus 3 campos están vacíos
            // (fila usada como separador visual entre bloques del plan).
            html += `
                <tr>
                    <td style="width:28%">${def(item.dimSegunPlano)}</td>
                    <td style="width:22%">${def(item.instrumento)}</td>
                    <td style="width:50%">${def(item.observaciones)}</td>
                </tr>`;
        });
        return html;
    }

    function armarHoja(itemsPagina, numeroHoja, totalHojas, esUltima) {
        const claseUltima = esUltima ? 'ultima-hoja' : '';
        const filas = armarFilas(itemsPagina);
        const observaciones = def(datos.observaciones).replace(/\n/g, '<br>');
        return `
            <div class="hoja-a4 ${claseUltima}">
                <table class="encabezado-titulo">
                    <tr>
                        <td class="celda-logo">
                            <div class="logo-empresa">metales</div>
                            <div class="logo-marca">ARCON</div>
                        </td>
                        <td class="celda-titulo">
                            <div class="titulo">PLAN DE INSPECCIÓN</div>
                        </td>
                        <td class="celda-fecha-hoja">
                            <table class="tabla-fecha-hoja">
                                <tr>
                                    <td class="etiqueta">FECHA</td>
                                    <td>${def(datos.fechaImpresionFormateada)}</td>
                                </tr>
                                <tr>
                                    <td class="etiqueta">HOJA</td>
                                    <td>${numeroHoja} DE ${totalHojas}</td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>

                <!-- Encabezado del producto: Denominación, Código, Revisión y Origen -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:10%">Denominación</td>
                        <td style="width:38%">${def(datos.denominacion)}</td>
                        <td class="etiqueta" style="width:9%">Código</td>
                        <td style="width:13%">${def(datos.codigo)}</td>
                        <td class="etiqueta" style="width:5%">Rev.</td>
                        <td style="width:6%">${def(datos.revision)}</td>
                        <td class="etiqueta" style="width:8%">Origen</td>
                        <td style="width:11%">${def(datos.origen)}</td>
                    </tr>
                </table>

                <!-- Encabezado del plan: AQL, Nivel, Frecuencia y Puntos de Detención -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:8%">AQL</td>
                        <td style="width:8%">${def(datos.aql)}</td>
                        <td class="etiqueta" style="width:14%">Nivel (%)</td>
                        <td style="width:10%">${def(datos.nivelPct)}</td>
                        <td class="etiqueta" style="width:18%">Frecuencia de Inspección</td>
                        <td style="width:20%">${def(datos.frecInspeccion)}</td>
                        <td class="etiqueta" style="width:11%">Puntos de Detención Obligatoria</td>
                        <td style="width:11%">${def(datos.puntoDetencion)}</td>
                    </tr>
                </table>

                <!-- Preparó / Aprobó / Fecha -->
                <table class="fila-formulario">
                    <tr>
                        <td class="etiqueta" style="width:12%">Preparó</td>
                        <td style="width:38%">${def(datos.nombrePreparo)}</td>
                        <td class="etiqueta" style="width:12%">Aprobó</td>
                        <td style="width:26%">${def(datos.nombreAprobo)}</td>
                        <td class="etiqueta" style="width:6%">Fecha</td>
                        <td style="width:6%">${def(datos.fechaAproboFormateada)}</td>
                    </tr>
                </table>

                <!-- El detalle crece dentro de un contenedor flexible para que el pie de
                     Observaciones quede siempre al final de la hoja, incluso si el detalle
                     tiene pocas filas. -->
                <div class="contenido-detalle">
                    <table class="detalle-plan">
                        <thead>
                            <tr>
                                <th style="width:28%">Dimensión s/Plano</th>
                                <th style="width:22%">Instrumento</th>
                                <th style="width:50%">Resultado</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${filas}
                        </tbody>
                    </table>
                </div>

                <!-- Observaciones generales del plan (Codiproc.Observaciones): se repiten
                     al pie de cada hoja generada. -->
                <div class="pie-observaciones">
                    <span class="etiqueta-pie">Observaciones:</span>
                    <span class="texto-pie">${observaciones}</span>
                </div>
            </div>
        `;
    }

    const totalHojas = paginas.length;
    const hojasHtml = paginas.map(function (pagina, indice) {
        return armarHoja(pagina, indice + 1, totalHojas, indice === totalHojas - 1);
    }).join('');

    ventanaImpresion.document.open();
    ventanaImpresion.document.write(`
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <title>Plan de Inspección ${def(datos.codigo)}</title>
                <style>
                    @page {
                        size: A4;
                        margin: 10mm 8mm 10mm 15mm;
                    }
                    * {
                        box-sizing: border-box;
                    }
                    html, body {
                        margin: 0;
                        padding: 0;
                        width: 100%;
                    }
                    body {
                        font-family: Arial, Helvetica, sans-serif;
                        font-size: 10pt;
                        color: #000;
                    }
                    .hoja-a4 {
                        width: 100%;
                        height: 277mm;
                        border: 1px solid #000;
                        padding: 4mm;
                        page-break-after: always;
                        display: flex;
                        flex-direction: column;
                    }
                    .hoja-a4.ultima-hoja {
                        page-break-after: auto;
                    }
                    .encabezado-titulo {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 6px;
                    }
                    .encabezado-titulo td {
                        border: 1px solid #000;
                        vertical-align: middle;
                    }
                    .celda-logo {
                        width: 15%;
                        text-align: center;
                        padding: 2px;
                    }
                    .logo-empresa {
                        font-size: 8pt;
                        color: #555;
                    }
                    .logo-marca {
                        font-weight: bold;
                        font-size: 14pt;
                        letter-spacing: 1px;
                    }
                    .celda-titulo {
                        width: 60%;
                        text-align: center;
                    }
                    .titulo {
                        font-weight: bold;
                        font-size: 16pt;
                        text-transform: uppercase;
                    }
                    .celda-fecha-hoja {
                        width: 25%;
                        padding: 0;
                    }
                    .tabla-fecha-hoja {
                        width: 100%;
                        height: 100%;
                        border-collapse: collapse;
                    }
                    .tabla-fecha-hoja td {
                        border: 1px solid #000;
                        padding: 2px 6px;
                        font-size: 9pt;
                    }
                    .fila-formulario {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 6px;
                        table-layout: auto;
                    }
                    .fila-formulario td {
                        border: 1px solid #000;
                        padding: 4px 6px;
                        vertical-align: middle;
                        word-wrap: break-word;
                        font-size: 9pt;
                    }
                    .etiqueta {
                        font-weight: bold;
                        background-color: #f0f0f0;
                    }
                    .contenido-detalle {
                        flex: 1 1 auto;
                    }
                    .detalle-plan {
                        width: 100%;
                        border-collapse: collapse;
                        table-layout: fixed;
                    }
                    .detalle-plan th, .detalle-plan td {
                        border: 1px solid #000;
                        padding: 3px 6px;
                        vertical-align: middle;
                        word-wrap: break-word;
                        font-size: 9pt;
                        height: 20px;
                    }
                    .detalle-plan th {
                        background-color: #f0f0f0;
                        font-weight: bold;
                        text-align: center;
                    }
                    .pie-observaciones {
                        flex: 0 0 auto;
                        border: 1px solid #000;
                        border-top: none;
                        padding: 4px 6px;
                        font-size: 9pt;
                        min-height: 14mm;
                    }
                    .etiqueta-pie {
                        font-weight: bold;
                        margin-right: 6px;
                    }
                </style>
            </head>
            <body>
                ${hojasHtml}
            </body>
        </html>
    `);
    ventanaImpresion.document.close();

    setTimeout(function () {
        try {
            ventanaImpresion.focus();
            ventanaImpresion.print();
        } catch (error) {
            console.error('[imprimirPlanInspeccionEnVentanaPreparada] Error al imprimir:', error);
        }
    }, 500);
};

// Hace scroll al final de un elemento por su id.
// Usado para posicionar la vista al final de la tabla de anotaciones ANOTAPED.
window.scrollToBottom = function (id) {
    const el = document.getElementById(id);
    if (el) {
        el.scrollTop = el.scrollHeight;
        // También hacer scroll al final de la página para mostrar el último grupo
        el.scrollIntoView({ behavior: 'smooth', block: 'end' });
    }
};
