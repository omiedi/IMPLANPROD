
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