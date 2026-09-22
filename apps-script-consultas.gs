/**
 * Fuerza Natural - Receptor de consultas del sitio web
 * Guarda cada consulta en la hoja "Consultas" y te avisa por mail.
 */
var MAIL_AVISO = 'contacto@fuerzanaturalbroker.com';

// Sirve para probar que el script está activo (abrir la URL en el navegador).
function doGet() {
  return ContentService.createTextOutput('El receptor de consultas está activo.');
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Consultas') || ss.insertSheet('Consultas');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Fecha', 'Nombre', 'Teléfono', 'Seguro', 'Detalles', 'Origen', 'Medio', 'Campaña', 'Página']);
    }

    var detalles = Object.keys(d.detalles || {})
      .filter(function (k) { return d.detalles[k]; })
      .map(function (k) { return k + ': ' + d.detalles[k]; })
      .join(' | ');

    // Limita el largo y evita que un texto se interprete como fórmula.
    function limpio(v) {
      v = String(v || '').slice(0, 500);
      return /^[=+\-@]/.test(v) ? "'" + v : v;
    }

    var fila = [d.nombre, d.telefono, d.ramo, detalles, d.origen, d.medio, d.campana, d.pagina].map(limpio);
    sh.appendRow([new Date()].concat(fila));

    MailApp.sendEmail({
      to: MAIL_AVISO,
      subject: 'Nueva consulta de ' + limpio(d.ramo) + ': ' + limpio(d.nombre),
      body: 'Nombre: ' + fila[0] + '\n' +
            'Teléfono: ' + fila[1] + '\n' +
            'Seguro: ' + fila[2] + '\n' +
            'Detalles: ' + fila[3] + '\n' +
            'Origen: ' + fila[4] + '\n\n' +
            'Todas las consultas están en tu hoja de Google, pestaña "Consultas".'
    });
  } catch (err) {
    // Si algo falla, no se interrumpe la experiencia del cliente.
  }
  return ContentService.createTextOutput('ok');
}
