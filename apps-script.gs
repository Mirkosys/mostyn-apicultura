/**
 * Mostyn Apicultura — Endpoint de pedidos
 *
 * Recibe POST del formulario de contacto y:
 *  1. Guarda el pedido como fila en la planilla "Pedidos"
 *  2. (Opcional) envía notificación por email
 *
 * Instalación: pegá este código en Extensions → Apps Script
 * desde una Google Sheet vacía, y publicá como Web App.
 */

const SHEET_NAME = 'Pedidos';
const NOTIFICATION_EMAIL = 'mostynmieles@gmail.com'; // poné '' para desactivar emails

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Crear la hoja con headers si no existe
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Fecha', 'Nombre', 'Email', 'Consulta', 'Mensaje', 'Estado', 'Notas internas'
      ]);
      sheet.getRange(1, 1, 1, 7)
        .setBackground('#F5A623')
        .setFontColor('#111111')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 140);
      sheet.setColumnWidth(2, 160);
      sheet.setColumnWidth(3, 200);
      sheet.setColumnWidth(4, 180);
      sheet.setColumnWidth(5, 320);
      sheet.setColumnWidth(6, 110);
      sheet.setColumnWidth(7, 200);
    }

    // Parsear datos: acepta form-encoded (e.parameter), JSON (postData) o URL-encoded body
    let data = {};
    if (e.parameter && e.parameter.nombre) {
      data = e.parameter;
    } else if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (_) {
        new URLSearchParams(e.postData.contents).forEach(function(v, k) { data[k] = v; });
      }
    }
    const fecha = new Date();

    // Agregar fila con el pedido
    sheet.appendRow([
      fecha,
      String(data.nombre || '').trim(),
      String(data.email || '').trim(),
      String(data.asunto || '').trim(),
      String(data.mensaje || '').trim(),
      'Pendiente',
      ''
    ]);

    // Aplicar validación de datos a la columna Estado (dropdown)
    const lastRow = sheet.getLastRow();
    const estadoCell = sheet.getRange(lastRow, 6);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente', 'En proceso', 'Atendido', 'Cancelado'], true)
      .build();
    estadoCell.setDataValidation(rule);

    // Notificación por email (si está configurada)
    if (NOTIFICATION_EMAIL && NOTIFICATION_EMAIL.indexOf('@') > 0) {
      MailApp.sendEmail({
        to: NOTIFICATION_EMAIL,
        subject: 'Nuevo pedido Mostyn — ' + (data.nombre || 'Sin nombre'),
        htmlBody:
          '<div style="font-family:system-ui,sans-serif;max-width:560px;margin:0 auto;padding:24px;background:#FBF8F3;border-radius:12px;">' +
          '<div style="border-left:4px solid #F5A623;padding-left:16px;margin-bottom:24px;">' +
          '<h2 style="margin:0;color:#111;">Nuevo pedido recibido</h2>' +
          '<p style="color:#666;margin:4px 0 0;">' + fecha.toLocaleString('es-AR') + '</p>' +
          '</div>' +
          '<table style="width:100%;border-collapse:collapse;">' +
          row('Nombre', data.nombre) +
          row('Email', data.email) +
          row('Consulta', data.asunto) +
          row('Mensaje', data.mensaje) +
          '</table>' +
          '<p style="margin-top:24px;font-size:13px;color:#999;">Pedido registrado en la planilla de Mostyn Apicultura.</p>' +
          '</div>'
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput('Mostyn Apicultura — Endpoint activo')
    .setMimeType(ContentService.MimeType.TEXT);
}

function row(label, value) {
  return '<tr>' +
    '<td style="padding:8px 12px 8px 0;color:#999;font-size:13px;text-transform:uppercase;letter-spacing:.06em;vertical-align:top;width:110px;">' + label + '</td>' +
    '<td style="padding:8px 0;color:#111;font-size:15px;">' + (value || '—') + '</td>' +
    '</tr>';
}
