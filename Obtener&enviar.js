const ss = SpreadsheetApp.getActiveSpreadsheet();
const ws = ss.getSheetByName('Setup - Tabla de resultados');

var mes = ws.getRange('a3').getValue();

var dia = ws.getRange('a4').getValue();

var titulo00 = ws.getRange('a6').getDisplayValue();
var titulo01 = ws.getRange('b6').getDisplayValue();
var titulo02 = ws.getRange('c6').getDisplayValue();
var titulo03 = ws.getRange('d6').getDisplayValue();

var ventas10 = ws.getRange('a7').getDisplayValue();
var ventas11 = ws.getRange('b7').getDisplayValue();
var ventas12 = ws.getRange('c7').getDisplayValue();
var ventas13 = ws.getRange('d7').getDisplayValue();

var inversion20 = ws.getRange('a8').getDisplayValue();
var inversion21 = ws.getRange('b8').getDisplayValue();
var inversion22 = ws.getRange('c8').getDisplayValue();
var inversion23 = ws.getRange('d8').getDisplayValue();

var cpa30 = ws.getRange('a9').getDisplayValue();
var cpa31 = ws.getRange('b9').getDisplayValue();
var cpa32 = ws.getRange('c9').getDisplayValue();
var cpa33 = ws.getRange('d9').getDisplayValue();
var analisis = ws.getRange('a13').getValue();

function enviar_correo() {
var html = HtmlService.createTemplateFromFile('Estructura del correo');
html.mes = mes;
html.dia = dia;
html.titulo00 = titulo00;
html.titulo01 = titulo01;
html.titulo02 = titulo02;
html.titulo03 = titulo03;
html.ventas10 = ventas10;
html.ventas11 = ventas11;
html.ventas12 = ventas12;
html.ventas13 = ventas13;
html.inversion20 = inversion20;
html.inversion21 = inversion21;
html.inversion22 = inversion22;
html.inversion23 = inversion23;
html.cpa30 = cpa30;
html.cpa31 = cpa31;
html.cpa32 = cpa32;
html.cpa33 = cpa33;
html.analisis = analisis;

const sheet = ss.getSheetByName('Setup - Correo');

var filas = sheet.getRange('b1').getValue();
const fila_de_inicio = 4; 

var cuerpo = sheet.getRange('d4').getValue();
html.cuerpo = cuerpo;

var cierre = sheet.getRange('e4').getValue();
html.cierre = cierre;

var firma = sheet.getRange('f4').getValue();
html.firma = firma;

var dataRange = sheet.getRange(fila_de_inicio, 1, filas, 4);
var data = dataRange.getValues();

   for (var x in data) {
    var columna = data[x];
    var correo = columna[0];
    var saludo = columna[1];
    html.saludo = saludo;
    var htmlEmail = html.evaluate().getContent();
    var asunto = columna[2];
    GmailApp.sendEmail(correo, asunto, saludo,{ htmlBody : htmlEmail}
    );
  }
} 
