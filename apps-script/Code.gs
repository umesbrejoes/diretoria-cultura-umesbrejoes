const CONFIG = { SHEET_ID: 'COLE_AQUI_O_ID_DA_PLANILHA', DEFAULT_SHEET: 'Respostas' };

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(data.sheet || CONFIG.DEFAULT_SHEET) || ss.insertSheet(data.sheet || CONFIG.DEFAULT_SHEET);
    const headers = ['timestamp','tipo','atividade','nome','escola','serie_curso','contato','mensagem','status'];
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
    sheet.appendRow([new Date(), data.tipo || 'participacao', data.atividade || '', data.nome || '', data.escola || '', data.serie_curso || '', data.contato || '', data.mensagem || '', 'NOVO']);
    return json({ok:true, message:'Recebido'});
  } catch (err) { return json({ok:false,error:String(err)}); }
}

function doGet() { return json({ok:true, service:'UMES Brejões — Central de Participação'}); }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// Execute uma vez no editor para criar as abas principais.
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  ['Respostas','Inscrições','Propostas','Trabalhos','Atividades','Certificados'].forEach(name => { if (!ss.getSheetByName(name)) ss.insertSheet(name); });
}
