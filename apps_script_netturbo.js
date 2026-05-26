// ══════════════════════════════════════════════════════════════
//  GOOGLE APPS SCRIPT — RFO Manutenção Netturbo
//  Cole este código em: script.google.com → Novo projeto
//  Depois: Implantar → Novo Implantação → App da Web
//  Acesso: "Qualquer pessoa" → Copiar a URL gerada
// ══════════════════════════════════════════════════════════════

const SHEET_NAME = 'RFO'; // Nome da aba na planilha

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // Criar aba se não existir e adicionar cabeçalho
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Timestamp', 'Empresa', 'Técnico', 'Data',
        'Protocolo O&M', 'Cidade', 'ID', 'Cliente',
        'Local do Defeito', 'GPS Coords',
        'Hora Início', 'Hora Término', 'SLA Total',
        'Hora Chegada', 'Tempo em Campo',
        'Tipo Ocorrência', 'Causa Principal', 'Motivo Carga Alta',
        'Obs. Causa', 'Materiais Utilizados', 'Solução Realizada', 'Melhoria em Campo'
      ]);
      // Formatar cabeçalho
      const header = sheet.getRange(1, 1, 1, 22);
      const hRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      hRange.setBackground('#1a1a1a');
      hRange.setFontColor('#8bc34a');
      hRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Adicionar linha de dados
    sheet.appendRow([
      data.timestamp     || '',
      data.empresa       || '',
      data.tecnico       || '',
      data.data          || '',
      data.protocolo     || '',
      data.cidade        || '',
      data.id            || '',
      data.cliente       || '',
      data.local_defeito || '',
      data.gps_coords    || '',
      data.hr_inicio     || '',
      data.hr_termino    || '',
      data.sla_total     || '',
      data.hr_chegada    || '',
      data.tempo_campo   || '',
      data.ocorrencia    || '',
      data.causa         || '',
      data.carga_motivo  || '',
      data.obs_extras    || '',
      data.materiais     || '',
      data.solucao       || '',
      data.melhoria      || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Teste via GET (opcional - para verificar se o script está funcionando)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'RFO Netturbo Script ativo!' }))
    .setMimeType(ContentService.MimeType.JSON);
}
