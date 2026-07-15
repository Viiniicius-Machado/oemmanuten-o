// ══════════════════════════════════════════════════════════════
//  GOOGLE APPS SCRIPT — Netturbo O&M Sistema v3
//  Abas: PENDENTES → RFO (aprovação) + FISCALIZACAO + ERROS
//
//  COMO ATUALIZAR:
//  1. script.google.com → abra o projeto
//  2. Selecione tudo (Ctrl+A) → cole este código → Salvar
//  3. Implantar → Gerenciar implantações → ✏️ Editar
//     → Versão: "Nova versão" → Implantar
//  A URL não muda — não precisa alterar nos HTMLs.
// ══════════════════════════════════════════════════════════════

const ABA_RFO          = 'RFO';
const ABA_PENDENTES    = 'PENDENTES';
const ABA_FISCALIZACAO = 'FISCALIZACAO';
const ABA_ERROS        = 'ERROS';

// ── CABEÇALHOS ─────────────────────────────────────────────────
const HEADERS_RFO = [
  'Timestamp Aprovação','Aprovado Por','Empresa','Técnico','Data',
  'Protocolo O&M','Cidade','ID','Cliente',
  'Local do Defeito','GPS',
  'Hora Início','Hora Término','SLA Total',
  'Hora Chegada','Tempo em Campo',
  'Tipo Ocorrência','Causa Principal','Motivo Carga Alta',
  'Detalhes / Obs','Materiais Utilizados',
  'Solução Realizada','Tipo CEO','Melhoria em Campo',
  'Obs. do Líder'
];

const HEADERS_PENDENTES = [
  'ID Pendente','Timestamp Envio','Empresa','Técnico','Data',
  'Protocolo O&M','Cidade','ID','Cliente',
  'Local do Defeito','GPS',
  'Hora Início','Hora Término','SLA Total',
  'Hora Chegada','Tempo em Campo',
  'Tipo Ocorrência','Causa Principal','Motivo Carga Alta',
  'Detalhes / Obs','Materiais Utilizados',
  'Solução Realizada','Tipo CEO','Melhoria em Campo',
  'Status'
];

const HEADERS_FISCALIZACAO = [
  'Timestamp','Tipo de Relatório','Protocolo / OS',
  'Fiscal Responsável','Prestador Fiscalizado',
  'Data','Cidade','Endereço','Descrição Geral',
  'Conformes','Atenção','Não Conformes',
  'Checklist Detalhado',
  'Ocorrências Identificadas','Providências Determinadas',
  'Prazo Correção','Status Geral',
  'Nome Fiscal','Nome Fiscalizado','Total Fotos'
];

const HEADERS_ERROS = [
  'Timestamp','Origem','Motivo','Dados'
];

// ── GARANTIR ABA ───────────────────────────────────────────────
function garantirAba(ss, nome, headers, corFundo, corTexto) {
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    sheet.appendRow(headers);
    const hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground(corFundo)
      .setFontColor(corTexto)
      .setFontWeight('bold')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 32);
    sheet.setFrozenRows(1);
    try { sheet.autoResizeColumns(1, headers.length); } catch(e) {}
  }
  return sheet;
}

// ── REGISTRAR ERRO ─────────────────────────────────────────────
function registrarErro(ss, motivo, dadosRaw) {
  try {
    const sheet = garantirAba(ss, ABA_ERROS, HEADERS_ERROS, '#3a0000', '#ff8a80');
    sheet.appendRow([new Date().toLocaleString('pt-BR'), 'App Web', motivo, dadosRaw.substring(0, 500)]);
  } catch(e) { Logger.log('Erro ao registrar: ' + e); }
}

// ── VALIDAÇÃO RFO ──────────────────────────────────────────────
function validarRFO(data) {
  const erros = [];
  if (!data.tecnico   || data.tecnico.trim()   === '') erros.push('Técnico vazio');
  if (!data.protocolo || data.protocolo.trim() === '' || data.protocolo === '0') erros.push('Protocolo vazio');
  if (!data.ocorrencia|| data.ocorrencia.trim()=== '') erros.push('Ocorrência vazia');
  if (!data.cidade    || data.cidade.trim()    === '') erros.push('Cidade vazia');
  return erros;
}

// ── HANDLER PRINCIPAL ──────────────────────────────────────────
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dadosRaw = '';
  try {
    dadosRaw = e.postData.contents;
    const data = JSON.parse(dadosRaw);

    if (data.aba === 'FISCALIZACAO') {
      gravarFiscalizacao(ss, data);
      return resposta('ok', 'Fiscalização salva');

    } else if (data.aba === 'APROVAR') {
      // Líder aprovou — move de PENDENTES para RFO
      return aprovarRFO(ss, data);

    } else if (data.aba === 'REPROVAR') {
      // Líder reprovou — atualiza status em PENDENTES
      return reprovarRFO(ss, data);

    } else if (data.aba === 'LISTAR_PENDENTES') {
      // Retorna lista de pendentes para a tela do líder
      return listarPendentes(ss);

    } else {
      // Técnico enviou RFO — vai para PENDENTES aguardando aprovação
      const erros = validarRFO(data);
      if (erros.length > 0) {
        registrarErro(ss, 'RFO inválido: ' + erros.join(', '), dadosRaw);
        return resposta('rejected', erros.join(', '));
      }
      return gravarPendente(ss, data);
    }

  } catch(err) {
    Logger.log('Erro crítico: ' + err.toString());
    registrarErro(ss, 'Erro: ' + err.toString(), dadosRaw);
    return resposta('error', err.toString());
  }
}

// ── GRAVAR EM PENDENTES ────────────────────────────────────────
function gravarPendente(ss, data) {
  const sheet = garantirAba(ss, ABA_PENDENTES, HEADERS_PENDENTES, '#1a1200', '#ffa000');

  // Gera ID único para este pendente
  const idPendente = 'RFO-' + new Date().getTime();

  sheet.appendRow([
    idPendente,
    data.timestamp      || new Date().toLocaleString('pt-BR'),
    data.empresa        || '',
    data.tecnico        || '',
    data.data           || '',
    data.protocolo      || '',
    data.cidade         || '',
    data.id             || '',
    data.cliente        || '',
    data.local_defeito  || '',
    data.gps_coords     || '',
    data.hr_inicio      || '',
    data.hr_termino     || '',
    data.sla_total      || '',
    data.hr_chegada     || '',
    data.tempo_campo    || '',
    data.ocorrencia     || '',
    data.causa          || '',
    data.carga_motivo   || '',
    data.obs_extras     || '',
    data.materiais      || '',
    data.solucao        || '',
    data.tipo_ceo       || '',
    data.melhoria       || '',
    'PENDENTE'          // Status
  ]);

  // Colorir status PENDENTE (col 25)
  const row = sheet.getLastRow();
  sheet.getRange(row, 25).setBackground('#ff6d00').setFontColor('#fff').setFontWeight('bold');

  Logger.log('RFO salvo em PENDENTES: ' + idPendente);
  return resposta('ok', idPendente);
}

// ── LISTAR PENDENTES ───────────────────────────────────────────
function listarPendentes(ss) {
  const sheet = ss.getSheetByName(ABA_PENDENTES);
  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', pendentes: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const pendentes = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[24]; // coluna Status
    if (status === 'PENDENTE') {
      pendentes.push({
        rowIndex:   i + 1, // 1-based para o sheet
        id:         row[0],
        timestamp:  row[1],
        empresa:    row[2],
        tecnico:    row[3],
        data:       row[4],
        protocolo:  row[5],
        cidade:     row[6],
        id_etiq:    row[7],
        cliente:    row[8],
        local:      row[9],
        gps:        row[10],
        hr_inicio:  row[11],
        hr_termino: row[12],
        sla_total:  row[13],
        hr_chegada: row[14],
        tempo_campo:row[15],
        ocorrencia: row[16],
        causa:      row[17],
        carga:      row[18],
        obs:        row[19],
        materiais:  row[20],
        solucao:    row[21],
        tipo_ceo:   row[22],
        melhoria:   row[23],
        status:     row[24]
      });
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', pendentes: pendentes }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── APROVAR RFO ────────────────────────────────────────────────
function aprovarRFO(ss, data) {
  const sheetPend = ss.getSheetByName(ABA_PENDENTES);
  if (!sheetPend) return resposta('error', 'Aba PENDENTES não encontrada');

  // Atualiza status para APROVADO em PENDENTES
  const rowIndex = parseInt(data.rowIndex);
  sheetPend.getRange(rowIndex, 25).setValue('APROVADO').setBackground('#2e7d32').setFontColor('#fff').setFontWeight('bold');

  // Grava na aba RFO definitiva
  const sheetRFO = garantirAba(ss, ABA_RFO, HEADERS_RFO, '#1a1a1a', '#8bc34a');
  sheetRFO.appendRow([
    new Date().toLocaleString('pt-BR'), // Timestamp Aprovação
    data.aprovado_por   || '',
    data.empresa        || '',
    data.tecnico        || '',
    data.data           || '',
    data.protocolo      || '',
    data.cidade         || '',
    data.id_etiq        || '',
    data.cliente        || '',
    data.local          || '',
    data.gps            || '',
    data.hr_inicio      || '',
    data.hr_termino     || '',
    data.sla_total      || '',
    data.hr_chegada     || '',
    data.tempo_campo    || '',
    data.ocorrencia     || '',
    data.causa          || '',
    data.carga          || '',
    data.obs            || '',
    data.materiais      || '',
    data.solucao        || '',
    data.tipo_ceo       || '',
    data.melhoria       || '',
    data.obs_lider      || ''  // Observação do líder
  ]);

  // Colorir Tipo Ocorrência (col 18 no RFO = índice 17+1)
  const row = sheetRFO.getLastRow();
  const tipoNorm = (data.ocorrencia||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const cell = sheetRFO.getRange(row, 18);
  cell.setFontWeight('bold');
  if      (tipoNorm.indexOf('ROMPIMENTO')     !== -1) cell.setBackground('#c62828').setFontColor('#fff');
  else if (tipoNorm.indexOf('MASSIVA')        !== -1) cell.setBackground('#e65100').setFontColor('#fff');
  else if (tipoNorm.indexOf('ACOMPANHAMENTO') !== -1) cell.setBackground('#2e7d32').setFontColor('#fff');
  else if (tipoNorm.indexOf('PREVENTIVA')     !== -1) cell.setBackground('#1565c0').setFontColor('#fff');

  Logger.log('RFO aprovado por ' + data.aprovado_por + ' — linha PENDENTES: ' + rowIndex);
  return resposta('ok', 'Aprovado e gravado no RFO');
}

// ── REPROVAR RFO ───────────────────────────────────────────────
function reprovarRFO(ss, data) {
  const sheetPend = ss.getSheetByName(ABA_PENDENTES);
  if (!sheetPend) return resposta('error', 'Aba PENDENTES não encontrada');

  const rowIndex = parseInt(data.rowIndex);
  sheetPend.getRange(rowIndex, 25).setValue('REPROVADO').setBackground('#c62828').setFontColor('#fff').setFontWeight('bold');

  // Salva obs do líder na col 26
  if (sheetPend.getLastColumn() < 26) {
    sheetPend.getRange(1, 26).setValue('Obs. Reprovação');
  }
  sheetPend.getRange(rowIndex, 26).setValue((data.obs_lider||'') + ' — Reprovado por: ' + (data.aprovado_por||''));

  Logger.log('RFO reprovado por ' + data.aprovado_por + ' — linha: ' + rowIndex);
  return resposta('ok', 'Reprovado');
}

// ── FISCALIZAÇÃO ───────────────────────────────────────────────
function gravarFiscalizacao(ss, data) {
  const sheet = garantirAba(ss, ABA_FISCALIZACAO, HEADERS_FISCALIZACAO, '#1a1500', '#d4a843');
  sheet.appendRow([
    data.timestamp || new Date().toLocaleString('pt-BR'),
    data.tipo_relatorio || '', data.protocolo || '',
    data.fiscal || '', data.prestador || '',
    data.data || '', data.cidade || '', data.endereco || '', data.descricao || '',
    Number(data.conformes||0), Number(data.atencao||0), Number(data.nao_conformes||0),
    data.checklist || '', data.ocorrencias || '', data.providencias || '',
    data.prazo || '', data.status_geral || '',
    data.fiscal_responsavel || '', data.fiscalizado || '', Number(data.total_fotos||0)
  ]);

  const row = sheet.getLastRow();
  const status = (data.status_geral||'').toUpperCase();
  const sc = sheet.getRange(row, 17);
  if      (status.includes('REPROVADO') && !status.includes('RESSALVAS')) sc.setBackground('#ffebee').setFontColor('#b71c1c').setFontWeight('bold');
  else if (status.includes('RESSALVAS')) sc.setBackground('#fff8e1').setFontColor('#e65100').setFontWeight('bold');
  else if (status.includes('APROVADO'))  sc.setBackground('#e8f5e9').setFontColor('#2e7d32').setFontWeight('bold');
  else if (status.includes('ANDAMENTO')) sc.setBackground('#e3f2fd').setFontColor('#1565c0').setFontWeight('bold');
}

// ── RESPOSTA ───────────────────────────────────────────────────
function resposta(status, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: status, message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TEST GET ───────────────────────────────────────────────────
function doGet(e) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ss.getSheets().map(s => ({ nome: s.getName(), linhas: s.getLastRow() }));
  return resposta('ok', JSON.stringify({ sistema: 'Netturbo O&M v3', abas: abas }));
}
