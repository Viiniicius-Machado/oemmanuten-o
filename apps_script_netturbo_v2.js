// ══════════════════════════════════════════════════════════════
//  GOOGLE APPS SCRIPT — Netturbo O&M Sistema v2
//  Abas: RFO (técnicos) + FISCALIZACAO (liderança)
//
//  COMO ATUALIZAR:
//  1. script.google.com → abra o projeto
//  2. Selecione tudo (Ctrl+A) → cole este código → Salvar
//  3. Implantar → Gerenciar implantações → ✏️ Editar
//     → Versão: "Nova versão" → Implantar
//  A URL não muda — não precisa alterar nos HTMLs.
// ══════════════════════════════════════════════════════════════

const ABA_RFO          = 'RFO';
const ABA_FISCALIZACAO = 'FISCALIZACAO';
const ABA_ERROS        = 'ERROS';      // Registra tentativas inválidas

// ── CABEÇALHOS ─────────────────────────────────────────────────
const HEADERS_RFO = [
  'Timestamp','Empresa','Técnico','Data',
  'Protocolo O&M','Cidade','ID','Cliente',
  'Local do Defeito','GPS',
  'Hora Início','Hora Término','SLA Total',
  'Hora Chegada','Tempo em Campo',
  'Tipo Ocorrência','Causa Principal','Motivo Carga Alta',
  'Detalhes / Obs','Materiais Utilizados',
  'Solução Realizada','Melhoria em Campo'
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
  'Timestamp','IP / Origem','Motivo da Rejeição','Dados Recebidos'
];

// ── VALIDAÇÃO RFO ──────────────────────────────────────────────
function validarRFO(data) {
  const erros = [];
  if (!data.tecnico   || data.tecnico.trim()   === '') erros.push('Técnico vazio');
  if (!data.protocolo || data.protocolo.trim() === '' || data.protocolo === '0') erros.push('Protocolo vazio ou zero');
  if (!data.ocorrencia|| data.ocorrencia.trim()=== '') erros.push('Tipo de ocorrência vazio');
  if (!data.cidade    || data.cidade.trim()    === '') erros.push('Cidade vazia');
  return erros;
}

// ── VALIDAÇÃO FISCALIZAÇÃO ─────────────────────────────────────
function validarFiscalizacao(data) {
  const erros = [];
  if (!data.fiscal     || data.fiscal.trim()     === '') erros.push('Fiscal vazio');
  if (!data.protocolo  || data.protocolo.trim()  === '') erros.push('Protocolo vazio');
  if (!data.tipo_relatorio || data.tipo_relatorio.trim() === '') erros.push('Tipo de relatório vazio');
  return erros;
}

// ── GARANTIR ABA ───────────────────────────────────────────────
function garantirAba(ss, nome, headers, corFundo, corTexto) {
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    const hr = sheet.getRange(1, 1, 1, headers.length);
    sheet.appendRow(headers);
    // Mover cabeçalho para linha 1 (appendRow vai para linha 1 se vazia)
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
    sheet.appendRow([
      new Date().toLocaleString('pt-BR'),
      'App Web',
      motivo,
      dadosRaw.substring(0, 500) // limita a 500 chars
    ]);
  } catch(e) {
    Logger.log('Erro ao registrar erro: ' + e.toString());
  }
}

// ── HANDLER PRINCIPAL ──────────────────────────────────────────
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dadosRaw = '';

  try {
    dadosRaw = e.postData.contents;
    const data = JSON.parse(dadosRaw);

    // Detecta qual módulo enviou
    if (data.aba === 'FISCALIZACAO') {
      // ── FISCALIZAÇÃO ──
      const erros = validarFiscalizacao(data);
      if (erros.length > 0) {
        registrarErro(ss, 'FISCALIZACAO inválida: ' + erros.join(', '), dadosRaw);
        return resposta('rejected', erros.join(', '));
      }
      gravarFiscalizacao(ss, data);
      return resposta('ok', 'Fiscalização salva');

    } else {
      // ── RFO TÉCNICO ──
      const erros = validarRFO(data);
      if (erros.length > 0) {
        registrarErro(ss, 'RFO inválido: ' + erros.join(', '), dadosRaw);
        return resposta('rejected', erros.join(', '));
      }
      gravarRFO(ss, data);
      return resposta('ok', 'RFO salvo');
    }

  } catch(err) {
    Logger.log('Erro crítico: ' + err.toString());
    registrarErro(ss, 'Erro de parse: ' + err.toString(), dadosRaw);
    return resposta('error', err.toString());
  }
}

// ── GRAVAR RFO ─────────────────────────────────────────────────
function gravarRFO(ss, data) {
  const sheet = garantirAba(ss, ABA_RFO, HEADERS_RFO, '#1a1a1a', '#8bc34a');

  sheet.appendRow([
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
    data.melhoria       || ''
  ]);

  // Colorir coluna de Tipo Ocorrência (col 16)
  const row = sheet.getLastRow();
  const tipo = (data.ocorrencia || '').toUpperCase();
  const corMap = {
    'ROMPIMENTO':    ['#ff1744','#fff'],
    'MASSIVA':       ['#ff6d00','#fff'],
    'ATENUAÇÃO':     ['#ffd600','#111'],
    'ACOMPANHAMENTO':['#00c853','#111'],
    'RADIO':         ['#2979ff','#fff']
  };
  for (const [key, [bg, fg]] of Object.entries(corMap)) {
    if (tipo.includes(key)) {
      sheet.getRange(row, 16).setBackground(bg).setFontColor(fg).setFontWeight('bold');
      break;
    }
  }

  // Zebra stripes suaves
  if (row % 2 === 0) {
    sheet.getRange(row, 1, 1, HEADERS_RFO.length).setBackground('#f9fbe7');
  }
}

// ── GRAVAR FISCALIZAÇÃO ────────────────────────────────────────
function gravarFiscalizacao(ss, data) {
  const sheet = garantirAba(
    ss, ABA_FISCALIZACAO, HEADERS_FISCALIZACAO, '#1a1500', '#d4a843'
  );

  sheet.appendRow([
    data.timestamp          || new Date().toLocaleString('pt-BR'),
    data.tipo_relatorio     || '',
    data.protocolo          || '',
    data.fiscal             || '',
    data.prestador          || '',
    data.data               || '',
    data.cidade             || '',
    data.endereco           || '',
    data.descricao          || '',
    Number(data.conformes   || 0),
    Number(data.atencao     || 0),
    Number(data.nao_conformes || 0),
    data.checklist          || '',
    data.ocorrencias        || '',
    data.providencias       || '',
    data.prazo              || '',
    data.status_geral       || '',
    data.fiscal_responsavel || '',
    data.fiscalizado        || '',
    Number(data.total_fotos || 0)
  ]);

  const row = sheet.getLastRow();

  // Colorir Status Geral (col 17)
  const status = (data.status_geral || '').toUpperCase();
  const statusCell = sheet.getRange(row, 17);
  if      (status.includes('REPROVADO') && !status.includes('RESSALVAS'))
    statusCell.setBackground('#ffebee').setFontColor('#b71c1c').setFontWeight('bold');
  else if (status.includes('RESSALVAS'))
    statusCell.setBackground('#fff8e1').setFontColor('#e65100').setFontWeight('bold');
  else if (status.includes('APROVADO'))
    statusCell.setBackground('#e8f5e9').setFontColor('#2e7d32').setFontWeight('bold');
  else if (status.includes('ANDAMENTO'))
    statusCell.setBackground('#e3f2fd').setFontColor('#1565c0').setFontWeight('bold');

  // Tipo de relatório colorido (col 2)
  const tipoCell = sheet.getRange(row, 2);
  const tipoRel = (data.tipo_relatorio || '').toUpperCase();
  if (tipoRel.includes('OBRA'))
    tipoCell.setBackground('#e8f5e9').setFontColor('#2e7d32').setFontWeight('bold');
  else if (tipoRel.includes('AUDITORIA'))
    tipoCell.setBackground('#fff3e0').setFontColor('#e65100').setFontWeight('bold');

  // Checklist — wrap text (col 13)
  sheet.getRange(row, 13).setWrap(true);

  // Zebra
  if (row % 2 === 0) {
    sheet.getRange(row, 1, 1, HEADERS_FISCALIZACAO.length).setBackground('#fffde7');
  }
}

// ── RESPOSTA PADRÃO ────────────────────────────────────────────
function resposta(status, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: status, message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TEST VIA GET ───────────────────────────────────────────────
function doGet(e) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ss.getSheets().map(s => ({
    nome:   s.getName(),
    linhas: s.getLastRow()
  }));
  return resposta('ok', JSON.stringify({
    sistema: 'Netturbo O&M v2',
    planilha: ss.getName(),
    abas: abas
  }));
}
