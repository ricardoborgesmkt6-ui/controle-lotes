/**
 * Controle de Lotes — banco de dados na planilha.
 * Cole este código em Extensões > Apps Script da planilha, salve,
 * recarregue a planilha e use o menu "Controle de Lotes" > "Preparar planilha".
 */

const ABA_LANC = 'Lançamentos';
const ABA_LOTES = 'Lotes';
const ABA_CONF = 'Config';
const ABA_RES = 'Resumo';
const H_LANC = ['ID', 'ID do lote', 'Lote', 'Tipo', 'Data', 'Qtd operações', 'Valor por operação', 'Valor', 'Situação', 'Observação', 'Criado em'];
const H_LOTES = ['ID', 'Nome', 'Data de criação', 'Observação', 'Criado em'];
const TIPOS = { investimento: 'Investimento', recolhimento: 'Recolhimento', aporte: 'Aporte', retirada: 'Retirada' };
const SITUACAO = { realizado: 'Realizado', previsto: 'Previsto' };
const FMT_DATA = 'dd/mm/yyyy';
const FMT_MOEDA = '"R$" #,##0.00';
const FMT_HORA = 'dd/mm/yyyy hh:mm';

/* ===================== MENU E PREPARAÇÃO ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Controle de Lotes')
    .addItem('Preparar planilha', 'prepararPlanilha')
    .addItem('Mostrar chave de acesso', 'mostrarChave')
    .addItem('Gerar nova chave', 'gerarNovaChave')
    .addToUi();
}

function prepararPlanilha() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  PropertiesService.getScriptProperties().setProperty('PLANILHA_ID', ss.getId());

  // Lançamentos
  const lanc = aba_(ss, ABA_LANC, H_LANC);
  lanc.getRange('E2:E').setNumberFormat(FMT_DATA);
  lanc.getRange('G2:H').setNumberFormat(FMT_MOEDA);
  lanc.getRange('K2:K').setNumberFormat(FMT_HORA);
  lanc.hideColumns(1, 2);
  [0, 0, 140, 120, 95, 110, 140, 120, 100, 240, 130].forEach((w, i) => { if (w) lanc.setColumnWidth(i + 1, w); });

  // Lotes (colunas F a J calculadas por fórmula)
  const lotes = aba_(ss, ABA_LOTES, H_LOTES);
  const L = "'" + ABA_LANC + "'!";
  lotes.getRange('F1').setFormula('={"Investido";ARRAYFORMULA(IF(A2:A="",,SUMIFS(' + L + 'H:H,' + L + 'B:B,A2:A,' + L + 'D:D,"Investimento")))}');
  lotes.getRange('G1').setFormula('={"Recolhido";ARRAYFORMULA(IF(A2:A="",,SUMIFS(' + L + 'H:H,' + L + 'B:B,A2:A,' + L + 'D:D,"Recolhimento",' + L + 'I:I,"Realizado")))}');
  lotes.getRange('H1').setFormula('={"A recolher";ARRAYFORMULA(IF(A2:A="",,SUMIFS(' + L + 'H:H,' + L + 'B:B,A2:A,' + L + 'D:D,"Recolhimento",' + L + 'I:I,"Previsto")))}');
  lotes.getRange('I1').setFormula('={"Resultado";ARRAYFORMULA(IF(A2:A="",,G2:G-F2:F))}');
  lotes.getRange('J1').setFormula('={"Situação";ARRAYFORMULA(IF(A2:A="",,IF(H2:H>0,"A recolher",IF(G2:G=0,"Aguardando",IF(I2:I>=0,"Lucro","Prejuízo")))))}');
  lotes.getRange('A1:J1').setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  lotes.getRange('C2:C').setNumberFormat(FMT_DATA);
  lotes.getRange('E2:E').setNumberFormat(FMT_HORA);
  lotes.getRange('F2:I').setNumberFormat(FMT_MOEDA);
  lotes.hideColumns(1);
  [0, 160, 110, 220, 130, 120, 120, 120, 120, 110].forEach((w, i) => { if (w) lotes.setColumnWidth(i + 1, w); });

  // Config
  const conf = aba_(ss, ABA_CONF, ['Chave', 'Valor']);
  if (!linhaPorId_(conf, 'capital')) conf.appendRow(['capital', 0]);
  if (!linhaPorId_(conf, 'inicio')) conf.appendRow(['inicio', '']);

  // Resumo (somente fórmulas)
  let res = ss.getSheetByName(ABA_RES);
  if (!res) res = ss.insertSheet(ABA_RES, 0);
  res.clear();
  const rows = [
    ['Resumo — Controle de Lotes', ''],
    ['Capital inicial', '=IFERROR(VLOOKUP("capital",Config!A:B,2,FALSE),0)'],
    ['Total investido', '=SUMIFS(' + L + 'H:H,' + L + 'D:D,"Investimento")'],
    ['Total recolhido', '=SUMIFS(' + L + 'H:H,' + L + 'D:D,"Recolhimento",' + L + 'I:I,"Realizado")'],
    ['Resultado (recolhido − investido)', '=B4-B3'],
    ['Retorno sobre o investido', '=IFERROR(B5/B3,0)'],
    ['Aportes', '=SUMIFS(' + L + 'H:H,' + L + 'D:D,"Aporte",' + L + 'I:I,"Realizado")'],
    ['Retiradas', '=SUMIFS(' + L + 'H:H,' + L + 'D:D,"Retirada",' + L + 'I:I,"Realizado")'],
    ['Saldo em caixa', '=B2+B7-B8-B3+B4'],
    ['A recolher (previsto)', '=SUMIFS(' + L + 'H:H,' + L + 'D:D,"Recolhimento",' + L + 'I:I,"Previsto")'],
    ['Saldo projetado', '=B9+B10'],
    ['Lotes criados', '=COUNTA(Lotes!A2:A)'],
  ];
  res.getRange(1, 1, rows.length, 2).setValues(rows);
  res.getRange('A1').setFontWeight('bold').setFontSize(13);
  res.getRange('A2:A12').setFontWeight('bold');
  res.getRange('B2:B11').setNumberFormat(FMT_MOEDA);
  res.getRange('B6').setNumberFormat('0.0%');
  res.getRange('B12').setNumberFormat('0');
  res.getRange('A9:B9').setBackground('#EEF2F8');
  res.setColumnWidth(1, 260); res.setColumnWidth(2, 160);
  res.getRange('A14').setValue('Esta aba é só leitura. Os lançamentos são feitos pelo painel; a aba Lançamentos guarda tudo.').setFontColor('#666666').setFontStyle('italic');

  // remove a aba padrão vazia
  ['Página1', 'Sheet1', 'Planilha1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  const chave = chave_();
  SpreadsheetApp.getUi().alert('Planilha pronta!',
    'Sua chave de acesso é:\n\n' + chave + '\n\nGuarde essa chave: você vai colar ela no painel junto com o link do app da Web.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function mostrarChave() {
  SpreadsheetApp.getUi().alert('Chave de acesso', chave_(), SpreadsheetApp.getUi().ButtonSet.OK);
}

function gerarNovaChave() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Gerar nova chave?', 'Os aparelhos conectados com a chave antiga vão precisar da nova.', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().deleteProperty('CHAVE');
  ui.alert('Nova chave', chave_(), ui.ButtonSet.OK);
}

function aba_(ss, nome, cab) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  return sh;
}

function chave_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('CHAVE');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 20); p.setProperty('CHAVE', k); }
  return k;
}

/* ===================== API (APP DA WEB) ===================== */

function doGet() {
  return ContentService.createTextOutput('Controle de Lotes: conexão ativa. Use este link no painel.');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return saida_({ ok: false, erro: 'Requisição inválida.' }); }

  const chave = PropertiesService.getScriptProperties().getProperty('CHAVE');
  if (!chave) return saida_({ ok: false, erro: 'A planilha ainda não foi preparada. Use o menu Controle de Lotes > Preparar planilha.' });
  if (req.chave !== chave) return saida_({ ok: false, codigo: 'chave', erro: 'Chave de acesso incorreta.' });

  const ss = planilha_();
  const lock = LockService.getScriptLock();
  const escrita = req.acao !== 'listar';
  try {
    if (escrita) lock.waitLock(20000);
    const p = req.dados || {};
    switch (req.acao) {
      case 'listar': break;
      case 'criarLote': criarLote_(ss, p); break;
      case 'lancar': lancar_(ss, p.lanc); break;
      case 'receber': receber_(ss, p.id, p.data); break;
      case 'excluirLancamento': excluirLinhas_(ss.getSheetByName(ABA_LANC), r => r[0] === p.id); break;
      case 'excluirLote': excluirLote_(ss, p.id); break;
      case 'config': config_(ss, p); break;
      default: throw new Error('Ação desconhecida.');
    }
    if (escrita) SpreadsheetApp.flush();
    return saida_({ ok: true, dados: dados_(ss) });
  } catch (err) {
    return saida_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    if (escrita) try { lock.releaseLock(); } catch (_) {}
  }
}

function planilha_() {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILHA_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function saida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- leitura ---------- */

function dados_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const conf = {};
  linhas_(ss.getSheetByName(ABA_CONF), 2).forEach(r => { conf[String(r[0])] = r[1]; });
  const lotes = linhas_(ss.getSheetByName(ABA_LOTES), 5).map(r => ({
    id: String(r[0]), nome: String(r[1]), data: iso_(r[2], tz), obs: String(r[3] || ''), createdAt: ms_(r[4])
  }));
  const lanc = linhas_(ss.getSheetByName(ABA_LANC), 11).map(r => ({
    id: String(r[0]), loteId: r[1] ? String(r[1]) : null, tipo: String(r[3]).toLowerCase(), data: iso_(r[4], tz),
    qtd: r[5] === '' ? null : Number(r[5]), valorOp: r[6] === '' ? null : Number(r[6]), valor: Number(r[7]) || 0,
    status: String(r[8]).toLowerCase() === 'previsto' ? 'previsto' : 'realizado', obs: String(r[9] || ''), createdAt: ms_(r[10])
  }));
  const temCapital = conf.capital !== undefined && conf.capital !== '' && (Number(conf.capital) > 0 || conf.inicio);
  return {
    config: temCapital ? { capital: Number(conf.capital) || 0, inicio: iso_(conf.inicio, tz) || null } : null,
    lotes, lanc, planilha: ss.getUrl()
  };
}

function linhas_(sh, n) {
  const last = ultimaLinha_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, n).getValues().filter(r => r[0] !== '');
}

function ultimaLinha_(sh) {
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) if (col[i][0] !== '') return i + 1;
  return 0;
}

function iso_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
}

function ms_(v) { return v instanceof Date ? v.getTime() : (Number(v) || 0); }

function data_(iso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso || ''))) throw new Error('Data inválida.');
  const p = iso.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
}

/* ---------- escrita ---------- */

function novoId_(prefixo) {
  return prefixo + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function criarLote_(ss, p) {
  const lote = p.lote || {};
  const nome = String(lote.nome || '').trim();
  if (!nome) throw new Error('Dê um nome ao lote.');
  const sh = ss.getSheetByName(ABA_LOTES);
  const existe = linhas_(sh, 2).some(r => String(r[1]).trim().toLowerCase() === nome.toLowerCase());
  if (existe) throw new Error('Já existe um lote chamado “' + nome + '”. Escolha outro nome.');
  const id = novoId_('L');
  const linha = ultimaLinha_(sh) + 1;
  sh.getRange(linha, 1, 1, 5).setValues([[id, nome, data_(lote.data), String(lote.obs || ''), new Date()]]);
  (p.lancamentos || []).forEach(l => lancar_(ss, Object.assign({}, l, { loteId: id })));
}

function lancar_(ss, l) {
  if (!l || !TIPOS[l.tipo]) throw new Error('Tipo de lançamento inválido.');
  const valor = Number(l.valor);
  if (!(valor > 0)) throw new Error('Informe um valor maior que zero.');
  let nomeLote = '';
  if (l.tipo === 'investimento' || l.tipo === 'recolhimento') {
    const lote = linhas_(ss.getSheetByName(ABA_LOTES), 2).find(r => r[0] === l.loteId);
    if (!lote) throw new Error('Lote não encontrado. Atualize o painel e tente de novo.');
    nomeLote = lote[1];
  }
  const sh = ss.getSheetByName(ABA_LANC);
  const linha = ultimaLinha_(sh) + 1;
  sh.getRange(linha, 1, 1, 11).setValues([[
    novoId_('M'), l.loteId || '', nomeLote || 'Caixa', TIPOS[l.tipo], data_(l.data),
    l.qtd ? Number(l.qtd) : '', l.valorOp ? Number(l.valorOp) : '', valor,
    l.status === 'previsto' && l.tipo === 'recolhimento' ? SITUACAO.previsto : SITUACAO.realizado,
    String(l.obs || ''), new Date()
  ]]);
}

function receber_(ss, id, iso) {
  const sh = ss.getSheetByName(ABA_LANC);
  const linha = linhaPorId_(sh, id);
  if (!linha) throw new Error('Lançamento não encontrado.');
  sh.getRange(linha, 9).setValue(SITUACAO.realizado);
  if (iso) sh.getRange(linha, 5).setValue(data_(iso));
}

function excluirLote_(ss, id) {
  excluirLinhas_(ss.getSheetByName(ABA_LANC), r => r[1] === id);
  excluirLinhas_(ss.getSheetByName(ABA_LOTES), r => r[0] === id);
}

function excluirLinhas_(sh, teste) {
  const last = ultimaLinha_(sh);
  if (last < 2) return;
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (teste(vals[i])) sh.deleteRow(i + 2);
}

function config_(ss, p) {
  const sh = ss.getSheetByName(ABA_CONF);
  const cap = Number(p.capital);
  if (!(cap >= 0)) throw new Error('Capital inválido.');
  setConf_(sh, 'capital', cap);
  setConf_(sh, 'inicio', data_(p.inicio));
}

function setConf_(sh, chave, valor) {
  let linha = linhaPorId_(sh, chave);
  if (!linha) { sh.appendRow([chave, '']); linha = ultimaLinha_(sh); }
  const cel = sh.getRange(linha, 2);
  cel.setValue(valor);
  if (valor instanceof Date) cel.setNumberFormat(FMT_DATA);
}

function linhaPorId_(sh, id) {
  const last = ultimaLinha_(sh);
  if (last < 2) return 0;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]) === String(id)) return i + 2;
  return 0;
}
