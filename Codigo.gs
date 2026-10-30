/**
 * Controle de Lotes — banco de dados na planilha, com login de admin e operadores.
 * Cole este código em Extensões > Apps Script da planilha e salve.
 * Primeira vez: recarregue a planilha e use o menu "Controle de Lotes" > "Preparar planilha".
 * Atualizando de uma versão anterior: cole, salve e faça Deploy > Manage deployments > editar > New version.
 */

const ABA_LANC = 'Lançamentos';
const ABA_LOTES = 'Lotes';
const ABA_USERS = 'Usuários';
const ABA_CONF = 'Config';
const ABA_RES = 'Resumo';
const H_LANC = ['ID', 'ID do lote', 'Lote', 'Tipo', 'Data', 'Qtd operações', 'Valor por operação', 'Valor', 'Situação', 'Observação', 'Criado em', 'ID operador', 'Operador'];
const H_LOTES = ['ID', 'Nome', 'Data de criação', 'Observação', 'Criado em'];          // F..J são fórmulas
const H_LOTES_OP = ['ID operador', 'Operador'];                                       // colunas K e L
const H_USERS = ['ID', 'Nome', 'Usuário', 'Senha (protegida)', 'Papel', 'Ativo', 'Capital inicial', 'Início', 'Criado em'];
const TIPOS = { investimento: 'Investimento', recolhimento: 'Recolhimento', aporte: 'Aporte', retirada: 'Retirada' };
const SITUACAO = { realizado: 'Realizado', previsto: 'Previsto' };
const FMT_DATA = 'dd/mm/yyyy';
const FMT_MOEDA = '"R$" #,##0.00';
const FMT_HORA = 'dd/mm/yyyy hh:mm';
const SESSAO_DIAS = 30;
const MAX_FALHAS = 5;
const BLOQUEIO_MIN = 10;

/* ===================== MENU E PREPARAÇÃO ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Controle de Lotes')
    .addItem('Preparar planilha', 'prepararPlanilha')
    .addItem('Mostrar chave de acesso', 'mostrarChave')
    .addItem('Redefinir senha de um usuário', 'redefinirSenha')
    .addItem('Gerar nova chave', 'gerarNovaChave')
    .addToUi();
}

function prepararPlanilha() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  PropertiesService.getScriptProperties().setProperty('PLANILHA_ID', ss.getId());
  garantir_(ss, true);
  ['Página1', 'Sheet1', 'Planilha1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  const chave = chave_();
  SpreadsheetApp.getUi().alert('Planilha pronta!',
    'Sua chave de acesso é:\n\n' + chave + '\n\nVocê vai usar essa chave uma vez, no painel, para criar o seu login de administrador.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function mostrarChave() {
  SpreadsheetApp.getUi().alert('Chave de acesso', chave_(), SpreadsheetApp.getUi().ButtonSet.OK);
}

function gerarNovaChave() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Gerar nova chave?', 'A chave só é usada para criar o primeiro administrador. Gerar uma nova também desconecta todos os aparelhos (todos precisarão entrar de novo).', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const p = PropertiesService.getScriptProperties();
  p.deleteProperty('CHAVE'); p.deleteProperty('SEGREDO');
  ui.alert('Nova chave', chave_(), ui.ButtonSet.OK);
}

function redefinirSenha() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  garantir_(ss, false);
  const r1 = ui.prompt('Redefinir senha', 'Digite o usuário (login) que vai receber a nova senha:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const sh = ss.getSheetByName(ABA_USERS);
  const u = usuarios_(ss).find(x => x.usuario === normUser_(r1.getResponseText()));
  if (!u) { ui.alert('Usuário não encontrado.'); return; }
  const r2 = ui.prompt('Nova senha para ' + u.nome, 'Mínimo de 6 caracteres:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const senha = r2.getResponseText();
  if (senha.length < 6) { ui.alert('A senha precisa ter pelo menos 6 caracteres.'); return; }
  sh.getRange(u.linha, 4).setValue(hashSenha_(senha));
  sh.getRange(u.linha, 6).setValue('Sim');
  PropertiesService.getScriptProperties().deleteProperty('FALHA_' + u.usuario);
  ui.alert('Senha redefinida. ' + u.nome + ' já pode entrar com a nova senha.');
}

/** Cria/atualiza abas, cabeçalhos e fórmulas. Seguro para rodar várias vezes. */
function garantir_(ss, formatar) {
  const L = "'" + ABA_LANC + "'!";
  // Lançamentos
  const lanc = aba_(ss, ABA_LANC, H_LANC);
  // Lotes
  const lotes = aba_(ss, ABA_LOTES, H_LOTES);
  const cabLotesOp = lotes.getRange(1, 11, 1, 2).getValues()[0];
  if (cabLotesOp[0] !== H_LOTES_OP[0]) lotes.getRange(1, 11, 1, 2).setValues([H_LOTES_OP]);
  const f1 = String(lotes.getRange('F1').getFormula() || '');
  if (!f1) {
    lotes.getRange('F1').setFormula('={"Investido";ARRAYFORMULA(IF(A2:A="",,SUMIFS(' + L + 'H:H,' + L + 'B:B,A2:A,' + L + 'D:D,"Investimento")))}');
    lotes.getRange('G1').setFormula('={"Recolhido";ARRAYFORMULA(IF(A2:A="",,SUMIFS(' + L + 'H:H,' + L + 'B:B,A2:A,' + L + 'D:D,"Recolhimento",' + L + 'I:I,"Realizado")))}');
    lotes.getRange('H1').setFormula('={"A recolher";ARRAYFORMULA(IF(A2:A="",,SUMIFS(' + L + 'H:H,' + L + 'B:B,A2:A,' + L + 'D:D,"Recolhimento",' + L + 'I:I,"Previsto")))}');
    lotes.getRange('I1').setFormula('={"Resultado";ARRAYFORMULA(IF(A2:A="",,G2:G-F2:F))}');
    lotes.getRange('J1').setFormula('={"Situação";ARRAYFORMULA(IF(A2:A="",,IF(H2:H>0,"A recolher",IF(G2:G=0,"Aguardando",IF(I2:I>=0,"Lucro","Prejuízo")))))}');
    formatar = true;
  }
  // Usuários
  const users = aba_(ss, ABA_USERS, H_USERS);
  // Config (legado: capital antes do login)
  const conf = aba_(ss, ABA_CONF, ['Chave', 'Valor']);
  // Resumo
  let res = ss.getSheetByName(ABA_RES);
  const resVelho = !res || String(res.getRange('B2').getFormula()).indexOf(ABA_USERS) < 0;
  if (resVelho) {
    if (!res) res = ss.insertSheet(ABA_RES, 0);
    res.clear();
    const U = "'" + ABA_USERS + "'!";
    const rows = [
      ['Resumo — Controle de Lotes (todos os operadores)', ''],
      ['Capital inicial', '=SUM(' + U + 'G2:G)'],
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
      ['Operadores ativos', '=COUNTIF(' + U + 'F2:F,"Sim")'],
    ];
    res.getRange(1, 1, rows.length, 2).setValues(rows);
    res.getRange('A1').setFontWeight('bold').setFontSize(13);
    res.getRange('A2:A13').setFontWeight('bold');
    res.getRange('B2:B11').setNumberFormat(FMT_MOEDA);
    res.getRange('B6').setNumberFormat('0.0%');
    res.getRange('B12:B13').setNumberFormat('0');
    res.getRange('A9:B9').setBackground('#EEF2F8');
    res.setColumnWidth(1, 300); res.setColumnWidth(2, 160);
    res.getRange('A15').setValue('Esta aba é só leitura. Os lançamentos são feitos pelo painel. O resultado por operador aparece no painel, na aba Operadores.').setFontColor('#666666').setFontStyle('italic');
  }
  if (formatar) {
    lanc.getRange('E2:E').setNumberFormat(FMT_DATA);
    lanc.getRange('G2:H').setNumberFormat(FMT_MOEDA);
    lanc.getRange('K2:K').setNumberFormat(FMT_HORA);
    lanc.hideColumns(1, 2); lanc.hideColumns(12);
    [0, 0, 140, 120, 95, 110, 140, 120, 100, 240, 130, 0, 140].forEach((w, i) => { if (w) lanc.setColumnWidth(i + 1, w); });
    lotes.getRange('A1:L1').setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    lotes.getRange('C2:C').setNumberFormat(FMT_DATA);
    lotes.getRange('E2:E').setNumberFormat(FMT_HORA);
    lotes.getRange('F2:I').setNumberFormat(FMT_MOEDA);
    lotes.hideColumns(1); lotes.hideColumns(11);
    [0, 160, 110, 220, 130, 120, 120, 120, 120, 110, 0, 140].forEach((w, i) => { if (w) lotes.setColumnWidth(i + 1, w); });
    users.getRange('G2:G').setNumberFormat(FMT_MOEDA);
    users.getRange('H2:H').setNumberFormat(FMT_DATA);
    users.getRange('I2:I').setNumberFormat(FMT_HORA);
    users.hideColumns(1); users.hideColumns(4);
    [0, 160, 120, 0, 90, 70, 130, 100, 130].forEach((w, i) => { if (w) users.setColumnWidth(i + 1, w); });
  }
}

function aba_(ss, nome, cab) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  const atual = sh.getRange(1, 1, 1, cab.length).getValues()[0];
  if (atual.join('|') !== cab.join('|')) {
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function chave_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('CHAVE');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 20); p.setProperty('CHAVE', k); }
  return k;
}

function segredo_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SEGREDO');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SEGREDO', s); }
  return s;
}

/* ===================== SENHAS E SESSÃO ===================== */

function hex_(bytes) { return bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }

function hashSenha_(senha, sal) {
  sal = sal || Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  let h = sal + ':' + senha;
  for (let i = 0; i < 300; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + sal, Utilities.Charset.UTF_8));
  return sal + '$' + h;
}

function confereSenha_(senha, guardado) {
  const sal = String(guardado || '').split('$')[0];
  return !!sal && hashSenha_(senha, sal) === guardado;
}

function assinar_(texto) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(texto, segredo_())).replace(/=+$/, '');
}

function emitirToken_(u) {
  const exp = Date.now() + SESSAO_DIAS * 864e5;
  const base = u.id + '.' + exp;
  return base + '.' + assinar_(base + '.' + String(u.senha).slice(-12));
}

function autenticar_(ss, token) {
  const partes = String(token || '').split('.');
  if (partes.length !== 3) throw erro_('sessao', 'Faça login para continuar.');
  const [id, exp] = partes;
  if (Number(exp) < Date.now()) throw erro_('sessao', 'Sua sessão expirou. Entre de novo.');
  const u = usuarios_(ss).find(x => x.id === id);
  if (!u || !u.ativo) throw erro_('sessao', 'Acesso desativado. Fale com o administrador.');
  if (assinar_(id + '.' + exp + '.' + String(u.senha).slice(-12)) !== partes[2]) throw erro_('sessao', 'Sessão inválida. Entre de novo.');
  return u;
}

function erro_(codigo, msg) { const e = new Error(msg); e.codigo = codigo; return e; }
function normUser_(s) { return String(s || '').trim().toLowerCase(); }

/* ===================== API (APP DA WEB) ===================== */

function doGet() {
  return ContentService.createTextOutput('Controle de Lotes: conexão ativa. Use este link no painel.');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return saida_({ ok: false, erro: 'Requisição inválida.' }); }

  const ss = planilha_();
  const lock = LockService.getScriptLock();
  const escrita = ['status', 'listar'].indexOf(req.acao) < 0;
  try {
    if (escrita) lock.waitLock(20000);
    garantir_(ss, false);
    const p = req.dados || {};

    if (req.acao === 'status') return saida_({ ok: true, temAdmin: usuarios_(ss).some(u => u.papel === 'admin') });
    if (req.acao === 'criarAdmin') return saida_(criarAdmin_(ss, p));
    if (req.acao === 'login') return saida_(login_(ss, p));

    const eu = autenticar_(ss, req.token);
    const admin = eu.papel === 'admin';
    const soAdmin = () => { if (!admin) throw new Error('Só o administrador pode fazer isso.'); };

    switch (req.acao) {
      case 'listar': break;
      case 'criarLote': criarLote_(ss, eu, p); break;
      case 'lancar': lancar_(ss, eu, p.lanc); break;
      case 'receber': receber_(ss, eu, p.id, p.data); break;
      case 'excluirLancamento': soAdmin(); excluirLinhas_(ss.getSheetByName(ABA_LANC), r => r[0] === p.id); break;
      case 'excluirLote': soAdmin(); excluirLinhas_(ss.getSheetByName(ABA_LANC), r => r[1] === p.id); excluirLinhas_(ss.getSheetByName(ABA_LOTES), r => r[0] === p.id); break;
      case 'salvarUsuario': soAdmin(); salvarUsuario_(ss, eu, p); break;
      case 'trocarSenha': trocarSenha_(ss, eu, p); break;
      default: throw new Error('Ação desconhecida.');
    }
    if (escrita) SpreadsheetApp.flush();
    const atual = usuarios_(ss).find(u => u.id === eu.id);
    const resp = { ok: true, dados: dados_(ss, atual) };
    if (req.acao === 'trocarSenha') resp.token = emitirToken_(atual);
    return saida_(resp);
  } catch (err) {
    return saida_({ ok: false, codigo: err.codigo || '', erro: String(err && err.message || err) });
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

/* ---------- login ---------- */

function criarAdmin_(ss, p) {
  const chave = PropertiesService.getScriptProperties().getProperty('CHAVE');
  if (!chave) throw new Error('A planilha ainda não foi preparada. Use o menu Controle de Lotes > Preparar planilha.');
  if (usuarios_(ss).some(u => u.papel === 'admin')) throw new Error('O administrador já foi criado. Entre com o seu usuário e senha.');
  if (String(p.chave || '').trim() !== chave) throw erro_('chave', 'Chave de acesso incorreta.');
  const u = novoUsuario_(ss, { nome: p.nome, usuario: p.usuario, senha: p.senha, papel: 'admin', capital: p.capital, inicio: p.inicio });
  // dados lançados antes do login passam a ser do administrador
  adotarOrfaos_(ss, u);
  const conf = linhas_(ss.getSheetByName(ABA_CONF), 2);
  const capLegado = conf.find(r => r[0] === 'capital'), iniLegado = conf.find(r => r[0] === 'inicio');
  if ((p.capital === '' || p.capital == null) && capLegado && Number(capLegado[1]) > 0) {
    const sh = ss.getSheetByName(ABA_USERS);
    sh.getRange(u.linha, 7).setValue(Number(capLegado[1]));
    if (iniLegado && iniLegado[1]) sh.getRange(u.linha, 8).setValue(iniLegado[1]);
  }
  SpreadsheetApp.flush();
  const atual = usuarios_(ss).find(x => x.id === u.id);
  return { ok: true, token: emitirToken_(atual), dados: dados_(ss, atual) };
}

function login_(ss, p) {
  const props = PropertiesService.getScriptProperties();
  const usuario = normUser_(p.usuario);
  const chaveFalha = 'FALHA_' + usuario;
  const falha = JSON.parse(props.getProperty(chaveFalha) || '{"n":0,"ate":0}');
  if (falha.ate > Date.now()) {
    const min = Math.ceil((falha.ate - Date.now()) / 60000);
    throw new Error('Muitas tentativas erradas. Tente de novo em ' + min + ' min.');
  }
  const u = usuarios_(ss).find(x => x.usuario === usuario);
  if (!u || !confereSenha_(String(p.senha || ''), u.senha)) {
    falha.n++;
    if (falha.n >= MAX_FALHAS) { falha.ate = Date.now() + BLOQUEIO_MIN * 60000; falha.n = 0; }
    props.setProperty(chaveFalha, JSON.stringify(falha));
    throw new Error('Usuário ou senha incorretos.');
  }
  if (!u.ativo) throw new Error('Este acesso está desativado. Fale com o administrador.');
  props.deleteProperty(chaveFalha);
  return { ok: true, token: emitirToken_(u), dados: dados_(ss, u) };
}

function trocarSenha_(ss, eu, p) {
  if (!confereSenha_(String(p.atual || ''), eu.senha)) throw new Error('A senha atual não confere.');
  validaSenha_(p.nova);
  ss.getSheetByName(ABA_USERS).getRange(eu.linha, 4).setValue(hashSenha_(p.nova));
}

function validaSenha_(s) { if (String(s || '').length < 6) throw new Error('A senha precisa ter pelo menos 6 caracteres.'); }

/* ---------- usuários ---------- */

function usuarios_(ss) {
  const sh = ss.getSheetByName(ABA_USERS);
  if (!sh) return [];
  const tz = ss.getSpreadsheetTimeZone();
  const last = ultimaLinha_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 9).getValues().map((r, i) => ({
    linha: i + 2, id: String(r[0]), nome: String(r[1]), usuario: normUser_(r[2]), senha: String(r[3]),
    papel: String(r[4]).toLowerCase() === 'admin' ? 'admin' : 'operador',
    ativo: String(r[5]).toLowerCase() !== 'não' && String(r[5]).toLowerCase() !== 'nao',
    capital: Number(r[6]) || 0, inicio: iso_(r[7], tz) || null, createdAt: ms_(r[8])
  })).filter(u => u.id);
}

function publico_(u) {
  return { id: u.id, nome: u.nome, usuario: u.usuario, papel: u.papel, ativo: u.ativo, capital: u.capital, inicio: u.inicio, createdAt: u.createdAt };
}

function novoUsuario_(ss, p) {
  const nome = String(p.nome || '').trim();
  const usuario = normUser_(p.usuario);
  if (!nome) throw new Error('Informe o nome.');
  if (!/^[a-z0-9._-]{3,30}$/.test(usuario)) throw new Error('O usuário deve ter de 3 a 30 caracteres: letras, números, ponto, hífen ou sublinhado, sem espaço.');
  if (usuarios_(ss).some(u => u.usuario === usuario)) throw new Error('Já existe alguém com o usuário “' + usuario + '”.');
  validaSenha_(p.senha);
  const cap = p.capital === '' || p.capital == null ? 0 : Number(p.capital);
  if (!(cap >= 0)) throw new Error('Capital inválido.');
  const sh = ss.getSheetByName(ABA_USERS);
  const linha = ultimaLinha_(sh) + 1;
  const id = novoId_('U');
  sh.getRange(linha, 1, 1, 9).setValues([[id, nome, usuario, hashSenha_(p.senha), p.papel === 'admin' ? 'Admin' : 'Operador', 'Sim', cap, p.inicio ? data_(p.inicio) : '', new Date()]]);
  return { id, linha, nome, usuario };
}

function salvarUsuario_(ss, eu, p) {
  if (!p.id) { novoUsuario_(ss, p); return; }
  const u = usuarios_(ss).find(x => x.id === p.id);
  if (!u) throw new Error('Usuário não encontrado.');
  const sh = ss.getSheetByName(ABA_USERS);
  const nome = String(p.nome || '').trim();
  if (!nome) throw new Error('Informe o nome.');
  if (u.id === eu.id && (p.ativo === false || p.papel === 'operador')) throw new Error('Você não pode desativar nem tirar o seu próprio acesso de administrador.');
  const cap = Number(p.capital);
  if (!(cap >= 0)) throw new Error('Capital inválido.');
  sh.getRange(u.linha, 2).setValue(nome);
  sh.getRange(u.linha, 5).setValue(p.papel === 'admin' ? 'Admin' : 'Operador');
  sh.getRange(u.linha, 6).setValue(p.ativo === false ? 'Não' : 'Sim');
  sh.getRange(u.linha, 7).setValue(cap);
  sh.getRange(u.linha, 8).setValue(p.inicio ? data_(p.inicio) : '');
  if (p.novaSenha) { validaSenha_(p.novaSenha); sh.getRange(u.linha, 4).setValue(hashSenha_(p.novaSenha)); }
  if (nome !== u.nome) renomearNasAbas_(ss, u.id, nome);
}

function renomearNasAbas_(ss, id, nome) {
  [[ABA_LOTES, 11], [ABA_LANC, 12]].forEach(([aba, col]) => {
    const sh = ss.getSheetByName(aba);
    const last = ultimaLinha_(sh);
    if (last < 2) return;
    const ids = sh.getRange(2, col, last - 1, 1).getValues();
    ids.forEach((r, i) => { if (r[0] === id) sh.getRange(i + 2, col + 1).setValue(nome); });
  });
}

function adotarOrfaos_(ss, u) {
  [[ABA_LOTES, 11], [ABA_LANC, 12]].forEach(([aba, col]) => {
    const sh = ss.getSheetByName(aba);
    const last = ultimaLinha_(sh);
    if (last < 2) return;
    const vals = sh.getRange(2, col, last - 1, 2).getValues();
    let mudou = false;
    vals.forEach(r => { if (!r[0]) { r[0] = u.id; r[1] = u.nome; mudou = true; } });
    if (mudou) sh.getRange(2, col, last - 1, 2).setValues(vals);
  });
}

/* ---------- leitura ---------- */

function dados_(ss, eu) {
  const tz = ss.getSpreadsheetTimeZone();
  const admin = eu.papel === 'admin';
  const users = usuarios_(ss);
  const lotes = linhas_(ss.getSheetByName(ABA_LOTES), 12)
    .filter(r => admin || r[10] === eu.id)
    .map(r => ({ id: String(r[0]), nome: String(r[1]), data: iso_(r[2], tz), obs: String(r[3] || ''), createdAt: ms_(r[4]), op: String(r[10] || '') }));
  const lanc = linhas_(ss.getSheetByName(ABA_LANC), 13)
    .filter(r => admin || r[11] === eu.id)
    .map(r => ({
      id: String(r[0]), loteId: r[1] ? String(r[1]) : null, tipo: String(r[3]).toLowerCase(), data: iso_(r[4], tz),
      qtd: r[5] === '' ? null : Number(r[5]), valorOp: r[6] === '' ? null : Number(r[6]), valor: Number(r[7]) || 0,
      status: String(r[8]).toLowerCase() === 'previsto' ? 'previsto' : 'realizado', obs: String(r[9] || ''), createdAt: ms_(r[10]),
      op: String(r[11] || '')
    }));
  return {
    eu: publico_(eu),
    usuarios: (admin ? users : users.filter(u => u.id === eu.id)).map(publico_),
    lotes, lanc,
    planilha: admin ? ss.getUrl() : ''
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

/** Operador sempre lança para si; admin pode escolher o operador. */
function alvo_(ss, eu, opId) {
  if (eu.papel !== 'admin' || !opId || opId === eu.id) return eu;
  const u = usuarios_(ss).find(x => x.id === opId);
  if (!u) throw new Error('Operador não encontrado.');
  if (!u.ativo) throw new Error(u.nome + ' está com o acesso desativado.');
  return u;
}

function criarLote_(ss, eu, p) {
  const lote = p.lote || {};
  const dono = alvo_(ss, eu, p.operadorId);
  const nome = String(lote.nome || '').trim();
  if (!nome) throw new Error('Dê um nome ao lote.');
  const sh = ss.getSheetByName(ABA_LOTES);
  const existe = linhas_(sh, 11).some(r => r[10] === dono.id && String(r[1]).trim().toLowerCase() === nome.toLowerCase());
  if (existe) throw new Error((dono.id === eu.id ? 'Você já tem' : dono.nome + ' já tem') + ' um lote chamado “' + nome + '”. Escolha outro nome.');
  const id = novoId_('L');
  const linha = ultimaLinha_(sh) + 1;
  sh.getRange(linha, 1, 1, 5).setValues([[id, nome, data_(lote.data), String(lote.obs || ''), new Date()]]);
  sh.getRange(linha, 11, 1, 2).setValues([[dono.id, dono.nome]]);
  (p.lancamentos || []).forEach(l => gravarLanc_(ss, dono, Object.assign({}, l, { loteId: id }), nome));
}

function lancar_(ss, eu, l) {
  if (!l || !TIPOS[l.tipo]) throw new Error('Tipo de lançamento inválido.');
  if (l.tipo === 'investimento' || l.tipo === 'recolhimento') {
    const lote = linhas_(ss.getSheetByName(ABA_LOTES), 11).find(r => r[0] === l.loteId);
    if (!lote) throw new Error('Lote não encontrado. Atualize o painel e tente de novo.');
    if (eu.papel !== 'admin' && lote[10] !== eu.id) throw new Error('Esse lote não é seu.');
    const dono = usuarios_(ss).find(u => u.id === lote[10]) || eu;
    gravarLanc_(ss, dono, l, lote[1]);
  } else {
    if (eu.papel !== 'admin') throw new Error('Só o administrador lança aportes e retiradas.');
    gravarLanc_(ss, alvo_(ss, eu, l.operadorId), Object.assign({}, l, { loteId: '' }), 'Caixa');
  }
}

function gravarLanc_(ss, dono, l, nomeLote) {
  if (!TIPOS[l.tipo]) throw new Error('Tipo de lançamento inválido.');
  const valor = Number(l.valor);
  if (!(valor > 0)) throw new Error('Informe um valor maior que zero.');
  const sh = ss.getSheetByName(ABA_LANC);
  const linha = ultimaLinha_(sh) + 1;
  sh.getRange(linha, 1, 1, 13).setValues([[
    novoId_('M'), l.loteId || '', nomeLote || 'Caixa', TIPOS[l.tipo], data_(l.data),
    l.qtd ? Number(l.qtd) : '', l.valorOp ? Number(l.valorOp) : '', valor,
    l.status === 'previsto' && l.tipo === 'recolhimento' ? SITUACAO.previsto : SITUACAO.realizado,
    String(l.obs || ''), new Date(), dono.id, dono.nome
  ]]);
}

function receber_(ss, eu, id, iso) {
  const sh = ss.getSheetByName(ABA_LANC);
  const linha = linhaPorId_(sh, id);
  if (!linha) throw new Error('Lançamento não encontrado.');
  if (eu.papel !== 'admin' && sh.getRange(linha, 12).getValue() !== eu.id) throw new Error('Esse lançamento não é seu.');
  sh.getRange(linha, 9).setValue(SITUACAO.realizado);
  if (iso) sh.getRange(linha, 5).setValue(data_(iso));
}

function excluirLinhas_(sh, teste) {
  const last = ultimaLinha_(sh);
  if (last < 2) return;
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (teste(vals[i])) sh.deleteRow(i + 2);
}

function linhaPorId_(sh, id) {
  const last = ultimaLinha_(sh);
  if (last < 2) return 0;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]) === String(id)) return i + 2;
  return 0;
}
