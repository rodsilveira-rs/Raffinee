/**
 * RAFFINÉE ANALYTICS — Backend (Google Apps Script)
 * ---------------------------------------------------------------
 * Projeto Apps Script VINCULADO à planilha "Raffinee Analytics".
 * O index (GitHub Pages) lê e grava por esta Web App.
 *
 * Implantação:
 *   1) Rode setup() uma vez (cria as abas e a CHAVE_ACESSO na aba CONFIG).
 *   2) Implantar > Nova implantação > App da Web
 *      Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 *   3) Copie a URL /exec e cole no index (Configurações), junto com a chave.
 *   Atualizações: sempre "Gerenciar implantações > Editar > Nova versão"
 *   na MESMA implantação (a URL não muda).
 */

var VERSAO_BACKEND = '1.1.0';

var ABAS = {
  BASE_VENDAS: ['ID_VENDA', 'DATA', 'ANO', 'MES', 'PEDIDO', 'COD_CLIENTE', 'CLIENTE', 'VENDEDOR', 'COD_PRODUTO', 'PRODUTO', 'CANAL', 'QTD', 'KG', 'VALOR', 'ORIGEM', 'IMPORTADO_EM'],
  CLIENTES:    ['COD_CLIENTE', 'CLIENTE', 'VENDEDOR', 'CANAL', 'CIDADE', 'UF', 'STATUS'],
  PRODUTOS:    ['COD_PRODUTO', 'PRODUTO', 'CATEGORIA', 'LINHA', 'PESO_KG', 'STATUS'],
  VENDEDORES:  ['COD_VENDEDOR', 'VENDEDOR', 'STATUS'],
  METAS:       ['ANO', 'MES', 'TIPO', 'CHAVE', 'META_VALOR', 'META_KG', 'ATUALIZADO_EM'],
  ARQUIVOS:    ['ID', 'TIPO', 'NOME', 'PERIODO', 'DATA_GERACAO', 'LINK_DRIVE', 'OBS'],
  CONFIG:      ['CHAVE', 'VALOR', 'DESCRICAO'],
  LOG:         ['DATA_HORA', 'ACAO', 'DETALHE']
};

// Colunas das tabelas lidas pelo index (ordem do JSON enviado)
var LEITURA = {
  vendas:     { aba: 'BASE_VENDAS', cols: ['DATA', 'PEDIDO', 'COD_CLIENTE', 'CLIENTE', 'VENDEDOR', 'COD_PRODUTO', 'PRODUTO', 'CANAL', 'QTD', 'KG', 'VALOR', 'ORIGEM'] },
  clientes:   { aba: 'CLIENTES',   cols: ABAS.CLIENTES },
  produtos:   { aba: 'PRODUTOS',   cols: ABAS.PRODUTOS },
  vendedores: { aba: 'VENDEDORES', cols: ABAS.VENDEDORES },
  metas:      { aba: 'METAS',      cols: ['ANO', 'MES', 'TIPO', 'CHAVE', 'META_VALOR', 'META_KG'] },
  arquivos:   { aba: 'ARQUIVOS',   cols: ABAS.ARQUIVOS }
};

var COR_VINHO = '#5A1E24';
var COR_CREME = '#F6EFE8';

/* ============================================================
 * SETUP — cria/ajusta a estrutura da planilha
 * ============================================================ */
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(ABAS).forEach(function (nome) {
    var sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
    var cab = ABAS[nome];
    sh.getRange(1, 1, 1, cab.length).setValues([cab])
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(COR_VINHO);
    sh.setFrozenRows(1);
    sh.setTabColor(nome === 'CONFIG' || nome === 'LOG' ? '#C9A46A' : COR_VINHO);
  });

  var cfg = ss.getSheetByName('CONFIG');
  var atuais = lerConfig_();
  var padrao = [
    ['CHAVE_ACESSO', Utilities.getUuid().split('-')[0] + Utilities.getUuid().split('-')[1], 'Chave usada pelo index. Troque quando quiser revogar o acesso.'],
    ['EMPRESA', 'Raffinée', 'Nome exibido no sistema'],
    ['PASTA_DRIVE_ID', '', 'ID da pasta do Drive onde os PDFs enviados pelo sistema serão guardados'],
    ['DIAS_INATIVO', '120', 'Dias sem compra para considerar cliente inativo'],
    ['FATOR_ATRASO', '1.5', 'Cliente em risco quando dias sem compra > frequência média × fator']
  ];
  padrao.forEach(function (l) {
    if (!(l[0] in atuais)) cfg.appendRow(l);
  });

  // Formatos úteis
  var v = ss.getSheetByName('BASE_VENDAS');
  v.getRange('B:B').setNumberFormat('dd/mm/yyyy');
  v.getRange('L:M').setNumberFormat('#,##0.00');
  v.getRange('N:N').setNumberFormat('R$ #,##0.00');
  ss.getSheetByName('METAS').getRange('E:E').setNumberFormat('R$ #,##0.00');

  // Remove a "Página1" vazia, se existir
  ['Página1', 'Sheet1', 'Planilha1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  log_('SETUP', 'Estrutura criada/atualizada (backend v' + VERSAO_BACKEND + ')');
  Logger.log('CHAVE_ACESSO = ' + lerConfig_().CHAVE_ACESSO);
}

/* ============================================================
 * ROTAS
 * ============================================================ */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'ping') return json_({ ok: true, versao: VERSAO_BACKEND });
    autenticar_(p.k);
    if (p.action === 'dados') return json_({ ok: true, versao: VERSAO_BACKEND, geradoEm: new Date().toISOString(), dados: lerTudo_() });
    return json_({ ok: false, erro: 'Ação inválida' });
  } catch (err) {
    return json_({ ok: false, erro: String(err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents || '{}');
    autenticar_(body.k);
    var p = body.payload || {};
    var r;
    switch (body.action) {
      case 'importarVendas':  r = importarVendas_(p.linhas || [], p.modo || 'acrescentar', p.origem || 'IMPORTACAO'); break;
      case 'lancarVenda':     r = importarVendas_([p], 'acrescentar', 'MANUAL'); break;
      case 'salvarMetas':     r = salvarMetas_(p); break;
      case 'importarPlanejamento': r = importarPlanejamento_(p); break;
      case 'salvarCadastro':  r = salvarCadastro_(p.tipo, p.registro); break;
      case 'salvarArquivo':   r = salvarArquivo_(p); break;
      case 'excluirArquivo':  r = excluirArquivo_(p.id); break;
      case 'enviarPdf':       r = enviarPdf_(p); break;
      default: throw new Error('Ação inválida: ' + body.action);
    }
    return json_({ ok: true, resultado: r });
  } catch (err) {
    return json_({ ok: false, erro: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ============================================================
 * LEITURA
 * ============================================================ */
function lerTudo_() {
  var out = {};
  Object.keys(LEITURA).forEach(function (k) {
    out[k] = lerTabela_(LEITURA[k].aba, LEITURA[k].cols);
  });
  return out;
}

/** Retorna { cols:[...], rows:[[...],...] } — formato compacto */
function lerTabela_(aba, cols) {
  var sh = SpreadsheetApp.getActive().getSheetByName(aba);
  if (!sh || sh.getLastRow() < 2) return { cols: cols, rows: [] };
  var vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var cab = vals[0].map(function (c) { return String(c).trim().toUpperCase(); });
  var idx = cols.map(function (c) { return cab.indexOf(c); });
  var tz = Session.getScriptTimeZone();
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var l = vals[i];
    if (l.join('') === '') continue;
    rows.push(idx.map(function (j) {
      if (j < 0) return '';
      var v = l[j];
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      return v;
    }));
  }
  return { cols: cols, rows: rows };
}

/* ============================================================
 * GRAVAÇÃO
 * ============================================================ */
function importarVendas_(linhas, modo, origem) {
  var sh = aba_('BASE_VENDAS');
  if (modo === 'substituir' && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  }
  // Chaves já existentes (PEDIDO|COD_PRODUTO|DATA) para não duplicar
  var existentes = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(function (l) {
      existentes[chaveVenda_(l[4], l[8], l[1])] = true;
    });
  }
  var agora = new Date();
  var novas = [], duplicadas = 0;
  linhas.forEach(function (x) {
    var data = paraData_(x.data);
    if (!data) return;
    var k = chaveVenda_(x.pedido, x.cod_produto, data);
    if (x.pedido && existentes[k]) { duplicadas++; return; }
    existentes[k] = true;
    novas.push([
      Utilities.getUuid().slice(0, 8), data, data.getFullYear(), data.getMonth() + 1,
      x.pedido || '', x.cod_cliente || '', x.cliente || '', x.vendedor || '',
      x.cod_produto || '', x.produto || '', x.canal || '',
      num_(x.qtd), num_(x.kg), num_(x.valor), origem, agora
    ]);
  });
  // Vendas reais substituem o resumo do planejamento nos meses em que chegaram
  if (origem !== 'PLANEJAMENTO' && novas.length) {
    var mesesNovos = {};
    novas.forEach(function (l) { mesesNovos[l[2] + '|' + l[3]] = true; });
    removerLinhas_(sh, function (l) { return String(l[14]) === 'PLANEJAMENTO' && mesesNovos[Number(l[2]) + '|' + Number(l[3])]; });
  }
  if (novas.length) sh.getRange(sh.getLastRow() + 1, 1, novas.length, novas[0].length).setValues(novas);
  log_('VENDAS', origem + ': ' + novas.length + ' linhas gravadas, ' + duplicadas + ' duplicadas ignoradas' + (modo === 'substituir' ? ' (base substituída)' : ''));
  return { gravadas: novas.length, duplicadas: duplicadas };
}

/** payload: { ano, tipo, chave, meses:[{mes, valor, kg}] } — substitui as linhas dessa meta */
function salvarMetas_(p) {
  var sh = aba_('METAS');
  var ano = Number(p.ano), tipo = String(p.tipo).toUpperCase(), chave = String(p.chave || 'EMPRESA');
  var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues() : [];
  var manter = vals.filter(function (l) {
    return !(Number(l[0]) === ano && String(l[2]).toUpperCase() === tipo && String(l[3]) === chave);
  });
  var agora = new Date();
  (p.meses || []).forEach(function (m) {
    if (num_(m.valor) || num_(m.kg)) manter.push([ano, Number(m.mes), tipo, chave, num_(m.valor), num_(m.kg), agora]);
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
  if (manter.length) sh.getRange(2, 1, manter.length, 7).setValues(manter);
  log_('METAS', tipo + ' / ' + chave + ' / ' + ano + ' atualizada');
  return { linhas: manter.length };
}

/**
 * Planilha de planejamento da Raffinée (lida e montada pelo index).
 * payload: { anoMeta, anoHist, metas[], produtos[], clientes[], vendedores[], historico[] }
 */
function importarPlanejamento_(p) {
  var anoMeta = Number(p.anoMeta), anoHist = Number(p.anoHist), agora = new Date();

  // Metas do ano: substitui tudo daquele ano
  var shM = aba_('METAS');
  var metas = shM.getLastRow() > 1 ? shM.getRange(2, 1, shM.getLastRow() - 1, 7).getValues() : [];
  metas = metas.filter(function (l) { return Number(l[0]) !== anoMeta; });
  (p.metas || []).forEach(function (m) { metas.push([anoMeta, Number(m.mes), String(m.tipo).toUpperCase(), String(m.chave), num_(m.valor), num_(m.kg), agora]); });
  if (shM.getLastRow() > 1) shM.getRange(2, 1, shM.getLastRow() - 1, 7).clearContent();
  if (metas.length) shM.getRange(2, 1, metas.length, 7).setValues(metas);

  // Cadastros
  var nProd = 0, nCli = 0, nVend = 0;
  (p.produtos || []).forEach(function (r) { if (inserirSeNaoExiste_('PRODUTOS', 0, r)) nProd++; });
  nCli = upsertLote_('CLIENTES', 0, p.clientes || []);
  (p.vendedores || []).forEach(function (r) { if (inserirSeNaoExiste_('VENDEDORES', 1, r)) nVend++; });

  // Histórico mensal por produto (resumo): troca o anterior do mesmo ano, sem mexer em meses que já têm vendas reais
  var shV = aba_('BASE_VENDAS');
  var reais = {};
  if (shV.getLastRow() > 1) shV.getRange(2, 1, shV.getLastRow() - 1, 15).getValues().forEach(function (l) {
    if (Number(l[2]) === anoHist && String(l[14]) !== 'PLANEJAMENTO') reais[Number(l[3])] = true;
  });
  removerLinhas_(shV, function (l) { return String(l[14]) === 'PLANEJAMENTO' && Number(l[2]) === anoHist; });
  var hist = (p.historico || []).filter(function (h) { return !reais[Number(String(h.data).slice(5, 7))]; });
  var r = hist.length ? importarVendas_(hist, 'acrescentar', 'PLANEJAMENTO') : { gravadas: 0 };

  log_('PLANEJAMENTO', 'Metas ' + anoMeta + ': ' + (p.metas || []).length + ' | clientes ' + nCli + ' | produtos novos ' + nProd + ' | vendedores novos ' + nVend + ' | histórico ' + anoHist + ': ' + r.gravadas + ' linhas');
  return { metas: (p.metas || []).length, clientes: nCli, produtosNovos: nProd, vendedoresNovos: nVend, historico: r.gravadas };
}

/** Insere o registro só se a coluna-chave (índice) ainda não existir na aba */
function inserirSeNaoExiste_(nomeAba, idx, reg) {
  var sh = aba_(nomeAba), cab = ABAS[nomeAba];
  var linha = cab.map(function (c) { return reg[c] !== undefined ? reg[c] : ''; });
  var n = sh.getLastRow();
  var ex = n > 1 ? sh.getRange(2, idx + 1, n - 1, 1).getValues().map(function (l) { return String(l[0]).trim().toUpperCase(); }) : [];
  if (ex.indexOf(String(linha[idx]).trim().toUpperCase()) >= 0) return false;
  sh.appendRow(linha);
  return true;
}

/** Atualiza ou insere vários registros de uma vez, pela coluna-chave (índice) */
function upsertLote_(nomeAba, idx, regs) {
  if (!regs.length) return 0;
  var sh = aba_(nomeAba), cab = ABAS[nomeAba], n = sh.getLastRow();
  var vals = n > 1 ? sh.getRange(2, 1, n - 1, cab.length).getValues() : [];
  var pos = {}; vals.forEach(function (l, i) { pos[String(l[idx]).trim().toUpperCase()] = i; });
  regs.forEach(function (reg) {
    var linha = cab.map(function (c) { return reg[c] !== undefined ? reg[c] : ''; });
    var k = String(linha[idx]).trim().toUpperCase();
    if (k in pos) vals[pos[k]] = vals[pos[k]].map(function (v, j) { return linha[j] !== '' ? linha[j] : v; });
    else { pos[k] = vals.length; vals.push(linha); }
  });
  sh.getRange(2, 1, vals.length, cab.length).setValues(vals);
  return regs.length;
}

/** Reescreve a aba sem as linhas em que remover(linha) for verdadeiro */
function removerLinhas_(sh, remover) {
  var n = sh.getLastRow(); if (n < 2) return 0;
  var w = sh.getLastColumn();
  var vals = sh.getRange(2, 1, n - 1, w).getValues();
  var manter = vals.filter(function (l) { return !remover(l); });
  var tirou = vals.length - manter.length;
  if (!tirou) return 0;
  sh.getRange(2, 1, n - 1, w).clearContent();
  if (manter.length) sh.getRange(2, 1, manter.length, w).setValues(manter);
  return tirou;
}

/** tipo: clientes | produtos | vendedores — upsert pela 1ª coluna (código) */
function salvarCadastro_(tipo, reg) {
  var nomeAba = { clientes: 'CLIENTES', produtos: 'PRODUTOS', vendedores: 'VENDEDORES' }[tipo];
  if (!nomeAba) throw new Error('Cadastro inválido');
  var sh = aba_(nomeAba), cab = ABAS[nomeAba];
  var linha = cab.map(function (c) { return reg[c] !== undefined ? reg[c] : (reg[c.toLowerCase()] || ''); });
  if (!linha[0]) throw new Error('Informe o código');
  var n = sh.getLastRow();
  var codigos = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(function (l) { return String(l[0]); }) : [];
  var pos = codigos.indexOf(String(linha[0]));
  if (pos >= 0) sh.getRange(pos + 2, 1, 1, cab.length).setValues([linha]);
  else sh.appendRow(linha);
  log_('CADASTRO', nomeAba + ' ' + linha[0] + (pos >= 0 ? ' atualizado' : ' criado'));
  return { codigo: linha[0], acao: pos >= 0 ? 'atualizado' : 'criado' };
}

function salvarArquivo_(p) {
  var sh = aba_('ARQUIVOS');
  var id = p.id || ('ARQ-' + Utilities.getUuid().slice(0, 6).toUpperCase());
  var linha = [id, p.tipo || 'Documento', p.nome || 'Sem nome', p.periodo || '', new Date(), p.link || '', p.obs || ''];
  var n = sh.getLastRow();
  var ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(function (l) { return String(l[0]); }) : [];
  var pos = ids.indexOf(id);
  if (pos >= 0) sh.getRange(pos + 2, 1, 1, linha.length).setValues([linha]); else sh.appendRow(linha);
  log_('ARQUIVO', 'Registrado: ' + linha[2]);
  return { id: id };
}

function excluirArquivo_(id) {
  var sh = aba_('ARQUIVOS');
  var n = sh.getLastRow();
  if (n < 2) return { removido: false };
  var ids = sh.getRange(2, 1, n - 1, 1).getValues().map(function (l) { return String(l[0]); });
  var pos = ids.indexOf(String(id));
  if (pos >= 0) { sh.deleteRow(pos + 2); log_('ARQUIVO', 'Removido da central: ' + id); }
  return { removido: pos >= 0 };
}

/** Recebe um PDF em base64, salva na pasta do Drive e registra na Central de Arquivos */
function enviarPdf_(p) {
  var cfg = lerConfig_();
  var pasta = cfg.PASTA_DRIVE_ID ? DriveApp.getFolderById(cfg.PASTA_DRIVE_ID) : DriveApp.getRootFolder();
  var blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mime || 'application/pdf', p.nomeArquivo || 'arquivo.pdf');
  var f = pasta.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return salvarArquivo_({ tipo: p.tipo, nome: p.nome || f.getName(), periodo: p.periodo, link: f.getUrl(), obs: p.obs });
}

/* ============================================================
 * UTILITÁRIOS
 * ============================================================ */
function autenticar_(k) {
  var chave = String(lerConfig_().CHAVE_ACESSO || '');
  if (!chave || String(k || '') !== chave) throw new Error('Acesso não autorizado');
}

function lerConfig_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('CONFIG');
  var o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (l) {
    if (l[0]) o[String(l[0]).trim()] = String(l[1]).trim();
  });
  return o;
}

function aba_(nome) {
  var sh = SpreadsheetApp.getActive().getSheetByName(nome);
  if (!sh) throw new Error('Aba ' + nome + ' não existe. Rode setup().');
  return sh;
}

function log_(acao, detalhe) {
  var sh = SpreadsheetApp.getActive().getSheetByName('LOG');
  if (sh) sh.appendRow([new Date(), acao, detalhe]);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function num_(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).replace(/R\$|\s/g, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.'); // formato pt-BR
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/** Aceita 2026-09-29, 29/09/2026 ou Date */
function paraData_(v) {
  if (v instanceof Date) return v;
  var s = String(v || '').trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function chaveVenda_(pedido, codProduto, data) {
  var d = data instanceof Date ? Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data);
  return [String(pedido), String(codProduto), d].join('|');
}
