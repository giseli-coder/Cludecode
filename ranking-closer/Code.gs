/**
 * RANKING CLOSER — V4 COMPANY · M&GA
 * Backend Google Apps Script
 *
 * REGRA:  quem tiver MAIS CASH (vendas ASSINADAS e PAGAS) fica em 1º. Top 5.
 * DESEMPATE: ordem alfabética.
 *
 * DE ONDE VÊM OS NÚMEROS
 *   Automático, da planilha "Metrics - Franchising", aba do mês (ex.: "Outubro-26").
 *   Cada linha é uma venda. Só entra a linha que tiver:
 *     - STATUS = "Assinado"
 *     - PAGO marcado (✔)
 *     - o closer entre parênteses na coluna SQUAD, ex.: "Danillo (Fink)" -> closer = Fink
 *   O valor somado é a coluna "VALOR (R$)".
 *   Venda de dois closers, ex.: "Oreste (Isis e Fink)": o valor é dividido igualmente
 *   (mude DIVIDIR_VENDA_COMPARTILHADA para false se cada um deve receber o valor inteiro).
 *
 * ESTA PLANILHA (a do ranking) só guarda:
 *   Aba "Imgs": A = nome do closer | B = link da foto no Google Drive
 *   (o nome não precisa ser igual: "Rodrigo" encontra "Rodrigo Monteiro")
 */

// ═══════════════ CONFIGURAÇÃO (é só aqui que você mexe) ═══════════════
const VENDAS_ID = '16qvgipLeodDaRb2gOjAF5KSSsWoMesQX5uCbjKRotHQ';  // trecho da URL entre /d/ e /edit
const GID_ABA   = 565140406;       // número que aparece depois de "gid=" na URL da aba do mês
const ABA_VENDAS = 'Outubro-26';   // só é usado se o GID não for encontrado (troque quando virar o mês)

// Quem participa da competição (o nome pode ser parcial e sem acento: "Junior" encontra "JUNIOR").
// Para entrar ou sair alguém, é só mexer nesta lista.
const PARTICIPANTES = ['Junior', 'Luiz', 'Fernando', 'Tayla', 'Isis'];

const DIVIDIR_VENDA_COMPARTILHADA = true;

const ABA_IMGS = 'Imgs';
const TOP      = 5;

// ═══════════════ SERVIR HTML ═══════════════
// O arquivo HTML do projeto precisa se chamar exatamente "index".
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ranking Closer — M&GA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════ UTILITÁRIOS ═══════════════
// Maiúsculas, sem acento e sem espaços repetidos: "  Conversão " -> "CONVERSAO"
function norm(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

// "Rodrigo" ~ "RODRIGO MONTEIRO"  ·  "Luiz" ~ "LUIZ"
function mesmoNome(a, b) {
  const x = norm(a), y = norm(b);
  if (!x || !y) return false;
  return x === y || x.indexOf(y + ' ') === 0 || y.indexOf(x + ' ') === 0;
}

// Aceita número, "R$ 1.234,56", "-R$ 500,00", "1234.5"
function dinheiro(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  const neg = /^\(.*\)$/.test(s) || s.indexOf('-') >= 0;
  s = s.replace(/[^\d.,]/g, '');
  if (!s) return 0;
  if (s.indexOf(',') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');                 // 1.234,56 -> 1234.56
  } else if (/^\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, '');                                   // 1.234 -> 1234
  }
  const n = parseFloat(s);
  if (!isFinite(n)) return 0;
  return neg ? -n : n;
}

function abrirVendas() {
  const id = String(VENDAS_ID || '').trim();
  if (!id) throw new Error('Cole o ID da planilha de vendas no topo do Code.gs (const VENDAS_ID).');
  return SpreadsheetApp.openById(id);
}

// Acha a aba pelo GID (o número depois de "gid=" na URL); se não achar, pelo nome.
function acharAba(ss, gid, nome) {
  const abas = ss.getSheets();
  for (let i = 0; i < abas.length; i++) {
    if (gid && abas[i].getSheetId() === Number(gid)) return abas[i];
  }
  const alvo = norm(nome);
  for (let i = 0; i < abas.length; i++) {
    if (norm(abas[i].getName()) === alvo) return abas[i];
  }
  throw new Error('Não achei a aba (gid ' + gid + ' / nome "' + nome + '") na planilha de vendas.');
}

// ═══════════════ LER AS VENDAS ═══════════════
// Procuro o cabeçalho "SQUAD" e, à direita dele, "VALOR", "STATUS" e "PAGO".
// Cada linha de baixo é uma venda: "Nome do cliente (Closer)".
function ehPago(v) {
  if (v === true) return true;
  const t = norm(v);
  return t === 'TRUE' || t === 'SIM' || t === 'X' || t === 'PAGO' || t === '✔' || t === '✓';
}

function ehAssinado(v) {
  const t = norm(v);
  return t.indexOf('ASSINAD') === 0;
}

// "Oreste (Isis e Fink)" -> ["Isis", "Fink"]   ·   "Wellington ( Luiz)" -> ["Luiz"]
function closersDaLinha(texto) {
  const m = String(texto == null ? '' : texto).match(/\(([^()]*)\)/g);
  if (!m) return [];
  const dentro = m[m.length - 1].replace(/^\(|\)$/g, '');
  return dentro.split(/\s*(?:,|;|\/|\+|&|\be\b)\s*/i).map(s => s.trim()).filter(Boolean);
}

function lerVendas(values) {
  // cabeçalho: linha onde existe a célula "SQUAD"
  let hr = -1, cSq = -1;
  for (let r = 0; r < Math.min(values.length, 20) && hr < 0; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (norm(values[r][c]) === 'SQUAD') { hr = r; cSq = c; break; }
    }
  }
  if (hr < 0) throw new Error('Não achei o cabeçalho "SQUAD" nas primeiras linhas da aba.');

  let cVal = -1, cSt = -1, cPg = -1;
  for (let c = cSq; c < Math.min(values[hr].length, cSq + 20); c++) {
    const t = norm(values[hr][c]);
    if (cVal < 0 && t.indexOf('VALOR') === 0) cVal = c;
    if (cSt < 0 && t === 'STATUS') cSt = c;
    if (cPg < 0 && t === 'PAGO') cPg = c;
  }
  if (cVal < 0 || cSt < 0 || cPg < 0) {
    throw new Error('No cabeçalho achei SQUAD, mas faltou: ' +
      [cVal < 0 ? 'VALOR (R$)' : '', cSt < 0 ? 'STATUS' : '', cPg < 0 ? 'PAGO' : ''].filter(Boolean).join(', ') + '.');
  }

  const vendas = [];
  for (let r = hr + 1; r < values.length; r++) {
    const closers = closersDaLinha(values[r][cSq]);
    if (!closers.length) continue;
    if (!ehAssinado(values[r][cSt]) || !ehPago(values[r][cPg])) continue;
    vendas.push({ cliente: String(values[r][cSq]).trim(), closers: closers, valor: dinheiro(values[r][cVal]) });
  }
  return { vendas: vendas, linhaCab: hr + 1 };
}

// ═══════════════ FOTOS (aba Imgs desta planilha) ═══════════════
function lerFotos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss ? ss.getSheetByName(ABA_IMGS) : null;
  const fotos = [];
  if (aba && aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues().forEach(l => {
      const n = String(l[0] || '').trim(), u = String(l[1] || '').trim();
      if (n && u) fotos.push({ nome: n, url: u });
    });
  }
  return fotos;
}

// ═══════════════ LER DADOS (chamado pela tela) ═══════════════
function getDados() {
  try {
    const aba = acharAba(abrirVendas(), GID_ABA, ABA_VENDAS);

    const nL = Math.min(aba.getLastRow(), 400);
    const nC = Math.min(aba.getLastColumn(), 80);
    const values = aba.getRange(1, 1, nL, nC).getValues();

    const lido = lerVendas(values);
    const avisos = [];

    const cash = {};
    PARTICIPANTES.forEach(p => { cash[p] = 0; });
    let semDono = 0;

    lido.vendas.forEach(v => {
      const donos = [];
      v.closers.forEach(c => {
        const p = PARTICIPANTES.find(x => mesmoNome(x, c));
        if (p && donos.indexOf(p) < 0) donos.push(p);
      });
      if (!donos.length) { semDono++; return; }
      // dividir pelo número de closers citados na linha (inclusive os que não competem)
      const partes = DIVIDIR_VENDA_COMPARTILHADA ? v.closers.length : 1;
      donos.forEach(p => { cash[p] += v.valor / partes; });
    });

    const fotos = lerFotos();
    const ranking = PARTICIPANTES.map(nome => {
      const f = fotos.find(x => mesmoNome(x.nome, nome));
      return { nome: nome, cash: Math.round(cash[nome] * 100) / 100, foto: f ? f.url : '' };
    });
    ranking.sort((a, b) => b.cash - a.cash || a.nome.localeCompare(b.nome, 'pt-BR'));

    return {
      success: true,
      ranking: ranking.slice(0, TOP),
      avisos:  avisos,
      fonte:   { aba: aba.getName(), vendas: lido.vendas.length, foraDoRanking: semDono }
    };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  }
}

// ═══════════════ AUTORIZAR (rodar 1x antes de publicar) ═══════════════
// Abre a planilha de vendas para o Google pedir a permissão. Não altera nada.
function autorizar() {
  const r = getDados();
  if (!r.success) throw new Error(r.error);
  r.ranking.forEach((p, i) => Logger.log((i + 1) + 'º ' + p.nome + ' — R$ ' + p.cash));
  Logger.log('Vendas assinadas e pagas: ' + r.fonte.vendas + ' (de outros closers: ' + r.fonte.foraDoRanking + ')');
}
