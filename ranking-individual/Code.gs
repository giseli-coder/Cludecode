/**
 * RANKING INDIVIDUAL SDR — STAAGE
 * Backend Google Apps Script
 *
 * PONTUAÇÃO:   RA = 1 ponto   ·   Venda (V) = 2 pontos
 * DESEMPATE:   mais vendas → mais RA → ordem alfabética
 *
 * DE ONDE VÊM OS NÚMEROS
 *   Automático, da planilha de METAS (coluna "Realizado" de cada SDR, linhas "RA" e "V"),
 *   na aba definida em ABA_METAS. Ninguém precisa lançar nada à mão.
 *
 * ESTA PLANILHA (a do ranking) só guarda:
 *   Aba "Equipes": coluna B = nomes de quem COMPETE  (se ficar vazia, entram todos os SDRs da metas)
 *   Aba "Imgs":    A = nome | B = link da foto no Google Drive
 *
 * O nome na Equipes não precisa ser igual ao da metas: "Tayla" encontra "TAYLA THE CREATOR",
 * "Matheus Lee" encontra "MATHEUS" (sem diferenciar maiúsculas/acentos).
 */

// ═══════════════ CONFIGURAÇÃO (é só aqui que você mexe) ═══════════════
const METAS_ID  = '17WLibwAUJb9Z6T-z1PPH8JU1h6umR-rw3efBV73MDeA';  // trecho da URL entre /d/ e /edit
const ABA_METAS = 'MENSAL SDR - OUTUBRO 26';              // troque quando virar o mês

const PTS_RA    = 1;
const PTS_VENDA = 2;

const ABA_EQUIPES = 'Equipes';
const ABA_IMGS    = 'Imgs';

// ═══════════════ SERVIR HTML ═══════════════
// O arquivo HTML do projeto precisa se chamar exatamente "index".
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ranking Individual SDR — STAAGE')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════ UTILITÁRIOS ═══════════════
// Maiúsculas, sem acento e sem espaços repetidos: "  Tayla  " -> "TAYLA"
function norm(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

// "Tayla" ~ "TAYLA THE CREATOR"  ·  "Matheus Lee" ~ "MATHEUS"  ·  "Pedro" ~ "PEDRO"
function mesmoNome(a, b) {
  const x = norm(a), y = norm(b);
  if (!x || !y) return false;
  return x === y || x.indexOf(y + ' ') === 0 || y.indexOf(x + ' ') === 0;
}

function num(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v == null ? '' : v).trim();
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');   // "1.234,5" -> 1234.5
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function abrirMetas() {
  const id = String(METAS_ID || '').trim();
  if (!id || id.indexOf('COLE_AQUI') === 0) {
    throw new Error('Cole o ID da planilha de METAS no topo do Code.gs (const METAS_ID).');
  }
  return SpreadsheetApp.openById(id);
}

// Acha a aba ignorando maiúsculas, acentos e espaços duplos.
function acharAba(ss, nome) {
  const abas = ss.getSheets();
  const alvo = norm(nome);
  for (let i = 0; i < abas.length; i++) {
    if (norm(abas[i].getName()) === alvo) return abas[i];
  }
  const parecidas = abas.map(a => a.getName()).filter(n => norm(n).indexOf('MENSAL SDR') === 0);
  throw new Error('Aba "' + nome + '" não encontrada na planilha de metas.' +
    (parecidas.length ? ' Abas parecidas: ' + parecidas.join(' | ') : ''));
}

// ═══════════════ LER OS BLOCOS DE CADA SDR NA ABA DE METAS ═══════════════
// Cada SDR é um quadro assim:
//        [  NOME DO SDR  ]                 <- nome (célula mesclada, vale a primeira)
//        | Projetado | Provisionado | Realizado
//   MQL
//   Conectados
//   ...
//   RA
//   ...
//   V
// Em vez de depender de endereços fixos, procuro o rótulo "MQL" com "Conectados" logo abaixo
// (assinatura exclusiva desses quadros), subo para achar o nome e desço para achar "RA" e "V".
function lerBlocos(values, avisos) {
  const blocos = [];

  for (let r = 1; r < values.length - 1; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (norm(values[r][c]) !== 'MQL' || norm(values[r + 1][c]) !== 'CONECTADOS') continue;

      // nome: primeira célula preenchida acima, na mesma coluna (pula a linha de cabeçalho)
      let nome = '';
      for (let k = r - 1; k >= 0 && k >= r - 4; k--) {
        const t = String(values[k][c] || '').trim();
        if (t && !/^(PROJETADO|PROVISIONADO|REALIZADO)$/.test(norm(t))) { nome = t; break; }
      }
      if (!nome) continue;

      // coluna "Realizado": procura no cabeçalho; se não achar, é a 3ª depois do rótulo
      let colReal = c + 3;
      for (let j = c + 1; j <= c + 4 && j < values[r - 1].length; j++) {
        if (norm(values[r - 1][j]) === 'REALIZADO') { colReal = j; break; }
      }

      let ra = null, v = null;
      for (let k = r; k < Math.min(r + 16, values.length); k++) {
        const rot = norm(values[k][c]);
        if (rot === 'RA' && ra === null) ra = num(values[k][colReal]);
        if (rot === 'V'  && v  === null) v  = num(values[k][colReal]);
      }
      if (ra === null) avisos.push('Não achei a linha "RA" no quadro de ' + nome + '.');
      if (v  === null) avisos.push('Não achei a linha "V" no quadro de ' + nome + '.');

      blocos.push({ nome: nome, ra: ra || 0, v: v || 0 });
    }
  }
  return blocos;
}

// Período (DATA INICIAL / DATA FINAL) — opcional, só para mostrar no rodapé
function lerPeriodo(values, tz) {
  let ini = '', fim = '';
  const fmt = x => Object.prototype.toString.call(x) === '[object Date]'
    ? Utilities.formatDate(x, tz, 'dd/MM/yyyy') : String(x || '').trim();
  for (let r = 0; r < Math.min(values.length - 1, 12); r++) {
    for (let c = 0; c < values[r].length; c++) {
      const t = norm(values[r][c]);
      if (t === 'DATA INICIAL') ini = fmt(values[r + 1][c]);
      if (t === 'DATA FINAL')   fim = fmt(values[r + 1][c]);
    }
  }
  return (ini && fim) ? ini + ' a ' + fim : '';
}

// ═══════════════ PLANILHA DO RANKING (Equipes + Imgs) ═══════════════
function lerParticipantes(ss) {
  const aba = ss.getSheetByName(ABA_EQUIPES);
  const nomes = [];
  if (aba && aba.getLastRow() > 1) {
    aba.getRange(2, 2, aba.getLastRow() - 1, 1).getValues().forEach(l => {
      const n = String(l[0] || '').trim();
      if (n && !nomes.some(x => norm(x) === norm(n))) nomes.push(n);
    });
  }
  return nomes;
}

function lerFotos(ss) {
  const aba = ss.getSheetByName(ABA_IMGS);
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
    const ssRank = SpreadsheetApp.getActiveSpreadsheet();
    const metas  = abrirMetas();
    const aba    = acharAba(metas, ABA_METAS);

    const nL = Math.min(aba.getLastRow(), 150);
    const nC = Math.min(aba.getLastColumn(), 40);
    const values = aba.getRange(1, 1, nL, nC).getValues();

    const avisos = [];
    const blocos = lerBlocos(values, avisos);
    if (!blocos.length) {
      throw new Error('Não encontrei os quadros dos SDRs na aba "' + aba.getName() +
        '". Confira se é a aba certa (cada quadro tem as linhas MQL, Conectados, ... RA, ... V).');
    }

    const participantes = lerParticipantes(ssRank);
    const fotos  = lerFotos(ssRank);
    const lista  = participantes.length ? participantes : blocos.map(b => b.nome);
    const nomesMetas = blocos.map(b => b.nome).join(', ');

    const ranking = lista.map(nome => {
      const b = blocos.find(x => mesmoNome(x.nome, nome));
      if (!b) {
        avisos.push('Não achei "' + nome + '" na aba de metas (nomes lá: ' + nomesMetas + ').');
      }
      const f = fotos.find(x => mesmoNome(x.nome, nome) || (b && mesmoNome(x.nome, b.nome)));
      const ra = b ? b.ra : 0, v = b ? b.v : 0;
      return {
        nome:    nome,
        ra:      ra,
        v:       v,
        pontos:  ra * PTS_RA + v * PTS_VENDA,
        foto:    f ? f.url : '',
        semDados: !b
      };
    });

    ranking.sort((a, b) =>
      b.pontos - a.pontos || b.v - a.v || b.ra - a.ra || a.nome.localeCompare(b.nome, 'pt-BR'));

    return {
      success: true,
      ranking: ranking,
      avisos:  avisos,
      fonte:   { aba: aba.getName(), periodo: lerPeriodo(values, metas.getSpreadsheetTimeZone()) },
      regras:  { ra: PTS_RA, venda: PTS_VENDA }
    };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  }
}

// ═══════════════ AUTORIZAR (rodar 1x antes de publicar) ═══════════════
// Abre a planilha de metas para o Google pedir a permissão. Não altera nada.
function autorizar() {
  const r = getDados();
  if (!r.success) throw new Error(r.error);
  r.ranking.forEach((p, i) => Logger.log((i + 1) + 'º ' + p.nome + ' — ' + p.pontos + ' pts (RA ' + p.ra + ', V ' + p.v + ')'));
  r.avisos.forEach(a => Logger.log('AVISO: ' + a));
}
