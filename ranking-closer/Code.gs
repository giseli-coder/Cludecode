/**
 * RANKING CLOSER — STAAGE
 * Backend Google Apps Script
 *
 * REGRA:  quem fizer MAIS CASH COLLECTED fica em 1º. Só os PARTICIPANTES competem.
 * DESEMPATE: ordem alfabética.
 *
 * DE ONDE VÊM OS NÚMEROS
 *   Automático, da planilha de METAS, na aba definida em ABA_METAS
 *   (coluna "Cash Collected" de cada closer). Ninguém lança nada à mão.
 *
 * ESTA PLANILHA (a do ranking) só guarda:
 *   Aba "Imgs": A = nome do closer | B = link da foto no Google Drive
 *   (o nome não precisa ser igual: "Rodrigo" encontra "Rodrigo Monteiro")
 */

// ═══════════════ CONFIGURAÇÃO (é só aqui que você mexe) ═══════════════
const METAS_ID  = '17WLibwAUJb9Z6T-z1PPH8JU1h6umR-rw3efBV73MDeA';  // trecho da URL entre /d/ e /edit
const ABA_METAS = 'CONVERSÃO CLOSER - OUTUBRO  26';                // troque quando virar o mês

// Quem participa da competição (o nome pode ser parcial e sem acento: "Junior" encontra "JUNIOR").
// Para entrar ou sair alguém, é só mexer nesta lista.
const PARTICIPANTES = ['Junior', 'Luiz', 'Fernando', 'Tayla'];

const ABA_IMGS = 'Imgs';

// ═══════════════ SERVIR HTML ═══════════════
// O arquivo HTML do projeto precisa se chamar exatamente "index".
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ranking Closer — STAAGE')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════ UTILITÁRIOS ═══════════════
// Maiúsculas, sem acento e sem espaços repetidos: "  Conversão " -> "CONVERSAO"
function norm(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

// "Rodrigo" ~ "RODRIGO MONTEIRO"  ·  "Leo Vianna" ~ "LEO VIANNA"
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
  const parecidas = abas.map(a => a.getName()).filter(n => norm(n).indexOf('CONVERSAO CLOSER') === 0);
  throw new Error('Aba "' + nome + '" não encontrada na planilha de metas.' +
    (parecidas.length ? ' Abas parecidas: ' + parecidas.join(' | ') : ''));
}

// ═══════════════ LER OS QUADROS DE CADA CLOSER ═══════════════
// Cada closer é um quadro assim:
//        [  NOME DO CLOSER  ]                <- nome (célula mesclada, vale a primeira)
//        | Projetado | Provisionado | Realizado
//   RM / RA / V / ...
//   Cash Collected   | R$ ...      | R$ ...     | R$ ...   <- uso a coluna "Realizado"
// Procuro o rótulo "Cash Collected", subo para achar o cabeçalho "Realizado" e o nome.
function lerBlocosCash(values) {
  const lista = [];
  let linha = 0;
  for (let r = 1; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (norm(values[r][c]) !== 'CASH COLLECTED') continue;

      // cabeçalho "Realizado" acima (até 14 linhas), à direita do rótulo
      let hr = -1, colReal = -1;
      for (let k = r - 1; k >= 0 && k >= r - 14 && hr < 0; k--) {
        for (let j = c + 1; j <= c + 5 && j < values[k].length; j++) {
          if (norm(values[k][j]) === 'REALIZADO') { hr = k; colReal = j; break; }
        }
      }
      if (hr < 0) continue;

      // nome: primeira célula preenchida acima do cabeçalho, na mesma coluna do rótulo
      let nome = '';
      for (let k = hr - 1; k >= 0 && k >= hr - 3; k--) {
        const t = String(values[k][c] == null ? '' : values[k][c]).trim();
        if (t && !/^(PROJETADO|PROVISIONADO|REALIZADO)$/.test(norm(t))) { nome = t; break; }
      }
      if (!nome) continue;

      if (!linha) linha = r + 1;
      lista.push({ nome: nome, cash: dinheiro(values[r][colReal]) });
    }
  }
  return { lista: lista, titulo: 'Cash Collected (Realizado)', linha: linha };
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
    const metas = abrirMetas();
    const aba   = acharAba(metas, ABA_METAS);

    const nL = Math.min(aba.getLastRow(), 200);
    const nC = Math.min(aba.getLastColumn(), 40);
    const values = aba.getRange(1, 1, nL, nC).getValues();

    const avisos = [];
    const tab = lerBlocosCash(values);
    if (!tab.lista.length) {
      throw new Error('Não encontrei os quadros dos closers na aba "' + aba.getName() +
        '". Cada quadro precisa ter o nome no topo, o cabeçalho "Realizado" e a linha "Cash Collected".');
    }

    const fotos = lerFotos();
    const ranking = PARTICIPANTES.map(nome => {
      const b = tab.lista.find(x => mesmoNome(x.nome, nome));
      if (!b) {
        avisos.push('Não achei "' + nome + '" na aba de metas (nomes lá: ' + tab.lista.map(x => x.nome).join(', ') + ').');
      }
      const f = fotos.find(x => mesmoNome(x.nome, nome));
      return { nome: nome, cash: b ? b.cash : 0, foto: f ? f.url : '' };
    });

    ranking.sort((a, b) => b.cash - a.cash || a.nome.localeCompare(b.nome, 'pt-BR'));

    return {
      success: true,
      ranking: ranking,
      avisos:  avisos,
      fonte:   { aba: aba.getName() }
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
  r.ranking.forEach((p, i) => Logger.log((i + 1) + 'º ' + p.nome + ' — R$ ' + p.cash));
  r.avisos.forEach(a => Logger.log('AVISO: ' + a));
}
