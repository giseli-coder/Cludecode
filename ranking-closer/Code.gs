/**
 * RANKING CLOSER — STAAGE
 * Backend Google Apps Script
 *
 * REGRA:  quem fizer MAIS CASH COLLECTED fica em 1º. Top 4.
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

// Se a aba tiver mais de uma coluna com "cash" no título e o ranking pegar a errada,
// escreva aqui o título EXATO da coluna certa. Vazio = o código escolhe sozinho.
const COLUNA_CASH = '';

const ABA_IMGS = 'Imgs';
const TOP      = 4;

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

// ═══════════════ LER A TABELA DE CASH COLLECTED ═══════════════
// Em vez de endereços fixos, procuro o título da coluna ("Cash Collected"),
// o título "Closer" perto dele e leio os nomes logo abaixo.
function acharColunaCash(values) {
  const alvoFixo = norm(COLUNA_CASH);
  let melhor = null;
  const maxL = Math.min(values.length, 80);
  for (let r = 0; r < maxL; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const t = norm(values[r][c]);
      if (!t) continue;
      if (alvoFixo) {
        if (t === alvoFixo) return { r: r, c: c, titulo: String(values[r][c]).trim() };
        continue;
      }
      if (t.indexOf('CASH') < 0) continue;
      if (/META|DELTA|FALTA|TX |CONV|%|PROJET|PREVIS/.test(t)) continue;
      const score = /COLLECT/.test(t) ? 2 : 1;
      if (!melhor || score > melhor.score) {
        melhor = { r: r, c: c, titulo: String(values[r][c]).trim(), score: score };
      }
    }
  }
  return melhor;
}

function rotulosDaAba(values) {
  const out = [];
  for (let r = 0; r < Math.min(values.length, 12); r++) {
    values[r].forEach(v => {
      const t = String(v == null ? '' : v).trim();
      if (t && isNaN(Number(t)) && out.length < 40 && out.indexOf(t) < 0) out.push(t);
    });
  }
  return out;
}

const NOMES_DE_COLUNA = ['CLOSER', 'NOME', 'VENDEDOR', 'RESPONSAVEL'];

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

function lerTabela(values, avisos) {
  const ach = acharColunaCash(values);
  if (!ach) {
    throw new Error('Não encontrei a coluna "Cash Collected" na aba. Títulos que vi no topo: ' +
      rotulosDaAba(values).join(' | '));
  }

  // coluna "Closer": procura o título perto do "Cash" (pode estar 1-2 linhas acima/abaixo)
  let linhaTit = ach.r, colNome = -1, dist = 99;
  for (let r = Math.max(0, ach.r - 2); r <= Math.min(values.length - 1, ach.r + 2); r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (c !== ach.c && NOMES_DE_COLUNA.indexOf(norm(values[r][c])) >= 0 && Math.abs(r - ach.r) < dist) {
        dist = Math.abs(r - ach.r); colNome = c; linhaTit = Math.max(ach.r, r);
      }
    }
  }

  // sem título "Closer": a coluna de texto mais à esquerda do cash
  if (colNome < 0) {
    for (let c = 0; c < ach.c && colNome < 0; c++) {
      for (let r = linhaTit + 1; r < Math.min(values.length, linhaTit + 6); r++) {
        const t = String(values[r][c] == null ? '' : values[r][c]).trim();
        if (t && !/^[\d.,R$\s%-]+$/.test(t)) { colNome = c; break; }
      }
    }
  }
  if (colNome < 0) {
    throw new Error('Achei a coluna "' + ach.titulo + '" mas não achei a coluna com os nomes dos closers (título "Closer").');
  }

  const lista = [];
  let vazias = 0;
  for (let r = linhaTit + 1; r < values.length; r++) {
    const nome = String(values[r][colNome] == null ? '' : values[r][colNome]).trim();
    if (!nome) { if (lista.length && ++vazias >= 2) break; continue; }
    vazias = 0;
    const n = norm(nome);
    if (/^(SQUAD|TOTAL|MEDIA|CLOSER|SOMA|TKM)\b/.test(n)) {
      if (/^TOTAL\b/.test(n) && lista.length) break;
      continue;
    }
    if (lista.some(x => norm(x.nome) === n)) {
      avisos.push('Nome repetido na tabela: "' + nome + '" (usei a primeira linha).');
      continue;
    }
    lista.push({ nome: nome, cash: dinheiro(values[r][ach.c]) });
  }
  return { lista: lista, titulo: ach.titulo, linha: ach.r + 1 };
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
    let tab = lerBlocosCash(values);                       // formato de quadros (um por closer)
    if (!tab.lista.length) tab = lerTabela(values, avisos); // formato de tabela com coluna "Cash Collected"
    if (!tab.lista.length) {
      throw new Error('Achei a coluna "' + tab.titulo + '", mas não há nomes de closers abaixo dela.');
    }

    const vistos = [];
    tab.lista = tab.lista.filter(p => {
      if (vistos.some(n => n === norm(p.nome))) { avisos.push('Nome repetido: "' + p.nome + '" (usei o primeiro quadro).'); return false; }
      vistos.push(norm(p.nome)); return true;
    });

    const fotos = lerFotos();
    const ranking = tab.lista.map(p => {
      const f = fotos.find(x => mesmoNome(x.nome, p.nome));
      return { nome: p.nome, cash: p.cash, foto: f ? f.url : '' };
    });

    ranking.sort((a, b) => b.cash - a.cash || a.nome.localeCompare(b.nome, 'pt-BR'));

    const total = ranking.reduce((s, p) => s + Math.max(0, p.cash), 0);
    const lider = ranking[0].cash;
    ranking.forEach(p => {
      p.parte = total > 0 ? Math.round(Math.max(0, p.cash) / total * 100) : 0;   // % do cash total
      p.falta = Math.max(0, lider - p.cash);                                      // distância até o 1º
    });

    return {
      success: true,
      ranking: ranking.slice(0, TOP),
      total:   total,
      qtd:     ranking.length,
      avisos:  avisos,
      fonte:   { aba: aba.getName(), coluna: tab.titulo, linha: tab.linha }
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
