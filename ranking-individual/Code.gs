/**
 * RANKING INDIVIDUAL SDR — STAAGE
 * Backend Google Apps Script
 *
 * PONTUAÇÃO:   RA (Reunião Acontecida) = 1 ponto
 * DESEMPATE:   ordem alfabética
 *
 * DE ONDE VÊM OS NÚMEROS
 *   Automático, da planilha "[Pré-vendas] - Métricas Individuais": para cada SDR somo a coluna
 *   "RA (Total)" de todos os grupos (FR, AS, M&GA) em todos os dias do mês. Ninguém lança nada à mão.
 *
 * ESTA PLANILHA (a do ranking) só guarda:
 *   Aba "Equipes": coluna B = nomes de quem COMPETE  (se ficar vazia, entram todos os SDRs da planilha)
 *   Aba "Imgs":    A = nome | B = link da foto no Google Drive
 *
 * O nome na Equipes não precisa ser igual ao da métricas: "Matheus" encontra "Matheus Lee"
 * (sem diferenciar maiúsculas/acentos).
 */

// ═══════════════ CONFIGURAÇÃO (é só aqui que você mexe) ═══════════════
const METRICAS_ID = '15c_Xf1K9_1ZPZkcDst3-FyfTqgcpETwVuQ1eE79dMaU';  // trecho da URL entre /d/ e /edit
const GID_ABA     = 491349697;           // número que aparece depois de "gid=" na URL da aba do mês
const ABA_METRICAS = 'Outubro26';        // só é usado se o GID não for encontrado (troque quando virar o mês)

const PTS_RA = 1;

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

function abrirMetricas() {
  const id = String(METRICAS_ID || '').trim();
  if (!id) throw new Error('Cole o ID da planilha de métricas no topo do Code.gs (const METRICAS_ID).');
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
  throw new Error('Não achei a aba (gid ' + gid + ' / nome "' + nome + '") na planilha de métricas. Abas que existem: ' +
    abas.map(a => a.getName()).join(' | '));
}

// ═══════════════ LER AS COLUNAS DE CADA SDR ═══════════════
// A planilha é assim (uma linha por dia):
//   linha dos NOMES   :  [ Pedro .......................................... ]  [ Jean ........ ]
//   linha dos GRUPOS  :  [ FR ........ ][ AS ........ ][ M&GA ........ ]
//   linha dos TÍTULOS :  RM | RA (Isis) | RA (Fink) | RA (Claud) | RA (Total) | NS | REM | ...
//   linhas dos DIAS   :  01/10/2026 | Quinta | ...
// Procuro a linha de títulos (a que tem "RA (Total)"), descubro de quem é cada coluna
// (o nome está duas linhas acima, na célula mesclada à esquerda) e somo os dias.
const RE_RA_TOTAL = /^RA\s*\(?\s*TOTAL\s*\)?$/;

function ehData(v) {
  return Object.prototype.toString.call(v) === '[object Date]' ||
    /^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(v == null ? '' : v).trim());
}

function textoCelula(v) {
  return String(v == null ? '' : v).trim();
}

// Dono da coluna c: primeiro nome encontrado à esquerda, parando na coluna separadora (título vazio)
function donoDaColuna(values, hr, c) {
  const ordem = [2, 3, 1];
  for (let o = 0; o < ordem.length; o++) {
    const r = hr - ordem[o];
    if (r < 0) continue;
    for (let j = c; j >= 0; j--) {
      if (!norm(values[hr][j])) break;
      const t = textoCelula(values[r][j]);
      if (t) return t;
    }
  }
  return '';
}

function lerBlocos(values, avisos) {
  // linha de títulos
  let hr = -1;
  for (let r = 0; r < Math.min(values.length, 15) && hr < 0; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (RE_RA_TOTAL.test(norm(values[r][c]))) { hr = r; break; }
    }
  }
  if (hr < 0) return { blocos: [], periodo: '' };

  // linhas dos dias (coluna A é uma data)
  const dias = [];
  for (let r = hr + 1; r < values.length; r++) {
    if (ehData(values[r][0])) dias.push(r);
  }
  if (!dias.length) avisos.push('Não encontrei linhas de data (coluna A) abaixo dos títulos.');

  const pessoas = [];
  const achar = nome => {
    let p = pessoas.find(x => norm(x.nome) === norm(nome));
    if (!p) { p = { nome: nome, ra: 0, rm: 0, ns: 0 }; pessoas.push(p); }
    return p;
  };

  for (let c = 0; c < values[hr].length; c++) {
    const titulo = norm(values[hr][c]);
    const campo = RE_RA_TOTAL.test(titulo) ? 'ra' : titulo === 'RM' ? 'rm' : titulo === 'NS' ? 'ns' : '';
    if (!campo) continue;
    const dono = donoDaColuna(values, hr, c);
    if (!dono) continue;
    const p = achar(dono);
    dias.forEach(r => { p[campo] += num(values[r][c]); });
  }

  let periodo = '';
  if (dias.length) {
    const fmt = v => Object.prototype.toString.call(v) === '[object Date]'
      ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : textoCelula(v).slice(0, 10);
    periodo = fmt(values[dias[0]][0]) + ' a ' + fmt(values[dias[dias.length - 1]][0]);
  }
  return { blocos: pessoas, periodo: periodo };
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
    const aba    = acharAba(abrirMetricas(), GID_ABA, ABA_METRICAS);

    const nL = Math.min(aba.getLastRow(), 120);
    const nC = Math.min(aba.getLastColumn(), 250);
    const values = aba.getRange(1, 1, nL, nC).getValues();

    const avisos = [];
    const lido   = lerBlocos(values, avisos);
    const blocos = lido.blocos;
    if (!blocos.length) {
      throw new Error('Não encontrei as colunas "RA (Total)" na aba "' + aba.getName() +
        '". Confira se o GID / nome da aba está certo.');
    }

    const participantes = lerParticipantes(ssRank);
    const fotos  = lerFotos(ssRank);
    const lista  = participantes.length ? participantes
                 : blocos.map(b => b.nome).filter(n => norm(n) !== 'GERAL');
    const nomesPlanilha = blocos.map(b => b.nome).join(', ');

    const ranking = lista.map(nome => {
      const b = blocos.find(x => mesmoNome(x.nome, nome));
      if (!b) {
        avisos.push('Não achei "' + nome + '" na planilha de métricas (nomes lá: ' + nomesPlanilha + ').');
      }
      const f = fotos.find(x => mesmoNome(x.nome, nome) || (b && mesmoNome(x.nome, b.nome)));
      const ra = b ? b.ra : 0;
      return {
        nome:    nome,
        ra:      ra,
        rm:      b ? b.rm : 0,
        ns:      b ? b.ns : 0,
        pontos:  ra * PTS_RA,
        foto:    f ? f.url : '',
        semDados: !b
      };
    });

    ranking.sort((a, b) => b.pontos - a.pontos || a.nome.localeCompare(b.nome, 'pt-BR'));

    return {
      success: true,
      ranking: ranking,
      avisos:  avisos,
      fonte:   { aba: aba.getName(), periodo: lido.periodo },
      regras:  { ra: PTS_RA }
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
  r.ranking.forEach((p, i) => Logger.log((i + 1) + 'º ' + p.nome + ' — ' + p.pontos + ' pts (RA ' + p.ra + ', RM ' + p.rm + ', NS ' + p.ns + ')'));
  r.avisos.forEach(a => Logger.log('AVISO: ' + a));
}
