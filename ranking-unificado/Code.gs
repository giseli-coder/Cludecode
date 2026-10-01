/**
 * RANKING SDR — RR · V4 COMPANY
 * UM backend para DOIS rankings (BLACK BOX + BDAY) na mesma tela.
 *
 * Cada RR lançada soma 1 no placar do SDR. O top 3 de cada ranking aparece no pódio.
 *
 * COMO LIGAR AS DUAS PLANILHAS
 *   - Este script fica no projeto da planilha BLACK BOX  → id: ''  (usa a planilha onde o script está)
 *   - Para a planilha BDAY, cole o ID dela abaixo.
 *     O ID é o trecho grande da URL:  docs.google.com/spreadsheets/d/  >>>ESTE_TRECHO<<<  /edit
 *
 * ESTRUTURA DE CADA PLANILHA (a mesma nas duas):
 *   Aba "Equipes":   A: Data/Hora (último lançamento) | B: sdr | C: RR (total)
 *   Aba "Historico": criada sozinha, uma linha por lançamento
 *   Aba "Imgs":      A: sdr (igual ao da Equipes) | B: link da foto no Google Drive
 */

const RANKINGS = [
  { key: 'blackbox', nome: 'BLACK BOX', id: '' },
  { key: 'bday',     nome: 'BDAY',      id: 'COLE_AQUI_O_ID_DA_PLANILHA_BDAY' }
];

const ABA_EQUIPES   = 'Equipes';
const ABA_HISTORICO = 'Historico';
const ABA_IMGS      = 'Imgs';

const COL = { DATA: 1, SDR: 2, RR: 3 };
const NUM_COLS = 3;

// ═══════════════ SERVIR HTML ═══════════════
// O arquivo HTML do projeto precisa se chamar exatamente "index".
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ranking SDR · V4 Company')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════ PLANILHAS ═══════════════
function abrirPlanilha(r) {
  const id = String(r.id || '').trim();
  if (id.indexOf('COLE_AQUI') === 0) {
    throw new Error('Cole o ID da planilha ' + r.nome + ' no topo do Code.gs (const RANKINGS).');
  }
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Este script não está dentro de uma planilha. Informe o ID de ' + r.nome + '.');
  return ss;
}

function acharRanking(key) {
  for (let i = 0; i < RANKINGS.length; i++) {
    if (RANKINGS[i].key === key) return RANKINGS[i];
  }
  return null;
}

// ═══════════════ LER DADOS ═══════════════
// Devolve os dois rankings. Se um falhar, o outro continua funcionando e o erro aparece no painel.
function getDados() {
  const rankings = RANKINGS.map(r => {
    try {
      return lerRanking(r);
    } catch (e) {
      return { key: r.key, nome: r.nome, sdrs: [], imgs: {}, error: String(e.message || e) };
    }
  });
  return { success: true, rankings: rankings };
}

function lerRanking(r) {
  const ss = abrirPlanilha(r);

  const abaEq = ss.getSheetByName(ABA_EQUIPES);
  if (!abaEq) throw new Error('Aba "' + ABA_EQUIPES + '" não encontrada na planilha ' + r.nome + '.');

  const sdrs = [];
  if (abaEq.getLastRow() > 1) {
    abaEq.getRange(2, 1, abaEq.getLastRow() - 1, NUM_COLS).getValues().forEach(linha => {
      const nome = String(linha[COL.SDR - 1] || '').trim();
      if (!nome) return;
      sdrs.push({ sdr: nome, rr: Number(linha[COL.RR - 1]) || 0 });
    });
  }
  sdrs.sort((a, b) => b.rr - a.rr);

  const imgs = {};
  const abaImgs = ss.getSheetByName(ABA_IMGS);
  if (abaImgs && abaImgs.getLastRow() > 1) {
    abaImgs.getRange(2, 1, abaImgs.getLastRow() - 1, 2).getValues().forEach(linha => {
      if (linha[0]) imgs[String(linha[0]).trim()] = String(linha[1] || '').trim();
    });
  }

  return { key: r.key, nome: r.nome, sdrs: sdrs, imgs: imgs };
}

// ═══════════════ LANÇAR ═══════════════
// dados = { ranking: 'blackbox' | 'bday', sdr: 'Nome', rr: 1 }
function lancar(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const r = acharRanking(dados && dados.ranking);
    if (!r) return { success: false, error: 'Ranking inválido.' };

    const sdr = String((dados && dados.sdr) || '').trim();
    const rr  = Math.floor(Number(dados && dados.rr) || 0);
    if (!sdr)    return { success: false, error: 'Selecione o SDR.' };
    if (rr <= 0) return { success: false, error: 'Informe uma quantidade de RR maior que zero.' };

    const ss    = abrirPlanilha(r);
    const sheet = ss.getSheetByName(ABA_EQUIPES);
    if (!sheet) return { success: false, error: 'Aba "' + ABA_EQUIPES + '" não encontrada em ' + r.nome + '.' };

    const last = sheet.getLastRow();
    if (last < 2) return { success: false, error: 'Nenhum SDR cadastrado em ' + r.nome + '.' };

    const nomes = sheet.getRange(2, COL.SDR, last - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < nomes.length; i++) {
      if (String(nomes[i][0]).trim() === sdr) { row = i + 2; break; }
    }
    if (row === -1) return { success: false, error: 'SDR não encontrado em ' + r.nome + ': ' + sdr };

    const atual = Number(sheet.getRange(row, COL.RR).getValue()) || 0;
    sheet.getRange(row, COL.RR).setValue(atual + rr);

    const cData = sheet.getRange(row, COL.DATA);
    cData.setValue(new Date());
    cData.setNumberFormat('dd/mm/yyyy hh:mm');

    registrarHistorico(ss, sdr, rr);
    SpreadsheetApp.flush();

    return { success: true, total: atual + rr };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ═══════════════ HISTÓRICO ═══════════════
function registrarHistorico(ss, sdr, rr) {
  let hist = ss.getSheetByName(ABA_HISTORICO);
  if (!hist) {
    hist = ss.insertSheet(ABA_HISTORICO);
    hist.getRange(1, 1, 1, 3).setValues([['Data/Hora', 'SDR', 'RR']])
      .setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#E31E24');
  }
  hist.appendRow([new Date(), sdr, rr]);
  hist.getRange(hist.getLastRow(), 1).setNumberFormat('dd/mm/yyyy hh:mm');
}

// ═══════════════ AUTORIZAR (rodar 1x antes de publicar) ═══════════════
// Abre as duas planilhas para o Google pedir a permissão de acesso. Não altera nada.
function autorizar() {
  RANKINGS.forEach(r => {
    const ss = abrirPlanilha(r);
    const rk = lerRanking(r);
    Logger.log(r.nome + ' → planilha "' + ss.getName() + '" · ' + rk.sdrs.length + ' SDRs');
  });
}

// ═══════════════ SETUP (opcional) ═══════════════
// Cria as abas que faltarem nas duas planilhas. Não apaga nem sobrescreve nada.
function setup() {
  RANKINGS.forEach(r => {
    const ss = abrirPlanilha(r);

    let eq = ss.getSheetByName(ABA_EQUIPES);
    if (!eq) eq = ss.insertSheet(ABA_EQUIPES);
    if (eq.getLastRow() === 0) {
      eq.getRange(1, 1, 1, NUM_COLS).setValues([['Data/Hora', 'sdr', 'RR']])
        .setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#E31E24');
    }

    if (!ss.getSheetByName(ABA_HISTORICO)) {
      const h = ss.insertSheet(ABA_HISTORICO);
      h.getRange(1, 1, 1, 3).setValues([['Data/Hora', 'SDR', 'RR']])
        .setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#E31E24');
    }

    let imgs = ss.getSheetByName(ABA_IMGS);
    if (!imgs) imgs = ss.insertSheet(ABA_IMGS);
    if (imgs.getLastRow() === 0) {
      imgs.getRange(1, 1, 1, 2).setValues([['sdr', 'URL IMG']])
        .setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#E31E24');
    }
  });
}
