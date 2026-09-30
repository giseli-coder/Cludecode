/**
 * RANKING SDR — RR · V4 COMPANY
 * Backend Google Apps Script
 *
 * Cada RR lançada soma 1 no placar do SDR.
 *
 * ABA "Equipes":
 *   A: Data/Hora (último lançamento) | B: sdr | C: RR (total acumulado)
 *
 * ABA "Historico" (criada automaticamente, uma linha por lançamento):
 *   A: Data/Hora | B: SDR | C: Quantidade de RR lançada
 *
 * ABA "Imgs":
 *   A: sdr (nome igual ao da aba Equipes) | B: URL da imagem (Google Drive)
 */

const ABA_EQUIPES   = 'Equipes';
const ABA_HISTORICO = 'Historico';
const ABA_IMGS      = 'Imgs';

const COL = { DATA: 1, SDR: 2, RR: 3 };
const NUM_COLS = 3;

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ═══════════════ SERVIR HTML ═══════════════
// O arquivo HTML do projeto precisa se chamar exatamente "index".
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ranking SDR · V4 Company')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════ LER DADOS ═══════════════
function getDados() {
  try {
    const ss = getSS();

    const abaEq = ss.getSheetByName(ABA_EQUIPES);
    if (!abaEq) return { success: false, error: 'Aba "' + ABA_EQUIPES + '" não encontrada.' };

    const sdrs = [];
    if (abaEq.getLastRow() > 1) {
      const dados = abaEq.getRange(2, 1, abaEq.getLastRow() - 1, NUM_COLS).getValues();
      dados.forEach(r => {
        const nome = String(r[COL.SDR - 1] || '').trim();
        if (!nome) return;
        sdrs.push({ sdr: nome, rr: Number(r[COL.RR - 1]) || 0 });
      });
    }
    sdrs.sort((a, b) => b.rr - a.rr);

    const imgs = {};
    const abaImgs = ss.getSheetByName(ABA_IMGS);
    if (abaImgs && abaImgs.getLastRow() > 1) {
      abaImgs.getRange(2, 1, abaImgs.getLastRow() - 1, 2).getValues().forEach(r => {
        if (r[0]) imgs[String(r[0]).trim()] = String(r[1] || '').trim();
      });
    }

    return { success: true, sdrs: sdrs, imgs: imgs };
  } catch (e) {
    return { success: false, error: String(e) };
  }
}

// ═══════════════ LANÇAR ═══════════════
// dados = { sdr: 'Nome', rr: 1 }
function lancar(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const sheet = getSS().getSheetByName(ABA_EQUIPES);
    if (!sheet) return { success: false, error: 'Aba "' + ABA_EQUIPES + '" não encontrada.' };

    const sdr = String((dados && dados.sdr) || '').trim();
    const rr  = Math.floor(Number(dados && dados.rr) || 0);

    if (!sdr)    return { success: false, error: 'Selecione o SDR.' };
    if (rr <= 0) return { success: false, error: 'Informe uma quantidade de RR maior que zero.' };

    const last = sheet.getLastRow();
    if (last < 2) return { success: false, error: 'Nenhum SDR cadastrado na aba Equipes.' };

    const nomes = sheet.getRange(2, COL.SDR, last - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < nomes.length; i++) {
      if (String(nomes[i][0]).trim() === sdr) { row = i + 2; break; }
    }
    if (row === -1) return { success: false, error: 'SDR não encontrado: ' + sdr };

    const atual = Number(sheet.getRange(row, COL.RR).getValue()) || 0;
    sheet.getRange(row, COL.RR).setValue(atual + rr);

    const cData = sheet.getRange(row, COL.DATA);
    cData.setValue(new Date());
    cData.setNumberFormat('dd/mm/yyyy hh:mm');

    registrarHistorico(sdr, rr);
    SpreadsheetApp.flush();

    return { success: true, total: atual + rr };
  } catch (e) {
    return { success: false, error: String(e) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ═══════════════ HISTÓRICO ═══════════════
function registrarHistorico(sdr, rr) {
  const ss = getSS();
  let hist = ss.getSheetByName(ABA_HISTORICO);
  if (!hist) {
    hist = ss.insertSheet(ABA_HISTORICO);
    hist.getRange(1, 1, 1, 3).setValues([['Data/Hora', 'SDR', 'RR']])
      .setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#E31E24');
  }
  hist.appendRow([new Date(), sdr, rr]);
  hist.getRange(hist.getLastRow(), 1).setNumberFormat('dd/mm/yyyy hh:mm');
}

// ═══════════════ SETUP (opcional, rodar 1x) ═══════════════
// Cria as abas que faltarem. Não apaga nada.
function setup() {
  const ss = getSS();

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
}
