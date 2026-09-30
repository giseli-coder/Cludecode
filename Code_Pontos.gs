/**
 * RANKING SDR — V4 COMPANY
 * Backend Google Apps Script — pontuação por regras da competição
 *
 * REGRAS:
 *   RM                     = +1 ponto
 *   No-show (NS)           = -2 pontos
 *   RA                     = +4 pontos
 *   Recuperação de NS      = +2 pontos
 *
 * ABA "Equipes":
 *   A: Data/Hora (último lançamento) | B: sdr | C: RM | D: NoShow | E: RA | F: RecNS | G: Pontos
 *
 * ABA "Historico" (uma linha por lançamento, nunca sobrescreve):
 *   A: Data/Hora | B: SDR | C: RM | D: NoShow | E: RA | F: RecNS | G: Pontos do lançamento
 *
 * ABA "Imgs":
 *   A: sdr (nome igual ao da aba Equipes) | B: URL IMG
 *
 * RANKING: ordenado pelos PONTOS (calculados aqui no servidor).
 */

const ABA_EQUIPES   = 'Equipes';
const ABA_HISTORICO = 'Historico';
const ABA_IMGS      = 'Imgs';

const COL = { DATA: 1, SDR: 2, RM: 3, NS: 4, RA: 5, REC: 6, PONTOS: 7 };
const NUM_COLS = 7;

const PONTOS = { RM: 1, NS: -2, RA: 4, REC: 2 };

const CABECALHO_EQUIPES   = ['Data/Hora', 'sdr', 'RM', 'No-show', 'RA', 'Rec. NS', 'Pontos'];
const CABECALHO_HISTORICO = ['Data/Hora', 'SDR', 'RM', 'No-show', 'RA', 'Rec. NS', 'Pontos'];

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function calcularPontos(rm, ns, ra, rec) {
  return (rm * PONTOS.RM) + (ns * PONTOS.NS) + (ra * PONTOS.RA) + (rec * PONTOS.REC);
}

function estilizarCabecalho(range) {
  range.setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#E31E24');
}

// Converte qualquer entrada ("2", "2,0", 2.7, null) em inteiro >= 0
function qtd(v) {
  const n = Math.floor(Number(String(v == null ? '' : v).replace(',', '.')));
  return isFinite(n) ? n : 0;
}

// Pega o primeiro campo existente entre os nomes aceitos
function pegar(obj, nomes) {
  for (let i = 0; i < nomes.length; i++) {
    if (obj[nomes[i]] !== undefined && obj[nomes[i]] !== null && obj[nomes[i]] !== '') return obj[nomes[i]];
  }
  return 0;
}

// ═══════════════ SERVIR HTML ═══════════════
// O arquivo HTML do projeto precisa se chamar exatamente "index".
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ranking SDR · V4 Company')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════ LER DADOS ═══════════════
// Ordena por pontos (desc). Desempate: mais RA, depois mais RM.
function getDados() {
  try {
    const ss = getSS();

    const abaEq = ss.getSheetByName(ABA_EQUIPES);
    if (!abaEq) return { success: false, error: 'Aba "' + ABA_EQUIPES + '" não encontrada.' };

    const sdrs = [];
    if (abaEq.getLastRow() > 1) {
      abaEq.getRange(2, 1, abaEq.getLastRow() - 1, NUM_COLS).getValues().forEach(r => {
        const nome = String(r[COL.SDR - 1] || '').trim();
        if (!nome) return;
        const rm  = qtd(r[COL.RM  - 1]);
        const ns  = qtd(r[COL.NS  - 1]);
        const ra  = qtd(r[COL.RA  - 1]);
        const rec = qtd(r[COL.REC - 1]);
        sdrs.push({ sdr: nome, rm: rm, ns: ns, ra: ra, rec: rec, pontos: calcularPontos(rm, ns, ra, rec) });
      });
    }
    sdrs.sort((a, b) => b.pontos - a.pontos || b.ra - a.ra || b.rm - a.rm);

    const imgs = {};
    const abaImgs = ss.getSheetByName(ABA_IMGS);
    if (abaImgs && abaImgs.getLastRow() > 1) {
      abaImgs.getRange(2, 1, abaImgs.getLastRow() - 1, 2).getValues().forEach(r => {
        if (r[0]) imgs[String(r[0]).trim()] = String(r[1] || '').trim();
      });
    }

    return { success: true, sdrs: sdrs, imgs: imgs, regras: PONTOS };
  } catch (e) {
    return { success: false, error: String(e) };
  }
}

// ═══════════════ LANÇAR ═══════════════
// Aceita { sdr, rm, ns, ra, rec } e também apelidos: noshow/no_show, recns/rec_ns/recuperacao.
function lancar(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!dados || typeof dados !== 'object') {
      return { success: false, error: 'Nenhum dado recebido do formulário.' };
    }

    const sdr = String(pegar(dados, ['sdr', 'nome', 'SDR']) || '').trim();
    const rm  = qtd(pegar(dados, ['rm', 'RM']));
    const ns  = qtd(pegar(dados, ['ns', 'NS', 'noshow', 'noShow', 'no_show']));
    const ra  = qtd(pegar(dados, ['ra', 'RA']));
    const rec = qtd(pegar(dados, ['rec', 'REC', 'recns', 'recNS', 'rec_ns', 'recuperacao']));

    if (!sdr) return { success: false, error: 'Selecione o SDR.' };
    if (rm < 0 || ns < 0 || ra < 0 || rec < 0) {
      return { success: false, error: 'As quantidades não podem ser negativas.' };
    }
    if (rm + ns + ra + rec === 0) {
      // mostra o que chegou para facilitar achar divergência de nomes de campo
      return { success: false, error: 'Informe ao menos uma quantidade maior que zero. (recebido: ' + JSON.stringify(dados) + ')' };
    }

    const sheet = getSS().getSheetByName(ABA_EQUIPES);
    if (!sheet) return { success: false, error: 'Aba "' + ABA_EQUIPES + '" não encontrada.' };
    const last = sheet.getLastRow();
    if (last < 2) return { success: false, error: 'Nenhum SDR cadastrado na aba Equipes.' };

    const nomes = sheet.getRange(2, COL.SDR, last - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < nomes.length; i++) {
      if (String(nomes[i][0]).trim() === sdr) { row = i + 2; break; }
    }
    if (row === -1) return { success: false, error: 'SDR não encontrado: ' + sdr };

    const v = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];
    const novoRM  = qtd(v[COL.RM  - 1]) + rm;
    const novoNS  = qtd(v[COL.NS  - 1]) + ns;
    const novoRA  = qtd(v[COL.RA  - 1]) + ra;
    const novoREC = qtd(v[COL.REC - 1]) + rec;

    sheet.getRange(row, COL.RM, 1, 5).setValues([[
      novoRM, novoNS, novoRA, novoREC, calcularPontos(novoRM, novoNS, novoRA, novoREC)
    ]]);

    const cData = sheet.getRange(row, COL.DATA);
    cData.setValue(new Date());
    cData.setNumberFormat('dd/mm/yyyy hh:mm');

    registrarHistorico(sdr, rm, ns, ra, rec);
    SpreadsheetApp.flush();

    return { success: true, pontosLancamento: calcularPontos(rm, ns, ra, rec) };
  } catch (e) {
    return { success: false, error: String(e) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ═══════════════ HISTÓRICO ═══════════════
function registrarHistorico(sdr, rm, ns, ra, rec) {
  const ss = getSS();
  let hist = ss.getSheetByName(ABA_HISTORICO);
  if (!hist) {
    hist = ss.insertSheet(ABA_HISTORICO);
    hist.getRange(1, 1, 1, NUM_COLS).setValues([CABECALHO_HISTORICO]);
    estilizarCabecalho(hist.getRange(1, 1, 1, NUM_COLS));
  }
  hist.appendRow([new Date(), sdr, rm, ns, ra, rec, calcularPontos(rm, ns, ra, rec)]);
  hist.getRange(hist.getLastRow(), 1).setNumberFormat('dd/mm/yyyy hh:mm');
}

// ═══════════════ RECALCULAR PONTOS ═══════════════
// Rode manualmente se mudar as regras (const PONTOS) ou editar quantidades na planilha.
function recalcularPontos() {
  const sheet = getSS().getSheetByName(ABA_EQUIPES);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const rng = sheet.getRange(2, 1, last - 1, NUM_COLS);
  const vals = rng.getValues();
  vals.forEach(r => {
    r[COL.PONTOS - 1] = calcularPontos(qtd(r[COL.RM - 1]), qtd(r[COL.NS - 1]), qtd(r[COL.RA - 1]), qtd(r[COL.REC - 1]));
  });
  rng.setValues(vals);
}

// ═══════════════ SETUP (rodar 1x) ═══════════════
// Cria as abas que faltarem e ajusta o cabeçalho da aba Equipes (linha 1). Não apaga dados.
function setup() {
  const ss = getSS();

  let eq = ss.getSheetByName(ABA_EQUIPES);
  if (!eq) eq = ss.insertSheet(ABA_EQUIPES);
  eq.getRange(1, 1, 1, NUM_COLS).setValues([CABECALHO_EQUIPES]);
  estilizarCabecalho(eq.getRange(1, 1, 1, NUM_COLS));

  if (!ss.getSheetByName(ABA_HISTORICO)) {
    const h = ss.insertSheet(ABA_HISTORICO);
    h.getRange(1, 1, 1, NUM_COLS).setValues([CABECALHO_HISTORICO]);
    estilizarCabecalho(h.getRange(1, 1, 1, NUM_COLS));
  }

  let imgs = ss.getSheetByName(ABA_IMGS);
  if (!imgs) imgs = ss.insertSheet(ABA_IMGS);
  if (imgs.getLastRow() === 0) {
    imgs.getRange(1, 1, 1, 2).setValues([['sdr', 'URL IMG']]);
    estilizarCabecalho(imgs.getRange(1, 1, 1, 2));
  }

  try {
    SpreadsheetApp.getUi().alert('✅ Setup concluído! Cabeçalho da aba Equipes ajustado. Agora faça o Deploy (nova versão).');
  } catch (_) {}
}
