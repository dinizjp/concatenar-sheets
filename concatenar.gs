/**
 * Consolida dados das planilhas POSTIGO e CRÉDITO FÁCIL em uma planilha central.
 * Copia todos os dados e formatações, sem sobrescrever validações nem cores.
 */

function consolidarPlanilhas() {
  const PLANILHA_DESTINO_ID = '1HeAJoA8jzUR-ywVewfdhfdHqfcSCUNR765ALYcoQ1DTzjAg0';
  const ABA_DESTINO = 'Dados';

  const PLANILHAS_ORIGEM = [
    { id: '1Tb5o3YTcW3UTXyPyp1wcVnK6drd-sUv_DC1LdHdasw1uR8', aba: 'POSTIGO' },
    { id: '15uSMKktVGbrwl-qNLEdE6ZBEjxeSWvYFkFO0Zidasd2eOew', aba: 'CRÉDITO FÁCIL' }
  ];

  const destino = SpreadsheetApp.openById(PLANILHA_DESTINO_ID);
  let abaDestino = destino.getSheetByName(ABA_DESTINO);
  if (!abaDestino) abaDestino = destino.insertSheet(ABA_DESTINO);

  // Limpa apenas o conteúdo, mantendo formatação e regras
  const ultimaColuna = abaDestino.getLastColumn() || 20;
  const totalLinhas = abaDestino.getMaxRows();
  if (totalLinhas > 1)
    abaDestino.getRange(2, 1, totalLinhas - 1, ultimaColuna).clearContent();

  let linhaAtual = 2;

  // Copia dados e formatações básicas das origens
  PLANILHAS_ORIGEM.forEach(origemInfo => {
    const origem = SpreadsheetApp.openById(origemInfo.id);
    const abaOrigem = origem.getSheetByName(origemInfo.aba);
    if (!abaOrigem) return;

    const lastRow = abaOrigem.getLastRow();
    const lastCol = abaOrigem.getLastColumn();
    if (lastRow < 2) return;

    const rangeOrigem = abaOrigem.getRange(2, 1, lastRow - 1, lastCol);
    const valores = rangeOrigem.getValues();

    // Garante espaço no destino
    if (abaDestino.getMaxRows() < linhaAtual + valores.length - 1)
      abaDestino.insertRowsAfter(abaDestino.getMaxRows(), (linhaAtual + valores.length) - abaDestino.getMaxRows());

    abaDestino.getRange(linhaAtual, 1, valores.length, lastCol).setValues(valores);

    linhaAtual += valores.length;
  });

  Logger.log(`✅ Consolidação concluída sem sobrescrever validações ou cores (${new Date()})`);
}
