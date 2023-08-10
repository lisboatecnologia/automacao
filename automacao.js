function onEdit(e) {
  const sheet = e.source.getSheetName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  if (sheet === "Oportunidades") {
    if (col === 4 && e.value === "NOVO") {
      moveRowToSheet("Novos", row);
    }

    if (col === 7) {
      handleStatusChange(range, row, e.value);
    }
  }

  if (sheet === "ListaClientes" && [9, 10, 11].includes(col)) {
    handleContactChange(range, row, col, e.value);
  }

  updateDtAlteracao(row);
}

function moveRowToSheet(sheetName, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Oportunidades");
  const targetSheet = ss.getSheetByName(sheetName);

  const values = sourceSheet.getRange(row, 1, 1, sourceSheet.getLastColumn()).getValues();
  targetSheet.appendRow(values[0]);
  sourceSheet.deleteRow(row);
}

function handleStatusChange(range, row, value) {
  const sheet = range.getSheet();
  const statusColumn = 7;
  const dtFinalColumn = 6;
  const situacaoColumn = 5;
  
  if (value === "CONCLUIDO") {
    range.setBackground("#b7e1cd"); // Verde claro
    sheet.getRange(row, dtFinalColumn).setValue(new Date());
  } else if (value === "REJEITADO") {
    range.setBackground("#ffbdbd"); // Vermelho claro
    sheet.getRange(row, dtFinalColumn).setValue(new Date());
    sheet.getRange(row, situacaoColumn).setValue("NÃO TEM INTERESSE");
  }
}

function handleContactChange(range, row, col, value) {
  const sheet = range.getSheet();
  const statusColumn = 7;

  if (value === "CLIENTE JA RENOVOU" || value === "CLIENTE NAO TEM INTERESSE") {
    updateStatus("CONCLUIDO", row, statusColumn);
  }

  updateDtAlteracao(row);
}

function updateStatus(newStatus, row, statusColumn) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(row, statusColumn).setValue(newStatus);
}

function updateDtAlteracao(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dtAlteracaoColumn = 8;
  sheet.getRange(row, dtAlteracaoColumn).setValue(new Date());
}
