// Função executada ao editar a planilha
function onEdit(e) {
  if (!e) return;

  const sheet = e.range.getSheet();
  const sheetName = "REGISTRO";
  const searchCell = "A6";
  const indicatorCell = "B6";
  const dataStartRow = 17;
  const dataEndRow = 400;
  const dataStartCol = 1;
  const dataEndCol = 5;

  if (sheet.getName() !== sheetName) return;

  if (e.range.getA1Notation() === searchCell) {
    const val = e.range.getValue().toString();

    const indicatorRange = sheet.getRange(indicatorCell);
    const currentIndicator = indicatorRange.getValue();

    // Atualiza o indicador "Pesquisar" conforme o conteúdo da barra de pesquisa
    if (val.trim() === "") {
      if (currentIndicator !== "Pesquisar") {
        indicatorRange.setValue("Pesquisar");
      }
    } else {
      if (currentIndicator !== "") {
        indicatorRange.clearContent();
      }
    }

    // Se o valor for exatamente "Pesquisar", apaga para facilitar digitação
    if (val === "Pesquisar") {
      e.range.setValue("");
      return;
    }

    // Aplica o filtro na tabela
    const searchTerm = val.toLowerCase().trim();
    const numRows = dataEndRow - dataStartRow + 1;
    const numCols = dataEndCol - dataStartCol + 1;
    const dataRange = sheet.getRange(dataStartRow, dataStartCol, numRows, numCols);
    const dataValues = dataRange.getValues();

    if (searchTerm === "") {
      sheet.showRows(dataStartRow, numRows);
      return;
    }

    sheet.hideRows(dataStartRow, numRows);

    for (let i = 0; i < dataValues.length; i++) {
      const row = dataValues[i];
      const rowText = row.map(cell => cell.toString().toLowerCase()).join(" ");
      if (rowText.includes(searchTerm)) {
        sheet.showRows(dataStartRow + i);
      }
    }
  }
}

// Função para aplicar estilos iniciais na barra de pesquisa e no indicador
function aplicarEstilosBarraPesquisa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("REGISTRO");
  if (!sheet) return;

  const searchCell = sheet.getRange("A6");
  const indicatorCell = sheet.getRange("B6");

  // Fundo azul claro na barra de pesquisa (A6)
  searchCell.setBackground("#ADD8E6");

  // Texto "Pesquisar" e cor preta no indicador (B6), se estiver vazio
  if (indicatorCell.getValue().trim() === "") {
    indicatorCell.setValue("Pesquisar");
  }
  indicatorCell.setFontColor("#000000");
}

// Função para aplicar estilos ao abrir a planilha
function onOpen() {
  aplicarEstilosBarraPesquisa();
}
