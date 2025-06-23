function pintarFundoPretoRespeitandoTabelas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Configurações das tabelas
  const tabelas = [
    {
      sheetName: "CADASTRO",
      intervalo: {startRow: 7, endRow: 11, startCol: 2, endCol: 3} // B7:C11
    },
    {
      sheetName: "REGISTRO",
      intervalo: {startRow: 17, endRow: 400, startCol: 1, endCol: 5} // A17:E400
    }
  ];

  tabelas.forEach(({sheetName, intervalo}) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Planilha '${sheetName}' não encontrada.`);
      return;
    }

    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();

    // Limpa o fundo de toda a planilha para garantir estado inicial
    sheet.getRange(1, 1, maxRows, maxCols).setBackground(null);

    // Pintar linhas acima do intervalo da tabela
    if (intervalo.startRow > 1) {
      sheet.getRange(1, 1, intervalo.startRow - 1, maxCols).setBackground("#000000");
    }

    // Pintar linhas abaixo do intervalo da tabela
    if (intervalo.endRow < maxRows) {
      sheet.getRange(intervalo.endRow + 1, 1, maxRows - intervalo.endRow, maxCols).setBackground("#000000");
    }

    // Pintar colunas à esquerda do intervalo da tabela
    if (intervalo.startCol > 1) {
      sheet.getRange(intervalo.startRow, 1, intervalo.endRow - intervalo.startRow + 1, intervalo.startCol - 1).setBackground("#000000");
    }

    // Pintar colunas à direita do intervalo da tabela
    if (intervalo.endCol < maxCols) {
      sheet.getRange(intervalo.startRow, intervalo.endCol + 1, intervalo.endRow - intervalo.startRow + 1, maxCols - intervalo.endCol).setBackground("#000000");
    }
  });
}


