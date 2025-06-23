// Classe para gerenciar iluminação
class Illuminator {
  constructor(sheetName, startRow, startCol, lastRow, lastCol, highlightFullRow, backgroundColor, fontColor) {
    this.sheetName = sheetName;
    this.startRow = startRow;
    this.startCol = startCol;
    this.lastRow = lastRow;
    this.lastCol = lastCol;
    this.highlightFullRow = highlightFullRow; // true para iluminar linha inteira, false para apenas a célula
    this.backgroundColor = backgroundColor || "#3c383f"; // Cor de iluminação padrão
    this.fontColor = fontColor || "#ffffff"; // Cor da fonte na iluminação (branco)
  }

  // Limpa o fundo e fonte do intervalo definido
  clearBackground(sheet) {
    let range = sheet.getRange(this.startRow, this.startCol, this.lastRow - this.startRow + 1, this.lastCol);
    range.setBackground(null);
    range.setFontColor(null);
  }

  // Ilumina a seleção conforme configuração
  highlightSelected(sheet, selectedRow, selectedCol) {
    if (this.sheetName === "CADASTRO") {
      // Para CADASTRO, ilumina somente a célula da coluna C (3) na linha selecionada
      let range = sheet.getRange(selectedRow, 3, 1, 1);
      range.setBackground(this.backgroundColor);
      range.setFontColor(this.fontColor);
    } else {
      // Para outras planilhas, segue o comportamento padrão
      if (this.highlightFullRow) {
        let range = sheet.getRange(selectedRow, this.startCol, 1, this.lastCol);
        range.setBackground(this.backgroundColor);
        range.setFontColor(this.fontColor);
      } else {
        let range = sheet.getRange(selectedRow, this.startCol, 1, 1);
        range.setBackground(this.backgroundColor);
        range.setFontColor(this.fontColor);
      }
    }
  }

  // Verifica se a seleção está no intervalo
  isInRange(selectedRow, selectedCol) {
    return selectedRow >= this.startRow && selectedRow <= this.lastRow &&
           selectedCol >= this.startCol && selectedCol <= this.lastCol;
  }
}

// Instâncias para diferentes planilhas com as cores e limites solicitados
const illuminators = [
  new Illuminator("REGISTRO", 17, 1, 200, 5, true, "#3c383f", "#ffffff"), // REGISTRO: linha inteira A:E
  new Illuminator("CADASTRO", 7, 2, 11, 2, false, "#3c383f", "#ffffff")   // CADASTRO: intervalo B7:B11, mas ilumina só coluna C
];

// Função única para mudança de seleção
function onSelectionChange(e) {
  if (!e) {
    Logger.log("Função onSelectionChange chamada sem evento.");
    return;
  }

  var sheet = e.source.getActiveSheet();
  if (!sheet) {
    Logger.log("Nenhuma planilha ativa encontrada.");
    return;
  }

  var sheetName = sheet.getName();
  var selectedCell = e.range;
  if (!selectedCell) {
    Logger.log("Nenhuma célula selecionada.");
    return;
  }

  var selectedRow = selectedCell.getRow();
  var selectedCol = selectedCell.getColumn();

  var illuminator = illuminators.find(i => i.sheetName === sheetName);
  if (!illuminator) {
    return;
  }

  // Limpar o fundo e fonte de todo o intervalo antes de aplicar novo highlight
  illuminator.clearBackground(sheet);

  // Verificar e iluminar apenas se a seleção estiver no intervalo
  if (illuminator.isInRange(selectedRow, selectedCol)) {
    illuminator.highlightSelected(sheet, selectedRow, selectedCol);
  }
}

// Função para pintar fundo externo em preto e texto preto
function colorirFundoExterno() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var configs = [
    {
      sheetName: "REGISTRO",
      startRow: 17,
      endRow: 200,
      startCol: 1,
      endCol: 5
    },
    {
      sheetName: "CADASTRO",
      startRow: 7,
      endRow: 11,
      startCol: 2,
      endCol: 2
    }
  ];

  configs.forEach(function(config) {
    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) {
      Logger.log("Aba '" + config.sheetName + "' não encontrada.");
      return;
    }

    var maxRows = sheet.getMaxRows();
    var maxCols = sheet.getMaxColumns();

    // Limpa toda a planilha para remover fundos e fontes anteriores
    sheet.getRange(1, 1, maxRows, maxCols).setBackground(null).setFontColor(null);

    // Pintar linhas acima do intervalo da tabela
    if (config.startRow > 1) {
      sheet.getRange(1, 1, config.startRow - 1, maxCols).setBackground("#000000").setFontColor("#000000");
    }

    // Pintar linhas abaixo do intervalo da tabela
    if (config.endRow < maxRows) {
      sheet.getRange(config.endRow + 1, 1, maxRows - config.endRow, maxCols).setBackground("#000000").setFontColor("#000000");
    }

    // Pintar colunas à esquerda do intervalo da tabela
    if (config.startCol > 1) {
      sheet.getRange(config.startRow, 1, config.endRow - config.startRow + 1, config.startCol - 1).setBackground("#000000").setFontColor("#000000");
    }

    // Pintar colunas à direita do intervalo da tabela
    if (config.endCol < maxCols) {
      sheet.getRange(config.startRow, config.endCol + 1, config.endRow - config.startRow + 1, maxCols - config.endCol).setBackground("#000000").setFontColor("#000000");
    }
  });
}
