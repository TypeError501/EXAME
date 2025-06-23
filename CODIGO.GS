function cadastrarDados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var origem = ss.getActiveSheet();              // Planilha atual (onde estão os dados)
  var destino = ss.getSheetByName("REGISTRO");   // Planilha "REGISTRO"

  // Pega os dados da tabela 1 (C7:C11)
  var dadosTabela1 = origem.getRange("C7:C11").getValues();

  // Verifica se há dados na tabela 1 (pelo menos uma célula não vazia)
  var temDados = dadosTabela1.some(function(linha) {
    return linha.some(function(celula) {
      return celula !== "" && celula !== null;
    });
  });

  if (!temDados) {
    SpreadsheetApp.getUi().alert("A tabela 1 está vazia. Nada para cadastrar.");
    return;
  }

  // Concatenar os dados da coluna C7:C11 em uma única linha de até 5 colunas (A até E)
  var dadosParaInserir = [];

  for (var i = 0; i < 5; i++) {
    dadosParaInserir.push(dadosTabela1[i] ? dadosTabela1[i][0] : "");
  }

  // Insere uma nova linha na posição 17 para empurrar as linhas existentes para baixo
  destino.insertRowBefore(17);

  // Agora insere os dados na linha 17, colunas A até E
  destino.getRange(17, 1, 1, 5).setValues([dadosParaInserir]);

  // Apagar os dados da tabela 1 na planilha origem
  origem.getRange("C7:C11").clearContent();

  SpreadsheetApp.getUi().alert("Dados cadastrados com sucesso na planilha REGISTRO!");
}
