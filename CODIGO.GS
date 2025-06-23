function cadastrarDados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var origem = ss.getActiveSheet();           // Planilha atual (onde estão os dados)
  var destino = ss.getSheetByName("REGISTRO"); // Planilha "REGISTRO"

  // Pega os dados da tabela 1 (B7:C11)
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

  // Concatenar os dados da tabela 1 em uma única lista (10 elementos)
  var dadosConcatenados = [];
  for (var i = 0; i < dadosTabela1.length; i++) {
    for (var j = 0; j < dadosTabela1[i].length; j++) {
      dadosConcatenados.push(dadosTabela1[i][j]);
    }
  }

  // Copiar os primeiros 5 dados para a tabela 2 na planilha REGISTRO (A17:E17)
  var dadosParaTabela2 = dadosConcatenados.slice(0, 5);
  destino.getRange("A17:E17").setValues([dadosParaTabela2]);

  // Apagar os dados da tabela 1 na planilha origem
  origem.getRange("C7:C11").clearContent();

  SpreadsheetApp.getUi().alert("Dados cadastrados com sucesso na planilha REGISTRO!");
}

