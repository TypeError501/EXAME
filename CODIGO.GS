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
  // Se precisar preencher as 5 colunas, pode ajustar aqui. Vou preencher da esquerda para direita.
  var dadosParaInserir = [];

  // Pega até 5 valores da lista (se tiver menos, preenche com vazio)
  for (var i = 0; i < 5; i++) {
    dadosParaInserir.push(dadosTabela1[i] ? dadosTabela1[i][0] : "");
  }

  // Encontrar a última linha preenchida na planilha "REGISTRO" a partir da linha 17 na coluna A
  var colunaA = destino.getRange("A17:A999").getValues();
  var ultimaLinhaRelativa = 0; // índice relativo dentro do intervalo A17:A999

  for (var i = colunaA.length - 1; i >= 0; i--) {
    if (colunaA[i][0] !== "" && colunaA[i][0] !== null) {
      ultimaLinhaRelativa = i + 1;
      break;
    }
  }

  // A linha absoluta para inserir os dados será 17 + ultimaLinhaRelativa
  var linhaInserir = 17 + ultimaLinhaRelativa;

  // Insere os dados na linha encontrada, colunas A até E
  destino.getRange(linhaInserir, 1, 1, 5).setValues([dadosParaInserir]);

  // Apagar os dados da tabela 1 na planilha origem
  origem.getRange("C7:C11").clearContent();

  SpreadsheetApp.getUi().alert("Dados cadastrados com sucesso na planilha REGISTRO!");
}
