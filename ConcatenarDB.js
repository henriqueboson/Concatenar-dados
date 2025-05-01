/*
  > função para concatenar nome dos projetos com nome das atividades
  
  > histórico de revisões
      - 20250430 - R02
        - autor: Henrique
        - observações:
        - 
*/

function concatenarMSP() {
    var masterPlan = SpreadsheetApp.getActiveSpreadsheet();
    
    var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    var mesAtual = meses[new Date().getMonth()];
    var abaMasterPlan = masterPlan.getSheetByName(mesAtual);
    
    if (!abaMasterPlan) {
      Logger.log("Erro: Aba do mês '" + mesAtual + "' não encontrada no MasterPlan.");
      return;
    }
    
    var dados = abaMasterPlan.getDataRange().getValues();
    var backgrounds = abaMasterPlan.getDataRange().getBackgrounds();
    var fontColors = abaMasterPlan.getRange(1, 7, dados.length, 1).getFontColors(); // Coluna G
    var fontColorsColunaA = abaMasterPlan.getRange(1, 1, dados.length, 1).getFontColors();

    
    var corEspecifica = "#e832ff";
    var corExcluir = "#375623"; // Cor para verificar na coluna A
    var valorGdaVez = "";
    var processar = false;
    var secoesConcatenadas = 0;
    var linhasConcatenadas = 0;
    var linhasExcluidas = 0;
    
    // Array para armazenar as linhas que serão excluídas
    var linhasParaExcluir = [];
    
    // Primeiro passo: identificar as linhas que têm a cor #275317 na coluna A
    for (var i = 0; i < backgrounds.length; i++) {
      if (fontColorsColunaA[i][0].toLowerCase() === corExcluir.toLowerCase()){
        linhasParaExcluir.push(i + 1); // +1 porque as linhas na planilha começam em 1
        Logger.log("Linha " + (i+1) + " marcada para exclusão (cor #275317 encontrada)");
      }
    }
    
    // Segundo passo: processar as concatenações como antes
    for (var i = 0; i < dados.length; i++) {
      // Ignorar processamento se a linha estiver na lista de exclusão
      if (linhasParaExcluir.indexOf(i + 1) !== -1) {
        continue;
      }
      
      Logger.log("Linha " + (i+1) + " - Cor da fonte em G: " + fontColors[i][0]);
      
      if (fontColors[i][0] && fontColors[i][0].toLowerCase() === corEspecifica.toLowerCase()) {
        valorGdaVez = dados[i][6]; // Coluna G
        processar = true;
        secoesConcatenadas++;
        Logger.log("Nova seção encontrada: " + valorGdaVez + " (linha " + (i+1) + ")");
        
        if (backgrounds[i][0].toLowerCase() === "#ffffff") {
          var valorA = dados[i][0];
          var novoValorH = (valorA ? valorA : "") + " - " + valorGdaVez;
          abaMasterPlan.getRange(i + 1, 8).setValue(novoValorH);
          linhasConcatenadas++;
          Logger.log("Concatenou linha " + (i+1) + ": " + novoValorH);
        }
        
        continue;
      }
      
      if (processar) {
        if (backgrounds[i][0].toLowerCase() === "#ffffff") {
          var valorA = dados[i][0];
          var novoValorH = (valorA ? valorA : "") + " - " + valorGdaVez;
          abaMasterPlan.getRange(i + 1, 8).setValue(novoValorH);
          linhasConcatenadas++;
          Logger.log("Concatenou linha " + (i+1) + ": " + novoValorH);
        }
      }
    }
    
    // Terceiro passo: excluir as linhas marcadas (de baixo para cima para não afetar os índices)
    if (linhasParaExcluir.length > 0) {
      linhasParaExcluir.sort(function(a, b) { return b - a; }); // Ordenar em ordem decrescente
      
      for (var i = 0; i < linhasParaExcluir.length; i++) {
        abaMasterPlan.deleteRow(linhasParaExcluir[i]);
        linhasExcluidas++;
      }
    }
    
    Logger.log("Processamento concluído com sucesso! Seções processadas: " + 
               secoesConcatenadas + ", Linhas concatenadas: " + linhasConcatenadas + 
               ", Linhas excluídas: " + linhasExcluidas);
  }