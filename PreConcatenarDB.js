/*
  > função para rodar antes da função principal de concatenação do MasterPlan
  
  > histórico de revisões
      - 20250505 - R01
        - autor: Henrique
        - observações:
        - 
*/

function preConcatenarMSP(){
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
    var fontColorsColunaA = abaMasterPlan.getRange(1, 1, dados.length, 1).getFontColors();

    var corConferencia = "#d0cece";
    var corEspecifica = "#e832ff"; 
    var corExcluir = "#375623"; 
    var linhasExcluidas = 0;
    var linhaCopiada = 0;

    // Array para armazenar as linhas que serão excluídas
    var linhasParaExcluir = [];
    
    // Identifica as linhas que têm a cor #275317 na coluna A
    for (var i = 0; i < dados.length; i++) {
      if (fontColorsColunaA[i][0].toLowerCase() === corExcluir.toLowerCase()){
        linhasParaExcluir.push(i + 1); // +1 porque as linhas na planilha começam em 1
        Logger.log("Linha " + (i+1) + " marcada para exclusão (cor #275317 encontrada)");
      }
    }

    // Identifica as linhas que têm a cor #d0cece na coluna A e copia o valor para a coluna G
    for (var i = 0; i < dados.length; i++){
      if (fontColorsColunaA[i][0].toLowerCase() === corConferencia) {
        var valorAconferencia = dados[i][0];
        abaMasterPlan.getRange(i + 1, 7).setValue(valorAconferencia).setFontColor(corEspecifica); // Coluna G = coluna 7
        linhaCopiada++;
        Logger.log("Conferência: copiou valor da coluna A para G na linha " + (i + 1));
      }
    }

    // Exclui as linhas marcadas (de baixo para cima para não afetar os índices)
    if (linhasParaExcluir.length > 0) {
      linhasParaExcluir.sort(function(a, b) { return b - a; }); // Ordenar em ordem decrescente
      
      for (var i = 0; i < linhasParaExcluir.length; i++) {
        abaMasterPlan.deleteRow(linhasParaExcluir[i]);
        linhasExcluidas++;
      }
    }

    Logger.log("Processamento concluído com sucesso! Linhas copiadas: " + linhaCopiada + 
        ", Linhas excluídas: " + linhasExcluidas);

}