/*
  > função para concatenar nome dos projetos com nome das atividades
  
  > histórico de revisões
      - 20250429
        - autor: Henrique
        - observações:
        - versão funcional
        
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

    var corEspecifica = "#e832ff";
    var valorGdaVez = "";
    var processar = false;
    var secoesConcatenadas = 0;
    var linhasConcatenadas = 0;

    for (var i = 0; i < dados.length; i++) {
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

    Logger.log("Processamento concluído com sucesso! Seções processadas: " +
               secoesConcatenadas + ", Linhas concatenadas: " + linhasConcatenadas);
}
