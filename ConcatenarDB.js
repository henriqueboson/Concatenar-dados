/*
  > função para concatenar nome dos projetos com nome das atividades
  
  > histórico de revisões
      - 20250429
        - autor: Henrique
        - observações:
        - em dev 
        
*/

function concatenarMSP() {
    
    var masterPlan = SpreadsheetApp.getActiveSpreadsheet();
    
    var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    var mesAtual = meses[new Date().getMonth()];
    var abaMasterPlan = masterPlan.getSheetByName(mesAtual);
    if (!abaMasterPlan) {
        Logger.log("Erro: Aba do mês '" + mesAtual + "' não encontrada no MasterPlan.");
        return;
    }
    
    // Buscar nas linhas do MasterPlan
    var dados = abaMasterPlan.getDataRange().getValues();
    var backgrounds = abaMasterPlan.getDataRange().getBackgrounds();
    var fontColors = abaMasterPlan.getRange("G1:G" + abaMasterPlan.getLastRow()).getFontColors();
    
    var corEspecifica = "#e832ff"; // Cor específica para identificar os textos na coluna G
    var valorGdaVez = ""; // Guarda o valor G da seção atual (texto roxo)
    var processar = false;
    var secoesConcatenadas = 0;
    var linhasConcatenadas = 0;
    
    for (var i = 0; i < dados.length; i++) {
        // Verificar se temos um texto com a cor específica na coluna G
        
        Logger.log("Linha " + (i+1) + " - Cor da fonte G: " + fontColors[i][0]);

        if (fontColors[i][0].toLowerCase() === corEspecifica.toLowerCase()){
            valorGdaVez = dados[i][6]; // Atualiza o valor G da seção atual
            processar = true;
            secoesConcatenadas++;
            Logger.log("Nova seção encontrada: " + valorGdaVez + " (linha " + (i+1) + ")");
            
            // Também concatena esta própria linha se tiver fundo branco na coluna A
            if (backgrounds[i][0] === "#ffffff" || backgrounds[i][0] === "#FFFFFF") {
                var valorA = dados[i][0]; // Coluna A (índice 0)
                var novoValorH = valorA + " - " + valorGdaVez;
                abaMasterPlan.getRange(i + 1, 8).setValue(novoValorH); // Coluna H = índice 8
                linhasConcatenadas++;
                Logger.log("Concatenou linha " + (i+1) + ": " + novoValorH);
            }
            
            continue;
        }
        
        // Se estamos processando uma seção
        if (processar) {
            // Verificar se a célula na coluna A tem fundo branco
            if (backgrounds[i][0] === "#ffffff" || backgrounds[i][0] === "#FFFFFF") {
                var valorA = dados[i][0]; // Coluna A (índice 0)
                
                // Concatenar valor A atual com o valor G da seção atual
                var novoValorH = valorA + " - " + valorGdaVez;
                abaMasterPlan.getRange(i + 1, 8).setValue(novoValorH); // Coluna H = índice 8
                linhasConcatenadas++;
                Logger.log("Concatenou linha " + (i+1) + ": " + novoValorH);
            }
        }
    }
    
    Logger.log("Processamento concluído com sucesso! Seções processadas: " + secoesConcatenadas + ", Linhas concatenadas: " + linhasConcatenadas);
}