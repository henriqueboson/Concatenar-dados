/*
  > função para concatenar nome dos projetos com nome das atividades
  
  > histórico de revisões
      - 20250428 - R01
        - autor: Henrique
        - observações:
        - em desenvolvimento
        
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

    // Quantidade de linhas na aba do MasterPlan
    var lastRow = abaMasterPlan.getLastRow();
    
    // Buscar nas linhas do MasterPlan
    var dados = abaMasterPlan.getDataRange().getValues();
    var backgrounds = abaMasterPlan.getDataRange().getBackgrounds();
    var fontColors = abaMasterPlan.getRange("G1:G" + abaMasterPlan.getLastRow()).getFontColors();
    
    var corEspecifica = "#e832ff"; // Cor específica para identificar os textos na coluna G
    var secaoAtual = null;
    var secoesConcatenadas = 0;
    var linhasConcatenadas = 0;
    
    // Loop através de todas as linhas
    for (var i = 0; i < lastRow; i++) {
        // Verificar se temos um texto com a cor específica na coluna G
        if (fontColors[i][0] === corEspecifica) {
            secaoAtual = dados[i][6]; // Valor da coluna G (índice 6)
            secoesConcatenadas++;
            Logger.log("Nova seção encontrada: " + secaoAtual + " (linha " + (i+1) + ")");
        }
        
        // Se temos uma seção ativa e a célula na coluna A tem fundo branco, concatenamos
        if (secaoAtual !== null && (backgrounds[i][0] === "#ffffff" || backgrounds[i][0] === "#FFFFFF")) {
            var valorA = dados[i][0]; // Coluna A (índice 0)
            var valorG = dados[i][6]; // Coluna G (índice 6)
            
            // Verifica se valorG existe e não é vazia
            if (valorG !== null && valorG !== undefined && valorG.toString().trim() !== "") {
                // Concatenar A com G e colocar em H
                var novoValorH = valorA + " - " + valorG;
                abaMasterPlan.getRange(i + 1, 8).setValue(novoValorH); // Coluna H = índice 8
                linhasConcatenadas++;
                Logger.log("Linha " + (i+1) + " concatenada: " + novoValorH);
            }
        }
    }
    
    Logger.log("Processamento concluído com sucesso! Seções processadas: " + secoesConcatenadas + ", Linhas concatenadas: " + linhasConcatenadas);
}