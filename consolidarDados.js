function consolidarDados() {
  var planilhaDestino = SpreadsheetApp.getActiveSpreadsheet();
  var abaDestino = planilhaDestino.getSheetByName("Consolidado");

  if (!abaDestino) {
    abaDestino = planilhaDestino.insertSheet("Consolidado");
  }

  abaDestino.clear(); // Limpa os dados antigos

  var urls = [
    { url: "xxxxxxxxx", dia: "25/03" },
    { url: "xxxxxxxxx", dia: "26/03" },
    { url: "xxxxxxxxx", dia: "27/03" }
  ];

  var nomes = {}; // Dicionário para armazenar nome, matrícula, ligação com a JMU e presenças por dia

  urls.forEach(function(entry) {
    try {
      var planilhaOrigem = SpreadsheetApp.openByUrl(entry.url);
      var abaOrigem = planilhaOrigem.getSheetByName("Respostas ao formulário 1"); // Ajuste o nome correto da aba
      
      if (abaOrigem) {
        var ultimaLinha = abaOrigem.getLastRow();
        if (ultimaLinha > 1) { // Evita erro caso a planilha esteja vazia
          var dados = abaOrigem.getRange("B2:D" + ultimaLinha).getValues();

          var nomesUnicos = new Map(); // Usar um Map para armazenar informações únicas dentro da mesma planilha

          dados.forEach(function(row) {
            var nome = row[0].toString().trim().toLowerCase();
            var matricula = row[1] ? row[1].toString().trim() : "N/A";
            var jmu = row[2] ? row[2].toString().trim() : "N/A";

            if (nome) {
              if (!nomesUnicos.has(nome)) {
                nomesUnicos.set(nome, { matricula, jmu });
              }
            }
          });

          // Atualizar a contagem e presença nos dias
          nomesUnicos.forEach(function(valor, nome) {
            if (!nomes[nome]) {
              nomes[nome] = { 
                frequencia: 0, 
                matricula: valor.matricula, 
                jmu: valor.jmu, 
                presenca: { "25/03": "", "26/03": "", "27/03": "" } // Inicializa como vazio
              };
            }
            nomes[nome].frequencia += 1;

            // Verificar se existe um valor válido para o dia antes de marcar a presença
            if (dados.some(row => row[0].toString().trim().toLowerCase() === nome)) {
              nomes[nome].presenca[entry.dia] = "X"; // Marca a presença no dia certo
            }
          });

        } else {
          Logger.log("! A aba 'Respostas ao formulário 1' da planilha " + entry.url + " está vazia!");
        }
      } else {
        Logger.log("! Aba 'Respostas ao formulário 1' não encontrada no arquivo: " + entry.url);
      }
    } catch (e) {
      Logger.log("X Erro ao acessar a planilha: " + entry.url + " - " + e.message);
    }
  });

  // Converter o dicionário em uma lista e ordenar alfabeticamente
  var resultado = [["Nome", "Matrícula", "Ligação com a JMU", "25/03", "26/03", "27/03", "Frequência"]];
  var listaOrdenada = Object.entries(nomes).sort((a, b) => a[0].localeCompare(b[0]));

  listaOrdenada.forEach(function(entry) {
    var nome = entry[0];
    var info = entry[1];
    resultado.push([nome, info.matricula, info.jmu, info.presenca["25/03"], info.presenca["26/03"], info.presenca["27/03"], info.frequencia]);
  });

  // Escrever os dados na aba "Consolidado"
  abaDestino.getRange(1, 1, resultado.length, 7).setValues(resultado);
}
