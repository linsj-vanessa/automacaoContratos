function capturarDadosFormulario() {
    var sheet = SpreadsheetApp.openById("SPREADSHEET_ID").getSheetByName("Respostas");
    var dadosTodasAsLinhas = sheet.getDataRange().getValues(); 
  
    var ultimaLinhaValida = 0;
    for (var i = dadosTodasAsLinhas.length - 1; i > 0; i--) {
      if (dadosTodasAsLinhas[i].some(valor => valor.toString().trim() !== "")) {
        ultimaLinhaValida = i;
        break;
      }
    }
  
    if (ultimaLinhaValida === 0) {
      Logger.log("Nenhuma resposta válida encontrada.");
      return;
    }
  
    var dados = dadosTodasAsLinhas[ultimaLinhaValida];
    Logger.log("Dados capturados corretamente: " + JSON.stringify(dados));
  }
  
  function criarContrato() {
    var docModeloId = "DOCUMENT_TEMPLATE_ID"; 
    var pastaDestinoId = "DESTINATION_FOLDER_ID"; 
  
    var sheet = SpreadsheetApp.openById("SPREADSHEET_ID").getSheetByName("Respostas");
    var dadosTodasAsLinhas = sheet.getDataRange().getValues(); 
  
    var ultimaLinhaValida = 0;
    for (var i = dadosTodasAsLinhas.length - 1; i > 0; i--) {
      if (dadosTodasAsLinhas[i].some(valor => valor.toString().trim() !== "")) {
        ultimaLinhaValida = i;
        break;
      }
    }
    if (ultimaLinhaValida === 0) {
      Logger.log("Nenhuma resposta válida encontrada.");
      return;
    }
  
    var dados = dadosTodasAsLinhas[ultimaLinhaValida];
    var nomeEmpresa = dados[1];
  
    if (!nomeEmpresa) {
      Logger.log("Nome da empresa não encontrado.");
      return;
    }
  
    var pastaDestino = DriveApp.getFolderById(pastaDestinoId);
    var pastas = pastaDestino.getFoldersByName("Contrato - " + nomeEmpresa);
    var pastaEmpresa = pastas.hasNext() ? pastas.next() : pastaDestino.createFolder("Contrato - " + nomeEmpresa);
  
    var docCopia = DriveApp.getFileById(docModeloId).makeCopy("Contrato - " + nomeEmpresa, pastaEmpresa);
    Logger.log("Contrato criado: " + docCopia.getUrl());
  
    return docCopia.getId();
  }
  
  function preencherContrato() {
    var contratoId = criarContrato(); 
    if (!contratoId) {
      Logger.log("Erro ao criar contrato.");
      return;
    }
  
    var sheet = SpreadsheetApp.openById("SPREADSHEET_ID").getSheetByName("Respostas");
    var dadosTodasAsLinhas = sheet.getDataRange().getValues(); 
  
    var ultimaLinhaValida = 0;
    for (var i = dadosTodasAsLinhas.length - 1; i > 0; i--) {
      if (dadosTodasAsLinhas[i].some(valor => valor.toString().trim() !== "")) {
        ultimaLinhaValida = i;
        break;
      }
    }
    if (ultimaLinhaValida === 0) {
      Logger.log("Nenhuma resposta válida encontrada.");
      return;
    }
  
    var dados = dadosTodasAsLinhas[ultimaLinhaValida];
    var nomeEmpresa = dados[1];
    var CNPJ = dados[2];
    var dataInicio = dados[3];
    var dataTermino = dados[4];
    var telefone = dados[9];
    var email = dados[10];
    var nomeProjeto = dados[11];
    
    var docEditavel = DocumentApp.openById(contratoId);
    var corpo = docEditavel.getBody();
  
    corpo.replaceText("{{empresa}}", nomeEmpresa);
    corpo.replaceText("{{CNPJ}}", CNPJ);
    corpo.replaceText("{{dataInicio}}", dataInicio);
    corpo.replaceText("{{dataTermino}}", dataTermino);
    corpo.replaceText("{{telefone}}", telefone);
    corpo.replaceText("{{email}}", email);
    corpo.replaceText("{{nomeProjeto}}", nomeProjeto);
  
    docEditavel.saveAndClose();
    Logger.log("Contrato preenchido com sucesso: " + docEditavel.getUrl());
  
    return contratoId;
  }
  