function sendEmails() {
  
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails").activate();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  
  var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1, 1).getValue();
  
  var email_Enviado = "email_Enviado";
  var cumpreRequisitos = "Sim";
  var campoVazio = "";  
  var cargaPadrao = 45;
  
  for (var i = 3; i<=lr;i++){
    
    if ( (ss.getRange(i, 19).getValue() != email_Enviado && ss.getRange(i, 19).getValue() == campoVazio) && ss.getRange(i, 18).getValue() == cumpreRequisitos) {
    
    var nomeUsuario = ss.getRange(i, 1).getValue();
    var currentEmail = ss.getRange(i, 2).getValue();
    var numeroMatricula = ss.getRange(i, 3).getValue();
    var numeroRegistro = ss.getRange(i,4).getValue();
    var numeroCurso1 = ss.getRange(i, 5).getValue();
    var cargaHoraria1 = ss.getRange(i, 6).getValue();
    var instituicao1 = ss.getRange(i, 7).getValue();
    var numeroCurso2 = ss.getRange(i, 8).getValue();
    var cargaHoraria2 = ss.getRange(i, 9).getValue();
    var instituicao2 = ss.getRange(i, 10).getValue();
    var numeroCurso3 = ss.getRange(i, 11).getValue();
    var cargaHoraria3 = ss.getRange(i, 12).getValue();
    var instituicao3 = ss.getRange(i, 13).getValue();
    var numeroCurso4 = ss.getRange(i, 14).getValue();
    var cargaHoraria4 = ss.getRange(i, 15).getValue();
    var instituicao4 = ss.getRange(i, 16).getValue();
    var totalCarga = ss.getRange(i, 17).getValue();
    var cumpreRequisitos = ss.getRange(i, 18).getValue();
    var emailEnviado = ss.getRange(i, 19).getValue();
    
      
      if (totalCarga > cargaPadrao ){ 
    
    var messageBody = templateText.replace("(nome)",nomeUsuario).replace("(matricula)",numeroMatricula).replace("(registro)",numeroRegistro).replace("(Curso1)",numeroCurso1).replace("(cargaHoraria1)",cargaHoraria1).replace("(Curso2)",numeroCurso2).replace("(cargaHoraria2)",cargaHoraria2).replace("(Curso3)",numeroCurso3).replace("(cargaHoraria3)",cargaHoraria3)
    .replace("(instituicao1)",instituicao1).replace("(instituicao2)",instituicao2).replace("(instituicao3)",instituicao3).replace("(Curso4)",numeroCurso4).replace("(cargaHoraria4)",cargaHoraria4).replace("(instituicao4)",instituicao4).replace("(totalCarga)",totalCarga).replace("(observacao)","Obs: A Carga horária excedente a 45 horas não será cumulativa para o próximo período.");
      
    }else if(totalCarga = cargaPadrao){
    var messageBody = templateText.replace("(nome)",nomeUsuario).replace("(matricula)",numeroMatricula).replace("(registro)",numeroRegistro).replace("(Curso1)",numeroCurso1).replace("(cargaHoraria1)",cargaHoraria1).replace("(Curso2)",numeroCurso2).replace("(cargaHoraria2)",cargaHoraria2).replace("(Curso3)",numeroCurso3).replace("(cargaHoraria3)",cargaHoraria3)
    .replace("(instituicao1)",instituicao1).replace("(instituicao2)",instituicao2).replace("(instituicao3)",instituicao3).replace("(Curso4)",numeroCurso4).replace("(cargaHoraria1)",cargaHoraria4).replace("(instituicao4)",instituicao4).replace("(totalCarga)",totalCarga).replace("(observacao)","");
      
      
    var subjectLine = "Prezado(a) " + nomeUsuario + ", matrícula: " +numeroMatricula +" sua análise de cursos para progressão foi realizada!";
      
    
    GmailApp.sendEmail(currentEmail, subjectLine, messageBody);
    
            
    ss.getRange(i, 19).setValue(email_Enviado);  // atualizacao de campo de condicional
        
    } else { 
    
      ss.getRange(i, 19).setValue("Carga Horaria INSUFICIENTE");  // indicacao de nao cumprimento de exigencia
    }
    
    }
    else {}
    }
        
  }
  
  

