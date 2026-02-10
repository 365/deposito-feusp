function forcarAutorizacao() {
  try {
    // Substitua pelo seu email real
    MailApp.sendEmail(
      "apmbraga@usp.br",  // ← COLOQUE SEU EMAIL AQUI
      "Teste de Autorização - Apps Script",
      "Este é um email de teste para autorizar o script a usar o MailApp."
    );
    
    Logger.log("Autorização concedida e email enviado com sucesso!");
    
  } catch (erro) {
    Logger.log("Erro: " + erro.toString());
  }
}