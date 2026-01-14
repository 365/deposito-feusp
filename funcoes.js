function doGet() {
  return HtmlService.createTemplateFromFile('web').evaluate().setTitle('Formulário de Depósito');
}

function obterDadosHtml(nome)
{
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

function onOpen(){
  DocumentApp.getUi().createAddonMenu('Opções Avançadas')
  .addItem('Formulário HTML', 'mostrarFormulario')
  .addToUi();
}

function mostrarFormulario() {
  var html = HtmlService.createHtmlOutputFromFile('pag.html')
  .setWidth(1000)
  .setHeight(700)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  DocumentApp.getUi().showModalDialog(html, "Formulário personalizado com Google Apps Script");
}

