/**
 * Função DUMMY apenas para forçar o Apps Script a pedir permissões do Calendar
 * Execute esta função manualmente UMA VEZ para autorizar
 */
function forcarAutorizacaoCalendar() {
  // Esta linha vai forçar o Google a pedir permissão de Calendar
  CalendarApp.getCalendarById('primary').getName();

  Logger.log("✅ Autorização concedida! Agora pode usar o calendário normalmente.");
}

// Force Auth: CalendarApp.getEvents(new Date(), new Date());

const ID_AGENDA_DEPOSITOS = 'c_f0c47043a5564c65f0ac0835c28e3b3fa13c3bf80618daa471d01679bc7a281d@group.calendar.google.com'

function doGet() {
  return HtmlService.createTemplateFromFile('web').evaluate().setTitle('Formulário de Depósito');
}

function obterDadosHtml(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

function onOpen() {
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

function consultarDisponibilidadeDataOLD(dataString) {
  try {
    // 1. CONFIGURAÇÃO DO NOVO CALENDÁRIO
    // Substitua pelo ID do calendário que você criou
    //const ID_CALENDARIO = "SEU_ID_AQUI@group.calendar.google.com"; 
    //const agenda = CalendarApp.getCalendarById(ID_CALENDARIO);
    const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);

    const dataRef = new Date(dataString + 'T00:00:00');

    // 2. Definir intervalos (Manhã: 08h-12h | Tarde: 13h-19h)
    const inicioDia = new Date(dataRef.getTime());
    inicioDia.setHours(8, 0, 0, 0);

    const meioDia = new Date(dataRef.getTime());
    meioDia.setHours(12, 0, 0, 0);

    const inicioTarde = new Date(dataRef.getTime());
    inicioTarde.setHours(13, 0, 0, 0);

    const fimDia = new Date(dataRef.getTime());
    fimDia.setHours(19, 0, 0, 0);

    // 3. Buscar eventos existentes no calendário específico
    const eventosManha = agenda.getEvents(inicioDia, meioDia);
    const eventosTarde = agenda.getEvents(inicioTarde, fimDia);

    const totalManha = eventosManha.length;
    const totalTarde = eventosTarde.length;

    // Regras de Negócio Aplicadas
    let disponivel = false;
    let mensagem = "";

    if (totalManha >= 3 && totalTarde >= 3) {
      mensagem = "Infelizmente este dia está totalmente lotado (3 manhã / 3 tarde).";
    } else {
      disponivel = true;
      mensagem = `Data disponível! No momento temos: ${totalManha} agendadas de manhã e ${totalTarde} à tarde.`;
    }

    return {
      disponivel: disponivel,
      mensagem: mensagem,
      totalManha: totalManha,
      totalTarde: totalTarde
    };

  } catch (e) {
    return { disponivel: false, mensagem: "Erro ao acessar o calendário: " + e.message };
  }
}


function consultarDisponibilidadeData(dataString) {
  try {
    // LOG 1: Ver o que está chegando
    console.log("📅 Data recebida:", dataString);

    const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);

    // LOG 2: Verificar se conseguiu acessar o calendário
    if (!agenda) {
      console.error("❌ Calendário não encontrado!");
      return {
        disponivel: false,
        mensagem: "Erro: Calendário não encontrado. Verifique o ID e as permissões."
      };
    }

    console.log("✅ Calendário acessado com sucesso");

    const dataRef = new Date(dataString + 'T00:00:00');
    console.log("📆 Data processada:", dataRef);

    // Definir intervalos
    const inicioDia = new Date(dataRef.getTime());
    inicioDia.setHours(8, 0, 0, 0);

    const meioDia = new Date(dataRef.getTime());
    meioDia.setHours(12, 0, 0, 0);

    const inicioTarde = new Date(dataRef.getTime());
    inicioTarde.setHours(13, 0, 0, 0);

    const fimDia = new Date(dataRef.getTime());
    fimDia.setHours(19, 0, 0, 0);

    console.log("⏰ Buscando eventos entre:", inicioDia, "e", fimDia);

    // Buscar eventos
    const eventosManha = agenda.getEvents(inicioDia, meioDia);
    const eventosTarde = agenda.getEvents(inicioTarde, fimDia);

    const totalManha = eventosManha.length;
    const totalTarde = eventosTarde.length;

    console.log("📊 Eventos encontrados - Manhã:", totalManha, "Tarde:", totalTarde);

    let disponivel = false;
    let mensagem = "";

    if (totalManha >= 3 && totalTarde >= 3) {
      mensagem = "Infelizmente este dia está totalmente lotado (3 manhã / 3 tarde).";
    } else {
      disponivel = true;
      mensagem = `Data disponível! No momento temos: ${totalManha} agendamento(s) de manhã e ${totalTarde} à tarde.`;
    }

    console.log("✅ Resposta:", { disponivel, mensagem, totalManha, totalTarde });

    return {
      disponivel: disponivel,
      mensagem: mensagem,
      totalManha: totalManha,
      totalTarde: totalTarde
    };

  } catch (e) {
    console.error("💥 ERRO CAPTURADO:", e.message);
    console.error("Stack:", e.stack);
    return {
      disponivel: false,
      mensagem: "Erro ao acessar o calendário: " + e.message
    };
  }
}


function testarAcessoCalendar() {
  try {
    const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);

    if (!agenda) {
      Logger.log("❌ Calendário NÃO encontrado com o ID: " + ID_AGENDA_DEPOSITOS);
      return "ERRO: Calendário não encontrado";
    }

    Logger.log("✅ Calendário encontrado: " + agenda.getName());

    const hoje = new Date();
    const amanha = new Date(hoje.getTime() + 24 * 60 * 60 * 1000);
    const eventos = agenda.getEvents(hoje, amanha);

    Logger.log("📅 Eventos encontrados: " + eventos.length);

    return "Teste OK! Calendário acessível.";

  } catch (e) {
    Logger.log("💥 ERRO: " + e.message);
    return "ERRO: " + e.message;
  }
}