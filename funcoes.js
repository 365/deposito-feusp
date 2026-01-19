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

// Agenda onde sera criada a data do evento
const ID_AGENDA_DEPOSITOS = 'c_f0c47043a5564c65f0ac0835c28e3b3fa13c3bf80618daa471d01679bc7a281d@group.calendar.google.com'

// Planilha onde os dados serao gravados
const ID_PLANILHA = '1yXdWwSiTsSbour4dQ-WhSl2r3LVzf_acxk3-EY2nV8E';
const NOME_ABA = 'Cadastro';

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

    // Lógica de verificação de lotação
    if (totalManha >= 3 && totalTarde >= 3) {
      disponivel = false; // Garante que o JS saiba que não pode agendar
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
    const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);
    const dataAlvo = new Date(dataString + 'T00:00:00');
    
    // Busca eventos do dia inteiro
    const inicioDia = new Date(dataAlvo.getTime());
    const fimDia = new Date(dataAlvo.getTime());
    fimDia.setHours(23, 59, 59);

    const eventos = agenda.getEvents(inicioDia, fimDia);

    let totalManha = 0;
    let totalTarde = 0;

    eventos.forEach(ev => {
      const hora = ev.getStartTime().getHours();
      if (hora < 12) totalManha++;
      else totalTarde++;
    });

    const totalGeral = totalManha + totalTarde;
    let disponivel = false;
    let mensagem = "";

    // TRAVA GLOBAL: Se já tem 6, não importa o horário, está lotado.
    if (totalGeral >= 6) {
      disponivel = false;
      mensagem = "Infelizmente esta data está totalmente lotada (limite de 6 agendamentos diários atingido).";
    } 
    // TRAVA POR PERÍODO: Se o usuário quer manhã mas já tem 3, ou tarde e já tem 3
    else if (totalManha >= 3 && totalTarde >= 3) {
       disponivel = false;
       mensagem = "Data indisponível: Ambos os períodos (manhã e tarde) já atingiram o limite de 3 cada.";
    }
    else {
      disponivel = true;
      mensagem = `Data disponível! No momento temos: ${totalManha} agendamento(s) de manhã e ${totalTarde} à tarde.`;
    }

    return {
      disponivel: disponivel,
      mensagem: mensagem,
      totalManha: totalManha,
      totalTarde: totalTarde
    };

  } catch (e) {
    return { disponivel: false, mensagem: "Erro ao consultar: " + e.message };
  }
}

function processarAgendamento(dados) {
  try {
    const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_ABA);

    // 1. Criar o Evento no Calendário
    // Combinamos data e hora para o agendamento
    const inicio = new Date(dados.dataDeposito + 'T' + dados.horaDeposito);
    const fim = new Date(inicio.getTime() + (60 * 60 * 1000)); // Duração de 1 hora

    const evento = agenda.createEvent(
      `Depósito: ${dados.nome}`,
      inicio,
      fim,
      { description: `Título: ${dados.tituloTese}\nNº USP: ${dados.nrUsp}\nE-mail: ${dados.emailUSP}` }
    );

    // 2. Gravar na Planilha (Dinâmico)
    // Buscamos os cabeçalhos para saber a ordem das colunas
    const headers = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
    const novaLinha = headers.map(header => dados[header] || ""); // Mapeia o dado para a coluna certa
    planilha.appendRow(novaLinha);

    return {
      sucesso: true,
      nome: dados.nome,
      data: dados.dataDeposito,
      hora: dados.horaDeposito,
      titulo: dados.tituloTese
    };

  } catch (e) {
    return { sucesso: false, erro: e.message };
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