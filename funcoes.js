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

function consultarDisponibilidadeData(dataString, horaString) {
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

    // Identifica o período escolhido pelo usuário
    const horaEscolhida = parseInt(horaString.split(':')[0]);
    const periodoEscolhido = horaEscolhida < 12 ? 'manha' : 'tarde';
    
    const totalGeral = totalManha + totalTarde;
    let disponivel = false;
    let mensagem = "";

    // VALIDAÇÃO 1: Limite total do dia (6 agendamentos)
    if (totalGeral >= 6) {
      disponivel = false;
      mensagem = "Infelizmente esta data está totalmente lotada (limite de 6 agendamentos diários atingido).<br><br>Por favor, escolha outra data.";
    } 
    // VALIDAÇÃO 2: Período escolhido está lotado?
    else if (periodoEscolhido === 'manha' && totalManha >= 3) {
      disponivel = false;
      mensagem = `O período da <strong>MANHÃ</strong> já está lotado (3/3 agendamentos).`;
      if (totalTarde < 3) {
        mensagem += `<br><br><div class="alert alert-warning mb-0 mt-2"><strong>💡 Sugestão:</strong> Ainda temos ${3 - totalTarde} vaga(s) disponível(is) no período da <strong>TARDE</strong>.<br>Altere o horário para após 13:00 e consulte novamente.</div>`;
      } else {
        mensagem += `<br><br>Por favor, escolha outra data.`;
      }
    }
    else if (periodoEscolhido === 'tarde' && totalTarde >= 3) {
      disponivel = false;
      mensagem = `O período da <strong>TARDE</strong> já está lotado (3/3 agendamentos).`;
      if (totalManha < 3) {
        mensagem += `<br><br><div class="alert alert-warning mb-0 mt-2"><strong>💡 Sugestão:</strong> Ainda temos ${3 - totalManha} vaga(s) disponível(is) no período da <strong>MANHÃ</strong>.<br>Altere o horário para antes de 12:00 e consulte novamente.</div>`;
      } else {
        mensagem += `<br><br>Por favor, escolha outra data.`;
      }
    }
    // VALIDAÇÃO 3: Está disponível!
    else {
      disponivel = true;
      const vagasPeriodo = periodoEscolhido === 'manha' ? (3 - totalManha) : (3 - totalTarde);
      const nomePeriodo = periodoEscolhido === 'manha' ? 'MANHÃ' : 'TARDE';
      
      mensagem = `✅ Data e horário disponíveis!<br><br>`;
      mensagem += `📊 <strong>Status atual:</strong><br>`;
      mensagem += `• Manhã: ${totalManha}/3 agendamentos<br>`;
      mensagem += `• Tarde: ${totalTarde}/3 agendamentos<br><br>`;
      mensagem += `Você escolheu o período da <strong>${nomePeriodo}</strong> (${vagasPeriodo} vaga(s) restante(s)).`;
    }

    return {
      disponivel: disponivel,
      mensagem: mensagem,
      totalManha: totalManha,
      totalTarde: totalTarde,
      periodoEscolhido: periodoEscolhido
    };

  } catch (e) {
    Logger.log("Erro em consultarDisponibilidadeData: " + e.message);
    return { 
      disponivel: false, 
      mensagem: "Erro ao consultar calendário: " + e.message
    };
  }
}

function processarAgendamento(dados) {
  try {
    const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_ABA);

    // 1. Criar o Evento no Calendário
    const inicio = new Date(dados.dataDeposito + 'T' + dados.horaDeposito);
    const fim = new Date(inicio.getTime() + (60 * 60 * 1000)); // Duração de 1 hora

    const evento = agenda.createEvent(
      `Depósito: ${dados.nome}`,
      inicio,
      fim,
      { description: `Título: ${dados.tituloTese}\nNº USP: ${dados.nrUsp}\nE-mail: ${dados.emailUSP}` }
    );

    // 2. Preparar dados para a planilha
    // Adiciona o campo "tipo" com valor fixo "ME" (Mestrado)
    dados.tipo = "ME";

    // 3. Gravar na Planilha de forma dinâmica
    const headers = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
    const novaLinha = headers.map(header => {
      // Se o cabeçalho for "tipo", retorna "ME"
      if (header.toLowerCase() === 'tipo') {
        return "ME";
      }
      // Caso contrário, retorna o valor do campo correspondente
      return dados[header] || "";
    });
    
    planilha.appendRow(novaLinha);

    return {
      sucesso: true,
      nome: dados.nome,
      data: dados.dataDeposito,
      hora: dados.horaDeposito,
      titulo: dados.tituloTese
    };

  } catch (e) {
    Logger.log("Erro em processarAgendamento: " + e.message);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Busca a lista de orientadores na aba 'Orientadores'
 */
function listarOrientadores() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("Orientadores");

    // Pega todos os dados da coluna A (pulando o cabeçalho)
    const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();

    // Converte de array de array [[nome1], [nome2]] para array simples [nome1, nome2]
    const listaSimples = valores.map(linha => linha[0]).filter(nome => nome !== "");

    return listaSimples.sort(); // Retorna em ordem alfabética
  } catch (e) {
    console.error("Erro ao listar orientadores: " + e.message);
    return [];
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