/**
 * Função DUMMY apenas para forçar o Apps Script a pedir permissões do Calendar
 */
function forcarAutorizacaoCalendar() {
  CalendarApp.getCalendarById('primary').getName();
  Logger.log("✅ Autorização concedida! Agora pode usar o calendário normalmente.");
}

/** ========================
     Configurações Globais  
    ======================== */
const ID_TIPO_PLANILHA = 'DOU'; // 1=QUALI, 2=ME, 3=DOU
const ID_AGENDA_DEPOSITOS = 'c_f0c47043a5564c65f0ac0835c28e3b3fa13c3bf80618daa471d01679bc7a281d@group.calendar.google.com'
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

/** =============================================
     CONSULTA DE DISPONIBILIDADE
    ============================================= */
function consultarDisponibilidadeData(dataString, horaString) {
  try {
    const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);
    const dataAlvo = new Date(dataString + 'T00:00:00');
    
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

    const horaEscolhida = parseInt(horaString.split(':')[0]);
    const periodoEscolhido = horaEscolhida < 12 ? 'manha' : 'tarde';
    
    const totalGeral = totalManha + totalTarde;
    let disponivel = false;
    let mensagem = "";

    if (totalGeral >= 6) {
      disponivel = false;
      mensagem = "Infelizmente esta data está totalmente lotada (limite de 6 agendamentos diários atingido).<br><br>Por favor, escolha outra data.";
    } 
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
    return { disponivel: false, mensagem: "Erro ao consultar calendário: " + e.message };
  }
}

/** =============================================
     PROCESSAR AGENDAMENTO
    ============================================= */
function processarAgendamento(dados) {
  try {
    const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const planilha = ss.getSheetByName(NOME_ABA);

    // 1. Criar o Evento no Calendário
    const inicio = new Date(dados.dataAgenda + 'T' + dados.horaAgenda);
    const fim = new Date(inicio.getTime() + (60 * 60 * 1000));

    agenda.createEvent(
      `Depósito: ${dados.nome}`,
      inicio,
      fim,
      { description: `Título: ${dados.tituloTese}\nNº USP: ${dados.nrUsp}\nE-mail: ${dados.emailAluno}` }
    );

    // 2. Lógica Coluna I (tipoDefesa) baseada nas marcações do frontend
    dados.tipoDefesa = calcularTipoDefesa(dados.listaMarcacoes);

    // 3. Gravar Planilha DINÂMICA
      const headers = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
      const novaLinha = headers.map(header => {
      const headerTrimmed = header.toString().trim();
      // REGRA NOVA: Se a coluna for a que criamos na Planilha, gera a data/hora agora
      if (headerTrimmed === 'Data do Deposito') return new Date();
      if (headerTrimmed === 'tipo') return ID_TIPO_PLANILHA;
      if (headerTrimmed === 'tipoDefesa') return dados.tipoDefesa || '';
      return dados[headerTrimmed] || '';
    });

    planilha.appendRow(novaLinha);

    return {
      sucesso: true, 
      nome: dados.nome, 
      data: dados.dataAgenda,
      hora: dados.horaAgenda,
      titulo: dados.tituloTese
    };
  } catch (e) {
    Logger.log('Erro em processarAgendamento: ' + e.message);
    return { sucesso: false, erro: e.message };
  }
}


/**
 * Lógica da Coluna I
 * Retorna vazio se não houver marcações enviadas
 */
function calcularTipoDefesa(marcacoes) {
  // Se não houver marcações ou o array estiver vazio, retorna vazio para a planilha
  if (!marcacoes || marcacoes.length === 0) return ""; 

  const total = marcacoes.length;
  const distancias = marcacoes.filter(m => m === 'Distancia').length;

  if (distancias === 0) return "Presencial";
  if (distancias === total) return "Distancia";
  return "Hibrido";
}

/** =====================================================
     Busca a lista de orientadores na aba 'Orientadores'
    ===================================================== */
function listarOrientadores() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName("Orientadores");
    if (!aba) return [];
    const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
    const listaSimples = valores.map(linha => linha[0]).filter(nome => nome !== "");
    return listaSimples.sort();
  } catch (e) {
    console.error("Erro ao listar orientadores: " + e.message);
    return [];
  }
}

/** =====================================================
     Teste de Acesso ao Calendário
    ===================================================== */
  function testarAcessoCalendar() {
    try {
      const agenda = CalendarApp.getCalendarById(ID_AGENDA_DEPOSITOS);
      if (!agenda) return "ERRO: Calendário não encontrado";
      const hoje = new Date();
      const amanha = new Date(hoje.getTime() + 24 * 60 * 60 * 1000);
      const eventos = agenda.getEvents(hoje, amanha);
      return "Teste OK! Eventos encontrados: " + eventos.length;
    } catch (e) {
      return "ERRO: " + e.message;
    }
  }

/** =====================================================
     Teste de Acesso ao Calendário
    ===================================================== */
function salvarDadosNaPlanilha(dados) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA); 
    const planilha = ss.getSheetByName(NOME_ABA);
    if (!planilha) throw new Error("Aba '" + NOME_ABA + "' não encontrada!");

    // 1. Captura tudo de uma vez para ganhar performance
    const rangeTotal = planilha.getDataRange();
    const valoresPlanilha = rangeTotal.getValues();
    const headers = valoresPlanilha[0].map(h => h.toString().trim());
    
    // 2. Localiza as colunas cruciais
    const idxNrUsp = headers.indexOf('nrUsp');
    const idxData = headers.indexOf('Data do Depósito');
    
    if (idxNrUsp === -1) throw new Error("Coluna 'nrUsp' não encontrada!");

    // 3. Procura a linha do aluno pelo nrUsp
    // Começamos de 1 para pular o header. findIndex retorna o índice do array (0-based)
    const rowIndex = valoresPlanilha.findIndex((linha, i) => i > 0 && linha[idxNrUsp] == dados.nrUsp);

    // 4. Constrói a linha de dados de forma inteligente
    const novaLinha = headers.map((h, colIdx) => {
      // Regra 1: Se for a coluna de data e já existir um valor lá (em caso de atualização)
      if (colIdx === idxData && rowIndex !== -1) {
        return valoresPlanilha[rowIndex][idxData]; // Mantém a data original
      }
      
      // Regra 2: Se for a coluna de data e for linha NOVA
      if (colIdx === idxData && rowIndex === -1) {
        return new Date(); // Grava a data agora apenas se for a primeira vez
      }

      // Regra 3: Tipo fixo
      if (h === 'tipo') return ID_TIPO_PLANILHA;

      // Regra 4: Mapeamento dinâmico do formulário
      return dados[h] !== undefined ? dados[h] : (rowIndex !== -1 ? valoresPlanilha[rowIndex][colIdx] : "");
    });

    // 5. Decide se Atualiza ou Insere
    if (rowIndex !== -1) {
      // Atualiza apenas a linha específica (rowIndex + 1 porque a planilha começa em 1)
      planilha.getRange(rowIndex + 1, 1, 1, novaLinha.length).setValues([novaLinha]);
      console.log(`✅ Sincronizado: nrUsp ${dados.nrUsp} atualizado na linha ${rowIndex + 1}`);
    } else {
      // Insere no final
      planilha.appendRow(novaLinha);
      console.log(`✅ Criado: Novo registro para nrUsp ${dados.nrUsp}`);
    }
    
    // 6. Dispara o e-mail de confirmação para o Aluno
    enviarEmailConfirmacao(dados);

    // 7. Retorna o HTML da página de sucesso para o navegador exibir
    return carregarPaginaSucesso(dados);

  } catch (erro) {
    console.error("Erro no salvamento: " + erro);
    throw new Error("Erro ao salvar: " + erro.message);
  }
}

/**
 * Função que envia o e-mail (Cole logo abaixo da salvarDadosNaPlanilha)
 */
function enviarEmailConfirmacao(dados) {
  const assunto = "Depósito Enviado - Sistema de Depósito Digital FEUSP";
  
  // Aqui usamos os nomes das chaves que estão no seu objeto window.dadosAluno
  const corpo = `
    Depósito Enviado com Sucesso!

    Seu depósito foi registrado no Sistema de Depósito Digital - FEUSP.

    Confirmação do Agendamento:
    --------------------------------------
    Nome: ${dados.nome}
    Data: ${dados.dataAgenda}
    Horário: ${dados.horaAgenda}
    Título: ${dados.tituloTese}
    --------------------------------------

    Verifique sua caixa de entrada e, caso não localize este e-mail, verifique também a pasta de Spams.

    Secretaria de Pós-Graduação – FEUSP
    Sistema de Depósito Digital
      `;
  
  // Envia para o e-mail do aluno (assumindo que a chave é dados.email)
  if (dados.email) {
    MailApp.sendEmail(dados.email, assunto, corpo);
  }
}


/* function carregarPaginaSucesso(dados) {
  // Cria o template a partir do arquivo sucesso.html
  var template = HtmlService.createTemplateFromFile('sucesso');
  
  // Injeta as variáveis que o mainSucesso.html vai usar
  template.nome = dados.nome;
  template.data = dados.dataAgenda;
  template.hora = dados.horaAgenda;
  template.titulo = dados.tituloTese;
  
  // Retorna o HTML final renderizado
  return template.evaluate().getContent();
} */

function carregarPaginaSucesso(dados) {
  // 1. Processa o conteúdo interno primeiro
  var templateMain = HtmlService.createTemplateFromFile('mainSucesso');
  templateMain.nome = dados.nome;
  templateMain.data = dados.dataAgenda;
  templateMain.hora = dados.horaAgenda;
  templateMain.titulo = dados.tituloTese;
  var mainProcessado = templateMain.evaluate().getContent();

  // 2. Cria a página principal (sucesso.html)
  var layout = HtmlService.createTemplateFromFile('sucesso');
  
  // 3. Precisamos de uma forma de passar o mainProcessado para o sucesso.html
  // No seu sucesso.html, no lugar de obterDadosHtml('mainSucesso'), 
  // você usaria: <?!= conteudoPrincipal ?>
  layout.conteudoPrincipal = mainProcessado;

  return layout.evaluate().getContent();
}

