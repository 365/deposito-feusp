/** ==============================================================================
     Função DUMMY apenas para forçar o Apps Script a pedir permissões do Calendar  
    ============================================================================== */
function forcarAutorizacaoCalendar() {
  CalendarApp.getCalendarById('primary').getName();
  Logger.log("✅ Autorização concedida! Agora pode usar o calendário normalmente.");
}

/** ========================================================================
     Função para forçar o Apps Script a pedir permissões de envio de E-mail  
    ======================================================================== */
function forcarAutorizacaoEmail() {
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

/** ====================================
     Testar a visualização dos e-mails   
    =================================== */
function testarVisualizacao() {
  const template = HtmlService.createTemplateFromFile('templatesEmails');
  
  // Criamos dados fictícios para o teste
  template.dados = {
    nome:       "Fulano de Tal",
    dataAgenda: "15/05/2026",
    horaAgenda: "14:00",
    tituloTese: "A Educação no Século XXI",
    nrUsp:      "1234567",
    emailAluno: "aluno@usp.br",
    orientador: "Prof. Dr. Orientador Exemplo"
  };
  
  template.logoFeusp = LOGO_FEUSP_BK_COMP; // Coloque o link da imagem
  
  // ESCOLHA O QUE QUER VER: 'ALUNO', 'SECRETARIA' ou 'ORIENTADOR'
  template.tipo = 'ALUNO'; 
  
  const htmlFinal = template.evaluate().getContent();
  Logger.log(htmlFinal); // Isso joga o código limpo no console
  
  // Isso abre uma janela no Google Apps Script para você ver
  const htmlOutput = HtmlService.createHtmlOutput(htmlFinal)
      .setWidth(650)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Teste de Visualização");
}

/** ========================
     Configurações Globais  
    ======================== */
const ID_TIPO_POS = 'ME'; // 1=QUALI, 2=ME, 3=DOU
const ID_AGENDA_DEPOSITOS = 'c_f0c47043a5564c65f0ac0835c28e3b3fa13c3bf80618daa471d01679bc7a281d@group.calendar.google.com';
const ID_PLANILHA = '1yXdWwSiTsSbour4dQ-WhSl2r3LVzf_acxk3-EY2nV8E';

const ID_MODELO_DOC_QUALI = '1GbWxl8gdalsJae1l-iQL1rm91YGGwuA6LnegJsL19iA';
const ID_MODELO_DOC_ME = '1e5BP2KptSFB7B-6cQfPUGCCi_ALLEZx_2k9DFeuVgvE';
const ID_MODELO_DOC_DOU = '164xGKfmthOr0MKWqSDJ3iguIo7o8b7-qvQ0E2zysbf4';

const ID_PASTA_PDFS = '1vExAeBv5NSvgtoRnXdoah6wX-uRQluEy';
const NOME_ABA_CADASTRO = 'Cadastro';
const NOME_ABA_ORIENTADORES = 'Orientadores';
const NOME_ABA_DOCENTES = 'Docentes';
const LOGO_FEUSP_BK_COMP = "https://365studio.com.br/wp-content/uploads/2026/02/Logo-FEUSP-Preto-comprido-RGB.png";
const LOGO_FEUSP_BK_RETANG = "https://365studio.com.br/wp-content/uploads/2026/02/Logo-FEUSP-Preto-retangular-RGB.png";

function doGet() {
  //return HtmlService.createTemplateFromFile('web').evaluate().setTitle('Formulário de Depósito');

  // 1. Cria o template a partir do arquivo principal (ex: index.html)
  const template = HtmlService.createTemplateFromFile('web');
  
  // 2. INJETA AS CONSTANTES GLOBAIS AQUI (Isso resolve o erro 404)
  template.LOGO_FEUSP_BK_COMP = LOGO_FEUSP_BK_COMP;
  template.LOGO_FEUSP_BK_RETANG = LOGO_FEUSP_BK_RETANG;
  template.ID_TIPO_POS = ID_TIPO_POS; // <--- ADICIONE ESTA LINHA
  
  // 3. Renderiza e define configurações de visualização
  return template.evaluate()
      .setTitle('Depósito Digital - FEUSP')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obterDadosHtml(nome) {
  const template = HtmlService.createTemplateFromFile(nome);
  
  // 1. Injeta os logos (Globais)
  template.LOGO_FEUSP_BK_COMP = LOGO_FEUSP_BK_COMP;
  template.LOGO_FEUSP_BK_RETANG = LOGO_FEUSP_BK_RETANG;
  
  // 2. Cria um objeto vazio para 'dados' caso o arquivo chame essa variável
  // Isso evita o erro "ReferenceError: dados is not defined"
  template.dados = {}; 

  try {
    return template.evaluate().getContent();
  } catch (e) {
    // Se ainda der erro, ele nos avisa qual arquivo está com problema
    return "";
  }
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
    const planilha = ss.getSheetByName(NOME_ABA_CADASTRO);

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
      if (headerTrimmed === 'tipo') return ID_TIPO_POS;
      if (headerTrimmed === 'tipoDefesa') return dados.tipoDefesa || '';
      return dados[headerTrimmed] || '';
    });

    planilha.appendRow(novaLinha);

    // Criamos a data formatada para exibição
    const dataExibicao = dados.dataAgenda.split('-').reverse().join('/');

    return {
      sucesso: true, 
      nome: dados.nome, 
      data: dataExibicao,
      hora: dados.horaAgenda,
      titulo: dados.tituloTese
    };
  } catch (e) {
    Logger.log('Erro em processarAgendamento: ' + e.message);
    return { sucesso: false, erro: e.message };
  }
}

/** ====================================================================
     LOGICA DA COLUNA I RETORNA VAZIO SE NAO HOUVER MARCACOES ENVIADAS
    ==================================================================== */
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
     BUSCA A LISTA DE ORIENTADORES NA ABA ORIENTADORES
    ===================================================== */
function listarOrientadores() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName(NOME_ABA_ORIENTADORES);
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
     LISTA DOCENTES DA ABA DOCENTES
    ===================================================== */
function listarDocentes() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName(NOME_ABA_DOCENTES);  // Nova aba
    if (!aba) return [];
    
    const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
    const listaSimples = valores.map(linha => linha[0]).filter(nome => nome !== "");
    return listaSimples.sort();
  } catch (e) {
    console.error("Erro ao listar docentes: " + e.message);
    return [];
  }
}

/** =====================================================
     TESTE DE ACESSO AO CALENDARIO
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

/** ============================
     SALVAR DADOS NA PLANILHA
    ============================ */
function salvarDadosNaPlanilha(dados) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA); 
    const planilha = ss.getSheetByName(NOME_ABA_CADASTRO);
    if (!planilha) throw new Error("Aba '" + NOME_ABA_CADASTRO + "' não encontrada!");

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
      if (h === 'tipo') return ID_TIPO_POS;

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
    //enviarEmailConfirmacao(dados);

    // 7. Retorna o HTML da página de sucesso para o navegador exibir
    return carregarPaginaSucesso(dados);

  } catch (erro) {
    console.error("Erro no salvamento: " + erro);
    throw new Error("Erro ao salvar: " + erro.message);
  }
}

/** ===============================
     ENVIAR E-MAIL DE CONFIRMAÇÃO
    =============================== */
function enviarEmailConfirmacao(dados) {
  // --- PASSO 1: TRADUÇÃO (Lógica no Servidor) ---
  // Aqui o servidor já conhece a sua const global ID_TIPO_POS
  const tradutor = {
    'QUALI': 'QUALIFICAÇÃO',
    'ME':    'MESTRADO',
    'DOU':   'DOUTORADO'
  };

  // Alimentamos o objeto 'dados' com o nome por extenso
  dados.cursoExtensao = tradutor[ID_TIPO_POS];

  // --- PASSO A: GERA O PDF ANTES DE TUDO ---
  const pdfOficial = gerarPdfPeloDoc(dados);

  const template = HtmlService.createTemplateFromFile('templatesEmails');
  
  // Passa os dados recebidos do front para o template
  template.dados = dados;
  template.logoFeusp = LOGO_FEUSP_BK_RETANG;

  // --- PASSO B: ORGANIZA OS ANEXOS ---
  let anexosSecretaria = [pdfOficial]; // Começa com o PDF que acabamos de criar

  // 1. Preparamos o anexo se ele existir
  var anexoTese = [];
  if (dados.pdfBlobTese) {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(dados.pdfBlobTese.conteudo), 
      dados.pdfBlobTese.mimeType, 
      dados.pdfBlobTese.nome
    );
    anexoTese.push(blob);
  }
  
  // Envio para o ALUNO
  template.tipo = 'ALUNO';
  const corpoAluno = template.evaluate().getContent();
  MailApp.sendEmail({
    to: dados.emailAluno,
    subject: "Confirmação de Depósito - FEUSP",
    htmlBody: corpoAluno
  });

  // Envio para o ORIENTADOR
  template.tipo = 'ORIENTADOR';
  const corpoOrientador = template.evaluate().getContent();
  MailApp.sendEmail({
    to: '365studiobr@gmail.com', // TESTE: seu e-mail fixo por enquanto
    subject: "Assinatura Necessária - Depósito de " + dados.nome,
    htmlBody: corpoOrientador,
    attachments: [pdfOficial] // <--- Ele vê o PDF gerado do Word
  });

  // Envio para a SECRETARIA
  template.tipo = 'SECRETARIA';
  const corpoSecretaria = template.evaluate().getContent();
  MailApp.sendEmail({
    to: "apmbraga@gmail.com",
    subject: "Novo Depósito Digital: " + dados.nome,
    htmlBody: corpoSecretaria,
    attachments: anexoTese // <--- O PDF entra aqui
  });

  return true; // Importante para o .withSuccessHandler do front saber que acabou
}

/** ===============================
     DISPARAR FLUXO DE E-MAILS PARA ALUNO E ORIENTADOR
     (TESTE MANUAL PARA VER SE O HTML CHEGA CORRETO)
    =============================== */
function dispararFluxoEmails(dadosFormulario) {
  const template = HtmlService.createTemplateFromFile('templatesEmails');
  
  // FORÇANDO O E-MAIL PARA TESTE (O pulo do gato)
  dadosFormulario.emailOrientador = 'alexandre.de.paul@gmail.com'; 
  // Se quiser testar o do aluno também no seu e-mail:
  // dadosFormulario.emailAluno = '365studiobr@gmail.com';

  template.dados = dadosFormulario;
  template.logoFeusp = LOGO_FEUSP_BK_RETANG;

  // --- 1. ENVIO ALUNO ---
  template.tipo = 'ALUNO';
  const corpoAluno = template.evaluate().getContent();
  MailApp.sendEmail({
    to: dadosFormulario.emailAluno,
    subject: "Confirmação de Depósito - FEUSP",
    htmlBody: corpoAluno
  });

  // --- 2. ENVIO ORIENTADOR ---
  template.tipo = 'ORIENTADOR';
  const corpoOrientador = template.evaluate().getContent();
  MailApp.sendEmail({
    to: dadosFormulario.emailOrientador, // Vai enviar para o 365studiobr
    subject: "Assinatura Necessária: Depósito de " + dadosFormulario.nome,
    htmlBody: corpoOrientador
  });

  Logger.log("Testes disparados para " + dadosFormulario.emailOrientador);
}

/** ==================================================================
     TESTE MANUAL DE FLUXO PARA VER SE O HTML CHEGA CORRETO NO E-MAIL
     (DISPARA PARA O SEU E-MAIL, VERIFICAR SE O HTML ESTÁ CORRETO)
    ================================================================== */
function testeManualDeFluxo() {
  // Criamos um objeto igual ao que o seu formulário cria
  const dadosFake = {
    nome: "Fulano de Tal",
    emailAluno: "365studiobr@gmail.com", // Mande para você mesmo
    emailOrientador: "365studiobr@gmail.com", // Mande para você mesmo
    dataAgenda: "15/05/2026",
    horaAgenda: "14:00",
    tituloTese: "A importância do design nos sistemas acadêmicos",
    nrUsp: "1234567",
    orientador: "Prof. Dr. Exemplo",
    areaConcentracao: "Educação e Tecnologia",
    arquivoPDF: "tese_final_v1.pdf"
  };

  // Chamamos a função principal que você montou
  dispararFluxoEmails(dadosFake);
}

/** ================================================
    CARREGAR PÁGINA DE SUCESSO COM DADOS DINÂMICOS
    ================================================ */
function carregarPaginaSucesso(dados) {
  // LOG DE RASTREIO
  console.log("LOG: Entrei na função carregarPaginaSucesso");
  console.log("LOG: Dados recebidos:", JSON.stringify(dados));

  try {
    // 1. Processa o conteúdo interno
    console.log("LOG: Tentando abrir mainSucesso...");
    var templateMain = HtmlService.createTemplateFromFile('mainSucesso');
    templateMain.dados = dados;
    var mainProcessado = templateMain.evaluate().getContent();

    // 2. Cria a página principal
    console.log("LOG: Tentando abrir moldura sucesso...");
    var layout = HtmlService.createTemplateFromFile('sucesso');
    
    layout.conteudoPrincipal = mainProcessado;
    
    console.log("LOG: Tudo certo! Retornando HTML final.");
    return layout.evaluate().getContent();

  } catch (erro) {
    // SE O ERRO APARECER AQUI, O LOG VAI DIZER EXATAMENTE QUAL ARQUIVO FALTOU
    console.error("ERRO DENTRO DA FUNÇÃO:", erro.toString());
    throw new Error("Erro ao montar página: " + erro.toString());
  }
}

/** ====================================================
    FUNÇÃO TESTE PARA VALIDAR O RECEBIMENTO DE ARQUIVOS
    ==================================================== */
function testarBlobAppsScript(dados) {
  try {
    // Tenta criar um Blob a partir dos dados recebidos para testar a integridade
    var blob = Utilities.newBlob(Utilities.base64Decode(dados.conteudo), dados.mimeType, dados.nome);
    
    return {
      status: "Sucesso",
      mensagem: "Recebido com sucesso! O arquivo '" + dados.nome + "' tem " + blob.getBytes().length + " bytes."
    };
  } catch (erro) {
    return {
      status: "Erro",
      mensagem: "Falha ao processar arquivo no servidor: " + erro.toString()
    };
  }
}

/** ======================================================================
    GERANDO PDF DINAMICAMENTE PELO DOC (TESTE DUMMY PARA VER SE FUNCIONA)
    ====================================================================== */
function gerarPdfPeloDoc(dados) {
  // --- A CORREÇÃO ESTÁ AQUI ---
  // Se o cursoExtensao não veio preenchido, nós traduzimos agora!
  if (!dados.cursoExtensao) {
    const tradutor = { 'QUALI': 'QUALIFICAÇÃO', 'ME': 'MESTRADO', 'DOU': 'DOUTORADO' };
    dados.cursoExtensao = tradutor[ID_TIPO_POS] || "Não Definido";
  }

  // --- O QUE VAMOS ADICIONAR LOGO ABAIXO ---
  // Criamos a const com a primeira letra maiúscula e o resto minúscula
  const tituloDeposito = dados.cursoExtensao.charAt(0).toUpperCase() + dados.cursoExtensao.slice(1).toLowerCase();
  
  console.log("DEBUG Replace - Curso (Após correção):", dados.cursoExtensao);
  console.log("=== INICIO DA GERAÇÃO DO PDF ===");
  console.log("Dados recebidos:", JSON.stringify(dados)); // Inspeciona o objeto completo

  const modeloId = ID_MODELO_DOC_DOU; 
  const pastaId = ID_PASTA_PDFS; 
  // Montando o nome do arquivo com o Tipo de Pós, Nº USP e Nome do Aluno para facilitar a identificação
  const nomeArquivo = "Deposito - " + (ID_TIPO_POS) + " - " + (dados.nrUsp) + " - " + (dados.nome);

  try {
    const modeloArquivo = DriveApp.getFileById(modeloId);
    const pastaDestino = DriveApp.getFolderById(pastaId);
    const agora = new Date(); // Captura o momento exato uma única vez
    
    // DECLARAÇÃO ÚNICA:
    const copia = modeloArquivo.makeCopy(nomeArquivo, pastaDestino); 
    
    const doc = DocumentApp.openById(copia.getId());
    const body = doc.getBody();

    // Logs individuais para monitorar o que vai entrar no replace
    console.log("DEBUG Replace - Nome:", dados.nome);
    console.log("DEBUG Replace - Titulo:", dados.tituloTese);
    console.log("DEBUG Replace - Curso:", dados.cursoExtensao);

    // Tratamento da Data da Defesa (YYYY-MM-DD para DD/MM/YYYY)
    let dataDefesaFormatada = dados.dataAgenda || "";
    if (dataDefesaFormatada.includes("-")) {
      dataDefesaFormatada = dataDefesaFormatada.split('-').reverse().join('/');
    }    

    // Substituições
    body.replaceText("{{TITULO_DEPOSITO}}", tituloDeposito);
    //=== identificação =========================================
    body.replaceText("{{NOME_ALUNO}}", dados.nome);
    body.replaceText("{{EMAIL_ALUNO}}", dados.emailAluno);
    body.replaceText("{{NR_USP}}", dados.nrUsp);
    //=== dados academicos ======================================
    body.replaceText("{{AREA_CONCENTRACAO}}", dados.areaConcentracao);
    body.replaceText("{{TITULO_TESE}}", dados.tituloTese);
    //=== agendamento ===========================================
    body.replaceText("{{DATA_DEFESA}}", dataDefesaFormatada);
    body.replaceText("{{HORA_DEFESA}}", dados.horaAgenda);

    //=== membros titulares =====================================
    body.replaceText("{{NOME_ORIENTADOR}}", dados.orientador);
    body.replaceText("{{EMAIL_ORIENTADOR}}", dados.emailOrientador);
    body.replaceText("{{VINCULO_ORIENTADOR}}", dados.vinculoOrientador);
    //=== segundo titular =======================================
    body.replaceText("{{NOME_TITULAR_2}}", dados.nomeTitular2);
    body.replaceText("{{EMAIL_TITULAR_2}}", dados.emailTitular2);
    body.replaceText("{{VINCULO_TITULAR_2}}", dados.vinculoTitular2);
    //=== terceiro titular ======================================
    body.replaceText("{{NOME_TITULAR_3}}", dados.nomeTitular3);
    body.replaceText("{{EMAIL_TITULAR_3}}", dados.emailTitular3);
    body.replaceText("{{VINCULO_TITULAR_3}}", dados.vinculoTitular3);
    //=== quarto titular ======================================
    body.replaceText("{{NOME_TITULAR_4}}", dados.nomeTitular4);
    body.replaceText("{{EMAIL_TITULAR_4}}", dados.emailTitular4);
    body.replaceText("{{VINCULO_TITULAR_4}}", dados.vinculoTitular4);
    //=== quinto titular ======================================
    body.replaceText("{{NOME_TITULAR_5}}", dados.nomeTitular5);
    body.replaceText("{{EMAIL_TITULAR_5}}", dados.emailTitular5);
    body.replaceText("{{VINCULO_TITULAR_5}}", dados.vinculoTitular5);
    
    //=== primeiro suplente ===================================
    body.replaceText("{{NOME_SUPLENTE_1}}", dados.nomeSuplente1);
    body.replaceText("{{EMAIL_SUPLENTE_1}}", dados.emailSuplente1);
    body.replaceText("{{VINCULO_SUPLENTE_1}}", dados.vinculoSuplente1);
    //=== segundo suplente ===================================
    body.replaceText("{{NOME_SUPLENTE_2}}", dados.nomeSuplente2);
    body.replaceText("{{EMAIL_SUPLENTE_2}}", dados.emailSuplente2);
    body.replaceText("{{VINCULO_SUPLENTE_2}}", dados.vinculoSuplente2);
    //=== terceiro suplente ===================================
    body.replaceText("{{NOME_SUPLENTE_3}}", dados.nomeSuplente3);
    body.replaceText("{{EMAIL_SUPLENTE_3}}", dados.emailSuplente3);
    body.replaceText("{{VINCULO_SUPLENTE_3}}", dados.vinculoSuplente3);
    //=== quarto suplente ===================================
    body.replaceText("{{NOME_SUPLENTE_4}}", dados.nomeSuplente4);
    body.replaceText("{{EMAIL_SUPLENTE_4}}", dados.emailSuplente4);
    body.replaceText("{{VINCULO_SUPLENTE_4}}", dados.vinculoSuplente4);
    //=== quinto suplente ===================================
    body.replaceText("{{NOME_SUPLENTE_5}}", dados.nomeSuplente5);
    body.replaceText("{{EMAIL_SUPLENTE_5}}", dados.emailSuplente5);
    body.replaceText("{{VINCULO_SUPLENTE_5}}", dados.vinculoSuplente5);

    body.replaceText("{{CURSO_EXTENSAO}}", dados.cursoExtensao); 
    body.replaceText("{{DATA_EMISSAO}}", Utilities.formatDate(agora, "GMT-3", "dd/MM/yyyy"));
    body.replaceText("{{HORA_EMISSAO}}", Utilities.formatDate(agora, "GMT-3", "HH:mm"));

    doc.saveAndClose();

    const pdfBlob = copia.getAs('application/pdf');
    pdfBlob.setName(nomeArquivo + ".pdf");
    
    pastaDestino.createFile(pdfBlob);
    copia.setTrashed(true);

    return pdfBlob;

  } catch (e) {
    console.error("ERRO DETALHADO NO PDF:", e.message);
    throw e; 
  }
}