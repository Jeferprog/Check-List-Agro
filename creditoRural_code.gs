/**
 * SISTEMA DE CONSULTA DE CRÉDITO RURAL - CRESOL
 * Gerenciamento de Linhas de Crédito e Regras de Enquadramento
 * Desenvolvido com Google Apps Script
 */

// ==================== CONFIGURAÇÕES GLOBAIS ====================
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_LINHAS = SS.getSheetByName("Linhas") || SS.insertSheet("Linhas");
const SHEET_CONFIG = SS.getSheetByName("Configurações") || SS.insertSheet("Configurações");
const SHEET_HISTORICO = SS.getSheetByName("Histórico") || SS.insertSheet("Histórico");

// ==================== INICIALIZAÇÃO DO SISTEMA ====================

function inicializarSistema() {
  inicializarSheetLinhas();
  inicializarSheetConfig();
  inicializarSheetHistorico();
  Logger.log("✓ Sistema inicializado com sucesso");
}

function inicializarSheetLinhas() {
  SHEET_LINHAS.clear();
  const headers = [
    "ID", "Nome Linha", "Órgão/Instituição", "Finalidade Principal",
    "Finalidades (tags)", "Enquadramento (Renda Min/Max)", "Taxa Mín (%)",
    "Taxa Máx (%)", "Prazo (meses)", "Carência (meses)", "Limite Min (R$)",
    "Limite Máx (R$)", "Requisitos", "Documentos Necessários",
    "Status (Ativa/Inativa)", "Data Atualização", "Observações",
    "Itens Financiáveis"
  ];

  SHEET_LINHAS.appendRow(headers);
  SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46");
  SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontColor("white");

  // Dados base das linhas
  const linhas = [
    ["L001", "PRONAF B (Microcrédito)", "Banco do Brasil/Caixa", "Microcrédito para pequenos produtores",
      "custeio,investimento", "Sem limite/R$ 500 mil", "0.5", "2.5", "24", "0", "1000", "25000",
      "DAP ativa, agricultor familiar", "DAP, RG, CPF, comprovante renda", "Ativa", new Date(), "Modalidade mais acessível",
      "Ferramentas e pequenos equipamentos; pequenos animais; insumos; pequenas agroindústrias; artesanato e atividades de geração de renda"],

    ["L002", "PRONAF Custeio", "Banco do Brasil/Caixa/BB", "Despesas do ciclo produtivo",
      "custeio", "Sem limite/R$ 500 mil", "2", "6", "12", "0", "5000", "500000",
      "DAP ativa, agricultor familiar", "DAP, RG, CPF, projeto técnico", "Ativa", new Date(), "Safra até 12 meses",
      "Sementes e mudas; fertilizantes e corretivos; defensivos; combustível; mão de obra; ração e insumos pecuários (vacinas, sais minerais); tratos culturais e colheita; antecipação de insumos"],

    ["L003", "PRONAF Investimento", "Banco do Brasil/Caixa/BB", "Máquinas, equipamentos, infraestrutura",
      "investimento,equipamento", "Sem limite/R$ 500 mil", "5", "5", "120", "36", "10000", "500000",
      "DAP ativa, agricultor familiar", "DAP, RG, CPF, projeto técnico, 3 últimos balanços", "Ativa", new Date(), "Carência de 3 anos",
      "Tratores, máquinas e implementos; benfeitorias e infraestrutura de produção; animais de produção; conectividade no campo (antenas, roteadores); prevenção e combate a incêndios"],

    ["L004", "PRONAF Agroecologia", "Banco do Brasil/Caixa", "Transição para sistemas sustentáveis",
      "investimento,agroecologia", "Sem limite/R$ 500 mil", "0.5", "0.5", "120", "36", "5000", "300000",
      "DAP ativa, agricultor familiar, projeto agroecológico", "DAP, RG, CPF, projeto agroecológico", "Ativa", new Date(), "Menor taxa - foco sustentabilidade",
      "Implantação de sistemas agroecológicos/orgânicos; insumos biológicos; adubação verde; certificação orgânica; recuperação de solo; biofábricas"],

    ["L005", "PRONAF Irrigação", "Banco do Brasil/Caixa", "Sistemas de irrigação eficiente",
      "investimento,irrigacao", "Sem limite/R$ 500 mil", "4", "4", "120", "36", "15000", "400000",
      "DAP ativa, agricultor familiar, projeto irrigação", "DAP, projeto técnico irrigação, orçamentos", "Ativa", new Date(), "Para modernizar irrigação",
      "Sistemas de irrigação (gotejamento, aspersão, pivô); motobombas; tubulações; reservatórios; automação hídrica"],

    ["L006", "PRONAF Mulher", "Banco do Brasil/Caixa", "Mulheres produtoras rurais",
      "investimento,custeio", "Sem limite/R$ 500 mil", "2", "5", "120", "36", "10000", "400000",
      "DAP ativa, mulher, agricultor familiar", "DAP, RG, CPF, projeto técnico", "Ativa", new Date(), "Exclusivo para mulheres",
      "Custeio e investimento da atividade (insumos, máquinas, animais, infraestrutura) em projetos liderados por mulheres agricultoras"],

    ["L007", "PRONAF Jovem", "Banco do Brasil/Caixa", "Beneficiários 16-29 anos",
      "investimento,custeio", "Sem limite/R$ 500 mil", "2", "5", "120", "36", "10000", "200000",
      "DAP ativa, jovem 16-29 anos", "DAP, RG, CPF, comprovante idade", "Ativa", new Date(), "Limite menor para jovens",
      "Estruturação da atividade produtiva por jovens: máquinas, animais, infraestrutura e custeio inicial"],

    ["L008", "PRONAMP Custeio", "Banco do Brasil/Caixa/BNDES", "Despesas do ciclo produtivo - Médios produtores",
      "custeio", "R$ 500 mil/R$ 3.5 mi", "10", "10", "12", "0", "100000", "2000000",
      "Mín 80% renda agrícola, renda até R$ 3.5 mi", "RG, CPF, últimos 2 balanços, comprovante renda", "Ativa", new Date(), "Para produtores médios",
      "Sementes, fertilizantes, defensivos, corretivos; combustível; mão de obra; insumos pecuários e demais despesas do ciclo produtivo"],

    ["L009", "PRONAMP Investimento", "Banco do Brasil/Caixa/BNDES", "Máquinas, equipamentos - Médios produtores",
      "investimento,equipamento", "R$ 500 mil/R$ 3.5 mi", "9", "10", "120", "24", "200000", "3500000",
      "Mín 80% renda agrícola, renda até R$ 3.5 mi", "RG, CPF, últimos 2 balanços, projeto técnico", "Ativa", new Date(), "Limite maior que PRONAF",
      "Tratores, máquinas e equipamentos; benfeitorias; infraestrutura; modernização e ampliação da propriedade"],

    ["L010", "Moderfrota", "BNDES/Bancos Credenciados", "Aquisição de máquinas e implementos",
      "investimento,equipamento,mecanizacao", "Conforme análise", "9", "13.5", "96", "12", "50000", "5000000",
      "Médio/grande produtor, documentação completa", "RG, CPF, últimos 3 balanços, orçamentos", "Ativa", new Date(), "Específico para máquinas agrícolas",
      "Tratores; colheitadeiras; plataformas de corte; pulverizadores; plantadeiras e semeadoras; equipamentos para beneficiamento de café (novos e usados)"],

    ["L011", "Moderagro", "BNDES/Banco do Brasil/Caixa", "Modernização e produtividade",
      "investimento,modernizacao", "Conforme análise", "8", "10", "144", "36", "100000", "2200000",
      "Propriedade registrada, projeto técnico", "Documentos de propriedade, projeto técnico, orçamentos", "Ativa", new Date(), "Foco modernização geral",
      "Suprimento de água, alimentação animal e tratamento de dejetos; frigoríficos e beneficiamento; equipamentos e embarcações de pesca/aquicultura; matrizes e reprodutores"],

    ["L012", "Programa ABC", "BNDES/Banco do Brasil", "Projetos de baixa emissão de carbono",
      "investimento,sustentabilidade,carbono", "Conforme análise", "8", "10", "120", "36", "50000", "2200000",
      "Propriedade registrada, projeto baixo carbono", "Projeto técnico ABC, documentos propriedade", "Ativa", new Date(), "Pivot, plantio direto, reflorestamento",
      "Recuperação de pastagens degradadas; plantio direto; integração lavoura-pecuária-floresta (ILPF); florestas comerciais; tratamento de dejetos"],

    ["L013", "PCA (Armazéns)", "BNDES", "Construção e ampliação de silos/armazéns",
      "investimento,infraestrutura,armazenagem", "Conforme análise", "8.5", "10", "120", "24", "100000", "200000000",
      "Capacidade até 12.000 ton ou cooperativa", "Documentos propriedade, projeto técnico, orçamentos", "Ativa", new Date(), "Para infraestrutura de armazenagem",
      "Construção, ampliação e modernização de armazéns e silos; equipamentos de secagem, climatização e movimentação de grãos"],

    ["L014", "Proirriga", "BNDES", "Sistemas de irrigação eficiente (geral)",
      "investimento,irrigacao", "Conforme análise", "10.5", "10.5", "120", "24", "20000", "500000",
      "Projeto irrigação, seguro obrigatório", "Projeto técnico, orçamentos, proponente habilitado", "Ativa", new Date(), "Obrigatório contratar seguro",
      "Sistemas de irrigação e cultivo protegido (estufas); equipamentos; automação e eficiência hídrica/energética"],

    ["L015", "RenovAgro", "BNDES", "Recuperação e projetos ambientais",
      "investimento,ambiental,sustentabilidade", "Conforme análise", "9", "9", "120", "24", "50000", "1500000",
      "Propriedade registrada, projeto ambiental", "Documentos propriedade, projeto técnico ambiental", "Ativa", new Date(), "Inclui prevenção de incêndios",
      "Recuperação de áreas degradadas; ILPF; energia renovável; tratamento de resíduos; recomposição ambiental; prevenção e combate a incêndios"],

    ["L016", "Funcafé", "BNDES", "Específica para cafeicultores",
      "investimento,custeio,cafe", "Conforme análise", "8.5", "8.5", "120", "36", "100000", "3000000",
      "Propriedade com café, projeto técnico", "Documentos propriedade, projeto técnico, comprovante atividade", "Ativa", new Date(), "Exclusiva para cafeicultura",
      "Custeio do cafezal e colheita; estocagem e aquisição de café; recuperação de cafezais (arranquio, decote, recepa); máquinas e beneficiamento"],

    ["L017", "Agricultura Empresarial", "Bancos Credenciados", "Custeio geral (Grande produtor)",
      "custeio", "Acima de R$ 3.5 mi", "14", "14", "12", "0", "500000", "999999999",
      "Renda acima R$ 3.5 milhões", "Documentos completos, últimos 3 balanços", "Ativa", new Date(), "Maior risco = maior taxa",
      "Custeio e investimento de grande porte: insumos, máquinas, infraestrutura e modernização, conforme análise da instituição"]
  ];

  linhas.forEach(linha => SHEET_LINHAS.appendRow(linha));

  // Formatar sheet
  SHEET_LINHAS.setColumnWidths(1, 18, 80);
  SHEET_LINHAS.getRange("O:O").setHorizontalAlignment("center");
}

function inicializarSheetConfig() {
  SHEET_CONFIG.clear();
  const headers = ["Parâmetro", "Valor", "Tipo", "Descrição"];
  SHEET_CONFIG.appendRow(headers);
  SHEET_CONFIG.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#005c46").setFontColor("white");

  const configs = [
    ["Versão Sistema", "1.0", "text", "Versão atual do sistema"],
    ["Data Atualização", new Date(), "date", "Última atualização de linhas"],
    ["Juros Base (Selic)", "15%", "percent", "Taxa Selic de referência"],
    ["Ano Fiscal", "2025/2026", "text", "Plano Safra atual"],
    ["Contato Suporte", "agro@cresol.com.br", "email", "E-mail para suporte"],
    ["Últimas Linhas Adicionadas", "Pronaf Irrigação", "text", "Informação das novas linhas"]
  ];

  configs.forEach(config => SHEET_CONFIG.appendRow(config));
  SHEET_CONFIG.setColumnWidth(1, 200);
  SHEET_CONFIG.setColumnWidth(2, 200);
}

function inicializarSheetHistorico() {
  SHEET_HISTORICO.clear();
  const headers = ["Data", "Tipo Consulta", "Produto/Finalidade", "Enquadramento", "Resultado", "Usuário"];
  SHEET_HISTORICO.appendRow(headers);
  SHEET_HISTORICO.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#005c46").setFontColor("white");
}

// ==================== MOTOR DE BUSCA E REGRAS ====================

function buscarLinhas(parametros) {
  try {
    const dados = SHEET_LINHAS.getDataRange().getValues();
    if (!dados || dados.length <= 1) return [];

    const headers = dados[0];
    const linhas = dados.slice(1);

    const resultado = linhas
      .filter(linha => {
        try {
          const statusIdx = headers.indexOf("Status (Ativa/Inativa)");
          if (statusIdx === -1 || linha[statusIdx] !== "Ativa") return false;

          const enquadramentoIdx = headers.indexOf("Enquadramento (Renda Min/Max)");
          if (enquadramentoIdx === -1) return false;

          if (!validarRenda(parametros.renda, linha[enquadramentoIdx])) return false;

          const finalidadesIdx = headers.indexOf("Finalidades (tags)");
          if (finalidadesIdx === -1) return false;

          if (!validarFinalidade(parametros.finalidade, linha[finalidadesIdx])) return false;

          // Filtro opcional por palavra-chave (Produto/Finalidade digitado)
          if (!validarProduto(parametros.produto, linha, headers)) return false;

          return true;
        } catch (e) {
          return false;
        }
      })
      .map(linha => {
        try {
          return {
            id: linha[headers.indexOf("ID")] || "",
            nome: linha[headers.indexOf("Nome Linha")] || "Sem nome",
            orgao: linha[headers.indexOf("Órgão/Instituição")] || "",
            finalidade: linha[headers.indexOf("Finalidade Principal")] || "",
            taxaMin: parseFloat(linha[headers.indexOf("Taxa Mín (%)")]) || 0,
            taxaMax: parseFloat(linha[headers.indexOf("Taxa Máx (%)")]) || 0,
            prazo: parseInt(linha[headers.indexOf("Prazo (meses)")]) || 0,
            carencia: parseInt(linha[headers.indexOf("Carência (meses)")]) || 0,
            limiteMin: parseInt(linha[headers.indexOf("Limite Min (R$)")]) || 0,
            limiteMax: parseInt(linha[headers.indexOf("Limite Máx (R$)")]) || 0,
            requisitos: linha[headers.indexOf("Requisitos")] || "",
            documentos: linha[headers.indexOf("Documentos Necessários")] || "",
            observacoes: linha[headers.indexOf("Observações")] || "",
            itensFinanciaveis: linha[headers.indexOf("Itens Financiáveis")] || ""
          };
        } catch (e) {
          return null;
        }
      })
      .filter(item => item !== null);

    registrarConsulta(parametros, resultado.length);
    return resultado;
  } catch (e) {
    Logger.log("Erro em buscarLinhas: " + e);
    return [];
  }
}

function validarRenda(renda, enquadramentoTexto) {
  try {
    if (!enquadramentoTexto || enquadramentoTexto === "") return true;
    if (typeof enquadramentoTexto !== "string") return true;

    if (enquadramentoTexto.includes("Conforme análise")) return true;

    const partes = enquadramentoTexto.split("/");
    if (partes.length < 2) return true;

    const minTexto = partes[0].trim();
    const maxTexto = partes[1].trim();

    if (minTexto === "Sem limite") {
      const max = parseInt(maxTexto.replace(/\D/g, "")) || Infinity;
      return renda <= max;
    }

    const min = parseInt(minTexto.replace(/\D/g, "")) || 0;
    const max = parseInt(maxTexto.replace(/\D/g, "")) || Infinity;

    return renda >= min && renda <= max;
  } catch (e) {
    return true;
  }
}

function validarProduto(produtoBuscado, linha, headers) {
  /**
   * Filtro opcional por palavra-chave. Se o colaborador digitou um
   * produto/finalidade (ex: "trator", "café", "silo"), a linha só passa se
   * o termo aparecer em algum campo relevante. Se vazio, não filtra nada.
   */
  try {
    if (!produtoBuscado || typeof produtoBuscado !== "string") return true;

    // Normaliza: minúsculas e remove acentos (ex: "café" -> "cafe")
    const normalizar = txt => String(txt || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");

    const termo = normalizar(produtoBuscado.trim());
    if (termo === "") return true;

    const camposRelevantes = [
      "Nome Linha", "Finalidade Principal", "Finalidades (tags)",
      "Itens Financiáveis", "Documentos Necessários", "Observações"
    ];
    const textoBusca = normalizar(camposRelevantes
      .map(c => {
        const idx = headers.indexOf(c);
        return idx === -1 ? "" : String(linha[idx] || "");
      })
      .join(" "));

    // Quebra a busca em palavras: basta uma palavra casar para aceitar
    const palavras = termo.split(/\s+/).filter(p => p.length >= 3);
    if (palavras.length === 0) return true;

    return palavras.some(p => textoBusca.includes(p));
  } catch (e) {
    return true;
  }
}

function validarFinalidade(finalidadeBuscada, finalidadeLinha) {
  try {
    if (!finalidadeBuscada || !finalidadeLinha) return true;
    if (typeof finalidadeBuscada !== "string" || typeof finalidadeLinha !== "string") return true;

    const tags = finalidadeLinha.toLowerCase().split(",").map(t => t.trim()).filter(t => t);
    const buscaTermos = finalidadeBuscada.toLowerCase().split(",").map(t => t.trim()).filter(t => t);

    if (tags.length === 0 || buscaTermos.length === 0) return true;

    return buscaTermos.some(termo => tags.some(tag => tag.includes(termo) || termo.includes(tag)));
  } catch (e) {
    return true;
  }
}

function registrarConsulta(parametros, qtdResultados) {
  try {
    if (!SHEET_HISTORICO) {
      Logger.log("SHEET_HISTORICO não inicializado");
      return;
    }

    const dataHora = new Date().toLocaleString('pt-BR');
    const finalidade = parametros.finalidade || "Não especificado";
    const enquadramento = parametros.enquadramento || "Não especificado";
    const resultado = `${qtdResultados} linha(s) encontrada(s)`;

    SHEET_HISTORICO.appendRow([
      dataHora,
      "Consulta Linha",
      finalidade,
      enquadramento,
      resultado,
      Session.getActiveUser().getEmail()
    ]);

    Logger.log("Consulta registrada: " + finalidade + " - " + resultado);
  } catch (e) {
    Logger.log("Erro ao registrar consulta: " + e.toString());
  }
}

function obterHistorico() {
  try {
    const dados = SHEET_HISTORICO.getDataRange().getValues();
    if (!dados || dados.length <= 1) return [];
    return dados.slice(1);
  } catch (e) {
    Logger.log("Erro ao obter histórico: " + e);
    return [];
  }
}

// ==================== FUNÇÕES ADMINISTRATIVAS ====================

function atualizarLinha(idLinha, novosDados) {
  /**
   * Atualiza dados de uma linha existente
   */
  const dados = SHEET_LINHAS.getDataRange().getValues();
  const headers = dados[0];
  const idIdx = headers.indexOf("ID");

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][idIdx] === idLinha) {
      for (const [chave, valor] of Object.entries(novosDados)) {
        const colIdx = headers.indexOf(chave);
        if (colIdx !== -1) {
          SHEET_LINHAS.getRange(i + 1, colIdx + 1).setValue(valor);
        }
      }
      SHEET_LINHAS.getRange(i + 1, headers.indexOf("Data Atualização") + 1).setValue(new Date());
      Logger.log(`✓ Linha ${idLinha} atualizada`);
      return true;
    }
  }
  return false;
}

function ativarDesativarLinha(idLinha, ativo) {
  /**
   * Ativa ou desativa uma linha de crédito.
   * Retorna o novo status para confirmação ao cliente.
   */
  const novoStatus = ativo ? "Ativa" : "Inativa";
  const sucesso = atualizarLinha(idLinha, {
    "Status (Ativa/Inativa)": novoStatus
  });
  return { sucesso: sucesso, status: novoStatus, id: idLinha };
}

function obterLinhaCompleta(idLinha) {
  try {
    const dados = SHEET_LINHAS.getDataRange().getValues();
    if (!dados || dados.length <= 1) return {};

    const headers = dados[0];
    const idIdx = headers.indexOf("ID");

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][idIdx] === idLinha) {
        const linha = {};
        headers.forEach((header, idx) => {
          linha[header] = sanitizarValor(dados[i][idx]);
        });
        return linha;
      }
    }
    return {};
  } catch (e) {
    Logger.log("Erro em obterLinhaCompleta: " + e);
    return {};
  }
}

/**
 * Converte qualquer valor (incluindo Date) em string/número simples,
 * evitando problemas de serialização do google.script.run.
 */
function sanitizarValor(valor) {
  if (valor === null || valor === undefined) return "";
  if (valor instanceof Date) return valor.toLocaleDateString('pt-BR');
  return valor;
}

function listarTodasAsLinhas() {
  /**
   * Retorna TODAS as linhas com todos os campos sanitizados em uma única
   * chamada, evitando segundo round-trip ao servidor e erros de serialização.
   */
  try {
    if (!SHEET_LINHAS) return [];

    const dados = SHEET_LINHAS.getDataRange().getValues();
    if (!dados || dados.length <= 1) return [];

    const headers = dados[0];
    if (!headers) return [];

    const resultado = [];
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      resultado.push({
        id: sanitizarValor(linha[headers.indexOf("ID")]),
        nome: sanitizarValor(linha[headers.indexOf("Nome Linha")]),
        orgao: sanitizarValor(linha[headers.indexOf("Órgão/Instituição")]),
        finalidadePrincipal: sanitizarValor(linha[headers.indexOf("Finalidade Principal")]),
        finalidades: sanitizarValor(linha[headers.indexOf("Finalidades (tags)")]),
        enquadramento: sanitizarValor(linha[headers.indexOf("Enquadramento (Renda Min/Max)")]),
        taxaMin: sanitizarValor(linha[headers.indexOf("Taxa Mín (%)")]),
        taxaMax: sanitizarValor(linha[headers.indexOf("Taxa Máx (%)")]),
        prazo: sanitizarValor(linha[headers.indexOf("Prazo (meses)")]),
        carencia: sanitizarValor(linha[headers.indexOf("Carência (meses)")]),
        limiteMin: sanitizarValor(linha[headers.indexOf("Limite Min (R$)")]),
        limiteMax: sanitizarValor(linha[headers.indexOf("Limite Máx (R$)")]),
        requisitos: sanitizarValor(linha[headers.indexOf("Requisitos")]),
        documentos: sanitizarValor(linha[headers.indexOf("Documentos Necessários")]),
        status: sanitizarValor(linha[headers.indexOf("Status (Ativa/Inativa)")]) || "Ativa",
        observacoes: sanitizarValor(linha[headers.indexOf("Observações")]),
        itensFinanciaveis: sanitizarValor(linha[headers.indexOf("Itens Financiáveis")])
      });
    }

    Logger.log("Retornando " + resultado.length + " linhas (versão completa)");
    return resultado;
  } catch (e) {
    Logger.log("Erro em listarTodasAsLinhas: " + e.toString());
    return [];
  }
}

/**
 * Adiciona uma nova linha de crédito gerando automaticamente o próximo ID.
 */
function adicionarLinha(dados) {
  try {
    const valores = SHEET_LINHAS.getDataRange().getValues();
    const headers = valores[0];

    // Gerar próximo ID no formato L0XX
    let maxNum = 0;
    for (let i = 1; i < valores.length; i++) {
      const id = String(valores[i][0] || "");
      const num = parseInt(id.replace(/\D/g, "")) || 0;
      if (num > maxNum) maxNum = num;
    }
    const novoId = "L" + String(maxNum + 1).padStart(3, "0");

    // Montar a linha na ordem correta dos headers
    const mapa = {
      "ID": novoId,
      "Nome Linha": dados.nome || "",
      "Órgão/Instituição": dados.orgao || "",
      "Finalidade Principal": dados.finalidadePrincipal || "",
      "Finalidades (tags)": dados.finalidades || "",
      "Enquadramento (Renda Min/Max)": dados.enquadramento || "Conforme análise",
      "Taxa Mín (%)": dados.taxaMin || "0",
      "Taxa Máx (%)": dados.taxaMax || "0",
      "Prazo (meses)": dados.prazo || "0",
      "Carência (meses)": dados.carencia || "0",
      "Limite Min (R$)": dados.limiteMin || "0",
      "Limite Máx (R$)": dados.limiteMax || "0",
      "Requisitos": dados.requisitos || "",
      "Documentos Necessários": dados.documentos || "",
      "Status (Ativa/Inativa)": dados.status || "Ativa",
      "Data Atualização": new Date(),
      "Observações": dados.observacoes || "",
      "Itens Financiáveis": dados.itensFinanciaveis || ""
    };

    const novaLinha = headers.map(h => (mapa[h] !== undefined ? mapa[h] : ""));
    SHEET_LINHAS.appendRow(novaLinha);

    Logger.log("✓ Nova linha adicionada: " + novoId);
    return { sucesso: true, id: novoId };
  } catch (e) {
    Logger.log("Erro em adicionarLinha: " + e.toString());
    return { sucesso: false, erro: e.toString() };
  }
}

// ==================== INTERFACE WEB ====================

function doGet() {
  return obterHTML();
}

function obterHTML() {
  const dataAtualizacao = new Date().toLocaleDateString('pt-BR');
  return HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sistema de Crédito Rural - CRESOL</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root {
  --primary-green: #005c46;
  --primary-green-dark: #004736;
  --accent-orange: #f58220;
  --accent-orange-dark: #d96f12;
  --bg-body: #eef1f0;
  --bg-card: #ffffff;
  --bg-soft: #f5f6f6;
  --text-dark: #2b2b2b;
  --text-muted: #666666;
  --border: #e0e0e0;
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Roboto', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: var(--bg-body); color: var(--text-dark); min-height: 100vh; padding: 20px; }
.container { max-width: 1250px; margin: 0 auto; background: var(--bg-card); border-radius: 10px; box-shadow: 0 6px 24px rgba(0,0,0,0.12); overflow: hidden; }
.header { background: var(--primary-green); color: white; padding: 30px; text-align: center; }
.header h1 { font-size: 28px; font-weight: 700; margin-bottom: 5px; }
.header p { font-size: 14px; font-weight: 400; opacity: 0.92; }
.header-badge { display: inline-block; background: var(--accent-orange); color: white; font-size: 12px; font-weight: 700; padding: 4px 14px; border-radius: 20px; margin-top: 12px; letter-spacing: 0.3px; }
.header-update { display: block; font-size: 12px; opacity: 0.8; margin-top: 8px; }
.tabs { display: flex; background: var(--bg-soft); border-bottom: 2px solid var(--border); }
.tab-btn { flex: 1; padding: 15px; background: none; border: none; cursor: pointer; font-size: 14px; font-weight: 500; font-family: inherit; color: var(--text-muted); transition: all 0.25s; border-bottom: 3px solid transparent; }
.tab-btn:hover { color: var(--primary-green); background: var(--bg-soft); box-shadow: none; transform: none; }
.tab-btn.active { color: var(--primary-green); border-bottom-color: var(--accent-orange); background: var(--bg-card); font-weight: 700; }
.tab-content { display: none; padding: 30px; }
.tab-content.active { display: block; }
.form-group { margin-bottom: 20px; }
label { display: block; margin-bottom: 8px; font-weight: 500; color: var(--text-dark); }
input, select, textarea { width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; font-family: inherit; color: var(--text-dark); }
input:focus, select:focus, textarea:focus { outline: none; border-color: var(--primary-green); box-shadow: 0 0 0 3px rgba(0, 92, 70, 0.12); }
button { background: var(--accent-orange); color: white; padding: 12px 30px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 700; font-family: inherit; transition: all 0.25s; }
button:hover { background: var(--accent-orange-dark); transform: translateY(-1px); box-shadow: 0 5px 14px rgba(245, 130, 32, 0.3); }
.resultado { background: var(--bg-soft); border-left: 4px solid var(--accent-orange); padding: 20px; margin-top: 30px; border-radius: 6px; display: none; }
.resultado.visible { display: block; }
.linha-card { background: var(--bg-card); border: 1px solid var(--border); border-left: 4px solid var(--accent-orange); border-radius: 6px; padding: 18px; margin-bottom: 15px; transition: all 0.25s; }
.linha-card:hover { box-shadow: 0 4px 14px rgba(0,0,0,0.1); }
.linha-card h3 { color: var(--primary-green); margin-bottom: 10px; font-size: 17px; font-weight: 700; }
.linha-info { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 10px; }
.info-item { font-size: 13px; }
.info-label { font-weight: 700; color: var(--text-muted); }
.info-value { color: var(--text-dark); margin-top: 3px; }
.alert { padding: 15px; border-radius: 6px; margin-bottom: 20px; }
.alert-info { background: #e6f2ee; color: var(--primary-green-dark); border-left: 4px solid var(--primary-green); }
.loading { text-align: center; padding: 20px; display: none; }
.spinner { border: 3px solid #e4e7e6; border-top: 3px solid var(--accent-orange); border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 0 auto; }
@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
.grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
table thead th { background: var(--primary-green) !important; color: white !important; border-bottom: none !important; }
table tbody tr:nth-child(even) { background: #f7f8f8; }
@media (max-width: 768px) { .grid-2 { grid-template-columns: 1fr; } .linha-info { grid-template-columns: 1fr; } .header h1 { font-size: 22px; } .tab-btn { font-size: 12px; padding: 12px 8px; } .tab-content { padding: 20px; } }
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>🌾 Sistema de Crédito Rural</h1>
<p>Consultar linhas de crédito para operações rurais - CRESOL</p>
<span class="header-badge">Plano Safra 2025/2026</span>
<span class="header-update">Última atualização: ${dataAtualizacao}</span>
</div>

<div class="tabs">
<button class="tab-btn active" id="btn-consulta" onclick="window.mudarAba(event, 'consulta')">Consultar</button>
<button class="tab-btn" id="btn-admin" onclick="window.mudarAba(event, 'admin')">Administrativo</button>
<button class="tab-btn" id="btn-historico" onclick="window.mudarAba(event, 'historico')">Histórico</button>
</div>

<div id="consulta" class="tab-content active">
<div class="alert alert-info">
<strong>ℹ️ Como usar:</strong> Preencha os campos com dados do associado para encontrar linhas disponíveis.
</div>
<div class="form-group">
<label>📦 Produto/Finalidade <span style="font-weight: 400; color: #888; font-size: 12px;">(opcional - refina a busca)</span></label>
<input type="text" id="produto" placeholder="Ex: trator, café, silo, irrigação...">
<small style="color: #666; display: block; margin-top: 5px;">
Digite uma palavra-chave para filtrar as linhas que mencionam esse produto. Deixe em branco para ver todas as linhas do tipo selecionado.
</small>
</div>
<div class="form-group">
<label>💰 Renda Bruta Anual (R$)</label>
<input type="number" id="renda" placeholder="Ex: 150000" min="0" oninput="window.atualizarEnquadramento()">
</div>
<div class="form-group">
<label>👤 Enquadramento Detectado</label>
<input type="text" id="enquadramento" readonly style="background-color: #f0f0f0; cursor: not-allowed;" placeholder="Preenchido automaticamente pela renda">
<small style="color: #666; display: block; margin-top: 5px;">
Até R$ 500 mil = PRONAF | R$ 500k a R$ 3,5M = PRONAMP | Acima R$ 3,5M = Agricultura Empresarial
</small>
</div>
<div class="form-group">
<label>🎯 Tipo de Operação</label>
<select id="finalidade" onchange="window.atualizarProdutosDisponiveis()">
<option value="">-- Selecione --</option>
<option value="custeio">Custeio (Despesas do ciclo)</option>
<option value="investimento">Investimento (Máquinas/Equipamentos)</option>
<option value="mecanizacao">Mecanização</option>
<option value="irrigacao">Irrigação</option>
<option value="agroecologia">Agroecologia/Sustentabilidade</option>
<option value="infraestrutura">Infraestrutura</option>
<option value="armazenagem">Armazenagem</option>
<option value="sustentabilidade">Sustentabilidade/Carbono</option>
<option value="cafe">Específico para Café</option>
</select>
</div>
<div id="produtosDisponiveis" style="display: none; background: #e6f2ee; padding: 12px; border-radius: 6px; margin-top: 10px; border-left: 4px solid #005c46; font-size: 13px;">
<strong style="color: #004736;">💡 Produtos financiáveis:</strong>
<span id="listaProdutos" style="display: block; margin-top: 5px; color: #333;"></span>
</div>
<button id="btnBuscar" onclick="window.buscar()">🔍 Buscar Linhas Disponíveis</button>
<div class="loading" id="loading">
<div class="spinner"></div>
<p style="margin-top: 10px; color: #666;">Buscando linhas...</p>
</div>
<div class="resultado" id="resultado">
<div id="resultadoConteudo"></div>
</div>
</div>

<div id="admin" class="tab-content">
<div class="alert alert-info">
<strong>⚙️ Área Administrativa:</strong> Gerenciar linhas de crédito - editar, incluir e ativar/inativar.
</div>
<div style="margin-bottom: 20px;">
<button onclick="window.abrirFormularioNovaLinha()" style="background: #28a745;">➕ Incluir Nova Linha</button>
</div>
<div id="formNovaLinha"></div>
<div id="edicaoConteudo"></div>
<div id="listaLinhasAdmin"></div>
</div>

<div id="historico" class="tab-content">
<div class="alert alert-info">
<strong>📊 Histórico de Consultas:</strong> Acompanhe todas as buscas realizadas.
</div>
<div id="tabelaHistorico"></div>
</div>
</div>

<script>
window.mudarAba = function(event, abaId) {
  document.querySelectorAll('.tab-content').forEach(aba => aba.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  document.getElementById(abaId).classList.add('active');
  event.target.classList.add('active');

  if (abaId === 'admin') {
    window.carregarLinhasAdministrativo();
  } else if (abaId === 'historico') {
    window.carregarHistorico();
  }
};

window.produtosPorTipo = {
  'custeio': ['Sementes', 'Fertilizantes', 'Defensivos', 'Combustível', 'Mão de obra', 'Aluguel de máquinas'],
  'investimento': ['Máquinas', 'Equipamentos', 'Implementos', 'Infraestrutura', 'Benfeitorias', 'Animais produtivos'],
  'mecanizacao': ['Tratores', 'Colheitadeiras', 'Plantadeiras', 'Pulverizadores', 'Implementos agrícolas'],
  'irrigacao': ['Sistemas de irrigação', 'Pivôs centrais', 'Gotejamento', 'Aspersão', 'Tubulação', 'Motores'],
  'agroecologia': ['Transição sustentável', 'Certificação orgânica', 'Sistemas agroecológicos', 'Insumos naturais'],
  'infraestrutura': ['Construções rurais', 'Galpões', 'Silos', 'Benfeitorias', 'Eletrificação rural'],
  'armazenagem': ['Silos', 'Armazéns', 'Sacos', 'Sistemas de armazenagem', 'Climatização'],
  'sustentabilidade': ['Plantio direto', 'Rotação de culturas', 'Reflorestamento', 'Recuperação ambiental', 'Energia renovável'],
  'cafe': ['Plantio de café', 'Renovação de cafezal', 'Máquinas de café', 'Sistemas de irrigação para café', 'Beneficiamento']
};

window.atualizarEnquadramento = function() {
  const renda = parseInt(document.getElementById('renda').value) || 0;
  let enquadramento = '';
  let enquadramentoDisplay = '';

  if (renda === 0) {
    enquadramento = '';
    enquadramentoDisplay = '';
  } else if (renda <= 500000) {
    enquadramento = 'pronaf';
    enquadramentoDisplay = '🌾 PRONAF (Agricultura Familiar)';
  } else if (renda <= 3500000) {
    enquadramento = 'pronamp';
    enquadramentoDisplay = '🌱 PRONAMP (Médio Produtor)';
  } else {
    enquadramento = 'empresarial';
    enquadramentoDisplay = '🏢 Agricultura Empresarial';
  }

  document.getElementById('enquadramento').value = enquadramentoDisplay;
  document.getElementById('enquadramento').dataset.value = enquadramento;
};

window.atualizarProdutosDisponiveis = function() {
  const tipoOperacao = document.getElementById('finalidade').value;
  const produtosDiv = document.getElementById('produtosDisponiveis');
  const listaProdutos = document.getElementById('listaProdutos');

  if (!tipoOperacao || !window.produtosPorTipo[tipoOperacao]) {
    produtosDiv.style.display = 'none';
    return;
  }

  const produtos = window.produtosPorTipo[tipoOperacao];
  listaProdutos.textContent = produtos.join(' • ');
  produtosDiv.style.display = 'block';
};

window.buscar = function() {
  const produto = document.getElementById('produto').value;
  const renda = parseInt(document.getElementById('renda').value) || 0;
  const enquadramento = document.getElementById('enquadramento').dataset.value || '';
  const finalidade = document.getElementById('finalidade').value;

  if (!renda || !enquadramento || !finalidade) {
    alert('Por favor, preencha todos os campos obrigatórios');
    return;
  }

  document.getElementById('loading').style.display = 'block';
  document.getElementById('resultado').classList.remove('visible');

  google.script.run.withSuccessHandler(window.mostrarResultados)
    .withFailureHandler(function(error) {
      alert('Erro na busca: ' + error);
      document.getElementById('loading').style.display = 'none';
    })
    .buscarLinhas({
      produto: produto,
      enquadramento: enquadramento,
      renda: renda,
      finalidade: finalidade
    });
};

window.mostrarResultados = function(linhas) {
  document.getElementById('loading').style.display = 'none';
  let html = '<h2 style="margin-bottom: 20px; color: #005c46;">Linhas Disponíveis (' + linhas.length + ')</h2>';

  if (linhas.length === 0) {
    html += '<div class="alert alert-info">Nenhuma linha encontrada. Ajuste os parâmetros.</div>';
  } else {
    linhas.forEach(linha => {
      html += '<div class="linha-card"><h3>' + linha.nome + '</h3>';
      html += '<p style="font-size: 12px; color: #999;">Instituição: ' + linha.orgao + '</p>';
      html += '<div class="linha-info">';
      html += '<div class="info-item"><span class="info-label">Taxa:</span><span class="info-value">' + linha.taxaMin + '% - ' + linha.taxaMax + '% a.a.</span></div>';
      html += '<div class="info-item"><span class="info-label">Prazo:</span><span class="info-value">Até ' + linha.prazo + ' meses</span></div>';
      html += '<div class="info-item"><span class="info-label">Carência:</span><span class="info-value">' + linha.carencia + ' meses</span></div>';
      html += '<div class="info-item"><span class="info-label">Limite:</span><span class="info-value">R$ ' + window.formatarMoeda(linha.limiteMin) + ' a R$ ' + window.formatarMoeda(linha.limiteMax) + '</span></div>';
      html += '<div class="info-item"><span class="info-label">Requisitos:</span><span class="info-value">' + linha.requisitos + '</span></div>';
      html += '<div class="info-item"><span class="info-label">Documentos:</span><span class="info-value">' + linha.documentos + '</span></div>';
      html += '</div>';
      if (linha.itensFinanciaveis) {
        html += '<div class="itens-financiaveis" style="margin-top: 12px; background: #eef6ee; border-left: 4px solid #28a745; padding: 10px 12px; border-radius: 4px;">' +
          '<span style="font-weight: 600; color: #1f6b1f; font-size: 13px;">✅ O que pode ser financiado:</span>' +
          '<div style="color: #333; font-size: 13px; margin-top: 4px;">' + linha.itensFinanciaveis + '</div>' +
          '</div>';
      }
      html += '<div style="margin-top: 15px; display: flex; gap: 10px;">';
      html += '<button onclick="window.abrirSimulador(' + "'" + linha.nome + "'" + ', ' + linha.taxaMin + ', ' + linha.prazo + ', ' + linha.carencia + ')" style="background: #005c46; color: white; padding: 8px 15px; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">💰 Simular Parcelas</button>';
      html += '<button onclick="window.exportarPDF()" style="background: #f58220; color: white; padding: 8px 15px; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">📄 Exportar PDF</button>';
      html += '</div></div>';
    });
  }
  document.getElementById('resultadoConteudo').innerHTML = html;
  document.getElementById('resultado').classList.add('visible');
};

window.abrirSimulador = function(nomeLinha, taxaMin, prazo, carencia) {
  const modal = document.createElement('div');
  modal.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 9999;';
  modal.id = 'modalSimulador';

  let html = '<div style="background: white; padding: 30px; border-radius: 10px; max-width: 500px; width: 90%; max-height: 80vh; overflow-y: auto;">' +
    '<h3 style="margin-bottom: 20px; color: #005c46;">Simulador de Parcelas</h3>' +
    '<p style="color: #666; margin-bottom: 15px;"><strong>Linha:</strong> ' + nomeLinha + '</p>' +
    '<p style="color: #666; margin-bottom: 15px;"><strong>Taxa:</strong> ' + taxaMin + '% a.a. | <strong>Prazo:</strong> até ' + prazo + ' meses</p>' +
    '<div style="margin-bottom: 15px;">' +
    '<label style="display: block; font-weight: 600; margin-bottom: 5px;">Valor do Crédito (R$)</label>' +
    '<input type="number" id="sim_valor" placeholder="100000" step="1000" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px;">' +
    '</div>' +
    '<div style="margin-bottom: 15px;">' +
    '<label style="display: block; font-weight: 600; margin-bottom: 5px;">Prazo Total (meses)</label>' +
    '<input type="number" id="sim_prazo" value="' + prazo + '" min="1" max="' + prazo + '" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px;">' +
    '</div>' +
    '<div style="margin-bottom: 15px;">' +
    '<label style="display: block; font-weight: 600; margin-bottom: 5px;">Periodicidade das Parcelas</label>' +
    '<select id="sim_periodicidade" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px;">' +
    '<option value="12" selected>Anual (típico no crédito rural)</option>' +
    '<option value="6">Semestral</option>' +
    '<option value="3">Trimestral</option>' +
    '<option value="1">Mensal</option>' +
    '</select>' +
    '</div>' +
    '<div style="margin-bottom: 15px;">' +
    '<label style="display: block; font-weight: 600; margin-bottom: 5px;">Taxa Anual (%)</label>' +
    '<input type="number" id="sim_taxa" value="' + taxaMin + '" step="0.1" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px;">' +
    '</div>' +
    '<button onclick="window.calcularParcelas()" style="width: 100%; background: #28a745; color: white; padding: 10px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; margin-bottom: 15px;">Calcular</button>' +
    '<div id="resultadoSimulador"></div>' +
    '<button onclick="document.getElementById(' + "'" + 'modalSimulador' + "'" + ').remove();" style="width: 100%; background: #6c757d; color: white; padding: 10px; border: none; border-radius: 4px; cursor: pointer;">Fechar</button>' +
    '</div>';

  modal.innerHTML = html;
  document.body.appendChild(modal);
};

window.calcularParcelas = function() {
  const valor = parseFloat(document.getElementById('sim_valor').value);
  const prazo = parseInt(document.getElementById('sim_prazo').value);
  const taxa = parseFloat(document.getElementById('sim_taxa').value);
  const periodicidade = parseInt(document.getElementById('sim_periodicidade').value) || 12;

  if (!valor || valor <= 0) {
    alert('Digite um valor de crédito válido');
    return;
  }
  if (!prazo || prazo <= 0) {
    alert('Digite um prazo válido');
    return;
  }

  // Quantidade de parcelas = prazo total dividido pela periodicidade
  const numParcelas = Math.max(1, Math.floor(prazo / periodicidade));

  // Taxa nominal do período (ex: anual = taxa/100; mensal = taxa/100/12)
  const taxaPeriodo = (taxa / 100) * (periodicidade / 12);

  let parcela;
  if (taxaPeriodo === 0) {
    parcela = valor / numParcelas;
  } else {
    const fator = Math.pow(1 + taxaPeriodo, numParcelas);
    parcela = (valor * taxaPeriodo * fator) / (fator - 1);
  }

  const totalPago = parcela * numParcelas;
  const totalJuros = totalPago - valor;

  const rotuloPeriodo = { 1: 'mensais', 3: 'trimestrais', 6: 'semestrais', 12: 'anuais' }[periodicidade] || '';

  let html = '<div style="background: #f9f9f9; padding: 15px; border-radius: 4px; margin-top: 15px;">' +
    '<h4 style="color: #005c46; margin-bottom: 10px;">Resultado:</h4>' +
    '<p><strong>Valor do Crédito:</strong> R$ ' + window.formatarMoeda(valor) + '</p>' +
    '<p><strong>Quantidade de Parcelas:</strong> ' + numParcelas + ' parcela(s) ' + rotuloPeriodo + '</p>' +
    '<p><strong>Valor de Cada Parcela:</strong> R$ ' + window.formatarMoeda(parcela) + '</p>' +
    '<p><strong>Total de Juros:</strong> R$ ' + window.formatarMoeda(totalJuros) + '</p>' +
    '<p><strong>Valor Total a Pagar:</strong> R$ ' + window.formatarMoeda(totalPago) + '</p>' +
    '<p style="color: #666; font-size: 12px; margin-top: 10px;"><em>*Simulação aproximada (Tabela Price), sem considerar carência, seguros ou outras taxas.</em></p>' +
    '</div>';

  document.getElementById('resultadoSimulador').innerHTML = html;
};

window.exportarPDF = function() {
  const resultadoDiv = document.getElementById('resultadoConteudo');
  if (!resultadoDiv || resultadoDiv.innerHTML.trim() === '') {
    alert('Nenhum resultado para exportar. Realize uma busca primeiro.');
    return;
  }

  const conteudo = resultadoDiv.innerHTML;
  const dataAtual = new Date().toLocaleDateString('pt-BR');
  const horaAtual = new Date().toLocaleTimeString('pt-BR');

  const htmlPDF = '<!DOCTYPE html>' +
    '<html lang="pt-BR">' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<title>Relatório de Crédito Rural</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '<style>' +
    '* { margin: 0; padding: 0; box-sizing: border-box; }' +
    'body { font-family: "Roboto", "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; background: white; padding: 20px; }' +
    '.header { background: #005c46; color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; page-break-after: avoid; }' +
    '.header h1 { font-size: 28px; margin-bottom: 10px; }' +
    '.header p { font-size: 14px; opacity: 0.9; }' +
    '.meta { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 10px; font-size: 13px; }' +
    '.meta-item { background: rgba(255,255,255,0.12); padding: 8px 12px; border-radius: 4px; }' +
    '.content { margin-top: 20px; }' +
    '.linha-card { background: white; border: 1px solid #ddd; border-left: 4px solid #f58220; border-radius: 5px; padding: 20px; margin-bottom: 20px; page-break-inside: avoid; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }' +
    '.linha-card h3 { color: #005c46; margin-bottom: 10px; font-size: 18px; border-bottom: 2px solid #005c46; padding-bottom: 8px; }' +
    '.linha-card p { margin: 8px 0; font-size: 13px; color: #666; }' +
    '.linha-info { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 15px; }' +
    '.info-item { background: #f9f9f9; padding: 10px; border-radius: 4px; font-size: 13px; }' +
    '.info-label { font-weight: 600; color: #005c46; display: block; margin-bottom: 4px; }' +
    '.info-value { color: #333; font-size: 13px; }' +
    '.itens-financiaveis { margin-top: 12px; background: #eef6ee !important; border-left: 4px solid #28a745 !important; padding: 12px 14px; border-radius: 4px; page-break-inside: avoid; }' +
    '.itens-financiaveis span { font-weight: 600; color: #1f6b1f; font-size: 13px; display: block; margin-bottom: 4px; }' +
    '.itens-financiaveis div { color: #333; font-size: 13px; }' +
    '.footer { margin-top: 30px; padding-top: 20px; border-top: 2px solid #ddd; text-align: center; font-size: 12px; color: #999; page-break-before: avoid; }' +
    '.alert { background: #e8f4f8; color: #0c5460; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 4px solid #0c5460; }' +
    '@media print { body { padding: 0; } .header { margin-bottom: 20px; } .linha-card { margin-bottom: 15px; } @page { margin: 15mm; } }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="header">' +
    '<h1>🌾 Relatório de Crédito Rural - CRESOL</h1>' +
    '<p>Consulta de Linhas Disponíveis para Operações Rurais</p>' +
    '<div class="meta">' +
    '<div class="meta-item"><strong>Data:</strong> ' + dataAtual + '</div>' +
    '<div class="meta-item"><strong>Hora:</strong> ' + horaAtual + '</div>' +
    '</div>' +
    '</div>' +
    '<div class="content">' +
    '<div class="alert">' +
    '<strong>ℹ️ Informações Importantes:</strong> Este relatório apresenta as linhas de crédito disponíveis conforme os critérios de busca. ' +
    'Para contratar um crédito, entre em contato com a agência da CRESOL mais próxima com toda a documentação necessária.' +
    '</div>' +
    conteudo +
    '</div>' +
    '<div class="footer">' +
    '<p><strong>Sistema de Crédito Rural - CRESOL</strong></p>' +
    '<p>Relatório gerado automaticamente. As informações contidas neste documento são baseadas no Plano Safra 2025/2026.</p>' +
    '<p>Para dúvidas, contate: agro@cresol.com.br | Telefone: (54) 3025-2000</p>' +
    '</div>' +
    '<scr' + 'ipt>' +
    'window.onload = function() { setTimeout(function() { window.print(); }, 500); };' +
    '<\/scr' + 'ipt>' +
    '</body>' +
    '</html>';

  const blob = new Blob([htmlPDF], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  window.open(url, '_blank');
};

window.linhasCache = [];

window.escaparHtml = function(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
};

window.notificar = function(mensagem, tipo) {
  const cores = tipo === 'erro'
    ? { bg: '#f8d7da', fg: '#721c24' }
    : { bg: '#d4edda', fg: '#155724' };
  const msg = document.createElement('div');
  msg.style.cssText = 'position: fixed; top: 20px; right: 20px; background: ' + cores.bg + '; color: ' + cores.fg + '; padding: 15px 20px; border-radius: 5px; box-shadow: 0 2px 10px rgba(0,0,0,0.2); z-index: 10000; max-width: 320px;';
  msg.innerHTML = mensagem;
  document.body.appendChild(msg);
  setTimeout(function() { msg.remove(); }, 3500);
};

window.carregarLinhasAdministrativo = function() {
  const lista = document.getElementById('listaLinhasAdmin');
  lista.innerHTML = '<p style="color: #666; padding: 10px;">Carregando linhas...</p>';

  google.script.run
    .withSuccessHandler(function(linhas) {
      if (!linhas || !Array.isArray(linhas) || linhas.length === 0) {
        lista.innerHTML = '<p style="color: #999; padding: 10px;">Nenhuma linha cadastrada.</p>';
        return;
      }
      window.linhasCache = linhas;
      window.renderizarListaLinhas();
      console.log('Linhas carregadas: ' + linhas.length);
    })
    .withFailureHandler(function(error) {
      console.error('Erro ao carregar linhas:', error);
      lista.innerHTML = '<p style="color: red; padding: 10px;">Erro ao carregar linhas: ' + error + '</p>';
    })
    .listarTodasAsLinhas();
};

window.renderizarListaLinhas = function() {
  const linhas = window.linhasCache;
  let html = '<h3 style="margin: 25px 0 15px; color: #005c46;">Linhas Cadastradas (' + linhas.length + ')</h3>';
  html += '<table style="width: 100%; border-collapse: collapse; font-size: 13px;">' +
    '<thead style="background: #f5f5f5;"><tr>' +
    '<th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Nome</th>' +
    '<th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Instituição</th>' +
    '<th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd;">Status</th>' +
    '<th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd;">Ações</th>' +
    '</tr></thead><tbody>';

  linhas.forEach(function(linha) {
    const ativa = linha.status === 'Ativa';
    const badge = ativa
      ? '<span style="background: #d4edda; color: #155724; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600;">● Ativa</span>'
      : '<span style="background: #f8d7da; color: #721c24; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600;">● Inativa</span>';

    const btnToggle = ativa
      ? '<button onclick="window.alternarStatusLinha(' + "'" + linha.id + "'" + ', false)" style="background: #ffc107; color: #333; padding: 6px 12px; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">⏸ Inativar</button>'
      : '<button onclick="window.alternarStatusLinha(' + "'" + linha.id + "'" + ', true)" style="background: #28a745; color: white; padding: 6px 12px; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">▶ Ativar</button>';

    html += '<tr style="border-bottom: 1px solid #eee;">' +
      '<td style="padding: 10px;">' + window.escaparHtml(linha.nome) + '</td>' +
      '<td style="padding: 10px; color: #666;">' + window.escaparHtml(linha.orgao) + '</td>' +
      '<td style="padding: 10px; text-align: center;">' + badge + '</td>' +
      '<td style="padding: 10px; text-align: center; white-space: nowrap;">' +
      '<button onclick="window.editarLinha(' + "'" + linha.id + "'" + ')" style="background: #005c46; color: white; padding: 6px 12px; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; margin-right: 5px;">✏️ Editar</button>' +
      btnToggle +
      '</td></tr>';
  });

  html += '</tbody></table>';
  document.getElementById('listaLinhasAdmin').innerHTML = html;
};

window.alternarStatusLinha = function(idLinha, ativar) {
  google.script.run
    .withSuccessHandler(function(resp) {
      const linha = window.linhasCache.find(function(l) { return l.id === idLinha; });
      if (linha) linha.status = ativar ? 'Ativa' : 'Inativa';
      window.renderizarListaLinhas();
      window.notificar('<strong>✓ Sucesso!</strong><br>Linha ' + (ativar ? 'ativada' : 'inativada') + '.');
    })
    .withFailureHandler(function(error) {
      window.notificar('<strong>✕ Erro:</strong><br>' + error, 'erro');
    })
    .ativarDesativarLinha(idLinha, ativar);
};

window.editarLinha = function(idLinha) {
  const linha = window.linhasCache.find(function(l) { return l.id === idLinha; });
  if (!linha) {
    window.notificar('<strong>✕ Erro:</strong><br>Linha não encontrada.', 'erro');
    return;
  }
  window.renderizarFormularioLinha(linha, false);
};

window.abrirFormularioNovaLinha = function() {
  const vazia = {
    id: '', nome: '', orgao: '', finalidadePrincipal: '', finalidades: '',
    enquadramento: 'Conforme análise', taxaMin: '', taxaMax: '', prazo: '',
    carencia: '', limiteMin: '', limiteMax: '', requisitos: '', documentos: '',
    status: 'Ativa', observacoes: '', itensFinanciaveis: ''
  };
  window.renderizarFormularioLinha(vazia, true);
};

/**
 * Renderiza o formulário de edição/inclusão. Se novaLinha=true, é inclusão.
 */
window.renderizarFormularioLinha = function(linha, novaLinha) {
  const e = window.escaparHtml;
  const containerId = novaLinha ? 'formNovaLinha' : 'edicaoConteudo';
  const outroId = novaLinha ? 'edicaoConteudo' : 'formNovaLinha';
  document.getElementById(outroId).innerHTML = '';

  const titulo = novaLinha ? '➕ Incluir Nova Linha' : '✏️ Editar: ' + e(linha.nome);

  const campo = function(id, label, valor, tipo, step) {
    const t = tipo || 'text';
    const s = step ? ' step="' + step + '"' : '';
    return '<div><label style="display: block; font-weight: 600; margin-bottom: 5px; font-size: 13px;">' + label + '</label>' +
      '<input type="' + t + '" id="' + id + '"' + s + ' value="' + e(valor) + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;"></div>';
  };
  const area = function(id, label, valor) {
    return '<div style="margin-top: 15px;"><label style="display: block; font-weight: 600; margin-bottom: 5px; font-size: 13px;">' + label + '</label>' +
      '<textarea id="' + id + '" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; height: 60px;">' + e(valor) + '</textarea></div>';
  };

  let html = '<div style="margin-top: 20px; border: 2px solid #005c46; padding: 20px; border-radius: 5px; background: #fafbff;">' +
    '<h3 style="margin-bottom: 20px; color: #005c46;">' + titulo + '</h3>' +
    '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">';

  html += campo('edit_nome', 'Nome Linha', linha.nome);
  html += campo('edit_orgao', 'Instituição', linha.orgao);
  html += campo('edit_finalidade_principal', 'Finalidade Principal', linha.finalidadePrincipal);
  html += campo('edit_finalidades', 'Finalidades (tags, separadas por vírgula)', linha.finalidades);
  html += campo('edit_enquadramento', 'Enquadramento (ex: Sem limite/R$ 500 mil)', linha.enquadramento);
  html += campo('edit_taxa_min', 'Taxa Mín (%)', linha.taxaMin, 'number', '0.1');
  html += campo('edit_taxa_max', 'Taxa Máx (%)', linha.taxaMax, 'number', '0.1');
  html += campo('edit_prazo', 'Prazo (meses)', linha.prazo, 'number');
  html += campo('edit_carencia', 'Carência (meses)', linha.carencia, 'number');
  html += campo('edit_limite_min', 'Limite Min (R$)', linha.limiteMin, 'number');
  html += campo('edit_limite_max', 'Limite Máx (R$)', linha.limiteMax, 'number');

  html += '<div><label style="display: block; font-weight: 600; margin-bottom: 5px; font-size: 13px;">Status</label>' +
    '<select id="edit_status" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">' +
    '<option value="Ativa"' + (linha.status === 'Ativa' ? ' selected' : '') + '>Ativa</option>' +
    '<option value="Inativa"' + (linha.status === 'Inativa' ? ' selected' : '') + '>Inativa</option>' +
    '</select></div>';

  html += '</div>';

  html += area('edit_itens_financiaveis', 'Itens Financiáveis (o que pode ser financiado)', linha.itensFinanciaveis);
  html += area('edit_documentos', 'Documentos Necessários', linha.documentos);
  html += area('edit_requisitos', 'Requisitos', linha.requisitos);
  html += area('edit_observacoes', 'Observações', linha.observacoes);

  const acao = novaLinha
    ? 'window.salvarNovaLinha()'
    : 'window.salvarEdicaoLinha(' + "'" + linha.id + "'" + ')';
  const rotuloSalvar = novaLinha ? '✓ Adicionar Linha' : '✓ Salvar Alterações';

  html += '<div style="margin-top: 20px; display: flex; gap: 10px;">' +
    '<button type="button" onclick="' + acao + '" style="background: #28a745; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600;">' + rotuloSalvar + '</button>' +
    '<button type="button" onclick="document.getElementById(' + "'" + containerId + "'" + ').innerHTML=' + "''" + ';" style="background: #6c757d; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600;">✕ Cancelar</button>' +
    '</div></div>';

  document.getElementById(containerId).innerHTML = html;
  document.getElementById(containerId).scrollIntoView({ behavior: 'smooth', block: 'start' });
};

window.coletarDadosFormulario = function() {
  return {
    nome: document.getElementById('edit_nome').value,
    orgao: document.getElementById('edit_orgao').value,
    finalidadePrincipal: document.getElementById('edit_finalidade_principal').value,
    finalidades: document.getElementById('edit_finalidades').value,
    enquadramento: document.getElementById('edit_enquadramento').value,
    taxaMin: document.getElementById('edit_taxa_min').value,
    taxaMax: document.getElementById('edit_taxa_max').value,
    prazo: document.getElementById('edit_prazo').value,
    carencia: document.getElementById('edit_carencia').value,
    limiteMin: document.getElementById('edit_limite_min').value,
    limiteMax: document.getElementById('edit_limite_max').value,
    documentos: document.getElementById('edit_documentos').value,
    requisitos: document.getElementById('edit_requisitos').value,
    observacoes: document.getElementById('edit_observacoes').value,
    itensFinanciaveis: document.getElementById('edit_itens_financiaveis').value,
    status: document.getElementById('edit_status').value
  };
};

window.salvarEdicaoLinha = function(idLinha) {
  const f = window.coletarDadosFormulario();
  if (!f.nome.trim()) {
    window.notificar('<strong>✕ Atenção:</strong><br>O nome da linha é obrigatório.', 'erro');
    return;
  }
  const dados = {
    'Nome Linha': f.nome,
    'Órgão/Instituição': f.orgao,
    'Finalidade Principal': f.finalidadePrincipal,
    'Finalidades (tags)': f.finalidades,
    'Enquadramento (Renda Min/Max)': f.enquadramento,
    'Taxa Mín (%)': f.taxaMin,
    'Taxa Máx (%)': f.taxaMax,
    'Prazo (meses)': f.prazo,
    'Carência (meses)': f.carencia,
    'Limite Min (R$)': f.limiteMin,
    'Limite Máx (R$)': f.limiteMax,
    'Documentos Necessários': f.documentos,
    'Requisitos': f.requisitos,
    'Observações': f.observacoes,
    'Itens Financiáveis': f.itensFinanciaveis,
    'Status (Ativa/Inativa)': f.status
  };

  google.script.run
    .withSuccessHandler(function() {
      document.getElementById('edicaoConteudo').innerHTML = '';
      window.notificar('<strong>✓ Sucesso!</strong><br>Linha atualizada com sucesso.');
      window.carregarLinhasAdministrativo();
    })
    .withFailureHandler(function(error) {
      window.notificar('<strong>✕ Erro ao salvar:</strong><br>' + error, 'erro');
    })
    .atualizarLinha(idLinha, dados);
};

window.salvarNovaLinha = function() {
  const dados = window.coletarDadosFormulario();
  if (!dados.nome.trim()) {
    window.notificar('<strong>✕ Atenção:</strong><br>O nome da linha é obrigatório.', 'erro');
    return;
  }

  google.script.run
    .withSuccessHandler(function(resp) {
      document.getElementById('formNovaLinha').innerHTML = '';
      if (resp && resp.sucesso) {
        window.notificar('<strong>✓ Sucesso!</strong><br>Nova linha adicionada (' + resp.id + ').');
      } else {
        window.notificar('<strong>✓ Linha adicionada.</strong>');
      }
      window.carregarLinhasAdministrativo();
    })
    .withFailureHandler(function(error) {
      window.notificar('<strong>✕ Erro ao adicionar:</strong><br>' + error, 'erro');
    })
    .adicionarLinha(dados);
};

window.carregarHistorico = function() {
  google.script.run
    .withSuccessHandler(function(historico) {
      let html = '<table style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px;">' +
        '<thead style="background: #f5f5f5;">' +
        '<tr><th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Data</th>' +
        '<th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Tipo</th>' +
        '<th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Finalidade</th>' +
        '<th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Resultado</th></tr>' +
        '</thead><tbody>';

      if (!historico || !Array.isArray(historico) || historico.length === 0) {
        html += '<tr><td colspan="4" style="padding: 20px; text-align: center; color: #999;">Nenhuma consulta registrada ainda</td></tr>';
      } else {
        historico.forEach(item => {
          html += '<tr style="border-bottom: 1px solid #eee;">' +
            '<td style="padding: 10px;">' + (item[0] || '-') + '</td>' +
            '<td style="padding: 10px;">' + (item[1] || '-') + '</td>' +
            '<td style="padding: 10px;">' + (item[2] || '-') + '</td>' +
            '<td style="padding: 10px;">' + (item[4] || '-') + '</td>' +
            '</tr>';
        });
      }

      html += '</tbody></table>';
      document.getElementById('tabelaHistorico').innerHTML = html;
    })
    .withFailureHandler(function(error) {
      console.error('Erro ao carregar histórico:', error);
      document.getElementById('tabelaHistorico').innerHTML = '<p style="color: red;">Erro ao carregar histórico: ' + error + '</p>';
    })
    .obterHistorico();
};

window.formatarMoeda = function(valor) {
  if (isNaN(valor)) return '0,00';
  return parseFloat(valor).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
};
</script>
</body>
</html>`).setWidth(1000).setHeight(800);
}
