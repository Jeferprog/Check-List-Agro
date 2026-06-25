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
    "Status (Ativa/Inativa)", "Data Atualização", "Observações"
  ];

  SHEET_LINHAS.appendRow(headers);
  SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1f4788");
  SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontColor("white");

  // Dados base das linhas
  const linhas = [
    ["L001", "PRONAF B (Microcrédito)", "Banco do Brasil/Caixa", "Microcrédito para pequenos produtores",
      "custeio,investimento", "Sem limite/R$ 500 mil", "0.5", "2.5", "24", "0", "1000", "25000",
      "DAP ativa, agricultor familiar", "DAP, RG, CPF, comprovante renda", "Ativa", new Date(), "Modalidade mais acessível"],

    ["L002", "PRONAF Custeio", "Banco do Brasil/Caixa/BB", "Despesas do ciclo produtivo",
      "custeio", "Sem limite/R$ 500 mil", "2", "6", "12", "0", "5000", "500000",
      "DAP ativa, agricultor familiar", "DAP, RG, CPF, projeto técnico", "Ativa", new Date(), "Safra até 12 meses"],

    ["L003", "PRONAF Investimento", "Banco do Brasil/Caixa/BB", "Máquinas, equipamentos, infraestrutura",
      "investimento,equipamento", "Sem limite/R$ 500 mil", "5", "5", "120", "36", "10000", "500000",
      "DAP ativa, agricultor familiar", "DAP, RG, CPF, projeto técnico, 3 últimos balanços", "Ativa", new Date(), "Carência de 3 anos"],

    ["L004", "PRONAF Agroecologia", "Banco do Brasil/Caixa", "Transição para sistemas sustentáveis",
      "investimento,agroecologia", "Sem limite/R$ 500 mil", "0.5", "0.5", "120", "36", "5000", "300000",
      "DAP ativa, agricultor familiar, projeto agroecológico", "DAP, RG, CPF, projeto agroecológico", "Ativa", new Date(), "Menor taxa - foco sustentabilidade"],

    ["L005", "PRONAF Irrigação", "Banco do Brasil/Caixa", "Sistemas de irrigação eficiente",
      "investimento,irrigacao", "Sem limite/R$ 500 mil", "4", "4", "120", "36", "15000", "400000",
      "DAP ativa, agricultor familiar, projeto irrigação", "DAP, projeto técnico irrigação, orçamentos", "Ativa", new Date(), "Para modernizar irrigação"],

    ["L006", "PRONAF Mulher", "Banco do Brasil/Caixa", "Mulheres produtoras rurais",
      "investimento,custeio", "Sem limite/R$ 500 mil", "2", "5", "120", "36", "10000", "400000",
      "DAP ativa, mulher, agricultor familiar", "DAP, RG, CPF, projeto técnico", "Ativa", new Date(), "Exclusivo para mulheres"],

    ["L007", "PRONAF Jovem", "Banco do Brasil/Caixa", "Beneficiários 16-29 anos",
      "investimento,custeio", "Sem limite/R$ 500 mil", "2", "5", "120", "36", "10000", "200000",
      "DAP ativa, jovem 16-29 anos", "DAP, RG, CPF, comprovante idade", "Ativa", new Date(), "Limite menor para jovens"],

    ["L008", "PRONAMP Custeio", "Banco do Brasil/Caixa/BNDES", "Despesas do ciclo produtivo - Médios produtores",
      "custeio", "R$ 500 mil/R$ 3.5 mi", "10", "10", "12", "0", "100000", "2000000",
      "Mín 80% renda agrícola, renda até R$ 3.5 mi", "RG, CPF, últimos 2 balanços, comprovante renda", "Ativa", new Date(), "Para produtores médios"],

    ["L009", "PRONAMP Investimento", "Banco do Brasil/Caixa/BNDES", "Máquinas, equipamentos - Médios produtores",
      "investimento,equipamento", "R$ 500 mil/R$ 3.5 mi", "9", "10", "120", "24", "200000", "3500000",
      "Mín 80% renda agrícola, renda até R$ 3.5 mi", "RG, CPF, últimos 2 balanços, projeto técnico", "Ativa", new Date(), "Limite maior que PRONAF"],

    ["L010", "Moderfrota", "BNDES/Bancos Credenciados", "Aquisição de máquinas e implementos",
      "investimento,equipamento,mecanizacao", "Conforme análise", "9", "13.5", "96", "12", "50000", "5000000",
      "Médio/grande produtor, documentação completa", "RG, CPF, últimos 3 balanços, orçamentos", "Ativa", new Date(), "Específico para máquinas agrícolas"],

    ["L011", "Moderagro", "BNDES/Banco do Brasil/Caixa", "Modernização e produtividade",
      "investimento,modernizacao", "Conforme análise", "8", "10", "144", "36", "100000", "2200000",
      "Propriedade registrada, projeto técnico", "Documentos de propriedade, projeto técnico, orçamentos", "Ativa", new Date(), "Foco modernização geral"],

    ["L012", "Programa ABC", "BNDES/Banco do Brasil", "Projetos de baixa emissão de carbono",
      "investimento,sustentabilidade,carbono", "Conforme análise", "8", "10", "120", "36", "50000", "2200000",
      "Propriedade registrada, projeto baixo carbono", "Projeto técnico ABC, documentos propriedade", "Ativa", new Date(), "Pivot, plantio direto, reflorestamento"],

    ["L013", "PCA (Armazéns)", "BNDES", "Construção e ampliação de silos/armazéns",
      "investimento,infraestrutura,armazenagem", "Conforme análise", "8.5", "10", "120", "24", "100000", "200000000",
      "Capacidade até 12.000 ton ou cooperativa", "Documentos propriedade, projeto técnico, orçamentos", "Ativa", new Date(), "Para infraestrutura de armazenagem"],

    ["L014", "Proirriga", "BNDES", "Sistemas de irrigação eficiente (geral)",
      "investimento,irrigacao", "Conforme análise", "10.5", "10.5", "120", "24", "20000", "500000",
      "Projeto irrigação, seguro obrigatório", "Projeto técnico, orçamentos, proponente habilitado", "Ativa", new Date(), "Obrigatório contratar seguro"],

    ["L015", "RenovAgro", "BNDES", "Recuperação e projetos ambientais",
      "investimento,ambiental,sustentabilidade", "Conforme análise", "9", "9", "120", "24", "50000", "1500000",
      "Propriedade registrada, projeto ambiental", "Documentos propriedade, projeto técnico ambiental", "Ativa", new Date(), "Inclui prevenção de incêndios"],

    ["L016", "Funcafé", "BNDES", "Específica para cafeicultores",
      "investimento,custeio,cafe", "Conforme análise", "8.5", "8.5", "120", "36", "100000", "3000000",
      "Propriedade com café, projeto técnico", "Documentos propriedade, projeto técnico, comprovante atividade", "Ativa", new Date(), "Exclusiva para cafeicultura"],

    ["L017", "Agricultura Empresarial", "Bancos Credenciados", "Custeio geral (Grande produtor)",
      "custeio", "Acima de R$ 3.5 mi", "14", "14", "12", "0", "500000", "999999999",
      "Renda acima R$ 3.5 milhões", "Documentos completos, últimos 3 balanços", "Ativa", new Date(), "Maior risco = maior taxa"]
  ];

  linhas.forEach(linha => SHEET_LINHAS.appendRow(linha));

  // Formatar sheet
  SHEET_LINHAS.setColumnWidths(1, 17, 80);
  SHEET_LINHAS.getRange("O:O").setHorizontalAlignment("center");
}

function inicializarSheetConfig() {
  SHEET_CONFIG.clear();
  const headers = ["Parâmetro", "Valor", "Tipo", "Descrição"];
  SHEET_CONFIG.appendRow(headers);
  SHEET_CONFIG.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#1f4788").setFontColor("white");

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
  SHEET_HISTORICO.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1f4788").setFontColor("white");
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
            observacoes: linha[headers.indexOf("Observações")] || ""
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
   * Ativa ou desativa uma linha de crédito
   */
  return atualizarLinha(idLinha, {
    "Status (Ativa/Inativa)": ativo ? "Ativa" : "Inativa"
  });
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
          linha[header] = dados[i][idx] || "";
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

function listarTodasAsLinhas() {
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
        id: linha[0] || "",
        nome: linha[1] || "",
        orgao: linha[2] || ""
      });
    }

    Logger.log("Retornando " + resultado.length + " linhas (versão simplificada)");
    return resultado;
  } catch (e) {
    Logger.log("Erro em listarTodasAsLinhas: " + e.toString());
    return [];
  }
}

// ==================== INTERFACE WEB ====================

function doGet() {
  return obterHTML();
}

function obterHTML() {
  return HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sistema de Crédito Rural - CRESOL</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #1f4788 0%, #2d5a9a 100%); min-height: 100vh; padding: 20px; }
.container { max-width: 1000px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); overflow: hidden; }
.header { background: linear-gradient(135deg, #1f4788 0%, #2d5a9a 100%); color: white; padding: 30px; text-align: center; }
.header h1 { font-size: 28px; margin-bottom: 5px; }
.header p { font-size: 14px; opacity: 0.9; }
.tabs { display: flex; background: #f5f5f5; border-bottom: 2px solid #ddd; }
.tab-btn { flex: 1; padding: 15px; background: none; border: none; cursor: pointer; font-size: 14px; font-weight: 500; color: #666; transition: all 0.3s; border-bottom: 3px solid transparent; }
.tab-btn.active { color: #1f4788; border-bottom-color: #1f4788; background: white; }
.tab-content { display: none; padding: 30px; }
.tab-content.active { display: block; }
.form-group { margin-bottom: 20px; }
label { display: block; margin-bottom: 8px; font-weight: 600; color: #333; }
input, select, textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 5px; font-size: 14px; font-family: inherit; }
input:focus, select:focus, textarea:focus { outline: none; border-color: #1f4788; box-shadow: 0 0 0 3px rgba(31, 71, 136, 0.1); }
button { background: linear-gradient(135deg, #1f4788 0%, #2d5a9a 100%); color: white; padding: 12px 30px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.3s; }
button:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(31, 71, 136, 0.3); }
.resultado { background: #f9f9f9; border-left: 4px solid #1f4788; padding: 20px; margin-top: 30px; border-radius: 5px; display: none; }
.resultado.visible { display: block; }
.linha-card { background: white; border: 1px solid #ddd; border-radius: 5px; padding: 15px; margin-bottom: 15px; transition: all 0.3s; }
.linha-card:hover { box-shadow: 0 3px 10px rgba(0,0,0,0.1); }
.linha-card h3 { color: #1f4788; margin-bottom: 10px; font-size: 16px; }
.linha-info { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 10px; }
.info-item { font-size: 13px; }
.info-label { font-weight: 600; color: #666; }
.info-value { color: #333; margin-top: 3px; }
.alert { padding: 15px; border-radius: 5px; margin-bottom: 20px; }
.alert-info { background: #e8f4f8; color: #0c5460; border-left: 4px solid #0c5460; }
.loading { text-align: center; padding: 20px; display: none; }
.spinner { border: 3px solid #f3f3f3; border-top: 3px solid #1f4788; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 0 auto; }
@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
.grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
@media (max-width: 600px) { .grid-2 { grid-template-columns: 1fr; } .linha-info { grid-template-columns: 1fr; } .header h1 { font-size: 20px; } .tab-btn { font-size: 12px; } }
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>🌾 Sistema de Crédito Rural</h1>
<p>Consultar linhas de crédito para operações rurais - CRESOL</p>
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
<label>📦 Produto/Finalidade</label>
<input type="text" id="produto" placeholder="Trator, custeio de safra...">
</div>
<div class="grid-2">
<div class="form-group">
<label>👤 Enquadramento</label>
<select id="enquadramento">
<option value="">-- Selecione --</option>
<option value="pronaf">PRONAF (Agricultura Familiar)</option>
<option value="pronamp">PRONAMP (Médio Produtor)</option>
<option value="empresarial">Agricultura Empresarial</option>
</select>
</div>
<div class="form-group">
<label>💰 Renda Bruta Anual (R$)</label>
<input type="number" id="renda" placeholder="Ex: 150000" min="0">
</div>
</div>
<div class="form-group">
<label>🎯 Tipo de Operação</label>
<select id="finalidade">
<option value="">-- Selecione --</option>
<option value="custeio">Custeio (Despesas do ciclo)</option>
<option value="investimento">Investimento (Máquinas/Equipamentos)</option>
<option value="mecanizacao">Mecanização</option>
<option value="irrigacao">Irrigação</option>
<option value="agroecologia">Agroecologia</option>
<option value="infraestrutura">Infraestrutura</option>
</select>
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
<strong>⚙️ Área Administrativa:</strong> Gerenciar linhas de crédito.
</div>
<div class="form-group">
<label>Selecione uma linha para editar:</label>
<select id="linhaParaEditar" onchange="window.carregarDadosLinha()">
<option value="">-- Selecione uma linha --</option>
</select>
</div>
<div id="edicaoConteudo"></div>
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

window.buscar = function() {
  const produto = document.getElementById('produto').value;
  const enquadramento = document.getElementById('enquadramento').value;
  const renda = parseInt(document.getElementById('renda').value) || 0;
  const finalidade = document.getElementById('finalidade').value;

  if (!enquadramento || !renda || !finalidade) {
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
  let html = '<h2 style="margin-bottom: 20px; color: #1f4788;">Linhas Disponíveis (' + linhas.length + ')</h2>';

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
      html += '</div></div>';
    });
  }
  document.getElementById('resultadoConteudo').innerHTML = html;
  document.getElementById('resultado').classList.add('visible');
};

window.carregarLinhasAdministrativo = function() {
  google.script.run
    .withSuccessHandler(function(linhas) {
      const select = document.getElementById('linhaParaEditar');
      select.innerHTML = '<option value="">-- Selecione uma linha --</option>';

      if (!linhas || !Array.isArray(linhas) || linhas.length === 0) {
        select.innerHTML += '<option disabled>Nenhuma linha disponível</option>';
        return;
      }

      linhas.forEach((linha) => {
        const option = document.createElement('option');
        option.value = linha.id || '';
        option.textContent = linha.nome || 'Sem nome';
        select.appendChild(option);
      });

      console.log('Linhas carregadas: ' + linhas.length);
    })
    .withFailureHandler(function(error) {
      console.error('Erro ao carregar linhas:', error);
      const select = document.getElementById('linhaParaEditar');
      select.innerHTML = '<option disabled>Erro ao carregar</option>';
    })
    .listarTodasAsLinhas();
};

window.carregarDadosLinha = function() {
  const selectElement = document.getElementById('linhaParaEditar');
  const idLinha = selectElement.value;

  if (!idLinha) {
    document.getElementById('edicaoConteudo').innerHTML = '';
    return;
  }

  const nomeLinhaElement = selectElement.options[selectElement.selectedIndex];
  const nomeLinha = nomeLinhaElement ? nomeLinhaElement.text : 'Desconhecido';

  const html = '<div style="margin-top: 20px; border: 1px solid #ddd; padding: 20px; border-radius: 5px;">' +
    '<h3 style="margin-bottom: 20px; color: #1f4788;">' + nomeLinha + '</h3>' +
    '<p><strong>ID:</strong> ' + idLinha + '</p>' +
    '<p style="color: #666; font-style: italic;">Clique em editar para modificar os dados...</p>' +
    '<button onclick="alert(\'Edição será implementada na próxima versão. ID: ' + idLinha + '\')" style="margin-top: 20px; background: #28a745;">✏️ Editar</button>' +
    '</div>';

  document.getElementById('edicaoConteudo').innerHTML = html;
};

window.mostrarFormularioEdicao = function(linha) {
  if (!linha || typeof linha !== 'object') {
    document.getElementById('edicaoConteudo').innerHTML = '<p style="color: red;">Nenhuma linha selecionada</p>';
    return;
  }

  let html = '<div style="margin-top: 20px; border: 1px solid #ddd; padding: 20px; border-radius: 5px;">';
  html += '<h3 style="margin-bottom: 20px; color: #1f4788;">' + (linha['Nome Linha'] || 'Sem nome') + '</h3>';
  html += '<p><strong>Finalidade:</strong> ' + (linha['Finalidade Principal'] || 'N/A') + '</p>';
  html += '<p><strong>Taxa:</strong> ' + (linha['Taxa Mín (%)'] || 'N/A') + '% - ' + (linha['Taxa Máx (%)'] || 'N/A') + '%</p>';
  html += '<p><strong>Prazo:</strong> ' + (linha['Prazo (meses)'] || 'N/A') + ' meses</p>';
  html += '<p><strong>Carência:</strong> ' + (linha['Carência (meses)'] || '0') + ' meses</p>';
  html += '<p><strong>Limite:</strong> R$ ' + window.formatarMoeda(linha['Limite Min (R$)']) + ' a R$ ' + window.formatarMoeda(linha['Limite Máx (R$)']) + '</p>';
  html += '<p><strong>Requisitos:</strong> ' + (linha['Requisitos'] || 'N/A') + '</p>';
  html += '<p><strong>Status:</strong> ' + (linha['Status (Ativa/Inativa)'] || 'Ativa') + '</p>';
  html += '</div>';
  document.getElementById('edicaoConteudo').innerHTML = html;
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
