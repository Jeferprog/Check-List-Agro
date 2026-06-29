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
const SHEET_BASE = SS.getSheetByName("Base") || SS.insertSheet("Base");
const SHEET_BASE_CREDITO = SS.getSheetByName("BaseCredito") || SS.insertSheet("BaseCredito");

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
    "Itens Financiáveis", "Culturas Financiadas"
  ];

  SHEET_LINHAS.appendRow(headers);
  SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46");
  SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontColor("white");

  // Dados base das linhas
  const linhas = [
    ["L001", "PRONAF CUSTEIO AGRÍCOLA Faixa I", "BNDES / Cresol", "Custeio", "agricola,custeio", "Sem limite/R$ 500 mil", "3", "3", "36", "0", "0", "250000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 3 anos para as culturas: Açafrão e Palmito; Até 2 anos para as culturas bienais e manejo florestal sustentável; Até 14 meses para as culturas permanentes; Até 11 meses para as demais culturas. | Norma: CIRCULAR SUP/ADIG No 47/2024-BNDES", "Itens de custeio relacionados à atividade agrícola; Despesas de soca e Ressoca de cana-de-açúcar; Aquisição antecipada de insumos; Aquisição de silos (bags). Limitada a 5% do valor do custeio; Até 15% do valor total do orçamento para: reparos ou reformas de bens de produção e de instalações, aquisição de animais de serviço, desmatamento, destoca e similares, inclusive aquisição, transporte, aplicação e incorporação de calcário agrícola.; Despesas de transporte e de frete.; Custos relativos à elaboração de projetos para outorga e licenciamento ambiental.; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "ABACAXI, ABOBORA(ORGÂNICA), ABOBORA-MORANGA(não orgânica), ABOBRINHA, AÇAÍ CULTIVADO, ALFACE, ALHO, ALHO NOBRE, ALHO PORÓ, ALMEIRÃO, AMENDOIM, ARROZ, BANANA, BATATA INGLESA, BATATA-DOCE, BERINJELA, BETERRABA, BROCOLIS, BUFALOS (BUBALINOS) LEITE, CACAU CULTIVADO, CARÁ, CEBOLA, CENOURA, CHÁ, CHICÓRIA/ESCAROLA, CHUCHU, COUVE, COUVE-FLOR, ERVAS AROMÁTICAS E CONDIMENTARES, ERVAS MEDICINAIS, FEIJÃO, FEIJÃO CAUPI, HORTALIÇAS, INHAME, LARANJA, MANDIOCA, MANDIOQUINHA (BAROA, SALSA, AIPO), MELANCIA, MELÃO, MILHO (até 25 mil), MORANGO, MOSTARDA, NABO, PAINÇO, PEPINO, PIMENTA, PIMENTÃO, QUIABO, RABANETE, REPOLHO, RÚCULA, SALSA, TANGERINA, TOMATE, TOMATE CEREJA, TOMATE RASTEIRO, TRIGO, URUCUM, VAGEM"],

    ["L002", "PRONAF CUSTEIO AGRÍCOLA Faixa II", "BNDES / Cresol", "Custeio", "agricola,custeio", "Sem limite/R$ 500 mil", "6.5", "6.5", "36", "0", "0", "250000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 3 anos para as culturas: Açafrão e Palmito; Até 2 anos para as culturas bienais e manejo florestal sustentável; Até 14 meses para as culturas permanentes; Até 11 meses para as demais culturas. | Norma: CIRCULAR SUP/ADIG No 47/2024-BNDES", "Itens de custeio relacionados à atividade agrícola; Despesas de soca e Ressoca de cana-de-açúcar; Aquisição antecipada de insumos; Aquisição de silos (bags). Limitada a 5% do valor do custeio; Até 15% do valor total do orçamento para: reparos ou reformas de bens de produção e de instalações, aquisição de animais de serviço, desmatamento, destoca e similares, inclusive aquisição, transporte, aplicação e incorporação de calcário agrícola.; Despesas de transporte e de frete.; Custos relativos à elaboração de projetos para outorga e licenciamento ambiental.; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "ABACATE, ACÁCIA NEGRA COM PALMEIRA REAL, ALFAFA - AZEVEM, AMEIXA, AMORA, AVEIA, CAFÉ, CANA-DE-AÇUCAR, CANOLA DE SEQUEIRO, CAQUI, CENTEIO, CEVADA, ERVILHACA, EUCALIPTO - TRATOS CULTURAIS, EUCALIPTO FLORESTAMENTO E REFLORESTAMENTO, FIGO, FLORES, GIRASSOL, GOIABA, GRAMA, KIWI, LIMÃO, MARACUJA, MAÇÃ, MILHO (acima de 25 mil), MUDAS DIVERSAS, NECTARINA, NOZ, PALMEIRA REAL, PASTAGEM, PERA, PINUS, PITAYA, PLANTAS ORNAMENTAIS, PORONGO (CUIA, CABAÇA), PÊSSEGO, SORGO, TRITICALE, UVA"],

    ["L003", "PRONAF CUSTEIO AGRÍCOLA Faixa III", "BNDES / Cresol", "Custeio", "agricola,custeio", "Sem limite/R$ 500 mil", "2", "2", "36", "0", "0", "250000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 3 anos para as culturas: Açafrão e Palmito; Até 2 anos para as culturas bienais e manejo florestal sustentável; Até 14 meses para as culturas permanentes; Até 11 meses para as demais culturas. | Norma: CIRCULAR SUP/ADIG No 47/2024-BNDES", "Itens de custeio relacionados à atividade agrícola; Despesas de soca e Ressoca de cana-de-açúcar; Aquisição antecipada de insumos; Aquisição de silos (bags). Limitada a 5% do valor do custeio; Até 15% do valor total do orçamento para: reparos ou reformas de bens de produção e de instalações, aquisição de animais de serviço, desmatamento, destoca e similares, inclusive aquisição, transporte, aplicação e incorporação de calcário agrícola.; Despesas de transporte e de frete.; Custos relativos à elaboração de projetos para outorga e licenciamento ambiental.; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "ABIU, ABOBORA(ORGÂNICA), ABOBORA-MORANGA (ORGÂNICA), AMORA-PRETA, ANDIROBA, ARATICUM, ARAÇÁ, ARAÇÁ-BOI, ARAÇÁ-PERA, AROEIRA-PIMENTEIRA, ARIÁ, ARUMBEVA, AÇAÍ EXTRATIVO, BABAÇU, BACABA, BACUPARI, BACURI, BARU, BATATA CREM, BELDROEGA, BIRIBÁ, BORRACHA EXTRATIVA, BURITI, BUTIÁ, CACAU EXTRATIVO, CAGAITA, CAJÁ, CAJU, CAJU-DO-CERRADO, CAMBUCI, CAMBUI, CAMU-CAMU, CARÁ AMAZÔNICO, CARÁ-DE-ESPINHO, CARNAÚBA, CASTANHA-DO-PARÁ/CASTANHA-DO-BRASIL, CASTANHA-DE-CUTIA, CASTANHA-DE-GALINHA, CENOURA (ORGÂNICA), CEREJA-DO-RIO-GRANDE, CHICHÁ, CHICÓRIA-DE-CABOCLO, COQUINHO-AZEDO, COPAÍBA, CROÁ, CUBIU, CUPUAÇU, ERVA MATE, FISALIS, GOIABA-SERRANA, GUABIROBA, GUARANÁ, GRUMIXAMA, GUEROBA, JABORANDI, JABUTICABA, JARACATIÁ, JAMBU, JATOBÁ, JENIPAPO, JUÇARA, LICURI, MACAÚBA, MAJOR-GOMES, MANDACARU, MANGABA, MAPATI, MINI-PEPININHO, MORANGA, MURICI, MURUMURU, ORA-PRO-NÓBIS, OSTRA-DE-MANGUE, PATAUÁ, PAJURÁ, PEQUI, PEPERÔMIA, PERA-DO-CERRADO, PIAÇAVA, PINHÃO, PIRARUCU DE MANEJO, PITANGA, PUPUNHA, PUXURI, SAPOTA, SETE-CAPOTES, SORVA, TAIOBA, TAPEREBÁ, TUCUMÃ, UMARI, UMBU, URUCUM, UVA (ORGÂNICA), UVAIA, UXI, MELIPONICULTURA"],

    ["L004", "PRONAF CUSTEIO AGRÍCOLA Faixa IV", "BNDES / Cresol", "Custeio", "agricola,custeio", "Sem limite/R$ 500 mil", "8", "8", "36", "0", "0", "250000", "Apresentação de DAP-Pronaf; Exploração de terra (proprietário, posseiro, arrendatário, comodatário, parceiro); Área de terra ≤ 4 módulos fiscais; Renda bruta familiar (últimos 12 meses) ≤ R$500.000,00; Mínimo 50% da renda bruta familiar de exploração agropecuária e não agropecuária", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 3 anos para Açafrão e Palmito; Até 2 anos para culturas bienais e manejo florestal sustentável; Até 14 meses para culturas permanentes; Até 11 meses para demais culturas. | Norma: CIRCULAR SUP/ADIG No 47/2024-BNDES", "Itens de custeio relacionados à atividade agrícola; Despesas de soca e ressoca de cana-de-açúcar; Aquisição antecipada de insumos; Aquisição de silos (bags) - limitada a 5% do valor do custeio; Até 15% do valor total do orçamento para: reparos ou reformas de bens de produção e instalações; aquisição de animais de serviço; desmatamento; destoca e similares; aquisição, transporte, aplicação e incorporação de calcário agrícola.; Despesas de transporte e frete; Custos relativos à elaboração de projetos para outorga e licenciamento ambiental; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, plataformas e soluções digitais", "Algodão, SOJA, BUFALOS (BUBALINOS) CARNE"],

    ["L005", "PRONAF CUSTEIO PECUÁRIO Faixa I", "BNDES / Cresol", "Custeio", "custeio,pecuaria", "Sem limite/R$ 500 mil", "3", "3", "0", "0", "0", "250000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 20 (vinte) meses, quando o financiamento se destinar a avicultura caipira de postura e até 10 (dez) meses nos demais financiamentos. | Norma: CIRCULAR SUP/ADIG No 47/2024-BNDES", "Itens de custeio relacionados com a atividade pecuária; Aquisição de animais para recria e engorda por exemplo compra de Alevinos que entram na Piscicultura; Aquisição de insumos (medicamentos, vacinas, antiparasitários, sais minerais, vitaminas, etc.); Despesas para colocação de brincos numerados e cápsulas de microchip; Limpeza e restauração de pastagens, fenação, silagem e formação de forragens periódicas; Até 15% do valor do orçamento para pequenas despesas de investimento.; Despesas de transporte e de frete de insumos; Custos relativos à elaboração de projetos para outorga de uso da água e para licenciamento ambiental; Despesas com aquisição de insumos para restauração e recuperação das áreas de reserva legal e das áreas de preservação permanente; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, plataformas e de soluções digitais", "APICULTURA, AQUICULTURA E PESCA, AVICULTURA DE POSTURA, BOVINOS LEITE, CAPRINOCULTURA, OVINOCULTURA"],

    ["L006", "PRONAF CUSTEIO PECUÁRIO Faixa II", "BNDES / Cresol", "Custeio", "custeio,pecuaria", "Sem limite/R$ 500 mil", "6.5", "6.5", "0", "0", "0", "250000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 10 (dez) meses nos demais financiamentos. | Norma: CIRCULAR SUP/ADIG No 47/2024-BNDES", "Itens de custeio relacionados com a atividade pecuária; Aquisição de insumos (medicamentos, vacinas, antiparasitários, sais minerais, vitaminas, etc.); Despesas para colocação de brincos numerados e cápsulas de microchip; Limpeza e restauração de pastagens, fenação, silagem e formação de forragens periódicas; Medicamentos, vacinas, antiparasitários, sais minerais, vitaminas entre outros; Até 15% do valor do orçamento pode incluir verbas de pequenas despesas conceituadas como investimento.; Despesas de transporte e de frete de insumos; Custos relativos à elaboração de projetos para outorga de uso da água e para licenciamento ambiental; Despesas com aquisição de insumos para restauração e recuperação das áreas de reserva legal e das áreas de preservação permanente; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, plataformas e de soluções digitais", "DEMAIS CRIAÇÕES NÃO ENQUADRADAS NAS FAIXAS I E IV, AVICULTURA PARA CORTE, SUINOS"],

    ["L007", "PRONAF CUSTEIO PECUÁRIO Faixa IV", "BNDES / Cresol", "Custeio", "custeio,pecuaria", "Sem limite/R$ 500 mil", "8", "8", "0", "0", "0", "250000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: I - até 6 (seis) meses, no financiamento para aquisição de bovinos e bubalinos para engorda em regime de confinamento;II - até 12 (doze) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para recria em regime extensivo;III - até 8 (oito) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para engorda em regime extensivo;IV - até 20 (vinte) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos, desde que a mesma operação abranja, necessariamente, a recria e a engord", "Itens de custeio relacionados com a atividade pecuária; Aquisição de animais (BOVINOS) para recria em engorda; Aquisição de insumos (medicamentos, vacinas, antiparasitários, sais minerais, vitaminas, etc.); Despesas para colocação de brincos numerados e cápsulas de microchip; Limpeza e restauração de pastagens, fenação, silagem e formação de forragens periódicas; Até 15% do valor do orçamento pode incluir verbas de pequenas despesas conceituadas como investimento.; Despesas de transporte e de frete de insumos; Custos relativos à elaboração de projetos para outorga de uso da água e para licenciamento ambiental; Despesas com aquisição de insumos para restauração e recuperação das áreas de reserva legal e das áreas de preservação permanente; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, plataformas e de soluções digitais", "BOVINOS CORTE (INCLUSIVE AQUISIÇÃO DE ANIMAIS DESTINADOS A RECRIA E ENGORDA)"],

    ["L008", "PRONAF CUSTEIO INDUSTRIALIZAÇÃO", "Cresol (Poupança Rural)", "Custeio", "custeio,infraestrutura", "Sem limite/R$ 500 mil", "8", "8", "12", "0", "0", "55000000", "Conforme política de crédito da Cresol", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Custeio | Sistemática: LCA/POUPANÇA | IOF: 0,38% | Prazo: Até 12 meses", "Beneficiamento e industrialização da produção (aquisição de embalagens, rótulos, condimentos, conservantes, adoçantes e outros insumos).; Formação de estoques de insumos, matéria-prima, produto final e serviços de apoio à comercialização.; Adiantamentos por conta do preço de produtos entregues para venda.; Financiamento da armazenagem e conservação de produtos para venda futura.; Aquisição de insumos pela cooperativa de produção.", ""],

    ["L009", "CUSTEIO DEMAIS PRODUTORES AGRÍCOLA", "BNDES / Cresol", "Custeio", "agricola,custeio", "Conforme análise", "14", "14", "36", "0", "0", "3000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 3 anos para as culturas: Açafrão e Palmito; Até 2 anos para as culturas bienais e manejo florestal sustentável; Até 14 meses para as culturas permanentes; Até 1 ano para as demais culturas. | Norma: CIRCULAR SUP/ADIG No 51/2024-BNDES", "Despesas de soca e Ressoca de cana-de-açúcar; Aquisição antecipada de insumos; Aquisição de silos (bags). Limitada a 5% do valor do custeio; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "ABACATE, ABACAXI, ABOBORA-MORANGA, ABOBORA-MORANGA (ORGÂNICA), ABOBRINHA, ACÁCIA NEGRA COM PALMEIRA REAL, ALFACE, ALFAFA - AZEVEM, ALHO, ALHO NOBRE, ALHO PORÓ, ALMEIRÃO, AMEIXA, AMENDOIM, AMORA, ARROZ, AVEIA, BANANA, BATATA INGLESA, BATATA-DOCE, BERINJELA, BETERRABA, BROCOLIS, CAFÉ, CANA-DE-AÇUCAR, CANOLA DE SEQUEIRO, CAQUI, CEBOLA, CENOURA, CENOURA (ORGÂNICA), CENTEIO, CEVADA, CHICÓRIA/ESCAROLA, CHUCHU, CHÁ, COUVE, COUVE-FLOR, ERVA MATE, ERVILHACA, EUCALIPTO - TRATOS CULTURAIS, EUCALIPTO FLORESTAMENTO E REFLORESTAMENTO, FEIJÃO, FIGO, FLORES, GIRASSOL, GOIABA, GRAMA, HORTALIÇAS, KIWI, LARANJA, LIMÃO, MANDIOCA, MANDIOQUINHA (BAROA, SALSA, AIPO), MARACUJA, MAÇÃ, MELANCIA, MELÃO, MILHO (acima de 25 mil), MILHO (até 25 mil), MORANGO, MOSTARDA, MUDAS DIVERSAS, NABO, NECTARINA, NOZ, PAINÇO, PALMEIRA REAL, PASTAGEM, PEPINO, PERA, PIMENTA, PIMENTÃO, PINUS, PITAYA, PLANTAS ORNAMENTAIS, PORONGO (CUIA, CABAÇA), PÊSSEGO, QUIABO, RABANETE, REPOLHO, RÚCULA, SALSA, SOJA, SORGO, TANGERINA, TOMATE CEREJA, TOMATE RASTEIRO, TRIGO, TRITICALE, URUCUM, UVA, UVA (ORGÂNICA), VAGEM"],

    ["L010", "CUSTEIO DEMAIS PRODUTORES AGRÍCOLA Sustentável", "BNDES / Cresol", "Custeio", "agricola,custeio,sustentabilidade", "Conforme análise", "13.5", "13.5", "36", "0", "0", "3000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 3 anos para as culturas: Açafrão e Palmito; Até 2 anos para as culturas bienais e manejo florestal sustentável; Até 14 meses para as culturas permanentes; Até 1 ano para as demais culturas. | Norma: CIRCULAR SUP/ADIG No 51/2024-BNDES", "Despesas de soca e Ressoca de cana-de-açúcar; Aquisição antecipada de insumos; Aquisição de silos (bags). Limitada a 5% do valor do custeio; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "ABACATE, ABACAXI, ABOBORA-MORANGA, ABOBORA-MORANGA (ORGÂNICA), ABOBRINHA, ACÁCIA NEGRA COM PALMEIRA REAL, ALFACE, ALFAFA - AZEVEM, ALHO, ALHO NOBRE, ALHO PORÓ, ALMEIRÃO, AMEIXA, AMENDOIM, AMORA, ARROZ, AVEIA, BANANA, BATATA INGLESA, BATATA-DOCE, BERINJELA, BETERRABA, BROCOLIS, CAFÉ, CANA-DE-AÇUCAR, CANOLA DE SEQUEIRO, CAQUI, CEBOLA, CENOURA, CENOURA (ORGÂNICA), CENTEIO, CEVADA, CHICÓRIA/ESCAROLA, CHUCHU, CHÁ, COUVE, COUVE-FLOR, ERVA MATE, ERVILHACA, EUCALIPTO - TRATOS CULTURAIS, EUCALIPTO FLORESTAMENTO E REFLORESTAMENTO, FEIJÃO, FIGO, FLORES, GIRASSOL, GOIABA, GRAMA, HORTALIÇAS, KIWI, LARANJA, LIMÃO, MANDIOCA, MANDIOQUINHA (BAROA, SALSA, AIPO), MARACUJA, MAÇÃ, MELANCIA, MELÃO, MILHO (acima de 25 mil), MILHO (até 25 mil), MORANGO, MOSTARDA, MUDAS DIVERSAS, NABO, NECTARINA, NOZ, PAINÇO, PALMEIRA REAL, PASTAGEM, PEPINO, PERA, PIMENTA, PIMENTÃO, PINUS, PITAYA, PLANTAS ORNAMENTAIS, PORONGO (CUIA, CABAÇA), PÊSSEGO, QUIABO, RABANETE, REPOLHO, RÚCULA, SALSA, SOJA, SORGO, TANGERINA, TOMATE CEREJA, TOMATE RASTEIRO, TRIGO, TRITICALE, URUCUM, UVA, UVA (ORGÂNICA), VAGEM"],

    ["L011", "CUSTEIO DEMAIS PRODUTORES PECUÁRIO", "BNDES / Cresol", "Custeio", "custeio,pecuaria", "Conforme análise", "14", "14", "0", "0", "0", "3000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: I - até 6 (seis) meses, no financiamento para aquisição de bovinos e bubalinos para engorda em regime de confinamento;II - até 12 (doze) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para recria em regime extensivo;III - até 8 (oito) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para engorda em regime extensivo;IV - até 20 (vinte) meses, quando o financiamento se destinar a avicultura caipira de postura ou quando o financiamento envolver a aquisição de bovinos e bubalinos, de", "Aquisição de animais para recria e engorda; Aquisição de insumos; Despesas para colocação de brincos numerados e cápsulas de microchip; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "APICULTURA, AVES, AVICULTURA, BOVINOS CORTE, BOVINOS LEITE, BUFALOS (BUBALINOS) CARNE E LEITE, CAPRINOS - CARNE E LEITE, OVINOS, PISCICULTURA, SUINOS"],

    ["L012", "CUSTEIO DEMAIS PRODUTORES PECUÁRIO Sustentável", "BNDES / Cresol", "Custeio", "custeio,pecuaria,sustentabilidade", "Conforme análise", "13.5", "13.5", "0", "0", "0", "3000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: I - até 6 (seis) meses, no financiamento para aquisição de bovinos e bubalinos para engorda em regime de confinamento;II - até 12 (doze) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para recria em regime extensivo;III - até 8 (oito) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para engorda em regime extensivo;IV - até 20 (vinte) meses, quando o financiamento se destinar a avicultura caipira de postura ou quando o financiamento envolver a aquisição de bovinos e bubalinos, de", "Aquisição de animais para recria e engorda; Aquisição de insumos; Despesas para colocação de brincos numerados e cápsulas de microchip; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "APICULTURA, AVES, AVICULTURA, BOVINOS CORTE, BOVINOS LEITE, BUFALOS (BUBALINOS) CARNE E LEITE, CAPRINOS - CARNE E LEITE, OVINOS, PISCICULTURA, SUINOS"],

    ["L013", "PRONAMP CUSTEIO AGRICOLA", "BNDES / Cresol", "Custeio", "agricola,custeio", "R$ 500 mil/R$ 3.5 mi", "10", "10", "36", "0", "0", "1500000", "Ser proprietário rural, posseiro, arrendatário ou parceiro.; Possuir renda bruta anual de até R$3.500.000,00.; Ter, no mínimo, 80% da renda bruta anual originária da atividade agropecuária.", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 36 meses para cultura de açafrão e palmito; Até 24 meses para as culturas bienais e manejo florestal sustentável; Até 14 meses, para as culturas permanentes; Até 11 meses para as demais culturas | Norma: CIRCULAR SUP/ADIG No 49/2024-BNDES", "Despesas de soca e Ressoca de cana-de-açúcar; Aquisição antecipada de insumos; Aquisição de silos (bags). Limitada a 5% do valor do custeio; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "ABACATE, ABACAXI, ABOBORA-MORANGA, ABOBORA-MORANGA (ORGÂNICA), ABOBRINHA, ACÁCIA NEGRA COM PALMEIRA REAL, ALFACE, ALFAFA - AZEVEM, ALHO, ALHO NOBRE, ALHO PORÓ, ALMEIRÃO, AMEIXA, AMENDOIM, AMORA, ARROZ, AVEIA, BANANA, BATATA INGLESA, BATATA-DOCE, BERINJELA, BETERRABA, BROCOLIS, CAFÉ, CANA-DE-AÇUCAR, CANOLA DE SEQUEIRO, CAQUI, CEBOLA, CENOURA, CENOURA (ORGÂNICA), CENTEIO, CEVADA, CHICÓRIA/ESCAROLA, CHUCHU, CHÁ, COUVE, COUVE-FLOR, ERVA MATE, ERVILHACA, EUCALIPTO - TRATOS CULTURAIS, EUCALIPTO FLORESTAMENTO E REFLORESTAMENTO, FEIJÃO, FIGO, FLORES, GIRASSOL, GOIABA, GRAMA, HORTALIÇAS, KIWI, LARANJA, LIMÃO, MANDIOCA, MANDIOQUINHA (BAROA, SALSA, AIPO), MARACUJA, MAÇÃ, MELANCIA, MELÃO, MILHO (acima de 25 mil), MILHO (até 25 mil), MORANGO, MOSTARDA, MUDAS DIVERSAS, NABO, NECTARINA, NOZ, PAINÇO, PALMEIRA REAL, PASTAGEM, PEPINO, PERA, PIMENTA, PIMENTÃO, PINUS, PITAYA, PLANTAS ORNAMENTAIS, PORONGO (CUIA, CABAÇA), PÊSSEGO, QUIABO, RABANETE, REPOLHO, RÚCULA, SALSA, SOJA, SORGO, TANGERINA, TOMATE CEREJA, TOMATE RASTEIRO, TRIGO, TRITICALE, URUCUM, UVA, UVA (ORGÂNICA), VAGEM"],

    ["L014", "PRONAMP CUSTEIO AGRICOLA Sustentável", "BNDES / Cresol", "Custeio", "agricola,custeio,sustentabilidade", "R$ 500 mil/R$ 3.5 mi", "9.5", "9.5", "36", "0", "0", "1500000", "Propriedade ou posse da terra (proprietário, posseiro, arrendatário ou parceiro); Renda bruta anual de até R$ 3.500.000,00; Mínimo 80% da renda bruta anual originária da atividade agropecuária", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: Até 36 meses para cultura de açafrão e palmito; Até 24 meses para as culturas bienais e manejo florestal sustentável; Até 14 meses, para as culturas permanentes; Até 11 meses para as demais culturas | Norma: CIRCULAR SUP/ADIG No 49/2024-BNDES", "Despesas de soca e Ressoca de cana-de-açúcar; Aquisição antecipada de insumos; Aquisição de silos (bags) - limitada a 5% do valor do custeio; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "ABACATE, ABACAXI, ABOBORA-MORANGA, ABOBORA-MORANGA (ORGÂNICA), ABOBRINHA, ACÁCIA NEGRA COM PALMEIRA REAL, ALFACE, ALFAFA - AZEVEM, ALHO, ALHO NOBRE, ALHO PORÓ, ALMEIRÃO, AMEIXA, AMENDOIM, AMORA, ARROZ, AVEIA, BANANA, BATATA INGLESA, BATATA-DOCE, BERINJELA, BETERRABA, BROCOLIS, CAFÉ, CANA-DE-AÇUCAR, CANOLA DE SEQUEIRO, CAQUI, CEBOLA, CENOURA, CENOURA (ORGÂNICA), CENTEIO, CEVADA, CHICÓRIA/ESCAROLA, CHUCHU, CHÁ, COUVE, COUVE-FLOR, ERVA MATE, ERVILHACA, EUCALIPTO - TRATOS CULTURAIS, EUCALIPTO FLORESTAMENTO E REFLORESTAMENTO, FEIJÃO, FIGO, FLORES, GIRASSOL, GOIABA, GRAMA, HORTALIÇAS, KIWI, LARANJA, LIMÃO, MANDIOCA, MANDIOQUINHA (BAROA, SALSA, AIPO), MARACUJA, MAÇÃ, MELANCIA, MELÃO, MILHO (acima de 25 mil), MILHO (até 25 mil), MORANGO, MOSTARDA, MUDAS DIVERSAS, NABO, NECTARINA, NOZ, PAINÇO, PALMEIRA REAL, PASTAGEM, PEPINO, PERA, PIMENTA, PIMENTÃO, PINUS, PITAYA, PLANTAS ORNAMENTAIS, PORONGO (CUIA, CABAÇA), PÊSSEGO, QUIABO, RABANETE, REPOLHO, RÚCULA, SALSA, SOJA, SORGO, TANGERINA, TOMATE CEREJA, TOMATE RASTEIRO, TRIGO, TRITICALE, URUCUM, UVA, UVA (ORGÂNICA), VAGEM"],

    ["L015", "PRONAMP CUSTEIO PECUÁRIO", "BNDES / Cresol", "Custeio", "custeio,pecuaria", "R$ 500 mil/R$ 3.5 mi", "10", "10", "0", "0", "0", "1500000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: I - até 6 (seis) meses, no financiamento para aquisição de bovinos e bubalinos para engorda em regime de confinamento;II - até 12 (doze) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para recria em regime extensivo;III - até 8 (oito) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para engorda em regime extensivo;IV - até 20 (vinte) meses, quando o financiamento se destinar a avicultura caipira de postura ou quando o financiamento envolver a aquisição de bovinos e bubalinos, de", "Aquisição de animais para recria e engorda; Aquisição de insumos; Despesas para colocação de brincos numerados e cápsulas de microchip; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "APICULTURA, AVES, AVICULTURA, BOVINOS CORTE, BOVINOS LEITE, BUFALOS (BUBALINOS) CARNE E LEITE, CAPRINOS - CARNE E LEITE, OVINOS, PISCICULTURA, SUINOS"],

    ["L016", "PRONAMP CUSTEIO PECUÁRIO Sustentável", "BNDES / Cresol", "Custeio", "custeio,pecuaria,sustentabilidade", "R$ 500 mil/R$ 3.5 mi", "9.5", "9.5", "0", "0", "0", "1500000", "Propriedade rural, posse, arrendamento ou parceria; Renda bruta anual até R$ 3.500.000,00; Mínimo 80% da renda bruta anual originária da atividade agropecuária", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | Sistemática: DIR/BNDES/POUPANÇA | IOF: 0,38% | Prazo: I - até 6 (seis) meses, no financiamento para aquisição de bovinos e bubalinos para engorda em regime de confinamento;II - até 12 (doze) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para recria em regime extensivo;III - até 8 (oito) meses, quando o financiamento envolver a aquisição de bovinos e bubalinos para engorda em regime extensivo;IV - até 20 (vinte) meses, quando o financiamento se destinar a avicultura caipira de postura ou quando o financiamento envolver a aquisição de bovinos e bubalinos, de", "Aquisição de animais para recria e engorda; Aquisição de insumos; Despesas para colocação de brincos numerados e cápsulas de microchip; Despesas com aquisição de insumos para restauração das áreas de reserva legal e preservação permanente.; Aquisição de bioinsumos; Despesas para manutenção de infraestrutura de rede, de plataformas e de soluções digitais.", "APICULTURA, AVES, AVICULTURA, BOVINOS CORTE, BOVINOS LEITE, BUFALOS (BUBALINOS) CARNE E LEITE, CAPRINOS - CARNE E LEITE, OVINOS, PISCICULTURA, SUINOS"],

    ["L017", "CUSTEIO AGRÍCOLA - POUPANÇA LIVRE", "Cresol", "Custeio", "agricola,custeio", "Conforme análise", "0", "0", "23", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: 0,38% | Prazo: Até 23 meses", "Atividade produtivas agrícolas diversas no imóvel rural.", "ABACATE, ABACAXI, ABOBORA-MORANGA, ABOBORA-MORANGA (ORGÂNICA), ABOBRINHA, ACÁCIA NEGRA COM PALMEIRA REAL, ALFACE, ALFAFA - AZEVEM, ALHO, ALHO NOBRE, ALHO PORÓ, ALMEIRÃO, AMEIXA, AMENDOIM, AMORA, ARROZ, AVEIA, BANANA, BATATA INGLESA, BATATA-DOCE, BERINJELA, BETERRABA, BROCOLIS, CAFÉ, CANA-DE-AÇUCAR, CANOLA DE SEQUEIRO, CAQUI, CEBOLA, CENOURA, CENOURA (ORGÂNICA), CENTEIO, CEVADA, CHICÓRIA/ESCAROLA, CHUCHU, CHÁ, COUVE, COUVE-FLOR, ERVA MATE, ERVILHACA, EUCALIPTO - TRATOS CULTURAIS, EUCALIPTO FLORESTAMENTO E REFLORESTAMENTO, FEIJÃO, FIGO, FLORES, FUMO, GIRASSOL, GOIABA, GRAMA, HORTALIÇAS, KIWI, LARANJA, LIMÃO, MANDIOCA, MANDIOQUINHA (BAROA, SALSA, AIPO), MARACUJA, MAÇÃ, MELANCIA, MELÃO, MILHO (acima de 25 mil), MILHO (até 25 mil), MORANGO, MOSTARDA, MUDAS DIVERSAS, NABO, NECTARINA, NOZ, PAINÇO, PALMEIRA REAL, PASTAGEM, PEPINO, PERA, PIMENTA, PIMENTÃO, PINUS, PITAYA, PLANTAS ORNAMENTAIS, PORONGO (CUIA, CABAÇA), PÊSSEGO, QUIABO, RABANETE, REPOLHO, RÚCULA, SALSA, SOJA, SORGO, TANGERINA, TOMATE CEREJA, TOMATE RASTEIRO, TRIGO, TRITICALE, URUCUM, UVA, UVA (ORGÂNICA), VAGEM"],

    ["L018", "CUSTEIO PECUÁRIO - POUPANÇA LIVRE", "Cresol", "Custeio", "custeio,pecuaria", "Conforme análise", "0", "0", "23", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: 0,38% | Prazo: Até 23 meses", "Atividade produtivas pecuárias diversas no imóvel rural.", "APICULTURA, AVES, AVICULTURA, BOVINOS CORTE, BOVINOS LEITE, BUFALOS (BUBALINOS) CARNE E LEITE, CAPRINOS - CARNE E LEITE, OVINOS, PISCICULTURA, SUINOS"],

    ["L019", "CUSTEIO AGRÍCOLA - CRÉDITO RURAL TFB", "BNDES / Cresol", "Custeio", "agricola,custeio", "Conforme análise", "1.35", "1.35", "36", "0", "0", "250000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: 0,38% | Prazo: Até 3 anos | Norma: CIRCULAR SUP/ADIG Nº 19/2024-BNDES", "Financiamentos destinados ao atendimento das despesas de custeio das atividades agrícolas.", "ABACATE, ABACAXI, ABOBORA-MORANGA, ABOBORA-MORANGA (ORGÂNICA), ABOBRINHA, ACÁCIA NEGRA COM PALMEIRA REAL, ALFACE, ALFAFA - AZEVEM, ALHO, ALHO NOBRE, ALHO PORÓ, ALMEIRÃO, AMEIXA, AMENDOIM, AMORA, ARROZ, AVEIA, BANANA, BATATA INGLESA, BATATA-DOCE, BERINJELA, BETERRABA, BROCOLIS, CAFÉ, CANA-DE-AÇUCAR, CANOLA DE SEQUEIRO, CAQUI, CEBOLA, CENOURA, CENOURA (ORGÂNICA), CENTEIO, CEVADA, CHICÓRIA/ESCAROLA, CHUCHU, CHÁ, COUVE, COUVE-FLOR, ERVA MATE, ERVILHACA, EUCALIPTO - TRATOS CULTURAIS, EUCALIPTO FLORESTAMENTO E REFLORESTAMENTO, FEIJÃO, FIGO, FLORES, FUMO, GIRASSOL, GOIABA, GRAMA, HORTALIÇAS, KIWI, LARANJA, LIMÃO, MANDIOCA, MANDIOQUINHA (BAROA, SALSA, AIPO), MARACUJA, MAÇÃ, MELANCIA, MELÃO, MILHO (acima de 25 mil), MILHO (até 25 mil), MORANGO, MOSTARDA, MUDAS DIVERSAS, NABO, NECTARINA, NOZ, PAINÇO, PALMEIRA REAL, PASTAGEM, PEPINO, PERA, PIMENTA, PIMENTÃO, PINUS, PITAYA, PLANTAS ORNAMENTAIS, PORONGO (CUIA, CABAÇA), PÊSSEGO, QUIABO, RABANETE, REPOLHO, RÚCULA, SALSA, SOJA, SORGO, TANGERINA, TOMATE CEREJA, TOMATE RASTEIRO, TRIGO, TRITICALE, URUCUM, UVA, UVA (ORGÂNICA), VAGEM"],

    ["L020", "CUSTEIO PECUÁRIO - CRÉDITO RURAL TFB", "BNDES / Cresol", "Custeio", "custeio,pecuaria", "Conforme análise", "1.35", "1.35", "36", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: 0,38% | Prazo: Até 3 anos | Norma: CIRCULAR SUP/ADIG Nº 19/2024-BNDES", "Financiamentos destinados ao atendimento das despesas de custeio das atividades pecuárias.", "APICULTURA, AVES, AVICULTURA, BOVINOS CORTE, BOVINOS LEITE, BUFALOS (BUBALINOS) CARNE E LEITE, CAPRINOS - CARNE E LEITE, OVINOS, PISCICULTURA, SUINOS"],

    ["L021", "CUSTEIO PECUARIO REG SICOR", "Cresol", "Custeio", "custeio,pecuaria", "Conforme análise", "0", "0", "23", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: 0,38% | Prazo: Até 23 meses", "Atividade produtivas pecuárias diversas no imóvel rural.", "APICULTURA, AVES, AVICULTURA, BOVINOS CORTE, BOVINOS LEITE, BUFALOS (BUBALINOS) CARNE E LEITE, CAPRINOS - CARNE E LEITE, OVINOS, PISCICULTURA, SUINOS"],

    ["L022", "CUSTEIO AGRICOLA REG SICOR", "Cresol", "Custeio", "agricola,custeio", "Conforme análise", "0", "0", "23", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: 0,38% | Prazo: Até 23 meses", "Atividade produtivas agrícolas diversas no imóvel rural.", "ABACATE, ABACAXI, ABOBORA-MORANGA, ABOBORA-MORANGA (ORGÂNICA), ABOBRINHA, ACÁCIA NEGRA COM PALMEIRA REAL, ALFACE, ALFAFA - AZEVEM, ALHO, ALHO NOBRE, ALHO PORÓ, ALMEIRÃO, AMEIXA, AMENDOIM, AMORA, ARROZ, AVEIA, BANANA, BATATA INGLESA, BATATA-DOCE, BERINJELA, BETERRABA, BROCOLIS, CAFÉ, CANA-DE-AÇUCAR, CANOLA DE SEQUEIRO, CAQUI, CEBOLA, CENOURA, CENOURA (ORGÂNICA), CENTEIO, CEVADA, CHICÓRIA/ESCAROLA, CHUCHU, CHÁ, COUVE, COUVE-FLOR, ERVA MATE, ERVILHACA, EUCALIPTO - TRATOS CULTURAIS, EUCALIPTO FLORESTAMENTO E REFLORESTAMENTO, FEIJÃO, FIGO, FLORES, FUMO, GIRASSOL, GOIABA, GRAMA, HORTALIÇAS, KIWI, LARANJA, LIMÃO, MANDIOCA, MANDIOQUINHA (BAROA, SALSA, AIPO), MARACUJA, MAÇÃ, MELANCIA, MELÃO, MILHO (acima de 25 mil), MILHO (até 25 mil), MORANGO, MOSTARDA, MUDAS DIVERSAS, NABO, NECTARINA, NOZ, PAINÇO, PALMEIRA REAL, PASTAGEM, PEPINO, PERA, PIMENTA, PIMENTÃO, PINUS, PITAYA, PLANTAS ORNAMENTAIS, PORONGO (CUIA, CABAÇA), PÊSSEGO, QUIABO, RABANETE, REPOLHO, RÚCULA, SALSA, SOJA, SORGO, TANGERINA, TOMATE CEREJA, TOMATE RASTEIRO, TRIGO, TRITICALE, URUCUM, UVA, UVA (ORGÂNICA), VAGEM"],

    ["L023", "FUNCAFÉ - CUSTEIO AGRICOLA", "Cresol", "Custeio", "agricola,cafe,custeio", "Conforme análise", "13", "13", "20", "0", "0", "3000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: Isento | Prazo: Até 20 meses", "Tratos culturais.; Colheita das lavouras, incluindo as despesas com a aquisição de insumos, mão de obra, operações com máquinas e equipamentos, arrumação, transporte para o terreiro e secagem.", "CAFÉ"],

    ["L024", "FUNCAFÉ - CUSTEIO PARA AQUISIÇÃO DE CAFÉ (FAC)", "Cresol", "Custeio", "cafe,custeio", "Conforme análise", "14.5", "14.5", "12", "0", "0", "3000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: Isento | Prazo: Até 12 meses", "Café verde adquirido diretamente de produtores rurais ou de suas cooperativas ou indiretamente de produtores rurais.; Taxas e Encargos:; Taxa de Juros Anual: 14,5% a.a.; IOF Complementar: Isento; Limites e Prazos:; Limite de Crédito por Beneficiário: Até R$ 3.000.000,00; Percentual de Financiamento: 100% do valor dos itens financiáveis; Prazo Total: Até 12 meses; Prazo de Carência: Sem carência; Modalidades e Códigos:; Modalidades Colmeia:; FUNCAFÉ; 20129; Normas e Regulamentações:; Norma Completa: MCR 1 - 9 - 3; Custeio; FUNCAFÉ - CAPITAL DE GIRO; Objetivo: Custeio; Taxa: 14,5% a.a.; Limite: Até R$ 3.000.000,00; Público: TODOS OS PÚBLICOS; Culturas: -; Ocultar detalhes", ""],

    ["L025", "FUNCAFÉ - CAPITAL DE GIRO", "Cresol", "Custeio", "cafe,custeio", "Conforme análise", "14.5", "14.5", "12", "0", "0", "3000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: Isento | Prazo: Até 12 meses", "Capital de giro; Taxas e Encargos:; Taxa de Juros Anual: 14,5% a.a.; IOF Complementar: Isento; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 3.000.000,00; Percentual de Financiamento: 100% do valor dos itens financiáveis; Prazo Total: Até 12 meses; Modalidades e Códigos:; Modalidades Colmeia:; FUNCAFÉ; 20130; Normas e Regulamentações:; Norma Completa: MCR 1 - 9 - 3; Custeio; FUNCAFÉ - COMERCIALIZAÇÃO; Objetivo: Custeio; Taxa: 13%; Limite: Até R$ 3.000.000,00; Público: TODOS OS PÚBLICOS; Culturas: -; Ocultar detalhes", ""],

    ["L026", "FUNCAFÉ - COMERCIALIZAÇÃO", "Cresol", "Custeio", "cafe,custeio", "Conforme análise", "13", "13", "12", "0", "0", "3000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Custeio | IOF: Isento | Prazo: Até 12 meses", "Despesas próprias da fase sucessiva à coleta da produção, inclusive estocagem; Taxas e Encargos:; Taxa de Juros Anual: 13%; IOF Complementar: Isento; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 3.000.000,00; Percentual de Financiamento: 100% do valor dos itens financiáveis; Prazo Total: Até 12 meses; Condições de Amortização: Pagamento em duas parcelas, a primeira em até 180 dias e a segunda em 360 dias; Modalidades e Códigos:; Modalidades Colmeia:; FUNCAFÉ; 20130; Normas e Regulamentações:; Norma Completa: MCR 1 - 9 - 3; Pronaf Mais Alimentos (Faixa I); Objetivo: Investimento; Taxa: 3%; Limite: Até R$ 250.000,00; Público: PRONAF; Culturas: -; Ocultar detalhes", ""],

    ["L027", "Pronaf Mais Alimentos (Faixa I)", "BNDES / Cresol", "Investimento", "investimento", "Sem limite/R$ 500 mil", "3", "3", "120", "0", "0", "450000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | Sistemática: BNDES, BRDE e Poupança Controlada | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Cultivo protegido - estufas; Silos e armazéns para grãos, frutas, tubérculos, bulbos, hortaliças e fibras; Resfriadores de leite e ordenhadeiras; Aquicultura e pesca; Avicultura de postura não integrada; Caprinocultura e Ovinocultura; Itens relacionados a conectividade no campo; Equipamentos adaptados a pessoas com deficiência, incluindo cadeiras de rodas motorizadas; Importados novos, desde que não haja similar nacional e com a documentação exigida pelo BNDES", ""],

    ["L028", "Pronaf Tratores e Colheitadeiras", "BNDES / Cresol", "Investimento", "investimento,mecanizacao", "Sem limite/R$ 500 mil", "5", "5", "84", "0", "0", "450000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 7 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Aquisição de tratores novos até 80 CV e implementos associados, colheitadeiras e suas plataformas de corte, assim como máquinas agrícolas autopropelidas para pulverização e adubação;; Quando se tratar de tratores, colheitadeiras e máquinas agrícolas autopropelidas para pulverização e adubação, também deve conter o código Mais Alimentos (MDA), observando a descrição mínima e valor máximo de cada item, além do CFI;; Importados novos desde que não haja similar nacional e com documentação exigida pelo BNDES;; Taxas e Encargos:; Taxa de Juros Anual: 5%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: Até R$ 250.000,00 para os demais empreendimentos e finalidades;Itens usados: até R$250.000,00 quando se tratar de colheitadeira automotriz, e de R$160.000,00 para os demais casos;Até R$ 450.000,00 para atividades de suinocultura, avicultura, aquicultura, carcinicul", ""],

    ["L029", "Pronaf Mais Alimentos (Faixa II)", "BNDES / Cresol", "Investimento", "investimento", "Sem limite/R$ 500 mil", "8", "8", "120", "0", "0", "450000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 5 anos (para aquisição de caminhonetes de carga e motocicletas adaptadas à atividade rural);Até 8 anos (para aquisição isolada de matrizes, reprodutores, animais de serviço, sêmen, óvulos e embriões);Até 10 anos (para os demais itens financiáveis) | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Reforma de trator/máquina agrícola; Galpão de armazenagem de máquinas e equipamentos; Construção de sala de ordenha; Máquinas e equipamentos; Caminhonetes de carga (exceto veículos de cabine dupla e motocicletas adequadas às condições rurais); Matrizes e reprodutores; Regularização Fundiária de Imóvel Rural; Importados novos (desde que não haja similar nacional e com documentação exigida pelo BNDES); Restrições:; Caminhonetes de passageiros, caminhonetes mistas e jipes não são financiáveis.; Para caminhões, é necessário código CFI e seguir as regras do Mais Alimentos.; Observar regras do MCR para aquisição de camionetes (modelo, rendas, declaração de 120 dias, NF de fabricante).; Taxas e Encargos:; Taxa de Juros Anual: 8%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: Até R$ 450.000,00 (para suinocultura, avicultura, aquicultura, carcinicultura e fruticu", ""],

    ["L030", "Pronaf Mais alimentos - Máquinas (Faixa III)", "BNDES / Cresol", "Investimento", "investimento,mecanizacao", "Sem limite/R$ 500 mil", "2.5", "2.5", "120", "0", "0", "100000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$150.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Aquisição de máquinas, equipamentos e implementos por Beneficiárias Finais cuja renda bruta familiar anual seja inferior a R$ 150.000,00 (cem mil reais); Taxas e Encargos:; Taxa de Juros Anual: 2,5%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: Até R$100.000,00 para máquinas e equipamentos de pequeno porte.; Prazo Total: Até 10 anos; Prazo de Carência: Até 3 anos para máquinas, equipamentos e implementos; Modalidades e Códigos:; Modalidades Colmeia:; BNDES 10038; BRDE 20253; Normas e Regulamentações:; Norma Completa: MCR 1-10 e 1-7-6; Circular BNDES: CIRCULAR SUP/ADIG Nº 64/2025-BNDES; Pronaf Habitação Rural (Faixa II); Objetivo: Investimento; Taxa: 8%; Limite: Até R$ 100.000,00; Público: PRONAF; Culturas: -; Ocultar detalhes", ""],

    ["L031", "Pronaf Habitação Rural (Faixa II)", "BNDES / Cresol", "Investimento", "investimento", "Sem limite/R$ 500 mil", "8", "8", "120", "0", "0", "100000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Construção ou reforma de moradia em imóvel rural de propriedade do mutuário ou de terceiro;; Devendo o CPF de ambos constar como titular em CAF válida, observado que cada mutuário somente pode ter uma operação “em ser” para essa finalidade;; Que seja definida no projeto técnico a viabilidade econômica das atividades desenvolvidas na unidade produtiva do mutuário para pagamento do crédito;; E que, no caso de o objeto do financiamento ser realizado em imóvel rural de terceiro, o proprietário deve avalizar a operação de crédito e concordar em ceder formalmente ao mutuário o local da construção ou a moradia a ser reformada, por prazo não inferior a 25 (vinte e cinco) anos.; Taxas e Encargos:; Taxa de Juros Anual: 8%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 100.000,00; Prazo Total: Até 10 anos; Prazo de Carência: Até 3 anos; Modalidades e Códigos:; M", ""],

    ["L032", "Pronaf Agroindústria (Faixa II)", "BNDES / Cresol", "Investimento", "infraestrutura,investimento", "Sem limite/R$ 500 mil", "8", "8", "120", "0", "0", "50000000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar (últimos 12 meses) até R$500.000,00; Mínimo 50% da renda bruta familiar de exploração agropecuária e não agropecuária; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | Sistemática: BNDES/Poupança | IOF: 0,38% | Prazo: Até 5 anos para financiamento de caminhonetes de carga;Até 10 anos para os demais empreendimentos. | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Investimentos em infraestrutura, beneficiamento, armazenagem, processamento e comercialização da produção agropecuária, produtos florestais e do extrativismo, ou de produtos artesanais, e a exploração de turismo rural (mínimo 80% da produção beneficiada, processada ou comercializada deve ser própria).; Aquisição de equipamentos e programas de informática.; Capital de giro associado (limitado a 35% do financiamento para investimentos).; Cooperativas da agricultura familiar, que apresentem DAP pessoa jurídica ativa ou RICAF ativo para essa forma de organização, e que comprovem que, no mínimo, 75% de seus participantes ativos são Beneficiárias Finais do PRONAF.", ""],

    ["L033", "Pronaf Bioeconomia (Silvicultura)", "BNDES / Cresol", "Investimento", "investimento,sustentabilidade", "Sem limite/R$ 500 mil", "8", "8", "144", "0", "0", "250000", "Apresentação de DAP-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido:; PRONAF; COOPERATIVAS", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | Prazo: Até 12 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Silvicultura (implantar ou manter povoamentos florestais geradores de diferentes produtos, madeireiros e não madeireiros); Taxas e Encargos:; Taxa de Juros Anual: 8%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 250.000,00; Prazo Total: Até 12 anos; Prazo de Carência: Até 8 anos; Modalidades e Códigos:; Normas e Regulamentações:; Norma Completa: MCR 1-10-16; Circular BNDES: CIRCULAR SUP/ADIG Nº 64/2025-BNDES; Pronaf Bioeconomia (Faixa I); Objetivo: Investimento; Taxa: 3%; Limite: Até R$ 250.000,00; Público: PRONAF,COOPERATIVAS; Culturas: -; Ocultar detalhes", ""],

    ["L034", "Pronaf Bioeconomia (Faixa I)", "BNDES / Cresol", "Investimento", "investimento,sustentabilidade", "Sem limite/R$ 500 mil", "3", "3", "120", "0", "0", "250000", "Apresentação de CAF-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$ 500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido:; PRONAF; COOPERATIVAS", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | Sistemática: BNDES, BRDE, Poupança Controlada | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Correção de solo; Turismo rural; Projetos de adequação ambiental; Formação e recuperação de pastagens destinadas à alimentação animal; Itens para irrigação; Infraestrutura de captação, armazenamento e distribuição de água (poço artesiano)", ""],

    ["L035", "Pronaf Bioeconomia (Sistemas Fotovoltaicos)", "BNDES / Cresol", "Investimento", "investimento,sustentabilidade", "Sem limite/R$ 500 mil", "3", "3", "120", "0", "0", "250000", "Apresentação de CAF-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | Sistemática: Finame | IOF: 0% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Aquisição de sistemas, painéis e geradores fotovoltaicos.; Novos que constem do Credenciamento Finame (CFI) do Sistema BNDES.; Importados novos sem similar de fabricação nacional ou usados.; Sistemática Finame, somente Nota fiscal do equipamento, não é permitido nota de serviços ou instalação separado", ""],

    ["L036", "Pronaf Mulher", "BNDES / Cresol", "Investimento", "investimento", "Sem limite/R$ 500 mil", "3", "3", "0", "0", "0", "500000", "Conforme política de crédito da Cresol", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Os mesmos previstos para a Linha PRONAF Mais Alimentos. | Norma: CIRCULAR SUP/ADIG Nº 48/2024-BNDES", "Atendimento de propostas de crédito de mulheres agricultoras, conforme projeto técnico ou proposta simplificada.; Finalidade: Investimento em atividades agrícolas.; Taxas e Encargos:; Taxa de Juros Anual: 3% a 8%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 50.000,00 (renda bruta familiar até R$100.000,00) e R$ 450.000,00 (renda familiar até R$ 500.000,00); Prazo Total: Os mesmos previstos para a Linha PRONAF Mais Alimentos.; Modalidades e Códigos:; Modalidades Colmeia:; BNDES 10023; BRDE 20195; BRDE Coletivo 20191; Normas e Regulamentações:; Norma Completa: MCR 1-10 e 1-7-6; CIRCULAR SUP/ADIG Nº 48/2024-BNDES; Circular BNDES: CIRCULAR SUP/ADIG Nº 48/2024-BNDES; Pronaf Microcrédito Grupo \"B\"; Objetivo: Investimento; Taxa: 0.5%; Limite: Até R$ 12.000,00; Público: PRONAF; Culturas: -; Ocultar detalhes", ""],

    ["L037", "Pronaf Microcrédito Grupo \"B\"", "BNDES / Cresol", "Investimento", "investimento", "Sem limite/R$ 500 mil", "0.5", "0.5", "36", "0", "0", "12000", "CAF-Pronaf enquadramento B.Beneficiários cuja renda bruta familiar anual não seja superior a R$50.000,00 e que não contratem trabalho assalariado permanente.; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | IOF: 0.38% | Prazo: Até 3 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Investimentos em atividades agropecuárias e não agropecuárias desenvolvidas no estabelecimento rural ou em áreas comunitárias rurais próximas para portadores de CAF \"B\".; Taxas e Encargos:; Taxa de Juros Anual: 0.5%; IOF Complementar: 0.38%; Limites e Prazos:; Limite de Crédito por Beneficiário: Até R$ 12.000,00 por ano agrícola, independente do nº de operações.; Prazo Total: Até 3 anos; Modalidades e Códigos:; Modalidades Colmeia:; BNDES 10026; Normas e Regulamentações:; Norma Completa: MCR 1-10 e 1-7-6; Circular BNDES: CIRCULAR SUP/ADIG Nº 64/2025-BNDES; Pronaf Agroecologia; Objetivo: Investimento; Taxa: 3%; Limite: Até R$ 250.000,00; Público: PRONAF; Culturas: -; Ocultar detalhes", ""],

    ["L038", "Pronaf Agroecologia", "BNDES / Cresol", "Investimento", "investimento,sustentabilidade", "Sem limite/R$ 500 mil", "3", "3", "0", "0", "0", "250000", "Apresentação de CAF-Pronaf; Exploração de terra em diferentes condições (proprietário, posseiro, etc.); Área de terra não superior a 4 módulos fiscais; Renda bruta familiar nos últimos 12 meses até R$500.000,00; Mínimo 50% da renda bruta familiar originada da exploração agropecuária e não agropecuária; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Os mesmos previstos para Linha PRONAF Mais Alimentos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Implantação e manutenção de sistemas de produção agroecológicos ou orgânicos; Taxas e Encargos:; Taxa de Juros Anual: 3%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 250.000,00 + Certificado de Propriedade Agroecológica; Prazo Total: Os mesmos previstos para Linha PRONAF Mais Alimentos; Modalidades e Códigos:; Modalidades Colmeia:; BNDES 10022; BRDE 20188; BRDE Coletivo 20193; Normas e Regulamentações:; Norma Completa: MCR 1-10 e 1-7-6; Circular BNDES: CIRCULAR SUP/ADIG Nº 64/2025-BNDES; Pronaf Jovem; Objetivo: Investimento; Taxa: 3%; Limite: Até R$ 35.000,00; Público: PRONAF; Culturas: -; Ocultar detalhes", ""],

    ["L039", "Pronaf Jovem", "BNDES / Cresol", "Investimento", "investimento", "Sem limite/R$ 500 mil", "3", "3", "120", "0", "0", "35000", "Maior de 16 anos e com até 29 anos de idade.; Integrante de unidade familiar.; Atender a uma ou mais condições do MCR 10-10.; Apresentação de CAF-Pronaf.; Público_resumido: PRONAF", "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 64/2025-BNDES", "Mesmas finalidades previstas no Pronaf Mais Alimentos; Taxas e Encargos:; Taxa de Juros Anual: 3%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 35.000,00 por ano agrícola; limitado a 3 financiamentos.; Prazo Total: Até 10 anos; Prazo de Carência: Até 2 anos, podendo ser elevado para até 5 anos, dependendo da atividade e comprovação técnica.; Modalidades e Códigos:; Modalidades Colmeia:; BNDES 10027; BRDE 20190; Normas e Regulamentações:; Circular BNDES: CIRCULAR SUP/ADIG Nº 64/2025-BNDES; Pronamp; Objetivo: Investimento; Taxa: 10%; Limite: Até R$ 600.000,00; Público: PRONAMP; Culturas: -; Ocultar detalhes", ""],

    ["L040", "Pronamp", "BNDES / Cresol", "Investimento", "investimento", "R$ 500 mil/R$ 3.5 mi", "10", "10", "96", "0", "0", "600000", "Ser produtor rural (proprietário, posseiro, arrendatário ou parceiro); Renda bruta anual até R$ 3.500.000,00; Mínimo 80% da renda bruta anual originária da atividade agropecuária; Público_resumido: PRONAMP", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 8 anos | Norma: CIRCULAR SUP/ADIG Nº 66/2025-BNDES", "Crédito para inversões fixas e semifixas em bens e serviços relacionados com a atividade agropecuária de médios produtores com renda bruta anual de até R$ 3.500.000,00, no mínimo 80% da renda agrícola comprovada.; Construção, reforma ou ampliação de benfeitorias e instalações permanentes;; Obras de irrigação, açudagem e drenagem;; Florestamento, reflorestamento, supressão de vegetação autorizada pelo órgão ambiental competente, destoca e manejo florestal sustentável;; Formação de lavouras permanentes;; Formação ou recuperação de pastagens;; Telefonia rural, e equipamentos e demais itens relacionados a sistemas de conectividade no campo;; Aquisição de equipamentos empregados na medição de lavouras;; Recuperação ou reforma de máquinas, tratores, embarcações, veículos e equipamentos, desde que destinados especificamente à atividade agropecuária;; Aquisição de veículos, observado o disposto ", ""],

    ["L041", "Pronamp Finame", "BNDES / Cresol", "Investimento", "investimento,mecanizacao", "R$ 500 mil/R$ 3.5 mi", "10", "10", "96", "0", "0", "600000", "Propriedade rural, posse, arrendamento ou parceria; Renda bruta anual até R$ 3.500.000,00; Mínimo 80% da renda bruta anual originária da atividade agropecuária; Público_resumido: PRONAMP", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | Sistemática: Finame (Validação da Nota Fiscal Antes da Liberação) | IOF: 0% | Prazo: Até 8 anos | Norma: CIRCULAR SUP/ADIG Nº 66/2025-BNDES", "Aquisição de máquinas e equipamentos novos com CFI, ou usados (desde que não se enquadrem no MODERFROTA), e importados novos (desde que não haja similar nacional e com a documentação exigida pelo BNDES).; Restrições:; Máquinas e equipamentos usados enquadrados no MODERFROTA; Equipamentos importados novos sem documentação exigida pelo BNDES; Equipamentos importados novos com similar nacional disponível", ""],

    ["L042", "RenovAgro Demais", "BNDES / Cresol", "Investimento", "investimento,sustentabilidade", "Conforme análise", "10", "10", "144", "0", "0", "5000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Prazo será de acordo com o empreendimento financiado:Até 12 anos (florestas);Até 5 anos (aquisição de animais);Até 10 anos (demais finalidades). | Norma: CIRCULAR SUP/ADIG Nº 69/2025-BNDES", "RenovAgro Manejo de solos: correção de solo; RenovAgro Orgânico: implantação e melhoramento de sistemas orgânicos de produção agropecuária; RenovAgro Florestas: implantação, manutenção e melhoramento do manejo de florestas comerciais; RenovAgro Manejo de Resíduos: sistemas de manejo de resíduos oriundos da produção animal para a geração de energia e compostagem; RenovAgro Integração: implantação e melhoramento de sistemas de integração lavoura-pecuária, lavoura-floresta, pecuária-floresta ou lavoura-pecuária-floresta e de sistemas agroflorestais; Taxas e Encargos:; Taxa de Juros Anual: 10%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 5.000.000,00; Percentual de Financiamento: Até 100%; Prazo Total: Prazo será de acordo com o empreendimento financiado:Até 12 anos (florestas);Até 5 anos (aquisição de animais);Até 10 anos (demais finalidades).; Prazo d", ""],

    ["L043", "RenovAgro Ambiental e RenovAgro Recuperação e Pastagem", "BNDES / Cresol", "Investimento", "investimento,sustentabilidade", "Conforme análise", "8.5", "8.5", "144", "0", "0", "5000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0.38% | Prazo: Até 12 anos (Florestas); Até 10 anos (Demais Finalidades) | Norma: CIRCULAR SUP/ADIG Nº 69/2025-BNDES", "Adequação à legislação ambiental.; Recuperação da reserva legal.; Recuperação de áreas de preservação permanente.; Recuperação de áreas degradadas.; Taxas e Encargos:; Taxa de Juros Anual: 8.5%; IOF Complementar: 0.38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 5.000.000,00; Percentual de Financiamento: Até 100% dos itens financiáveis; Prazo Total: Até 12 anos (Florestas); Até 10 anos (Demais Finalidades); Prazo de Carência: Até 8 anos (Florestas); Até 5 anos (Demais Finalidades); Condições de Amortização: Prazo será de acordo com o investimento; Modalidades e Códigos:; Modalidades Colmeia:; BRDE 20211; BNDES 20292; Normas e Regulamentações:; Circular BNDES: CIRCULAR SUP/ADIG Nº 69/2025-BNDES; Inovagro; Objetivo: Investimento; Taxa: 12.5%; Limite: Até R$ 4.000.000,00; Público: DEMAIS PRODUTORES,COOPERATIVAS; Culturas: -; Ocultar detalhes", ""],

    ["L044", "Inovagro", "BNDES / Cresol", "Investimento", "investimento", "Conforme análise", "12.5", "12.5", "120", "0", "0", "4000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | Sistemática: Finame | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 71/2025-BNDES", "Investimentos para incorporação de inovação tecnológica nas propriedades rurais.; Placas para energia solar com Finame.; Máquinas e equipamentos com código Finame.; Instalações (sala de ordenha, aviários e pocilgas) relacionadas à bovinocultura de leite, suínos e aves.; Construção e ampliação de instalações para guarda de máquinas, implementos agrícolas e estocagem de insumos agropecuários.", ""],

    ["L045", "Moderfrota", "BNDES / Cresol", "Investimento", "investimento,mecanizacao", "Conforme análise", "13.5", "13.5", "84", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0% | Prazo: Até 7 anos para itens novos;Até 4 anos para itens usados | Norma: CIRCULAR SUP/ADIG Nº 73/2025-BNDES", "Itens novos com código CFI: tratores e implementos associados, colheitadeiras e suas plataformas de corte, equipamentos para preparo, secagem e beneficiamento de café, máquinas agrícolas autopropelidas para pulverização e adubação.; Itens usados: tratores e colheitadeiras com idade máxima de 8 e 10 anos, respectivamente, isolados ou associados com sua plataforma de corte, máquinas agrícolas autopropelidas para pulverização e adubação, plantadeiras usadas e semeadoras usadas com idade máxima de 5 anos.; Taxas e Encargos:; Taxa de Juros Anual: 13,5%; IOF Complementar: 0%; Limites e Prazos:; Percentual de Financiamento: Até 85% do valor dos bens objeto do financiamento. No caso de maquinário que utilize biometano como combustível, este limite poderá ser elevado para até 100%.; Prazo Total: Até 7 anos para itens novos;Até 4 anos para itens usados; Prazo de Carência: Até 12 meses; Modalidades", ""],

    ["L046", "Moderfrota Pronamp", "BNDES / Cresol", "Investimento", "investimento,mecanizacao", "R$ 500 mil/R$ 3.5 mi", "12.5", "12.5", "84", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0% | Prazo: Até 7 anos para itens novos;Até 4 anos para itens usados | Norma: CIRCULAR SUP/ADIG Nº 73/2025-BNDES", "Itens novos com código CFI: tratores e implementos associados, colheitadeiras e suas plataformas de corte, equipamentos para preparo, secagem e beneficiamento de café, máquinas agrícolas autopropelidas para pulverização e adubação; Itens usados: tratores e colheitadeiras com idade máxima de 8 e 10 anos, respectivamente, isolados ou associados com sua plataforma de corte, máquinas agrícolas autopropelidas para pulverização e adubação, plantadeiras usadas e semeadoras usadas com idade máxima de 5 anos; Taxas e Encargos:; Taxa de Juros Anual: 12,5%; IOF Complementar: 0%; Limites e Prazos:; Percentual de Financiamento: 100%; Prazo Total: Até 7 anos para itens novos;Até 4 anos para itens usados; Prazo de Carência: Até 12 meses; Modalidades e Códigos:; Modalidades Colmeia:; BNDES 10012; BRDE 20185; LCA Controlada 20695; Normas e Regulamentações:; Circular BNDES: CIRCULAR SUP/ADIG Nº 73/2025-BN", ""],

    ["L047", "Proirriga", "BNDES / Cresol", "Investimento", "investimento,irrigacao", "Conforme análise", "12.5", "12.5", "96", "0", "0", "3500000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 8 anos | Norma: CIRCULAR SUP/ADIG N° 72/2025-BNDES", "Sistemas de irrigação (inclusive infraestrutura elétrica, reserva de água e equipamento para monitoramento da umidade no solo); Pivôs de Irrigação (contendo código Finame); Cultivo protegido; Taxas e Encargos:; Taxa de Juros Anual: 12,5%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 3.500.000,00; Limite de Crédito por Empreendimento Coletivo: R$ 10.500.000,00; Percentual de Financiamento: Até 100% dos itens financiáveis; Prazo Total: Até 8 anos; Prazo de Carência: Até 1 ano; Modalidades e Códigos:; Modalidades Colmeia:; BNDES 20326; BRDE 20212; Normas e Regulamentações:; Circular BNDES: CIRCULAR SUP/ADIG N° 72/2025-BNDES; PCA Grãos; Objetivo: Investimento; Taxa: 8,5% para financiamento de uma ou mais unidades de armazenagem de grãos que somadas não ultrapassem 12.000 toneladas; Limite: Até R$ 50.000.000,00; Público: DEMAIS PRODUTORES,COOPERATIVAS; Cu", ""],

    ["L048", "PCA Grãos", "BNDES / Cresol", "Investimento", "armazenagem,investimento", "Conforme análise", "8.5", "8.5", "120", "0", "0", "50000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 70/2025-BNDES", "Ampliação, modernização, reforma e construção de novos armazéns destinados à guarda de grãos de até 12.000 toneladas, frutas, tubérculos, bulbos, hortaliças, fibras e açúcar.", "Grãos, Frutas, Tubérculos, Bulbos, Hortaliças, Fibras, Açúcar"],

    ["L049", "PCA Demais", "BNDES / Cresol", "Investimento", "armazenagem,investimento", "Conforme análise", "10", "10", "120", "0", "0", "200000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 70/2025-BNDES", "Ampliação, modernização, reforma e à construção de novos armazéns destinados à guarda de grãos que ultrapasse 12.000 toneladas, frutas, tubérculos, bulbos, hortaliças, fibras e açúcar.; Taxas e Encargos:; Taxa de Juros Anual: 10%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 50.000.000,00 (para investimentos em armazenagem de grãos);R$ 25.000.000,00 (para demais itens financiáveis);R$ 200.000.000,00 (para armazenagem de grãos em cooperativas).; Limite de Crédito por Empreendimento Coletivo: R$ 200.000.000,00 (para armazenagem de grãos em cooperativas); Percentual de Financiamento: Até 100% dos itens financiáveis; Prazo Total: Até 10 anos; Prazo de Carência: Até 2 anos; Modalidades e Códigos:; Modalidades Colmeia:; BNDES 10102; BRDE 20201; BRDE Finame 20512; Normas e Regulamentações:; Circular BNDES: CIRCULAR SUP/ADIG Nº 70/2025-BNDES; Investimento; P", ""],

    ["L050", "PRODECOOP", "BNDES / Cresol", "Investimento", "investimento", "Conforme análise", "13.5", "13.5", "120", "0", "0", "150000000", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 74/2025-BNDES", "Projetos que visem incrementar a competitividade do complexo agroindustrial das cooperativas brasileiras, por meio da modernização dos sistemas produtivos e de comercialização.; Taxas e Encargos:; Taxa de Juros Anual: 13,5%; IOF Complementar: 0,38%; Limites e Prazos:; Limite de Crédito por Beneficiário: R$ 150.000.000,00 por cooperativa, em uma ou mais operações.; Percentual de Financiamento: Até 90% do valor do projeto.; Prazo Total: Até 10 anos; Prazo de Carência: Até 2 anos; Modalidades e Códigos:; Normas e Regulamentações:; Circular BNDES: CIRCULAR SUP/ADIG Nº 74/2025-BNDES; CREDITO RURAL INVESTIMENTO - TFB; Objetivo: Investimento; Taxa: 2,8% + 0,95% + TFB; Limite: Não informado; Público: TODOS OS PÚBLICOS; Culturas: -; Ocultar detalhes", ""],

    ["L051", "CREDITO RURAL INVESTIMENTO - TFB", "BNDES / Cresol", "Investimento", "investimento", "Conforme análise", "2.8", "2.8", "120", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 19/2024-BNDES", "Construção, reforma ou ampliação de benfeitorias e instalações permanentes.; Obras de irrigação, açudagem, drenagem.; Florestamento, reflorestamento, destoca e manejo florestal sustentável.; Formação de lavouras permanentes.; Formação ou recuperação de pastagens.; Eletrificação, inclusive implantação de sistemas para geração e distribuição de energia renovável para consumo próprio.; Telefonia rural e equipamentos e itens relacionados a sistemas de conectividade no campo.; Práticas conservacionistas de uso, manejo e proteção do sistema solo-água-planta, incluindo correção de acidez e fertilidade do solo, e aquisição, transporte, aplicação e incorporação de insumos.; Aquisição de animais para reprodução ou cria.; Restrições:; Não financia máquinas, residência, e turismo; Taxas e Encargos:; Taxa de Juros Anual: 2,8% + 0,95% + TFB; IOF Complementar: 0,38%; Limites e Prazos:; Percentual de Fi", ""],

    ["L052", "CRÉDITO RURAL MÁQUINAS E EQUIPAMENTOS - TFB", "BNDES / Cresol", "Investimento", "investimento,mecanizacao", "Conforme análise", "2.1", "2.1", "120", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | Sistemática: Finame | IOF: 0% | Prazo: Até 10 anos | Norma: CIRCULAR SUP/ADIG Nº 19/2024-BNDES", "Aquisição isolada de máquinas e equipamentos novos para uso na atividade agropecuária. (Somente Novos com Finame)", ""],

    ["L053", "CRÉDITO RURAL COOPERATIVAS - TFB", "BNDES / Cresol", "Investimento", "investimento", "Conforme análise", "4.3", "4.3", "24", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: 0,38% | Prazo: Até 2 anos | Norma: CIRCULAR SUP/ADIG Nº 19/2024-BNDES", "Capital de giro para atender às necessidades operacionais de cooperativas de produtores rurais.; Taxas e Encargos:; Taxa de Juros Anual: 4,3% + 1,35% + TFB; IOF Complementar: 0,38%; Limites e Prazos:; Percentual de Financiamento: Até 100% dos itens financiáveis; Prazo Total: Até 2 anos; Prazo de Carência: Até 6 meses; Modalidades e Códigos:; Modalidades Colmeia:; BNDES Carência trimestral 20348; BNDES Carência Semestral 20349; Normas e Regulamentações:; Circular BNDES: CIRCULAR SUP/ADIG Nº 19/2024-BNDES; FCO RURAL TAXA FIXA - MICRO, PEQUENO E PEQUENO-MÉDIO; Objetivo: Investimento; Taxa: 9.05%; Limite: Não informado; Público: TODOS OS PÚBLICOS; Culturas: -; Ocultar detalhes", ""],

    ["L054", "FCO RURAL TAXA FIXA - MICRO, PEQUENO E PEQUENO-MÉDIO", "FCO / Cresol", "Investimento", "investimento", "Conforme análise", "9.05", "9.05", "120", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | Prazo: Até 10 anos", "", ""],

    ["L055", "INVESTIMENTO RECURSOS LIVRES COM REGISTRO NO SICOR", "Cresol", "Investimento", "investimento", "Conforme análise", "9", "9", "120", "0", "0", "0", "Conforme política de crédito da Cresol", "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda", "Ativa", new Date(), "Investimento | IOF: IOF complementar | Prazo: Poupança até 10 anos", "Os créditos concedidos com recursos livres podem ter por objeto operações de custeio, de investimento, de comercialização ou de industrialização, envolvendo quaisquer produtos de origem vegetal ou animal, inclusive os obtidos em atividades extrativistas.; Taxas e Encargos:; Taxa de Juros Anual: Poupança: de 9% a 25% a.a.RP: de 12% a 45% a.a.; IOF Complementar: IOF complementar; Limites e Prazos:; Percentual de Financiamento: 100% dos itens financiáveis, conforme projeto/orçamento; Prazo Total: Poupança até 10 anos; Prazo de Carência: Até 12 meses; Condições de Amortização: periodicidade mensal ou anual; Modalidades e Códigos:; Modalidades Colmeia:; POUPANÇA LIVRE 20156; Recursos Próprios com Registro no Sicor 7062; Normas e Regulamentações:; Cresol Empresarial BNDES (BNDES Pequenas Empresas); Objetivo: Investimento; Taxa: BNDES: Spread do BNDES: 1,35% a.a. + Spread da Cresol: de 3% a.a a", ""]
  ];

  linhas.forEach(linha => SHEET_LINHAS.appendRow(linha));

  // Formatar sheet
  SHEET_LINHAS.setColumnWidths(1, 19, 80);
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
    ["Últimas Linhas Adicionadas", "Pronaf Irrigação", "text", "Informação das novas linhas"],
    ["Link Consulta Crédito (SICOR/CACR)", "https://www.gov.br/pt-br/servicos/acessar-as-informacoes-de-operacoes-de-credito-rural-cacr", "url", "Link para consulta de crédito já tomado (SICOR/CACR)"]
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
            itensFinanciaveis: linha[headers.indexOf("Itens Financiáveis")] || "",
            culturas: linha[headers.indexOf("Culturas Financiadas")] || "",
            limiteDisponivel: Math.max(0, (parseInt(linha[headers.indexOf("Limite Máx (R$)")]) || 0) - (parametros.valorTomado || 0)),
            valorTomado: parametros.valorTomado || 0
          };
        } catch (e) {
          return null;
        }
      })
      .filter(item => item !== null)
      .filter(item => {
        // Respeita o limite já tomado: oculta linhas cujo teto já foi
        // atingido pelo valor já tomado pelo associado.
        if ((parametros.valorTomado || 0) > 0 && item.limiteMax > 0 && item.limiteDisponivel <= 0) {
          return false;
        }
        return true;
      });

    registrarConsulta(parametros, resultado.length);
    return resultado;
  } catch (e) {
    Logger.log("Erro em buscarLinhas: " + e);
    return [];
  }
}

/**
 * Converte texto monetário em número, interpretando "mil" (x1.000) e
 * "mi"/"milhão"/"milhões" (x1.000.000). Ex: "R$ 500 mil" -> 500000;
 * "R$ 3.5 mi" -> 3500000.
 */
function parseValorRenda(parte) {
  try {
    if (!parte) return null;
    const t = String(parte).toLowerCase().trim();

    const m = t.match(/([\d.,]+)/);
    if (!m) return null;

    let num = parseFloat(m[1].replace(/\s/g, "").replace(",", "."));
    if (isNaN(num)) return null;

    // Detecta a unidade. Verifica milhão antes de mil (pois "milhões"
    // começa com "mil"). "mi" como palavra isolada também é milhão.
    if (/milh|\bmi\b/.test(t)) {
      num = num * 1000000;
    } else if (/\bmil\b/.test(t)) {
      num = num * 1000;
    }
    return num;
  } catch (e) {
    return null;
  }
}

function validarRenda(renda, enquadramentoTexto) {
  try {
    if (!enquadramentoTexto || enquadramentoTexto === "") return true;
    if (typeof enquadramentoTexto !== "string") return true;

    const t = enquadramentoTexto.toLowerCase();

    // "Conforme análise": não há faixa fixa, não bloqueia
    if (t.includes("conforme")) return true;

    // Formato faixa "min/max" (ex: "Sem limite/R$ 500 mil")
    if (enquadramentoTexto.includes("/")) {
      const partes = enquadramentoTexto.split("/");
      const minTexto = partes[0].trim();
      const maxTexto = partes[1].trim();

      const min = minTexto.toLowerCase().includes("sem limite")
        ? 0 : (parseValorRenda(minTexto) || 0);
      const max = maxTexto.toLowerCase().includes("sem limite")
        ? Infinity : (parseValorRenda(maxTexto) || Infinity);

      return renda >= min && renda <= max;
    }

    // Formato "Acima de X" (ex: "Acima de R$ 3.5 mi")
    if (t.includes("acima")) {
      const min = parseValorRenda(enquadramentoTexto);
      if (min === null) return true;
      return renda > min;
    }

    // Formato "Até X"
    if (t.includes("até") || t.includes("ate")) {
      const max = parseValorRenda(enquadramentoTexto);
      if (max === null) return true;
      return renda <= max;
    }

    // Formato não reconhecido: não bloqueia
    return true;
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
      "Itens Financiáveis", "Culturas Financiadas", "Documentos Necessários", "Observações"
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
        itensFinanciaveis: sanitizarValor(linha[headers.indexOf("Itens Financiáveis")]),
        culturas: sanitizarValor(linha[headers.indexOf("Culturas Financiadas")])
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
      "Itens Financiáveis": dados.itensFinanciaveis || "",
      "Culturas Financiadas": dados.culturas || ""
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

// Parâmetro de configuração que guarda o link de consulta (SICOR/CACR)
const PARAM_LINK_CONSULTA = "Link Consulta Crédito (SICOR/CACR)";

/**
 * Retorna o link de consulta cadastrado na aba Configurações.
 * Se não houver, devolve o link oficial do CACR como padrão.
 */
function obterLinkConsulta() {
  try {
    const padrao = "https://www.gov.br/pt-br/servicos/acessar-as-informacoes-de-operacoes-de-credito-rural-cacr";
    if (!SHEET_CONFIG) return padrao;

    const dados = SHEET_CONFIG.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === PARAM_LINK_CONSULTA) {
        const valor = String(dados[i][1] || "").trim();
        return valor || padrao;
      }
    }
    return padrao;
  } catch (e) {
    Logger.log("Erro em obterLinkConsulta: " + e.toString());
    return "";
  }
}

/**
 * Salva (ou cria) o link de consulta na aba Configurações.
 */
function salvarLinkConsulta(url) {
  try {
    const dados = SHEET_CONFIG.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === PARAM_LINK_CONSULTA) {
        SHEET_CONFIG.getRange(i + 1, 2).setValue(url);
        return { sucesso: true };
      }
    }
    // Não existe ainda: cria a linha de parâmetro
    SHEET_CONFIG.appendRow([PARAM_LINK_CONSULTA, url, "url", "Link para consulta de crédito já tomado (SICOR/CACR)"]);
    return { sucesso: true };
  } catch (e) {
    Logger.log("Erro em salvarLinkConsulta: " + e.toString());
    return { sucesso: false, erro: e.toString() };
  }
}

// ============ ATUALIZAÇÃO DA BASE VIA ARQUIVO DA CRESOL (.docx) ============

/**
 * Recebe o formulário de upload (campo "arquivo" = .docx da Cresol),
 * extrai as linhas de crédito e reconstrói a aba Linhas.
 */
function processarArquivoCresol(form) {
  try {
    const blob = form && form.arquivo;
    if (!blob) return { sucesso: false, erro: "Nenhum arquivo recebido." };

    blob.setContentType("application/zip");
    const arquivos = Utilities.unzip(blob);
    let docXml = null;
    for (let i = 0; i < arquivos.length; i++) {
      if (arquivos[i].getName() === "word/document.xml") {
        docXml = arquivos[i].getDataAsString("UTF-8");
        break;
      }
    }
    if (!docXml) return { sucesso: false, erro: "Arquivo .docx inválido (document.xml não encontrado)." };

    // Extrai texto por parágrafo e decodifica entidades XML
    const texto = docXml
      .replace(/<\/w:p>/g, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&apos;/g, "'");
    const linhasTxt = texto.split("\n").map(s => s.trim());

    const blocos = _cresolMontarBlocos(linhasTxt);
    if (blocos.length === 0) {
      return { sucesso: false, erro: "Nenhuma linha encontrada no arquivo (formato inesperado)." };
    }

    // Processa TODAS as linhas (rurais e não-rurais). A escolha de quais
    // incluir é feita pelo gestor na tela de confirmação.
    const rows = [];
    const itens = [];
    let idNum = 0;
    for (let k = 0; k < blocos.length; k++) {
      const rec = _cresolParseBloco(blocos[k]);
      if (!rec.nome) continue;
      idNum++;
      rows.push(_cresolMapear(rec, idNum));
      itens.push({ idx: rows.length - 1, nome: rec.nome, rural: _cresolEhRural(rec.nome) });
    }
    if (rows.length === 0) return { sucesso: false, erro: "Nenhuma linha encontrada no arquivo." };

    // Grava em uma aba temporária (staging) - NÃO altera a base ainda.
    const staging = _cresolStagingSheet(true);
    _cresolEscreverEmSheet(staging, rows);

    return {
      sucesso: true,
      total: rows.length,
      qtdRural: itens.filter(function (i) { return i.rural; }).length,
      qtdNaoRural: itens.filter(function (i) { return !i.rural; }).length,
      itens: itens
    };
  } catch (e) {
    Logger.log("Erro em processarArquivoCresol: " + e.toString());
    return { sucesso: false, erro: e.toString() };
  }
}

/**
 * Aplica a atualização pendente (staging) na aba Linhas, criando antes um
 * backup da base atual. Mantém apenas o backup mais recente.
 */
function aplicarAtualizacaoCresol(selecionados) {
  try {
    const staging = SS.getSheetByName("Linhas_Staging");
    if (!staging) return { sucesso: false, erro: "Nenhuma atualização pendente. Faça o upload do arquivo primeiro." };
    const vals = staging.getDataRange().getValues();
    if (!vals || vals.length <= 1) return { sucesso: false, erro: "A pré-visualização está vazia. Refaça o upload." };

    if (!selecionados || !selecionados.length) {
      return { sucesso: false, erro: "Selecione ao menos uma linha para incluir." };
    }

    const headers = vals[0];
    const dados = vals.slice(1);
    // Filtra pelos índices escolhidos e re-sequencia os IDs (L001, L002, ...)
    const escolhidas = [];
    selecionados.forEach(function (i) {
      if (i >= 0 && i < dados.length) escolhidas.push(dados[i].slice());
    });
    if (escolhidas.length === 0) return { sucesso: false, erro: "Seleção inválida." };
    escolhidas.forEach(function (r, i) { r[0] = "L" + ("000" + (i + 1)).slice(-3); });

    // Backup: remove backups antigos e copia a base atual
    const sheets = SS.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().indexOf("Linhas_Backup") === 0) SS.deleteSheet(sheets[i]);
    }
    const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
    const backup = SHEET_LINHAS.copyTo(SS);
    backup.setName("Linhas_Backup " + carimbo);

    // Aplica na aba Linhas: cabeçalho + linhas escolhidas
    SHEET_LINHAS.clear();
    SHEET_LINHAS.getRange(1, 1, 1, headers.length).setValues([headers]);
    SHEET_LINHAS.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");
    SHEET_LINHAS.getRange(2, 1, escolhidas.length, headers.length).setValues(escolhidas);
    SHEET_LINHAS.setColumnWidths(1, headers.length, 80);

    SS.deleteSheet(staging);
    return { sucesso: true, linhas: escolhidas.length, backup: backup.getName() };
  } catch (e) {
    Logger.log("Erro em aplicarAtualizacaoCresol: " + e.toString());
    return { sucesso: false, erro: e.toString() };
  }
}

/**
 * Descarta a atualização pendente (remove o staging) sem alterar a base.
 */
function cancelarAtualizacaoCresol() {
  try {
    const staging = SS.getSheetByName("Linhas_Staging");
    if (staging) SS.deleteSheet(staging);
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

function _cresolStagingSheet(criarSeNao) {
  let sh = SS.getSheetByName("Linhas_Staging");
  if (!sh && criarSeNao) sh = SS.insertSheet("Linhas_Staging");
  return sh;
}

function _cresolMontarBlocos(linhas) {
  const idxs = [];
  for (let i = 0; i < linhas.length; i++) {
    if (linhas[i].indexOf("Nome da Linha de Crédito:") === 0) idxs.push(i);
  }
  const blocos = [];
  for (let k = 0; k < idxs.length; k++) {
    const start = idxs[k];
    const end = (k + 1 < idxs.length) ? idxs[k + 1] : linhas.length;
    blocos.push(linhas.slice(start, end));
  }
  return blocos;
}

function _cresolFieldAfter(block, label) {
  for (let i = 0; i < block.length; i++) {
    if (block[i].indexOf(label) === 0) return block[i].substring(label.length).trim();
  }
  return "";
}

function _cresolSection(block, startLabel, endLabels) {
  const out = [];
  let cap = false;
  for (let i = 0; i < block.length; i++) {
    const l = block[i];
    if (!cap && l.indexOf(startLabel) === 0) {
      cap = true;
      const rest = l.substring(startLabel.length).trim();
      if (rest) out.push(rest);
      continue;
    }
    if (cap) {
      let stop = false;
      for (let j = 0; j < endLabels.length; j++) { if (l.indexOf(endLabels[j]) === 0) { stop = true; break; } }
      if (stop) break;
      if (l.trim()) out.push(l);
    }
  }
  return out;
}

function _cresolParseBloco(b) {
  return {
    nome: _cresolFieldAfter(b, "Nome da Linha de Crédito:"),
    tipo: _cresolFieldAfter(b, "Tipo de Linha:"),
    objetivo: _cresolFieldAfter(b, "Objetivo:"),
    publico: _cresolFieldAfter(b, "Público_resumido:"),
    sistematica: _cresolFieldAfter(b, "Sistemática:"),
    taxa: _cresolFieldAfter(b, "Taxa de Juros Anual:"),
    iof: _cresolFieldAfter(b, "IOF Complementar:") || _cresolFieldAfter(b, "IOF:"),
    limite: _cresolFieldAfter(b, "Limite de Crédito por Beneficiário:"),
    prazo: _cresolFieldAfter(b, "Prazo Total:"),
    circular: _cresolFieldAfter(b, "Circular BNDES:"),
    requisitos: _cresolSection(b, "Requisitos:", ["Tipos:", "Financiamento:"]),
    financia: _cresolSection(b, "O que financia:", ["Produtos Beneficiados:", "Sistemática:", "Garantias:"]),
    produtos: _cresolSection(b, "Produtos Beneficiados:", ["Sistemática:", "Taxas e Encargos:", "Garantias:"])
  };
}

function _cresolEhRural(nome) {
  const n = nome.toUpperCase();
  const exclui = ["FINEP", "FUNGETUR", "PROMOVE SUL", "FUNDO CLIMA", "PROCAPCRED",
    "INVESTIMENTO EMPRESARIAL", "CRESOL EMPRESARIAL BNDES", "BNDES FINAME (FINAME BK"];
  for (let i = 0; i < exclui.length; i++) { if (n.indexOf(exclui[i]) !== -1) return false; }
  return true;
}

function _cresolNumTaxa(t) {
  if (!t) return "0";
  const m = String(t).match(/(\d+(?:[.,]\d+)?)/);
  return m ? m[1].replace(",", ".") : "0";
}

function _cresolNumLimite(t) {
  if (!t) return "0";
  const re = /R\$\s*([\d.]+(?:,\d+)?)/g;
  let m, max = 0, achou = false;
  while ((m = re.exec(t))) {
    const v = m[1].replace(/\./g, "").replace(",", ".");
    const n = parseInt(parseFloat(v), 10);
    if (!isNaN(n)) { achou = true; if (n > max) max = n; }
  }
  return achou ? String(max) : "0";
}

function _cresolPrazoMeses(t) {
  if (!t) return "0";
  let meses = 0, m;
  const reAnos = /(\d+)\s*anos?/g;
  while ((m = reAnos.exec(t))) { const v = parseInt(m[1], 10) * 12; if (v > meses) meses = v; }
  const reMes = /(\d+)\s*mes/g;
  while ((m = reMes.exec(t))) { const v = parseInt(m[1], 10); if (v > meses) meses = v; }
  return String(meses);
}

function _cresolEnquadramento(pub, nome) {
  const p = (pub || "").toUpperCase();
  const n = nome.toUpperCase();
  if (p.indexOf("PRONAF") !== -1 || n.indexOf("PRONAF") !== -1) return "Sem limite/R$ 500 mil";
  if (p.indexOf("PRONAMP") !== -1 || n.indexOf("PRONAMP") !== -1) return "R$ 500 mil/R$ 3.5 mi";
  return "Conforme análise";
}

function _cresolTags(nome, tipo, objetivo) {
  const s = (nome + " " + tipo + " " + objetivo).toLowerCase();
  const t = {};
  if (s.indexOf("custeio") !== -1) t["custeio"] = 1;
  const inv = ["investimento", "tratores", "colheitadeira", "máquina", "maquina", "finame", "moderfrota", "inovagro", "agroind", "habita", "bioeconomia", "pca", "prodecoop", "irriga", "renovagro"];
  for (let i = 0; i < inv.length; i++) { if (s.indexOf(inv[i]) !== -1) { t["investimento"] = 1; break; } }
  if (s.indexOf("pecuár") !== -1 || s.indexOf("pecuar") !== -1) t["pecuaria"] = 1;
  if (s.indexOf("agrícola") !== -1 || s.indexOf("agricola") !== -1) t["agricola"] = 1;
  if (s.indexOf("café") !== -1 || s.indexOf("cafe") !== -1 || s.indexOf("funcaf") !== -1) t["cafe"] = 1;
  if (s.indexOf("irriga") !== -1) t["irrigacao"] = 1;
  const mec = ["tratores", "colheitadeira", "máquina", "maquina", "moderfrota", "finame"];
  for (let j = 0; j < mec.length; j++) { if (s.indexOf(mec[j]) !== -1) { t["mecanizacao"] = 1; break; } }
  if (s.indexOf("pca") !== -1 || s.indexOf("armaz") !== -1) t["armazenagem"] = 1;
  const sus = ["renovagro", "sustentável", "sustentavel", "agroecolog", "bioeconomia", "ambiental", "clima"];
  for (let k = 0; k < sus.length; k++) { if (s.indexOf(sus[k]) !== -1) { t["sustentabilidade"] = 1; break; } }
  if (s.indexOf("agroind") !== -1 || s.indexOf("industrial") !== -1) t["infraestrutura"] = 1;
  let keys = Object.keys(t);
  if (keys.length === 0) keys = ["investimento"];
  keys.sort();
  return keys.join(",");
}

function _cresolOrgao(rec) {
  const src = ((rec.sistematica || "") + " " + (rec.circular || "")).toUpperCase();
  if (src.indexOf("BNDES") !== -1) return "BNDES / Cresol";
  if (src.indexOf("FCO") !== -1 || rec.nome.toUpperCase().indexOf("FCO") !== -1) return "FCO / Cresol";
  if (src.indexOf("POUPAN") !== -1) return "Cresol (Poupança Rural)";
  return "Cresol";
}

function _cresolDocumentos(pub) {
  if ((pub || "").toUpperCase().indexOf("PRONAF") !== -1) return "CAF/DAP-Pronaf, RG, CPF, projeto técnico, comprovante de renda";
  return "RG, CPF, documentação da propriedade, projeto técnico, comprovantes de renda";
}

function _cresolObs(rec) {
  const p = [];
  if (rec.objetivo) p.push(rec.objetivo);
  if (rec.sistematica) p.push("Sistemática: " + rec.sistematica);
  if (rec.iof) p.push("IOF: " + rec.iof);
  if (rec.prazo) p.push("Prazo: " + rec.prazo);
  if (rec.circular) p.push("Norma: " + rec.circular);
  return p.join(" | ").substring(0, 600);
}

function _cresolMapear(rec, idNum) {
  const rid = "L" + ("000" + idNum).slice(-3);
  const status = (rec.nome.toLowerCase().indexOf("fechado") !== -1) ? "Inativa" : "Ativa";
  const taxa = _cresolNumTaxa(rec.taxa);
  return [
    rid, rec.nome, _cresolOrgao(rec),
    (rec.objetivo || rec.tipo || "Crédito rural"),
    _cresolTags(rec.nome, rec.tipo, rec.objetivo),
    _cresolEnquadramento(rec.publico, rec.nome),
    taxa, taxa,
    _cresolPrazoMeses(rec.prazo), "0",
    "0", _cresolNumLimite(rec.limite),
    (rec.requisitos.join("; ")).substring(0, 600) || "Conforme política de crédito da Cresol",
    _cresolDocumentos(rec.publico),
    status, new Date(), _cresolObs(rec),
    (rec.financia.join("; ")).substring(0, 900),
    (rec.produtos.join(", ")).substring(0, 1500)
  ];
}

function _cresolEscreverEmSheet(sheet, rows) {
  const headers = [
    "ID", "Nome Linha", "Órgão/Instituição", "Finalidade Principal",
    "Finalidades (tags)", "Enquadramento (Renda Min/Max)", "Taxa Mín (%)",
    "Taxa Máx (%)", "Prazo (meses)", "Carência (meses)", "Limite Min (R$)",
    "Limite Máx (R$)", "Requisitos", "Documentos Necessários",
    "Status (Ativa/Inativa)", "Data Atualização", "Observações",
    "Itens Financiáveis", "Culturas Financiadas"
  ];
  sheet.clear();
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.setColumnWidths(1, headers.length, 80);
}

// ============ BASE DE ASSOCIADOS (basedepessoas.csv) ============

const PARAM_LINK_BASE = "Link Pasta Base de Associados";

function obterLinkBase() {
  try {
    const dados = SHEET_CONFIG.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === PARAM_LINK_BASE) return String(dados[i][1] || "").trim();
    }
    return "";
  } catch (e) { return ""; }
}

function salvarLinkBase(url) {
  try {
    const dados = SHEET_CONFIG.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === PARAM_LINK_BASE) {
        SHEET_CONFIG.getRange(i + 1, 2).setValue(url);
        return { sucesso: true };
      }
    }
    SHEET_CONFIG.appendRow([PARAM_LINK_BASE, url, "url", "Pasta/arquivo (Drive) com basedepessoas.csv"]);
    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.toString() }; }
}

/**
 * Lê o conteúdo do CSV a partir de um link de pasta ou arquivo do Drive,
 * ou de uma URL direta.
 */
function _lerCsvDoLink(link, filename) {
  // Link de pasta do Drive
  const mFolder = link.match(/folders\/([a-zA-Z0-9_\-]+)/);
  if (mFolder) {
    const folder = DriveApp.getFolderById(mFolder[1]);
    const it = folder.getFilesByName(filename);
    if (it.hasNext()) return it.next().getBlob().getDataAsString("UTF-8");
    const its = folder.getFiles();
    while (its.hasNext()) {
      const f = its.next();
      if (f.getName().toLowerCase().indexOf(".csv") !== -1) return f.getBlob().getDataAsString("UTF-8");
    }
    return null;
  }
  // Link de arquivo do Drive (/d/<id> ou ?id=<id>)
  const idm = link.match(/(?:\/d\/|id=)([a-zA-Z0-9_\-]+)/);
  if (idm) return DriveApp.getFileById(idm[1]).getBlob().getDataAsString("UTF-8");
  // URL direta
  if (/^https?:\/\//.test(link)) return UrlFetchApp.fetch(link).getContentText();
  // Talvez seja só um ID
  const mId = link.match(/^[a-zA-Z0-9_\-]{20,}$/);
  if (mId) return DriveApp.getFileById(link).getBlob().getDataAsString("UTF-8");
  return null;
}

/**
 * Busca o basedepessoas.csv no link configurado e regrava a aba Base.
 * Também é chamada pelo trigger automático.
 */
function atualizarBaseAssociados() {
  try {
    const link = obterLinkBase();
    if (!link) return { sucesso: false, erro: "Configure o link da pasta/arquivo da base na aba Administrativo." };

    const conteudo = _lerCsvDoLink(link, "basedepessoas.csv");
    if (!conteudo) return { sucesso: false, erro: "Arquivo basedepessoas.csv não encontrado no link informado." };

    const primeiraLinha = conteudo.split("\n")[0] || "";
    const delim = (primeiraLinha.split(";").length > primeiraLinha.split(",").length) ? ";" : ",";
    const dados = Utilities.parseCsv(conteudo, delim);
    if (!dados || dados.length < 2) return { sucesso: false, erro: "CSV vazio ou inválido." };

    // Normaliza largura das linhas
    const largura = dados[0].length;
    const norm = dados.map(function (r) {
      const linha = r.slice(0, largura);
      while (linha.length < largura) linha.push("");
      return linha;
    });

    SHEET_BASE.clear();
    SHEET_BASE.getRange(1, 1, norm.length, largura).setValues(norm);
    SHEET_BASE.getRange(1, 1, 1, largura).setFontWeight("bold").setBackground("#005c46").setFontColor("white");

    return { sucesso: true, registros: norm.length - 1, atualizado: new Date().toLocaleString("pt-BR") };
  } catch (e) {
    Logger.log("Erro em atualizarBaseAssociados: " + e.toString());
    return { sucesso: false, erro: e.toString() };
  }
}

function _numBR(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  let s = String(v).replace(/[^\d.,\-]/g, "");
  if (s.indexOf(",") !== -1) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/**
 * Busca um associado na aba Base por conta ou CPF/CNPJ (somente dígitos).
 * vl_anual_fonte_renda_total é a renda MENSAL (apesar do nome) -> anualiza.
 */
function buscarAssociado(termo) {
  try {
    if (!termo) return { sucesso: false, erro: "Informe a conta ou o CPF/CNPJ." };
    const alvo = String(termo).replace(/\D/g, "");
    if (!alvo) return { sucesso: false, erro: "Informe um número válido." };

    const dados = SHEET_BASE.getDataRange().getValues();
    if (!dados || dados.length < 2) {
      return { sucesso: false, erro: "Base de associados vazia. Atualize a base na aba Administrativo." };
    }
    const H = dados[0].map(function (h) { return String(h).trim(); });
    const iCpf = H.indexOf("nr_cpf_cnpj");
    const iConta = H.indexOf("nr_conta_corrente");
    const iNome = H.indexOf("nm_nome");
    const iRenda = H.indexOf("vl_anual_fonte_renda_total");
    const iTipo = H.indexOf("ds_pessoa_tipo");

    for (let r = 1; r < dados.length; r++) {
      const cpf = iCpf !== -1 ? String(dados[r][iCpf] || "").replace(/\D/g, "") : "";
      const conta = iConta !== -1 ? String(dados[r][iConta] || "").replace(/\D/g, "") : "";
      if ((cpf && cpf === alvo) || (conta && conta === alvo)) {
        const rendaMensal = iRenda !== -1 ? _numBR(dados[r][iRenda]) : 0;
        return {
          sucesso: true,
          nome: iNome !== -1 ? (dados[r][iNome] || "") : "",
          cpfCnpj: iCpf !== -1 ? (dados[r][iCpf] || "") : "",
          conta: iConta !== -1 ? (dados[r][iConta] || "") : "",
          tipo: iTipo !== -1 ? (dados[r][iTipo] || "") : "",
          rendaMensal: rendaMensal,
          rendaAnual: Math.round(rendaMensal * 12)
        };
      }
    }
    return { sucesso: false, erro: "Associado não encontrado na base." };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

// ---- Trigger automático de atualização da base ----

function criarTriggerBaseAssociados() {
  try {
    const ts = ScriptApp.getProjectTriggers();
    for (let i = 0; i < ts.length; i++) {
      if (ts[i].getHandlerFunction() === "atualizarBaseAssociados") ScriptApp.deleteTrigger(ts[i]);
    }
    ScriptApp.newTrigger("atualizarBaseAssociados").timeBased().everyDays(1).atHour(5).create();
    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.toString() }; }
}

function removerTriggerBaseAssociados() {
  try {
    const ts = ScriptApp.getProjectTriggers();
    let n = 0;
    for (let i = 0; i < ts.length; i++) {
      if (ts[i].getHandlerFunction() === "atualizarBaseAssociados") { ScriptApp.deleteTrigger(ts[i]); n++; }
    }
    return { sucesso: true, removidos: n };
  } catch (e) { return { sucesso: false, erro: e.toString() }; }
}

function statusTriggerBaseAssociados() {
  try {
    const ts = ScriptApp.getProjectTriggers();
    for (let i = 0; i < ts.length; i++) {
      if (ts[i].getHandlerFunction() === "atualizarBaseAssociados") return true;
    }
    return false;
  } catch (e) { return false; }
}

// ============ BASE DE CRÉDITO TOMADO / LIMITE (XLSX) ============

/**
 * Converte um blob XLSX em matriz de valores, criando um Google Sheet
 * temporário via Drive API e lendo a primeira aba.
 */
function _xlsxParaValores(blob) {
  const boundary = "xlsxbnd" + Date.now();
  const metadata = { name: "tmp_credito_" + Date.now(), mimeType: "application/vnd.google-apps.spreadsheet" };
  const ct = blob.getContentType() || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const pre = "--" + boundary + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) + "\r\n--" + boundary + "\r\nContent-Type: " + ct + "\r\n\r\n";
  const post = "\r\n--" + boundary + "--";
  const bytes = Utilities.newBlob(pre).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(post).getBytes());

  const res = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true", {
    method: "post",
    contentType: "multipart/related; boundary=" + boundary,
    payload: bytes,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("Falha ao converter XLSX (HTTP " + code + "): " + res.getContentText().substring(0, 200));
  }
  const fileId = JSON.parse(res.getContentText()).id;
  try {
    return SpreadsheetApp.openById(fileId).getSheets()[0].getDataRange().getValues();
  } finally {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
  }
}

/**
 * Importa o XLSX de crédito tomado e grava na aba BaseCredito as colunas
 * relevantes (mapeadas por posição):
 * D=Valor financiado, G=Produto, S=Atividade, W=CPF/CNPJ,
 * AB=Alíquota ProAgro, AC=Limite Custeio Disponível (C=Ano Safra).
 */
function processarArquivoCreditoBase(form) {
  try {
    const blob = form && form.arquivo;
    if (!blob) return { sucesso: false, erro: "Nenhum arquivo recebido." };

    const valores = _xlsxParaValores(blob);
    if (!valores || valores.length < 2) return { sucesso: false, erro: "Planilha vazia ou inválida." };

    const C = 2, D = 3, G = 6, S = 18, W = 22, AB = 27, AC = 28;
    const out = [["CPF/CNPJ", "Ano Safra", "Produto/Finalidade", "Atividade", "Valor Financiado", "Alíquota ProAgro", "Limite Custeio Disponível"]];
    for (let r = 1; r < valores.length; r++) {
      const row = valores[r];
      if (!row || row.length === 0) continue;
      const cpf = String(row[W] || "").trim();
      if (!cpf) continue;
      out.push([cpf, row[C] || "", row[G] || "", row[S] || "", _numBR(row[D]), row[AB] || "", _numBR(row[AC])]);
    }
    if (out.length < 2) return { sucesso: false, erro: "Nenhuma linha com CPF/CNPJ encontrada (confira o layout/colunas do arquivo)." };

    SHEET_BASE_CREDITO.clear();
    SHEET_BASE_CREDITO.getRange(1, 1, out.length, out[0].length).setValues(out);
    SHEET_BASE_CREDITO.getRange(1, 1, 1, out[0].length).setFontWeight("bold").setBackground("#005c46").setFontColor("white");

    return { sucesso: true, registros: out.length - 1, atualizado: new Date().toLocaleString("pt-BR") };
  } catch (e) {
    Logger.log("Erro em processarArquivoCreditoBase: " + e.toString());
    return { sucesso: false, erro: e.toString() };
  }
}

/**
 * Retorna os financiamentos já tomados do CPF/CNPJ informado.
 */
function buscarCreditoTomado(cpfCnpj) {
  try {
    if (!cpfCnpj) return { sucesso: false, itens: [] };
    const alvo = String(cpfCnpj).replace(/\D/g, "");
    const dados = SHEET_BASE_CREDITO.getDataRange().getValues();
    if (!dados || dados.length < 2) return { sucesso: false, erro: "Base de crédito tomado vazia.", itens: [] };

    const itens = [];
    let totalTomado = 0, limiteMax = 0;
    for (let r = 1; r < dados.length; r++) {
      const cpf = String(dados[r][0] || "").replace(/\D/g, "");
      if (cpf && cpf === alvo) {
        const vf = _numBR(dados[r][4]);            // valor financiado (bruto)
        const lim = _numBR(dados[r][6]);           // limite custeio disponível
        const aliqPct = _numBR(dados[r][5]);       // alíquota ProAgro (% a descontar)
        // ProAgro só existe em custeio agrícola: quando há alíquota, o valor
        // já tomado é o financiado descontado o percentual da alíquota.
        const fator = aliqPct > 0 ? (1 - aliqPct / 100) : 1;
        const vfTomado = vf * fator;
        totalTomado += vfTomado;
        if (lim > limiteMax) limiteMax = lim;
        itens.push({
          anoSafra: dados[r][1] || "",
          produto: dados[r][2] || "",
          atividade: dados[r][3] || "",
          valorFinanciado: vf,
          aliquotaProagro: dados[r][5] || "",
          valorTomado: vfTomado,
          limiteDisponivel: lim
        });
      }
    }
    return { sucesso: itens.length > 0, itens: itens, totalFinanciado: totalTomado, limiteDisponivel: limiteMax };
  } catch (e) {
    return { sucesso: false, erro: e.toString(), itens: [] };
  }
}

// ==================== INTERFACE WEB ====================

function doGet() {
  return obterHTML();
}

function obterHTML() {
  const dataAtualizacao = new Date().toLocaleDateString('pt-BR');
  const linkConsulta = obterLinkConsulta();
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
<strong>ℹ️ Como usar:</strong> Informe a conta ou o CPF/CNPJ para carregar os dados do associado, ou preencha manualmente.
</div>
<div style="border: 1px solid #e0e0e0; border-radius: 6px; padding: 16px; margin-bottom: 20px; background: #fafbfb;">
<strong style="color: #005c46; display: block; margin-bottom: 10px;">👤 Dados do Associado</strong>
<div class="grid-2">
<div class="form-group" style="margin-bottom: 12px;">
<label>🔢 Conta</label>
<input type="text" id="conta" placeholder="Nº da conta" onchange="window.buscarAssociado('conta')">
</div>
<div class="form-group" style="margin-bottom: 12px;">
<label>🆔 CPF/CNPJ</label>
<input type="text" id="cpfcnpj" placeholder="Somente números" onchange="window.buscarAssociado('cpf')">
</div>
</div>
<div class="form-group" style="margin-bottom: 10px;">
<label>📛 Nome do Associado</label>
<input type="text" id="nomeAssociado" readonly style="background-color: #f0f0f0;" placeholder="Preenchido automaticamente pela busca">
</div>
<button type="button" onclick="window.buscarAssociado('manual')" style="background: none; color: #005c46; border: 1px solid #005c46; padding: 8px 14px; border-radius: 6px; font-size: 13px; font-weight: 500;">🔍 Buscar Associado</button>
<div id="assocStatus" style="margin-top: 8px;"></div>
</div>
<div class="form-group">
<label>📦 Produto/Finalidade <span style="font-weight: 400; color: #888; font-size: 12px;">(opcional - refina a busca)</span></label>
<input type="text" id="produto" placeholder="Ex: trator, café, silo, irrigação...">
<small style="color: #666; display: block; margin-top: 5px;">
Digite uma palavra-chave para filtrar as linhas que mencionam esse produto. Deixe em branco para ver todas as linhas do tipo selecionado.
</small>
</div>
<div class="grid-2">
<div class="form-group">
<label>💰 Renda Bruta Anual (R$)</label>
<input type="number" id="renda" placeholder="Ex: 150000" min="0" oninput="window.atualizarEnquadramento()">
</div>
<div class="form-group">
<label>🌾 Valor já tomado na cultura (R$) <span style="font-weight: 400; color: #888; font-size: 12px;">(opcional)</span></label>
<input type="number" id="valorTomado" placeholder="Ex: 50000" min="0">
<button type="button" onclick="window.abrirConsultaSicor()" style="margin-top: 8px; background: none; color: #005c46; border: 1px solid #005c46; padding: 8px 14px; border-radius: 6px; font-size: 13px; font-weight: 500; width: 100%;">🔎 Consultar valor já financiado</button>
</div>
</div>
<div id="creditoTomadoBox" style="margin-bottom: 20px;"></div>
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
<option value="pecuaria">Pecuária (custeio/investimento)</option>
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
<div style="margin-bottom: 25px; padding: 18px; border: 1px solid #e0e0e0; border-radius: 6px; background: #fafbfb;">
<label style="display: block; font-weight: 600; margin-bottom: 8px; color: #005c46;">🔗 Link de Consulta de Crédito Tomado (SICOR/CACR)</label>
<input type="text" id="linkConsultaInput" placeholder="https://..." style="margin-bottom: 10px;">
<small style="color: #666; display: block; margin-bottom: 10px;">Este link é usado pelo botão "Consultar valor já financiado" na aba Consultar.</small>
<button type="button" onclick="window.salvarLinkConsulta()" style="background: #005c46; padding: 8px 18px; font-size: 13px;">💾 Salvar Link</button>
</div>
<div style="margin-bottom: 25px; padding: 18px; border: 1px solid #e0e0e0; border-radius: 6px; background: #fafbfb;">
<label style="display: block; font-weight: 600; margin-bottom: 8px; color: #005c46;">📤 Atualizar base de linhas com o arquivo da Cresol (.docx)</label>
<small style="color: #666; display: block; margin-bottom: 10px;">Envie o documento oficial mais recente. O sistema lê as modalidades e mostra uma <strong>pré-visualização</strong> para você escolher quais linhas incluir antes de substituir a base (com backup automático).</small>
<form id="formUploadCresol">
<input type="file" name="arquivo" id="arquivoCresol" accept=".docx" style="margin-bottom: 10px;">
</form>
<button type="button" onclick="window.enviarArquivoCresol()" style="background: #f58220; padding: 8px 18px; font-size: 13px;">⬆️ Processar e Atualizar Linhas</button>
<div id="uploadStatus"></div>
</div>
<div style="margin-bottom: 25px; padding: 18px; border: 1px solid #e0e0e0; border-radius: 6px; background: #fafbfb;">
<label style="display: block; font-weight: 600; margin-bottom: 8px; color: #005c46;">👥 Base de Associados (basedepessoas.csv)</label>
<small style="color: #666; display: block; margin-bottom: 10px;">Informe o link da pasta (ou arquivo) do Google Drive onde o <strong>basedepessoas.csv</strong> é disponibilizado. O sistema busca o arquivo e atualiza a aba "Base".</small>
<input type="text" id="linkBaseInput" placeholder="https://drive.google.com/drive/folders/..." style="margin-bottom: 10px;">
<div style="display: flex; gap: 8px; flex-wrap: wrap; align-items: center;">
<button type="button" onclick="window.salvarLinkBase()" style="background: #005c46; padding: 8px 16px; font-size: 13px;">💾 Salvar Link</button>
<button type="button" onclick="window.atualizarBaseAgora()" style="background: #f58220; padding: 8px 16px; font-size: 13px;">🔄 Atualizar Base Agora</button>
<button type="button" id="btnTriggerBase" onclick="window.alternarTriggerBase()" style="background: #6c757d; padding: 8px 16px; font-size: 13px;">⏱️ Ativar atualização automática</button>
</div>
<div id="baseStatus" style="margin-top: 10px;"></div>
</div>
<div style="margin-bottom: 25px; padding: 18px; border: 1px solid #e0e0e0; border-radius: 6px; background: #fafbfb;">
<label style="display: block; font-weight: 600; margin-bottom: 8px; color: #005c46;">💳 Base de Crédito Tomado / Limite Disponível (.xlsx)</label>
<small style="color: #666; display: block; margin-bottom: 10px;">Importação manual. Envie o arquivo XLSX; o sistema busca pelo <strong>CPF/CNPJ</strong> e grava na aba "BaseCredito". Colunas usadas: D (valor financiado), G (produto), S (atividade), W (CPF/CNPJ), AB (alíquota ProAgro), AC (limite custeio disponível).</small>
<form id="formUploadCredito">
<input type="file" name="arquivo" id="arquivoCredito" accept=".xlsx" style="margin-bottom: 10px;">
</form>
<button type="button" onclick="window.enviarArquivoCredito()" style="background: #f58220; padding: 8px 18px; font-size: 13px;">⬆️ Importar Base de Crédito</button>
<div id="creditoUploadStatus"></div>
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
window.LINK_CONSULTA = '${linkConsulta}';

window.abrirConsultaSicor = function() {
  if (window.LINK_CONSULTA && window.LINK_CONSULTA.trim() !== '') {
    window.open(window.LINK_CONSULTA, '_blank');
  } else {
    alert('O link de consulta ainda não foi configurado.\\nConfigure-o na aba Administrativo.');
  }
};

window.mudarAba = function(event, abaId) {
  document.querySelectorAll('.tab-content').forEach(aba => aba.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  document.getElementById(abaId).classList.add('active');
  event.target.classList.add('active');

  if (abaId === 'admin') {
    window.carregarLinhasAdministrativo();
    window.carregarLinkConsulta();
    window.carregarConfigBase();
  } else if (abaId === 'historico') {
    window.carregarHistorico();
  }
};

window.carregarConfigBase = function() {
  google.script.run
    .withSuccessHandler(function(link) {
      const input = document.getElementById('linkBaseInput');
      if (input) input.value = link || '';
    })
    .withFailureHandler(function(e) { console.error('Erro ao carregar link base:', e); })
    .obterLinkBase();

  google.script.run
    .withSuccessHandler(function(ativo) {
      const btn = document.getElementById('btnTriggerBase');
      if (!btn) return;
      if (ativo) {
        btn.textContent = '⏱️ Atualização automática ATIVA (desativar)';
        btn.style.background = '#28a745';
      } else {
        btn.textContent = '⏱️ Ativar atualização automática';
        btn.style.background = '#6c757d';
      }
      btn.dataset.ativo = ativo ? '1' : '0';
    })
    .withFailureHandler(function(e) { console.error(e); })
    .statusTriggerBaseAssociados();
};

window.salvarLinkBase = function() {
  const url = document.getElementById('linkBaseInput').value.trim();
  google.script.run
    .withSuccessHandler(function() {
      window.notificar('<strong>✓ Sucesso!</strong><br>Link da base salvo.');
    })
    .withFailureHandler(function(error) {
      window.notificar('<strong>✕ Erro ao salvar:</strong><br>' + error, 'erro');
    })
    .salvarLinkBase(url);
};

window.atualizarBaseAgora = function() {
  const st = document.getElementById('baseStatus');
  st.innerHTML = '<p style="color:#666;">⏳ Buscando e atualizando a base de associados...</p>';
  google.script.run
    .withSuccessHandler(function(resp) {
      if (resp && resp.sucesso) {
        st.innerHTML = '<div style="background:#eef6ee; border-left:4px solid #28a745; padding:10px 12px; border-radius:4px;">' +
          '<strong style="color:#1f6b1f;">✓ Base atualizada!</strong><br>Registros: <strong>' + resp.registros + '</strong><br>Em: ' + resp.atualizado + '</div>';
      } else {
        st.innerHTML = '<div style="background:#fdecea; border-left:4px solid #dc3545; padding:10px 12px; border-radius:4px;">' +
          '<strong style="color:#a4282b;">✕ Erro:</strong> ' + (resp ? resp.erro : 'desconhecido') + '</div>';
      }
    })
    .withFailureHandler(function(error) {
      st.innerHTML = '<div style="color:red;">✕ Falha: ' + error + '</div>';
    })
    .atualizarBaseAssociados();
};

window.enviarArquivoCredito = function() {
  const input = document.getElementById('arquivoCredito');
  const st = document.getElementById('creditoUploadStatus');
  if (!input.files || input.files.length === 0) {
    window.notificar('<strong>✕</strong> Selecione um arquivo .xlsx', 'erro');
    return;
  }
  st.innerHTML = '<p style="color:#666; margin-top:8px;">⏳ Importando e convertendo a planilha (pode levar alguns segundos)...</p>';
  google.script.run
    .withSuccessHandler(function(resp) {
      if (resp && resp.sucesso) {
        st.innerHTML = '<div style="margin-top:8px; background:#eef6ee; border-left:4px solid #28a745; padding:10px 12px; border-radius:4px;">' +
          '<strong style="color:#1f6b1f;">✓ Base de crédito importada!</strong><br>Registros: <strong>' + resp.registros + '</strong><br>Em: ' + resp.atualizado + '</div>';
        document.getElementById('formUploadCredito').reset();
      } else {
        st.innerHTML = '<div style="margin-top:8px; background:#fdecea; border-left:4px solid #dc3545; padding:10px 12px; border-radius:4px;">' +
          '<strong style="color:#a4282b;">✕ Erro:</strong> ' + (resp ? resp.erro : 'desconhecido') + '</div>';
      }
    })
    .withFailureHandler(function(error) {
      st.innerHTML = '<div style="margin-top:8px; color:red;">✕ Falha: ' + error + '</div>';
    })
    .processarArquivoCreditoBase(document.getElementById('formUploadCredito'));
};

window.alternarTriggerBase = function() {
  const btn = document.getElementById('btnTriggerBase');
  const ativo = btn.dataset.ativo === '1';
  const st = document.getElementById('baseStatus');
  st.innerHTML = '<p style="color:#666;">⏳ ' + (ativo ? 'Desativando' : 'Ativando') + ' atualização automática...</p>';
  const fn = ativo ? 'removerTriggerBaseAssociados' : 'criarTriggerBaseAssociados';
  google.script.run
    .withSuccessHandler(function(resp) {
      if (resp && resp.sucesso) {
        window.notificar('<strong>✓</strong> Atualização automática ' + (ativo ? 'desativada' : 'ativada (diária, ~5h)') + '.');
        window.carregarConfigBase();
        st.innerHTML = '';
      } else {
        st.innerHTML = '<div style="color:red;">✕ ' + (resp ? resp.erro : 'erro') + '</div>';
      }
    })
    .withFailureHandler(function(error) { st.innerHTML = '<div style="color:red;">✕ ' + error + '</div>'; })
    [fn]();
};

window.carregarLinkConsulta = function() {
  google.script.run
    .withSuccessHandler(function(link) {
      const input = document.getElementById('linkConsultaInput');
      if (input) input.value = link || '';
    })
    .withFailureHandler(function(error) {
      console.error('Erro ao carregar link:', error);
    })
    .obterLinkConsulta();
};

window.salvarLinkConsulta = function() {
  const url = document.getElementById('linkConsultaInput').value.trim();
  google.script.run
    .withSuccessHandler(function() {
      window.LINK_CONSULTA = url;
      window.notificar('<strong>✓ Sucesso!</strong><br>Link de consulta salvo.');
    })
    .withFailureHandler(function(error) {
      window.notificar('<strong>✕ Erro ao salvar:</strong><br>' + error, 'erro');
    })
    .salvarLinkConsulta(url);
};

window.enviarArquivoCresol = function() {
  const input = document.getElementById('arquivoCresol');
  const status = document.getElementById('uploadStatus');
  if (!input.files || input.files.length === 0) {
    window.notificar('<strong>✕</strong> Selecione um arquivo .docx primeiro.', 'erro');
    return;
  }
  const nome = input.files[0].name || '';
  if (nome.toLowerCase().indexOf('.docx') === -1) {
    window.notificar('<strong>✕</strong> O arquivo deve ser .docx.', 'erro');
    return;
  }

  status.innerHTML = '<p style="color:#666; margin-top:12px;">⏳ Processando o arquivo, aguarde (pode levar alguns segundos)...</p>';

  google.script.run
    .withSuccessHandler(function(resp) {
      if (resp && resp.sucesso) {
        window.mostrarPreviaAtualizacao(resp);
      } else {
        status.innerHTML = '<div style="margin-top:12px; background:#fdecea; border-left:4px solid #dc3545; padding:12px 14px; border-radius:4px;">' +
          '<strong style="color:#a4282b;">✕ Erro:</strong> ' + (resp ? resp.erro : 'desconhecido') + '</div>';
      }
    })
    .withFailureHandler(function(error) {
      status.innerHTML = '<div style="margin-top:12px; background:#fdecea; border-left:4px solid #dc3545; padding:12px 14px; border-radius:4px;">' +
        '<strong style="color:#a4282b;">✕ Falha no envio:</strong> ' + error + '</div>';
    })
    .processarArquivoCresol(document.getElementById('formUploadCresol'));
};

window.mostrarPreviaAtualizacao = function(resp) {
  const status = document.getElementById('uploadStatus');
  const itens = resp.itens || [];
  const rurais = itens.filter(function(i){ return i.rural; });
  const naoRurais = itens.filter(function(i){ return !i.rural; });

  function item(it, checked) {
    return '<label style="display:block; font-size:13px; margin:3px 0; cursor:pointer;">' +
      '<input type="checkbox" class="cresol-sel" value="' + it.idx + '"' + (checked ? ' checked' : '') +
      ' style="width:auto; margin-right:7px; vertical-align:middle;">' +
      window.escaparHtml(it.nome) + '</label>';
  }

  const htmlRural = rurais.map(function(it){ return item(it, true); }).join('');
  const htmlNao = naoRurais.map(function(it){ return item(it, false); }).join('');

  let bloco =
    '<div style="margin-top:12px; background:#fff8ee; border:1px solid #f3c98b; border-left:4px solid #f58220; padding:14px; border-radius:6px;">' +
    '<strong style="color:#b3590f;">⚠️ Selecione as linhas que deseja incluir</strong>' +
    '<p style="margin:8px 0; font-size:13px;">Detectadas <strong>' + resp.total + '</strong> linhas no arquivo. ' +
    'Marque as que devem entrar na base. Ao confirmar, a base atual será <strong>substituída</strong> (backup automático).</p>' +
    '<div style="margin:8px 0;"><strong style="color:#1f6b1f;">🌾 Crédito rural (' + rurais.length + ') — recomendadas</strong>' +
    '<div style="max-height:180px; overflow-y:auto; background:#fff; border:1px solid #eee; border-radius:4px; padding:8px;">' + (htmlRural || '<em style="color:#999;">nenhuma</em>') + '</div></div>';

  if (naoRurais.length) {
    bloco += '<div style="margin:8px 0;"><strong style="color:#b3590f;">🏢 Outras linhas / não-rurais (' + naoRurais.length + ') — opcionais</strong>' +
      '<div style="max-height:160px; overflow-y:auto; background:#fff; border:1px solid #eee; border-radius:4px; padding:8px;">' + htmlNao + '</div></div>';
  }

  bloco += '<div style="margin-top:6px; font-size:12px;">' +
    '<a href="#" onclick="window.cresolMarcar(true);return false;" style="color:#005c46;">Marcar todas</a> &nbsp;·&nbsp; ' +
    '<a href="#" onclick="window.cresolMarcar(false);return false;" style="color:#005c46;">Desmarcar todas</a></div>' +
    '<div style="margin-top:12px; display:flex; gap:10px;">' +
    '<button type="button" onclick="window.confirmarAtualizacaoCresol()" style="background:#28a745; padding:8px 18px; font-size:13px;">✓ Confirmar e Aplicar</button>' +
    '<button type="button" onclick="window.cancelarAtualizacaoCresol()" style="background:#6c757d; padding:8px 18px; font-size:13px;">✕ Cancelar</button>' +
    '</div></div>';

  status.innerHTML = bloco;
};

window.cresolMarcar = function(val) {
  const cbs = document.querySelectorAll('.cresol-sel');
  for (let i = 0; i < cbs.length; i++) cbs[i].checked = val;
};

window.confirmarAtualizacaoCresol = function() {
  const status = document.getElementById('uploadStatus');
  const sel = [];
  const cbs = document.querySelectorAll('.cresol-sel');
  for (let i = 0; i < cbs.length; i++) { if (cbs[i].checked) sel.push(parseInt(cbs[i].value, 10)); }
  if (sel.length === 0) {
    window.notificar('<strong>✕</strong> Selecione ao menos uma linha para incluir.', 'erro');
    return;
  }
  status.innerHTML = '<p style="color:#666; margin-top:12px;">⏳ Aplicando atualização e gerando backup...</p>';
  google.script.run
    .withSuccessHandler(function(resp) {
      if (resp && resp.sucesso) {
        status.innerHTML = '<div style="margin-top:12px; background:#eef6ee; border-left:4px solid #28a745; padding:12px 14px; border-radius:4px;">' +
          '<strong style="color:#1f6b1f;">✓ Base atualizada!</strong><br>' +
          'Linhas incluídas na base: <strong>' + resp.linhas + '</strong><br>' +
          'Backup criado: <em>' + window.escaparHtml(resp.backup) + '</em></div>';
        document.getElementById('formUploadCresol').reset();
        window.carregarLinhasAdministrativo();
      } else {
        status.innerHTML = '<div style="margin-top:12px; background:#fdecea; border-left:4px solid #dc3545; padding:12px 14px; border-radius:4px;">' +
          '<strong style="color:#a4282b;">✕ Erro ao aplicar:</strong> ' + (resp ? resp.erro : 'desconhecido') + '</div>';
      }
    })
    .withFailureHandler(function(error) {
      status.innerHTML = '<div style="margin-top:12px; color:red;">✕ Falha ao aplicar: ' + error + '</div>';
    })
    .aplicarAtualizacaoCresol(sel);
};

window.cancelarAtualizacaoCresol = function() {
  const status = document.getElementById('uploadStatus');
  google.script.run
    .withSuccessHandler(function() {
      status.innerHTML = '<p style="color:#666; margin-top:12px;">Atualização cancelada. A base não foi alterada.</p>';
      document.getElementById('formUploadCresol').reset();
    })
    .withFailureHandler(function(error) {
      status.innerHTML = '<div style="margin-top:12px; color:red;">✕ ' + error + '</div>';
    })
    .cancelarAtualizacaoCresol();
};

window.produtosPorTipo = {
  'custeio': ['Sementes', 'Fertilizantes', 'Defensivos', 'Combustível', 'Mão de obra', 'Animais (recria/engorda)', 'Ração e sais minerais', 'Vacinas e medicamentos', 'Manutenção de pastagens'],
  'pecuaria': ['Animais (recria/engorda)', 'Ração e sais minerais', 'Vacinas e medicamentos', 'Manutenção de pastagens', 'Bovino, suíno, aves', 'Piscicultura/aquicultura', 'Apicultura'],
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

window.buscarAssociado = function(origem) {
  const conta = document.getElementById('conta').value.trim();
  const cpf = document.getElementById('cpfcnpj').value.trim();
  let termo = '';
  if (origem === 'conta') termo = conta;
  else if (origem === 'cpf') termo = cpf;
  else termo = cpf || conta;
  if (!termo) return;

  const st = document.getElementById('assocStatus');
  st.innerHTML = '<span style="color:#666; font-size:12px;">Buscando associado...</span>';

  google.script.run
    .withSuccessHandler(function(resp) {
      if (resp && resp.sucesso) {
        document.getElementById('nomeAssociado').value = resp.nome || '';
        if (resp.conta) document.getElementById('conta').value = resp.conta;
        if (resp.cpfCnpj) document.getElementById('cpfcnpj').value = resp.cpfCnpj;
        if (resp.rendaAnual) {
          document.getElementById('renda').value = resp.rendaAnual;
          window.atualizarEnquadramento();
        }
        st.innerHTML = '<span style="color:#1f6b1f; font-size:12px;">✓ ' + window.escaparHtml(resp.nome || 'Associado') +
          ' — renda mensal R$ ' + window.formatarMoeda(resp.rendaMensal) +
          ' (anualizada para enquadramento: R$ ' + window.formatarMoeda(resp.rendaAnual) + ').</span>';
        // Busca também o crédito já tomado pelo CPF/CNPJ
        const cpfBusca = (resp.cpfCnpj || cpf || '');
        if (cpfBusca) window.carregarCreditoTomado(cpfBusca);
      } else {
        st.innerHTML = '<span style="color:#a4282b; font-size:12px;">✕ ' + (resp ? resp.erro : 'Erro na busca') + '</span>';
      }
    })
    .withFailureHandler(function(error) {
      st.innerHTML = '<span style="color:red; font-size:12px;">✕ ' + error + '</span>';
    })
    .buscarAssociado(termo);
};

window.carregarCreditoTomado = function(cpf) {
  const box = document.getElementById('creditoTomadoBox');
  if (!cpf) { box.innerHTML = ''; return; }
  google.script.run
    .withSuccessHandler(function(resp) {
      if (resp && resp.sucesso && resp.itens && resp.itens.length) {
        // Preenche automaticamente o "Valor já tomado" com o total já tomado
        // (financiado descontada a alíquota do ProAgro, quando houver).
        document.getElementById('valorTomado').value = Math.round(resp.totalFinanciado);
        let linhas = '';
        resp.itens.forEach(function(it) {
          const aliq = String(it.aliquotaProagro || '').trim();
          linhas += '<tr style="border-bottom:1px solid #eee;">' +
            '<td style="padding:4px 8px;">' + window.escaparHtml(it.produto) + '</td>' +
            '<td style="padding:4px 8px;">' + window.escaparHtml(it.atividade) + '</td>' +
            '<td style="padding:4px 8px; text-align:center;">' + window.escaparHtml(String(it.anoSafra)) + '</td>' +
            '<td style="padding:4px 8px; text-align:right;">R$ ' + window.formatarMoeda(it.valorFinanciado) + '</td>' +
            '<td style="padding:4px 8px; text-align:center;">' + (aliq ? window.escaparHtml(aliq) : '—') + '</td>' +
            '<td style="padding:4px 8px; text-align:right;"><strong>R$ ' + window.formatarMoeda(it.valorTomado) + '</strong></td>' +
            '<td style="padding:4px 8px; text-align:right;">R$ ' + window.formatarMoeda(it.limiteDisponivel) + '</td></tr>';
        });
        box.innerHTML =
          '<div style="background:#fff8ee; border:1px solid #f3c98b; border-left:4px solid #f58220; padding:12px; border-radius:6px;">' +
          '<strong style="color:#b3590f;">💳 Crédito já tomado (base importada)</strong>' +
          '<div style="overflow-x:auto; margin-top:8px;"><table style="width:100%; border-collapse:collapse; font-size:12px;">' +
          '<thead><tr style="background:#005c46; color:#fff;">' +
          '<th style="padding:4px 8px; text-align:left;">Produto/Finalidade</th>' +
          '<th style="padding:4px 8px; text-align:left;">Atividade</th>' +
          '<th style="padding:4px 8px;">Safra</th>' +
          '<th style="padding:4px 8px;">Valor financiado</th>' +
          '<th style="padding:4px 8px;">Alíq. ProAgro</th>' +
          '<th style="padding:4px 8px;">Valor já tomado</th>' +
          '<th style="padding:4px 8px;">Limite custeio disp.</th></tr></thead>' +
          '<tbody>' + linhas + '</tbody></table></div>' +
          '<p style="margin-top:8px; font-size:13px;"><strong>Total já tomado:</strong> R$ ' + window.formatarMoeda(resp.totalFinanciado) +
          ' &nbsp;|&nbsp; <strong>Limite custeio disponível:</strong> R$ ' + window.formatarMoeda(resp.limiteDisponivel) + '</p>' +
          '<small style="color:#666;">Valor já tomado = valor financiado menos a alíquota do ProAgro (somente custeio agrícola). O campo "Valor já tomado na cultura" foi preenchido com o total. Ajuste se necessário.</small></div>';
      } else {
        box.innerHTML = '<p style="color:#888; font-size:12px;">Sem registros de crédito tomado para este CPF/CNPJ na base.</p>';
      }
    })
    .withFailureHandler(function(error) {
      box.innerHTML = '<p style="color:red; font-size:12px;">✕ ' + error + '</p>';
    })
    .buscarCreditoTomado(cpf);
};

window.buscar = function() {
  const produto = document.getElementById('produto').value;
  const renda = parseInt(document.getElementById('renda').value) || 0;
  const enquadramento = document.getElementById('enquadramento').dataset.value || '';
  const finalidade = document.getElementById('finalidade').value;
  const valorTomado = parseFloat(document.getElementById('valorTomado').value) || 0;

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
      finalidade: finalidade,
      valorTomado: valorTomado
    });
};

window.mostrarResultados = function(linhas) {
  document.getElementById('loading').style.display = 'none';
  let html = '<h2 style="margin-bottom: 20px; color: #005c46;">Linhas Disponíveis (' + linhas.length + ')</h2>';

  const valorTomado = parseFloat(document.getElementById('valorTomado').value) || 0;
  if (valorTomado > 0) {
    html += '<div class="alert" style="background:#fff8ee; border-left:4px solid #f58220; color:#b3590f; font-size:13px;">' +
      'ℹ️ Considerando o valor já tomado de <strong>R$ ' + window.formatarMoeda(valorTomado) + '</strong>: ' +
      'linhas cujo limite já foi atingido foram <strong>ocultadas</strong>, e as demais mostram o limite ainda disponível.</div>';
  }

  if (linhas.length === 0) {
    html += '<div class="alert alert-info">Nenhuma linha encontrada. Ajuste os parâmetros' + (valorTomado > 0 ? ' (pode não haver linha com limite disponível para o valor já tomado)' : '') + '.</div>';
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
      if (linha.culturas) {
        html += '<div class="culturas-financiadas" style="margin-top: 8px; background: #fdf2e6; border-left: 4px solid #f58220; padding: 10px 12px; border-radius: 4px;">' +
          '<span style="font-weight: 600; color: #b3590f; font-size: 13px;">🌱 Culturas e atividades financiadas:</span>' +
          '<div style="color: #333; font-size: 13px; margin-top: 4px;">' + linha.culturas + '</div>' +
          '</div>';
      }
      if (linha.valorTomado > 0) {
        const esgotado = linha.limiteDisponivel <= 0;
        const bg = esgotado ? '#fdecea' : '#eef6ee';
        const cor = esgotado ? '#a4282b' : '#1f6b1f';
        const borda = esgotado ? '#dc3545' : '#28a745';
        const titulo = esgotado ? '⚠️ Limite possivelmente esgotado' : '💳 Limite disponível para a cultura';
        let corpo = 'Limite máximo da linha: R$ ' + window.formatarMoeda(linha.limiteMax) +
          ' &nbsp;|&nbsp; Já tomado: R$ ' + window.formatarMoeda(linha.valorTomado) +
          '<br><strong>Disponível: R$ ' + window.formatarMoeda(linha.limiteDisponivel) + '</strong>';
        if (esgotado) {
          corpo += '<br><em style="font-size: 12px;">O valor já tomado atinge o teto desta linha para a cultura.</em>';
        }
        html += '<div class="limite-disponivel" style="margin-top: 8px; background: ' + bg + '; border-left: 4px solid ' + borda + '; padding: 10px 12px; border-radius: 4px;">' +
          '<span style="font-weight: 600; color: ' + cor + '; font-size: 13px; display: block; margin-bottom: 4px;">' + titulo + '</span>' +
          '<div style="color: #333; font-size: 13px;">' + corpo + '</div>' +
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
    '.culturas-financiadas { margin-top: 8px; background: #fdf2e6 !important; border-left: 4px solid #f58220 !important; padding: 12px 14px; border-radius: 4px; page-break-inside: avoid; }' +
    '.culturas-financiadas span { font-weight: 600; color: #b3590f; font-size: 13px; display: block; margin-bottom: 4px; }' +
    '.culturas-financiadas div { color: #333; font-size: 13px; }' +
    '.limite-disponivel { margin-top: 8px; padding: 12px 14px; border-radius: 4px; page-break-inside: avoid; }' +
    '.limite-disponivel span { font-size: 13px; display: block; margin-bottom: 4px; }' +
    '.limite-disponivel div { color: #333; font-size: 13px; }' +
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
    status: 'Ativa', observacoes: '', itensFinanciaveis: '', culturas: ''
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
  html += area('edit_culturas', 'Culturas e Atividades Financiadas (custeio - lavouras e pecuária)', linha.culturas);
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
    culturas: document.getElementById('edit_culturas').value,
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
    'Culturas Financiadas': f.culturas,
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
