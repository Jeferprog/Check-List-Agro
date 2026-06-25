# 📋 SISTEMA DE CONSULTA DE CRÉDITO RURAL - GUIA DE IMPLEMENTAÇÃO

## 🎯 Visão Geral

Sistema integrado para consulta inteligente de linhas de crédito rural, desenvolvido em **Google Apps Script (GAS)** com interface HTML5 responsiva. Funciona como um guia consultivo para colaboradores identificarem qual linha de crédito enquadra a operação solicitada.

---

## 🚀 INSTALAÇÃO E CONFIGURAÇÃO

### Passo 1: Criar Google Sheet Base

1. Acesse [Google Sheets](https://sheets.google.com)
2. Crie uma nova planilha com o nome: **"Sistema Crédito Rural - CRESOL"**
3. Mantenha a aba padrão com o nome "Linhas"

### Passo 2: Deploy em Google Apps Script

1. Na planilha, clique em **Extensões > Apps Script**
2. Delete o código padrão do `Code.gs`
3. Cole todo o conteúdo do arquivo `creditoRural_code.gs`
4. Salve o projeto com o nome: `Sistema Crédito Rural`

### Passo 3: Inicializar o Sistema

1. No Apps Script, clique em **Executar > inicializarSistema**
2. Autorize as permissões necessárias:
   - Acessar a planilha
   - Obter informações da sessão (e-mail do usuário)
   - Criar abas adicionais

3. Aguarde a conclusão (leva 5-10 segundos)
4. Verifique na planilha se as abas foram criadas:
   - ✓ Linhas (com 17 linhas pré-cadastradas)
   - ✓ Configurações
   - ✓ Histórico

### Passo 4: Publicar a Interface Web

1. No Apps Script, clique em **Implantar > Novo Implantação**
2. Tipo: **Aplicativo Web**
3. Executar como: **Sua conta (seu e-mail)**
4. Quem tem acesso: **Qualquer pessoa dentro da organização** (ou restrito conforme política)
5. Copie o URL gerado e compartilhe com os colaboradores

---

## 📊 ESTRUTURA DO BANCO DE DADOS

### Tabela "Linhas"

Cada linha de crédito possui os seguintes campos:

| Campo | Descrição | Tipo | Exemplo |
|-------|-----------|------|---------|
| ID | Identificador único | Texto | L001 |
| Nome Linha | Nome da modalidade | Texto | PRONAF B (Microcrédito) |
| Órgão/Instituição | Instituição responsável | Texto | Banco do Brasil/Caixa |
| Finalidade Principal | Objetivo principal | Texto | Microcrédito para pequenos produtores |
| Finalidades (tags) | Tags para busca | Texto | custeio,investimento |
| Enquadramento (Renda Min/Max) | Faixa de renda | Texto | Sem limite/R$ 500 mil |
| Taxa Mín (%) | Taxa mínima anual | Número | 0.5 |
| Taxa Máx (%) | Taxa máxima anual | Número | 2.5 |
| Prazo (meses) | Prazo máximo | Número | 24 |
| Carência (meses) | Período de carência | Número | 0 |
| Limite Min (R$) | Limite mínimo | Número | 1000 |
| Limite Máx (R$) | Limite máximo | Número | 25000 |
| Requisitos | Requisitos para enquadramento | Texto | DAP ativa, agricultor familiar |
| Documentos Necessários | Lista de documentos | Texto | DAP, RG, CPF, comprovante renda |
| Status (Ativa/Inativa) | Se aparece nas buscas | Texto | Ativa ou Inativa |
| Data Atualização | Última atualização | Data | 25/06/2026 |
| Observações | Notas gerais | Texto | Modalidade mais acessível |

### Tabela "Configurações"

Armazena parâmetros do sistema:
- Versão
- Data de Atualização
- Taxa Selic base
- Ano Fiscal
- Contato de Suporte
- Últimas Alterações

### Tabela "Histórico"

Registra todas as consultas realizadas:
- Data/Hora
- Tipo de Consulta
- Produto Buscado
- Enquadramento Utilizado
- Quantidade de Resultados
- E-mail do Consultor

---

## 🔍 COMO USAR - VISÃO DO COLABORADOR

### Aba "Consultar"

**Passo 1:** Preencha o formulário com os dados do associado:

```
Campo: "Produto/Finalidade"
Exemplo: "Trator", "Sistema de Irrigação", "Custeio de Safra"
```

```
Campo: "Enquadramento"
Opções:
- PRONAF (renda até R$ 500 mil)
- PRONAMP (renda de R$ 500 mil a R$ 3,5 milhões)
- Agricultura Empresarial (acima de R$ 3,5 milhões)
```

```
Campo: "Renda Bruta Anual"
Exemplo: 250000 (para R$ 250 mil)
```

```
Campo: "Tipo de Operação"
Opções: Custeio, Investimento, Mecanização, Irrigação, Agroecologia, Infraestrutura
```

**Passo 2:** Clique em "🔍 Buscar Linhas Disponíveis"

**Passo 3:** O sistema retornará:
- ✓ Todas as linhas compatíveis
- ✓ Taxa de juros (mín-máx)
- ✓ Prazo total e carência
- ✓ Limites de crédito
- ✓ Requisitos específicos
- ✓ Documentos necessários
- ✓ Observações importantes

### Exemplo Prático

**Cenário:** Associado quer financiar um trator

```
Produto: "Trator"
Enquadramento: "PRONAF"
Renda: 180000
Tipo: "Investimento/Mecanização"
```

**Resultado esperado:**
- PRONAF Investimento (5% a.a., até 10 anos)
- Moderfrota (9-13,5% a.a., até 8 anos)
- PRONAMP Investimento (se renda permitir)

---

## ⚙️ COMO USAR - VISÃO ADMINISTRATIVA

### Aba "Administrativo"

#### 1. Atualizar Informações de uma Linha

- Selecione a linha na lista suspensa
- Os campos aparecem automaticamente
- Edite os dados desejados:
  - Taxa (mín/máx)
  - Prazo
  - Carência
  - Limite de crédito
  - Requisitos
  - Observações

#### 2. Ativar/Desativar Linha

- No campo "Status", mude entre:
  - **Ativa** → Aparece nas buscas
  - **Inativa** → NÃO aparece nas buscas

**Uso:** Quando uma linha é descontinuada ou temporariamente indisponível

#### 3. Adicionar Nova Linha

Para adicionar uma nova linha de crédito:

1. Acesse a planilha diretamente
2. Aba "Linhas"
3. Clique na primeira linha vazia após os dados existentes
4. Preencha todos os campos conforme o padrão
5. Certifique-se de que "Status" está "Ativa" se deseja que apareça
6. Salve

### Aba "Histórico"

Visualize todas as consultas realizadas:
- Data e hora da consulta
- Produto buscado
- Enquadramento utilizado
- Quantas linhas foram encontradas
- Qual colaborador fez a busca

**Uso:** Análise de padrões, auditoria, e acompanhamento de demandas

---

## 🔧 MANUTENÇÃO E ATUALIZAÇÕES

### Atualizar Taxas de Juros

Quando houver mudança nas taxas (por mudança da Selic ou políticas do banco):

1. Acesse **Administrativo**
2. Selecione a linha
3. Atualize "Taxa Mín (%)" e "Taxa Máx (%)"
4. Salve
5. A data de atualização será registrada automaticamente

### Atualizar Limites de Crédito

Mesmo processo acima, editando "Limite Máx (R$)"

### Desativar Linha Temporariamente

1. Selecione a linha
2. Mude "Status" para "Inativa"
3. Salve
4. Não aparecerá mais nas buscas

### Criar Backup

Regularmente (ex: mensalmente):
1. Abra a planilha
2. Menu **Arquivo > Fazer download > Microsoft Excel**
3. Guarde o arquivo

---

## 📱 COMPATIBILIDADE

- ✓ Desktop (Chrome, Firefox, Safari, Edge)
- ✓ Tablet (iPad, Android)
- ✓ Mobile (Android, iOS)
- ✓ Qualquer navegador moderno com JavaScript

---

## 🎨 CUSTOMIZAÇÃO VISUAL (Padrão CRESOL)

Para adaptar ao layout CRESOL:

### Editar Cores (no `obterHTML()`)

Localize a seção `<style>` e altere:

```css
/* Cor primária CRESOL */
background: linear-gradient(135deg, #1f4788 0%, #2d5a9a 100%);

/* Mudar para cores da Cresol */
background: linear-gradient(135deg, #SEU_COR1 0%, #SEU_COR2 100%);
```

### Adicionar Logo CRESOL

No `<div class="header">`, após `<h1>`, adicione:

```html
<img src="URL_LOGO_CRESOL" alt="Cresol" style="height: 40px; margin: 10px 0;">
```

### Adicionar Rodapé

Antes de `</div>` (fechamento de container), adicione:

```html
<div style="background: #f5f5f5; padding: 20px; text-align: center; color: #666; font-size: 12px;">
  <p>© 2026 CRESOL - Sistema de Crédito Rural | Versão 1.0</p>
  <p>Para suporte, contate: agro@cresol.com.br</p>
</div>
```

---

## 🔐 PERMISSÕES E COMPARTILHAMENTO

### Compartilhar com Colaboradores

1. No Apps Script, clique em **Implantar > Novo Implantação**
2. Configure o acesso:
   - **Qualquer pessoa dentro da organização**: Todos podem usar
   - **Usuários específicos**: Especifique os e-mails

### Permissões Necessárias

O sistema solicita permissão para:
- ✓ Ler e escrever na planilha
- ✓ Obter informações da sessão (apenas e-mail)
- ✓ Nenhuma integração com sistemas externos

---

## 📞 SUPORTE E TROUBLESHOOTING

### Erro: "Permissão Negada"

**Causa:** Usuário não tem acesso à planilha ou Apps Script

**Solução:**
1. Verifique compartilhamento da planilha
2. Verifique acesso no Apps Script (Implantar > Novo Implantação)

### Erro: "Abas não encontradas"

**Causa:** `inicializarSistema()` não foi executado

**Solução:**
1. No Apps Script, clique em **Executar > inicializarSistema**
2. Autorize as permissões

### Busca retorna "0 linhas"

**Causa:** Parâmetros não correspondem a nenhuma linha

**Solução:**
1. Tente enquadramentos diferentes
2. Verifique se as linhas estão marcadas como "Ativa"
3. Verifique no "Administrativo" qual a renda mínima/máxima esperada

### Interface não carrega

**Causa:** Problema de rede ou cache do navegador

**Solução:**
1. Limpe cache do navegador
2. Tente em navegador privado
3. Aguarde 2 minutos e tente novamente

---

## 🎓 FLUXO RECOMENDADO NA CRESOL

```
ASSOCIADO SOLICITA CRÉDITO
        ↓
COLABORADOR ACESSA SISTEMA
        ↓
PREENCHE DADOS DO ASSOCIADO
        ↓
SISTEMA RETORNA LINHAS COMPATÍVEIS
        ↓
COLABORADOR ANALISA OPÇÕES
(Taxa, Prazo, Documentos)
        ↓
ESCOLHE MELHOR LINHA
        ↓
SOLICITA DOCUMENTAÇÃO NECESSÁRIA
        ↓
ENCAMINHA PARA ANÁLISE DE CRÉDITO
        ↓
GERENTE APROVA/NEGA
```

---

## 📈 FUNCIONALIDADES FUTURAS

Melhorias planejadas para versões futuras:

- [ ] Integração com dados de renda do associado
- [ ] Simulador de parcelas (com juros compostos)
- [ ] Relatórios de demandas por linha
- [ ] Dashboard com estatísticas
- [ ] Envio automático de propostas por e-mail
- [ ] Integração com sistema de CRM da Cresol
- [ ] Cálculo automático de adequação de renda
- [ ] Exportar resultado para PDF

---

## 📋 CHECKLIST DE IMPLEMENTAÇÃO

- [ ] Google Sheet criado
- [ ] Apps Script instalado
- [ ] `inicializarSistema()` executado
- [ ] Abas criadas com sucesso
- [ ] Interface web publicada
- [ ] URL copiada e testada
- [ ] Colaboradores têm acesso
- [ ] Primeiras buscas testadas
- [ ] Histórico funcionando
- [ ] Cores/logos customizadas (opcional)

---

## 📞 CONTATO E FEEDBACK

Para sugestões, melhorias ou relatar problemas:
- E-mail: agro@cresol.com.br
- Sistema desenvolvido em: 2026
- Versão: 1.0

---

**Última atualização:** 25/06/2026
