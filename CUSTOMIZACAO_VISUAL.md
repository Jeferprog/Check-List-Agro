# 🎨 GUIA DE CUSTOMIZAÇÃO VISUAL - PADRÃO CRESOL

## Adaptação Visual do Sistema para o Layout da Cresol

Este guia mostra como customizar o sistema para seguir o padrão visual da Cresol.

---

## 1. CORES CORPORATIVAS

### Paleta Padrão (Azul)
Se a Cresol usa azul como cor primária:

```css
/* Cores atuais no sistema */
--cor-primaria: #1f4788;      /* Azul escuro */
--cor-primaria-claro: #2d5a9a; /* Azul médio */
--cor-destaque: #e8f4f8;        /* Azul muito claro */
```

### Customizar para Cores da Cresol

Se você tiver uma paleta específica (ex: verde, vermelho), substitua no `code.gs`:

**Localize esta seção no `obterHTML()`:**

```html
<style>
  body {
    background: linear-gradient(135deg, #1f4788 0%, #2d5a9a 100%);
  }
  
  .header {
    background: linear-gradient(135deg, #1f4788 0%, #2d5a9a 100%);
  }
  
  .tab-btn.active {
    color: #1f4788;
    border-bottom-color: #1f4788;
  }
</style>
```

**Substitua pelos valores da Cresol:**

```html
<style>
  body {
    background: linear-gradient(135deg, #SEU_COR_PRIMARIA 0%, #SEU_COR_SECUNDARIA 100%);
  }
  
  .header {
    background: linear-gradient(135deg, #SEU_COR_PRIMARIA 0%, #SEU_COR_SECUNDARIA 100%);
  }
  
  .tab-btn.active {
    color: #SEU_COR_PRIMARIA;
    border-bottom-color: #SEU_COR_PRIMARIA;
  }
</style>
```

### Exemplo com Verde (Tema Ambiental)

Se Cresol usar verde:

```css
--cor-primaria: #2d6a3e;       /* Verde escuro */
--cor-primaria-claro: #3d8a52;  /* Verde médio */
--cor-destaque: #d4f0d9;        /* Verde claro */
```

---

## 2. LOGO E CABEÇALHO

### Adicionar Logo da Cresol

**No Apps Script, procure por:**

```html
<div class="header">
  <h1>🌾 Sistema de Crédito Rural</h1>
  <p>Consultar linhas de crédito para operações rurais - CRESOL</p>
</div>
```

**Substitua por:**

```html
<div class="header">
  <img src="https://www.cresol.com.br/logo.png" alt="Cresol" style="height: 50px; margin-bottom: 15px;">
  <h1>Sistema de Crédito Rural</h1>
  <p>Consultar linhas de crédito para operações rurais</p>
</div>
```

⚠️ **Nota:** Substitua `https://www.cresol.com.br/logo.png` pelo URL real da logo

### Opção: Fazer Upload no Google Drive

1. Faça upload da logo Cresol para Google Drive
2. Compartilhe com "acesso público"
3. Copie o ID do arquivo (URL: `/d/FILE_ID/view`)
4. Use: `https://drive.google.com/uc?export=view&id=FILE_ID`

Exemplo:
```html
<img src="https://drive.google.com/uc?export=view&id=1a2b3c4d5e6f7g8h9i0j" alt="Cresol">
```

---

## 3. TIPOGRAFIA

### Fonte Corporativa

Se Cresol tem fonte específica, customize no `<style>`:

**Padrão:**
```css
font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
```

**Com font customizada:**
```html
<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;600&display=swap" rel="stylesheet">

<style>
  body {
    font-family: 'Roboto', sans-serif;
  }
</style>
```

**Outras opções:**
- `'Inter'` - Moderna e legível
- `'Poppins'` - Amigável
- `'Source Sans Pro'` - Profissional
- `'Montserrat'` - Elegante

---

## 4. RODAPÉ

### Adicionar Rodapé Corporativo

**Localize antes do `</div>` final (container):**

```html
<div style="background: #f5f5f5; padding: 20px; text-align: center; color: #666; font-size: 12px;">
  <p>© 2026 CRESOL - Sistema de Crédito Rural | Versão 1.0</p>
  <p>Para suporte, contate: agro@cresol.com.br</p>
</div>
```

**Customize para:**

```html
<div style="background: linear-gradient(to right, #f5f5f5, #efefef); padding: 30px; text-align: center; color: #666; font-size: 12px; border-top: 2px solid #ddd;">
  <div style="margin-bottom: 15px;">
    <img src="LOGO_PEQUENA" alt="Cresol" style="height: 30px; margin-right: 10px;">
  </div>
  <p style="margin: 5px 0;"><strong>CRESOL - Sistema de Crédito Rural</strong></p>
  <p style="margin: 5px 0; font-size: 11px;">Versão 1.0 | 2025/2026</p>
  <p style="margin: 10px 0 0 0;">
    📧 <strong>agro@cresol.com.br</strong> | 
    📱 <strong>(XX) XXXXX-XXXX</strong>
  </p>
  <p style="margin-top: 10px; font-size: 10px; color: #999;">
    © 2026 CRESOL - Todos os direitos reservados
  </p>
</div>
```

---

## 5. ÍCONES E EMOJIS

### Substituir Emojis por Ícones (Opcional)

Se preferir ícones mais profissionais:

**Adicione Font Awesome no `<head>`:**

```html
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
```

**Substitua emojis:**

| Emoji | Ícone Font Awesome | Código |
|-------|-------------------|--------|
| 🔍 | Lupa | `<i class="fas fa-search"></i>` |
| 📋 | Checklist | `<i class="fas fa-clipboard-list"></i>` |
| ⚙️ | Engrenagem | `<i class="fas fa-gear"></i>` |
| 📊 | Gráfico | `<i class="fas fa-chart-bar"></i>` |
| 🌾 | Milho | `<i class="fas fa-wheat"></i>` |

**Exemplo:**
```html
<!-- Antes -->
<button onclick="buscar()">🔍 Buscar Linhas Disponíveis</button>

<!-- Depois -->
<button onclick="buscar()"><i class="fas fa-search"></i> Buscar Linhas Disponíveis</button>
```

---

## 6. CARDS E COMPONENTES

### Customizar Estilo dos Cards

**Cards das Linhas (padrão):**

```css
.linha-card {
  background: white;
  border: 1px solid #ddd;
  border-radius: 5px;
  padding: 15px;
  margin-bottom: 15px;
}
```

**Versão customizada (mais elegante):**

```css
.linha-card {
  background: white;
  border-left: 4px solid #2d5a9a;  /* Barra lateral colorida */
  border-radius: 8px;
  padding: 20px;
  margin-bottom: 15px;
  box-shadow: 0 2px 8px rgba(45, 90, 154, 0.1);
  transition: all 0.3s ease;
}

.linha-card:hover {
  box-shadow: 0 6px 16px rgba(45, 90, 154, 0.15);
  transform: translateY(-2px);
}
```

---

## 7. ALERTAS E NOTIFICAÇÕES

### Customizar Cores dos Alertas

**Padrão:**
```css
.alert-info {
  background: #e8f4f8;
  color: #0c5460;
  border-left: 4px solid #0c5460;
}
```

**Com tema Cresol:**
```css
.alert-info {
  background: linear-gradient(135deg, #e8f4f8 0%, #d0e8f2 100%);
  color: #0c5460;
  border-left: 4px solid #2d5a9a;
  border-radius: 4px;
}

.alert-success {
  background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
  color: #155724;
  border-left: 4px solid #28a745;
}
```

---

## 8. TEMA ESCURO (Opcional)

### Implementar Modo Noturno

Se Cresol quiser opção de tema escuro:

```html
<button id="toggleTheme" style="position: absolute; top: 20px; right: 20px;">
  🌙 Tema Escuro
</button>

<script>
  const toggleTheme = () => {
    document.body.classList.toggle('dark-theme');
    localStorage.setItem('theme', document.body.classList.contains('dark-theme') ? 'dark' : 'light');
  };

  document.getElementById('toggleTheme').addEventListener('click', toggleTheme);

  // Aplicar tema salvo
  if (localStorage.getItem('theme') === 'dark') {
    document.body.classList.add('dark-theme');
  }
</script>

<style>
  body.dark-theme {
    background: #1a1a1a;
    color: #e0e0e0;
  }

  body.dark-theme .container {
    background: #2d2d2d;
    color: #e0e0e0;
  }

  body.dark-theme .header {
    background: linear-gradient(135deg, #0a2847 0%, #133657 100%);
  }

  body.dark-theme .linha-card {
    background: #3a3a3a;
    border-color: #555;
    color: #e0e0e0;
  }
</style>
```

---

## 9. LAYOUT RESPONSIVO

### Otimizar para Dispositivos Cresol

Se a Cresol usa tablets/mesas de caixa:

```css
@media (max-width: 1024px) {
  .grid-2 {
    grid-template-columns: 1fr;
  }

  .container {
    max-width: 95%;
  }

  .header h1 {
    font-size: 22px;
  }
}

@media (max-width: 768px) {
  .tab-btn {
    font-size: 13px;
    padding: 12px;
  }

  .linha-card {
    padding: 12px;
  }
}

/* Para uso em quiosque (1920x1080) */
@media (min-width: 1920px) {
  .container {
    max-width: 1200px;
    font-size: 16px;
  }

  .linha-card {
    padding: 25px;
  }
}
```

---

## 10. INTEGRAÇÃO COM CRESOL

### Adicionar Links e Contatos

Customize a seção de contato:

```html
<div class="alert alert-info">
  <strong>📞 Suporte Cresol:</strong><br>
  Central de Atendimento: (54) 3025-2000<br>
  WhatsApp: (54) 99999-9999<br>
  E-mail: agro@cresol.com.br<br>
  <a href="https://www.cresol.com.br" target="_blank" style="color: #1f4788; text-decoration: none;">
    Visite nosso site
  </a>
</div>
```

---

## 11. INSTRUÇÕES DE IMPLEMENTAÇÃO

### Passo a Passo

1. **Copie o código atual do `code.gs`**
2. **Abra o Apps Script**
3. **Localize a função `obterHTML()`**
4. **Encontre a seção `<style>`**
5. **Faça as alterações desejadas**
6. **Teste na interface web**
7. **Ajuste conforme necessário**

### Teste Rápido

Após cada customização:
1. Salve no Apps Script
2. Clique em **Implantar > Novo Implantação**
3. Abra a URL no navegador
4. Limpe cache (Ctrl+Shift+Delete)
5. Veja as mudanças

---

## 12. EXEMPLO COMPLETO: TEMA CRESOL VERDE

Se Cresol usa verde como cor primária:

**Substitua a seção `<style>` completa por:**

```css
<style>
  * {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
  }

  body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: linear-gradient(135deg, #1e5631 0%, #3a7d44 100%);
    min-height: 100vh;
    padding: 20px;
  }

  .container {
    max-width: 1000px;
    margin: 0 auto;
    background: white;
    border-radius: 12px;
    box-shadow: 0 12px 48px rgba(0,0,0,0.15);
    overflow: hidden;
  }

  .header {
    background: linear-gradient(135deg, #1e5631 0%, #3a7d44 100%);
    color: white;
    padding: 40px 30px;
    text-align: center;
  }

  .header h1 {
    font-size: 28px;
    margin-bottom: 8px;
  }

  .header p {
    font-size: 14px;
    opacity: 0.95;
  }

  .tabs {
    display: flex;
    background: #f0f0f0;
    border-bottom: 2px solid #ccc;
  }

  .tab-btn {
    flex: 1;
    padding: 15px;
    background: none;
    border: none;
    cursor: pointer;
    font-size: 14px;
    font-weight: 500;
    color: #666;
    transition: all 0.3s;
  }

  .tab-btn.active {
    color: #1e5631;
    border-bottom: 3px solid #1e5631;
    background: white;
  }

  .tab-content {
    display: none;
    padding: 30px;
  }

  .tab-content.active {
    display: block;
  }

  button {
    background: linear-gradient(135deg, #1e5631 0%, #3a7d44 100%);
    color: white;
    padding: 12px 30px;
    border: none;
    border-radius: 6px;
    cursor: pointer;
    font-weight: 600;
    transition: all 0.3s;
  }

  button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(30, 86, 49, 0.3);
  }

  .linha-card {
    background: white;
    border-left: 4px solid #3a7d44;
    border-radius: 8px;
    padding: 20px;
    margin-bottom: 15px;
    box-shadow: 0 2px 8px rgba(30, 86, 49, 0.1);
  }

  .linha-card h3 {
    color: #1e5631;
    margin-bottom: 10px;
  }

  .alert-info {
    background: linear-gradient(135deg, #d0e8d9 0%, #c3e6cb 100%);
    color: #1e5631;
    border-left: 4px solid #1e5631;
    padding: 15px;
    border-radius: 6px;
    margin-bottom: 20px;
  }
</style>
```

---

## 📋 CHECKLIST DE CUSTOMIZAÇÃO

- [ ] Cores corporativas alteradas
- [ ] Logo Cresol adicionada
- [ ] Tipografia customizada
- [ ] Rodapé adicionado
- [ ] Contatos/telefone inclusos
- [ ] Links para site Cresol
- [ ] Testado em desktop
- [ ] Testado em tablet
- [ ] Testado em mobile
- [ ] Cache limpo
- [ ] Compartilhado com equipe

---

## 💡 DICAS FINAIS

1. **Sempre faça teste antes de compartilhar**
2. **Use Firefox DevTools (F12) para ajustar em tempo real**
3. **Mantenha contraste bom (cores acessíveis)**
4. **Teste com diferentes tamanhos de tela**
5. **Não exagere em animações (podem deixar lento)**
6. **Documente todas as mudanças feitas**

---

**Precisa de ajuda com cores específicas?** Compartilhe o layout/guideline visual da Cresol para customizar com exatidão!
