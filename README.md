# 🎵 Slides Harmonia Divina

Gerador automático de slides para músicas da Missa, a partir de dados do Notion.  
Produz arquivos **HTML** (apresentação interativa no browser) e **PPTX** (PowerPoint) prontos para projeção.

---

## 📋 Sumário

- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Como usar](#como-usar)
- [Estrutura de pastas](#estrutura-de-pastas)
- [Estrutura do banco de dados no Notion](#estrutura-do-banco-de-dados-no-notion)
- [Lógica de geração dos slides](#lógica-de-geração-dos-slides)
- [Exceções litúrgicas](#exceções-litúrgicas)
- [Saída gerada](#saída-gerada)

---

## Pré-requisitos

- Python 3.9+
- Conta no Notion com acesso à integração

---

## Instalação

```bash
# Clone o repositório
git clone https://github.com/gpereira12/slides-harmonia-divina.git
cd slides-harmonia-divina

# Instale as dependências
pip install requests python-dotenv python-pptx
```

---

## Configuração

Copie o arquivo de exemplo e preencha com suas credenciais do Notion:

```bash
cp .env.example .env
```

Edite o `.env`:

```env
NOTION_TOKEN=secret_xxxxxxxxxxxxxxxxxxxxxxxxx
NOTION_DATABASE_ID=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

> **Como obter o token:** Acesse [notion.so/my-integrations](https://www.notion.so/my-integrations), crie uma integração e copie o token.  
> **Como obter o database ID:** Abra o banco de dados no Notion, copie a URL — o ID é o trecho entre `/` e `?` no final.

---

## Como usar

```bash
python3 generate_slides.py "Nome da Missa"
```

**Exemplos:**

```bash
python3 generate_slides.py "Quaresma I"
python3 generate_slides.py "Quaresma II"
python3 generate_slides.py "Cinzas"
```

Os arquivos gerados são salvos automaticamente em `missas/Quaresma/`.

---

## Estrutura de pastas

```
slides-harmonia-divina/
├── generate_slides.py      # Script principal
├── .env                    # Credenciais (não versionado)
├── .env.example            # Template das variáveis de ambiente
├── .gitignore
├── README.md
└── missas/
    └── Quaresma/
        ├── Quaresma_I.html
        ├── Quaresma_I.pptx
        ├── Quaresma_II.html
        ├── Quaresma_II.pptx
        ├── Cinzas.html
        └── Cinzas.pptx
```

---

## Estrutura do banco de dados no Notion

O banco de dados deve ter as seguintes propriedades:

| Propriedade | Tipo | Descrição |
|---|---|---|
| `Name` | Título | Nome da música |
| `Missa` | Multi-select | Missas às quais a música pertence (ex: `Quaresma II`) |
| `Momento` | Select | Posição na missa (ex: `01. Entrada`, `5.2 Cinzas`) |
| `Letra (Traduzida)` | Rich Text | Letra da música (negrito = refrão) |

> Músicas sem letra traduzida são **ignoradas automaticamente**.

---

## Lógica de geração dos slides

### Regra geral
- **Texto em negrito** → Refrão (repetido após cada estrofe)
- **Texto normal** → Estrofes (cada parágrafo em slide separado)

### Filtragem automática
| Condição | Ação |
|---|---|
| Sem letra traduzida | Ignorado |
| Momento contém `fora da liturgia` | Ignorado |
| Momento contém `Aspersão` | Ignorado |

### Limpeza do campo Momento
Prefixos numéricos são removidos automaticamente:
- `01. Entrada` → **Entrada**
- `5.2 Cinzas` → **Cinzas**
- `99. Aspersão` → ignorado

---

## Exceções litúrgicas

Certos momentos têm renderização especial:

| Momento | Comportamento |
|---|---|
| **Ato Penitencial** (Fórmula 2) | Solo + Resposta do povo agrupados no mesmo slide |
| **Ato Penitencial** (Fórmula 3) | Renderização sequencial exata do Notion |
| **Aclamação ao Evangelho / Aleluia** | Apenas o texto em negrito é exibido |
| **Santo** | Renderização sequencial sem repetição do refrão global |
| **Cordeiro de Deus** | Renderização sequencial (mesma regra do Santo) |

---

## Saída gerada

### HTML
- Apresentação fullscreen com scroll-snap
- Navegação por teclado (← → ↑ ↓ Espaço PageUp/Down) e clique
- Fontes litúrgicas elegantes (Cormorant Garamond + Source Serif 4)
- Slide preto de transição entre cada música

### PPTX
- Formato 16:9 (1920×1080)
- Fundo creme (#FAF9F7) com texto escuro
- Texto centralizado vertical e horizontalmente
- Refrão em cinza escuro bold, estrofes em cinza médio
- Subtítulo do momento em vermelho itálico
- Slides pretos entre músicas para transição

---

## Licença

Uso interno — Coral Harmonia Divina.
