# Anhangá Radar

App desktop para registro inteligente de contatos com vereadores durante a prospecção comercial.

Desenvolvido por **[@Jitterkkk](https://github.com/Jitterkkk)** • Instagram: **[@jitterkkk](https://instagram.com/jitterkkk)**

O usuário descreve o contato em linguagem natural — a IA extrai os dados automaticamente e os salva numa planilha Excel organizada por abas.

---

## Como funciona

1. Descreva o contato em texto livre (ex: _"liguei pro João Silva de Campinas, não atendeu"_)
2. Clique em **Processar com IA** ou pressione `Ctrl+Enter`
3. Revise os campos extraídos (nome, cidade, status etc.)
4. Clique em **Salvar na planilha**

Os registros são salvos em `data/contatos.xlsx` com duas abas:

| Aba | Quando usar |
|-----|-------------|
| **Falhas e Sem Contato** | Não atendeu, caixa postal, número errado, ocupado |
| **Sem Interesse** | Vereador atendeu mas recusou ou demonstrou desinteresse |

---

## Pré-requisitos

- Python **3.8 ou superior** — [python.org/downloads](https://www.python.org/downloads/)
- Chave de API do Groq — [console.groq.com](https://console.groq.com) (gratuito, sem cartão)
- Git — [git-scm.com](https://git-scm.com)

---

## Instalação

```bash
# 1. Clone o repositório
git clone https://github.com/Jitterkkk/anhanga-radar.git
cd anhanga-radar

# 2. Instale as dependências
pip install -r requirements.txt
```

---

## Obtendo a API Key do Groq (gratuito)

1. Acesse [console.groq.com](https://console.groq.com) e crie uma conta (gratuita, sem cartão)
2. No menu lateral, clique em **API Keys**
3. Clique em **Create API Key**, dê um nome (ex: `anhanga-radar`) e copie o valor gerado
4. Cole a chave no campo **Chave da API** dentro do app — ela será salva automaticamente para as próximas sessões

Limites gratuitos: **30 req/min** e **14.400 req/dia** (Llama 3.3 70B).

> A chave fica armazenada localmente em `.config` na raiz do projeto e nunca é enviada ao repositório.

---

## Como rodar

```bash
python main.py
```

A janela do **Anhangá Radar** será aberta. Na primeira execução, informe sua chave da API no campo do topo.

---

## Exemplos de frases

Cole qualquer descrição em linguagem natural no campo de texto. A IA entende variações:

```
Liguei pro João Silva de Campinas, não atendeu. Número (19) 99999-1234.
```
```
Tentei contato via WhatsApp com a vereadora Maria Souza de Ribeirão Preto — sem interesse, disse que já tem assessoria.
```
```
Carlos Pereira, Sorocaba, liguei duas vezes, caixa postal nas duas. Fonte: lista da Câmara.
```
```
Mandei e-mail pro gabinete do Dr. Antônio Lima (SP capital). Responderam que não têm interesse no momento.
```
```
Visita presencial ao vereador Roberto Dias em Bauru. Número errado no cadastro, não foi possível contato.
```

---

## Estrutura do projeto

```
anhanga-radar/
├── main.py              # Ponto de entrada — verifica Python e inicia o app
├── requirements.txt     # Dependências: groq e openpyxl
├── .gitignore
├── src/
│   ├── app.py           # Interface gráfica (tkinter)
│   ├── ia_processor.py  # Integração com a API do Groq (Llama 3.3)
│   ├── excel_manager.py # Criação e escrita da planilha Excel
│   └── config.py        # Persistência da API key em .config
├── data/
│   └── contatos.xlsx    # Gerado automaticamente na primeira execução
└── tests/
    └── test_excel.py    # Testes do módulo excel_manager
```

---

## Rodando os testes

```bash
python -m unittest tests.test_excel -v
```

---

## Licença

Desenvolvido por [@Jitterkkk](https://github.com/Jitterkkk) © 2026
