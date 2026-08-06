# Diagnóstico Financeiro Consolidado (app fora do Claude)

App independente (não usa Claude/Cowork), genérico — serve para **qualquer
empresa ou grupo**, não só para um cliente específico. O time sobe os
balancetes de quem for (uma empresa isolada, ou várias empresas de um mesmo
grupo) e o app gera o Excel (Balanço + DRE + DFC) e o relatório Word
automaticamente. Roda de duas formas: **no computador de qualquer pessoa do
time** (modo desktop) ou **hospedado, com um link** para todo mundo acessar
pelo navegador.

## O que o app faz

1. Na barra lateral, você digita o nome da empresa/grupo (aparece nos
   relatórios) e sobe os arquivos `.xls`/`.xlsx` de balancete de verificação
   de cada empresa (de qualquer período).
2. O app identifica cada empresa pelo CNPJ e cada período pela data do
   próprio balancete — não precisa organizar em pastas nem renomear nada.
   Para uma empresa isolada, suba só os arquivos dela; para um grupo, suba os
   balancetes de todas as empresas do grupo juntos.
3. Consolida Balanço, DRE e DFC por soma simples (sem eliminar saldos entre
   empresas, se for mais de uma).
4. Mostra os indicadores financeiros na tela (liquidez, endividamento,
   alavancagem, margens, fluxo de caixa) e sinaliza automaticamente pontos de
   atenção (ex.: alavancagem acima de 3,5x, cobertura de juros abaixo de
   2,0x — ajustável no arquivo `docx_report.py`, dicionário `THRESHOLDS`).
5. Gera para download: o Excel consolidado, um relatório Word de
   diagnóstico e uma apresentação PowerPoint (mesma paleta navy/azul-gelo
   usada no deck feito para o Grupo CLA/Claudio, sem marca institucional) —
   todos já com o nome da empresa/grupo no título.

**Sobre usar com uma empresa nova**: a lógica de leitura do balancete (quais
códigos de classificação, ex. `1.1.01`, correspondem a Caixa, Fornecedores
etc.) foi calibrada e validada com o Grupo CLA/Claudio. Ela tende a
funcionar bem para outras empresas que usem uma estrutura de plano de contas
parecida (comum quando o mesmo escritório de contabilidade padroniza a
codificação entre os clientes), mas isso não é garantido. Sempre que usar
com uma empresa pela primeira vez, olhe a tabela de **Checagem de
consistência** dentro do app — se algum valor estiver bem diferente de
0,00, o app avisa em vermelho e é sinal de que o plano de contas dessa
empresa é diferente do esperado.

**Importante**: o relatório é gerado automaticamente a partir das regras de
referência configuradas — ele não substitui a leitura de alguém do time,
principalmente para explicar casos específicos (uma reclassificação
contábil, uma baixa de fornecedor, uma correção pontual). Para esse tipo de
análise mais fina, quem tiver acesso ao Claude pode usar a skill
"diagnostico-bancario-grupo-cla-claudio" em conjunto com este app — os dois
usam exatamente a mesma lógica de consolidação, então os números sempre vão
bater.

## Opção 1 — Rodar no computador (modo desktop, sem instalar nada em nuvem)

Pré-requisitos: Python 3.9+ instalado. Em alguns casos (arquivos `.xls`
antigos), também precisa do LibreOffice instalado (gratuito,
libreoffice.org) — o Excel/Word do Office não é necessário.

```bash
cd app
pip install -r requirements.txt
streamlit run app.py
```

Isso abre automaticamente uma aba no navegador (`http://localhost:8501`) — o
app roda 100% localmente, os balancetes não saem do computador. Para
qualquer pessoa do time usar, é só repetir esses 3 comandos na pasta `app`
depois de copiá-la para o computador dela (ou compartilhar a pasta inteira).

## Opção 2 — Hospedar com um link compartilhável (grátis)

Não tenho como criar/manter um servidor com link público permanente direto
deste ambiente de trabalho — isso precisa ser hospedado em algum serviço
(mesmo que gratuito) vinculado a uma conta de alguém do seu time. O caminho
mais simples e sem custo é o **Streamlit Community Cloud**:

1. Crie uma conta gratuita em [share.streamlit.io](https://share.streamlit.io)
   (pode entrar com uma conta do GitHub).
2. Se ainda não tiver, crie um repositório no GitHub (pode ser privado) e
   suba os 6 arquivos desta pasta (`app.py`, `pipeline.py`, `docx_report.py`,
   `pptx_report.py`, `requirements.txt`, `packages.txt`).
3. No Streamlit Community Cloud, clique em "New app", escolha o repositório
   e o arquivo `app.py`, e clique em "Deploy".
4. Em alguns minutos, o Streamlit te dá um link (algo como
   `https://seu-app.streamlit.app`) que você compartilha com o time — todo
   mundo acessa pelo navegador, sem instalar nada.

O arquivo `packages.txt` (incluído nesta pasta) já avisa o Streamlit Cloud
para instalar o LibreOffice automaticamente, necessário para ler arquivos
`.xls` antigos — não precisa fazer nada além de subir o arquivo junto.

Alternativas igualmente gratuitas, se preferirem não usar GitHub/Streamlit
Cloud: **Hugging Face Spaces** (aceita apps Streamlit direto, sem GitHub
obrigatório) ou **Render.com** (free tier, um pouco mais manual). Todas
seguem o mesmo princípio: subir estes 6 arquivos e apontar para `app.py`.

## Arquivos desta pasta

- `app.py` — interface Streamlit (upload, tabela de indicadores, botões de download).
- `pipeline.py` — motor de consolidação (leitura de balancete, cálculo de
  Balanço/DRE/DFC, indicadores). Mesma lógica usada na skill do Claude
  `diagnostico-bancario-grupo-cla-claudio`, então os dois sempre produzem os
  mesmos números para os mesmos balancetes.
- `docx_report.py` — gera o relatório Word (python-docx puro, sem depender
  de Node/Office).
- `pptx_report.py` — gera a apresentação PowerPoint (python-pptx puro, mesma
  paleta e estrutura de slides do deck original do Grupo CLA/Claudio: capa,
  sumário executivo, gráfico de trajetória, indicadores detalhados, pontos
  de atenção e nota metodológica).
- `requirements.txt` / `packages.txt` — dependências para instalar
  localmente ou no serviço de hospedagem escolhido.

## Segurança

Os balancetes contêm dados financeiros sensíveis. No modo desktop, nada saí
do computador de quem está usando. Se optarem por hospedar com link público
(Opção 2), qualquer pessoa com o link consegue abrir o app (embora precise
enviar seus próprios balancetes para ver algum dado) — se isso for uma
preocupação, o Streamlit Community Cloud permite restringir o acesso ao
app por e-mail/domínio nas configurações de compartilhamento.
