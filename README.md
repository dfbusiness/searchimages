# SearchImages

Ferramenta em Python para gerar uma planilha Excel com **imagens reais de produtos** a partir dos links de fornecedores informados em uma planilha padrão.

O objetivo principal é evitar o erro comum de inserir imagens genéricas, ícones, placeholders ou imagens ilustrativas. A ferramenta tenta capturar a imagem real do produto vinculado ao link utilizado na precificação.

## O que a ferramenta faz

1. Lê uma planilha de entrada, por exemplo `MEMORIA DE CALCULO.xlsx`.
2. Usa os links dos fornecedores, por exemplo a coluna `U`.
3. Para links do Mercado Livre, tenta obter a imagem pela API pública do Mercado Livre.
4. Para outros sites, usa Playwright para abrir a página em navegador real e localizar a imagem principal do produto.
5. Gera uma nova planilha Excel com:
   - Item;
   - Descrição do Item;
   - Imagem;
   - UN;
   - Qtde;
   - Valor Unitário;
   - Valor Total.
6. Cria uma aba `Fontes` com link, marca identificada, arquivo de imagem, status e observações.
7. Cria um `log_execucao.csv` para conferência.

## Estrutura recomendada do repositório

```text
searchimages/
├── dados/
│   ├── MEMORIA DE CALCULO.xlsx
│   └── TR 21 HXI - Material apoio.pdf
├── config.example.yaml
├── requirements.txt
├── searchimages.py
└── README.md
```

## Como colocar os arquivos de entrada

Crie uma pasta chamada `dados` no repositório e coloque dentro dela:

```text
MEMORIA DE CALCULO.xlsx
TR 21 HXI - Material apoio.pdf
```

Para criar a pasta pelo GitHub:

1. Clique em `Add file > Create new file`.
2. Digite: `dados/.gitkeep`.
3. Clique em `Commit changes`.
4. Entre na pasta `dados`.
5. Clique em `Add file > Upload files`.
6. Envie o PDF e a planilha Excel.
7. Clique em `Commit changes`.

## Instalação local

No terminal, dentro da pasta do repositório:

```bash
python -m venv .venv
```

Ativar ambiente virtual no Windows:

```bash
.venv\Scripts\activate
```

Ativar ambiente virtual no macOS/Linux:

```bash
source .venv/bin/activate
```

Instalar dependências:

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

## Executar com a configuração padrão

```bash
python searchimages.py run --config config.example.yaml
```

A saída será gerada em:

```text
outputs/Planilha_TR_HXI_21_Material_Apoio_CORRIGIDA.xlsx
outputs/log_execucao.csv
outputs/imagens/
```

## Configuração padrão

O arquivo `config.example.yaml` está ajustado para a estrutura da planilha `MEMORIA DE CALCULO.xlsx`:

```yaml
input:
  excel_path: "dados/MEMORIA DE CALCULO.xlsx"
  pdf_path: "dados/TR 21 HXI - Material apoio.pdf"
  sheet_name: "Planilha1"
  first_data_row: 5
  last_data_row: 55

columns:
  item: "A"
  description_memory: "B"
  quantity: "F"
  unit_price: "S"
  total_price: "T"
  url: "U"
```

Se uma futura planilha tiver colunas diferentes, basta alterar as letras das colunas no YAML.

## Criar uma planilha modelo para futuras pesquisas

```bash
python searchimages.py template --output modelo_pesquisa_produtos.xlsx
```

Esse comando cria uma planilha simples com as colunas essenciais para futuras pesquisas.

## Critério de segurança das imagens

A ferramenta foi desenhada para **não inserir imagem genérica**. Quando não consegue capturar a imagem real com segurança, ela:

1. deixa a célula de imagem em branco;
2. marca o item como `REVISAR` no log;
3. registra a observação na aba `Fontes`.

Isso é intencional. É melhor deixar o item para revisão do que inserir uma imagem errada.

## Observações importantes

- Alguns sites podem exibir cookies, banners, carrosséis ou bloqueios. Por isso, o modo padrão do navegador é `headless: false`, permitindo intervenção visual se necessário.
- Para Mercado Livre, a ferramenta tenta primeiro a API, pois é mais confiável do que capturar a página visual.
- Para fornecedores externos, o Playwright abre a página e tenta identificar a maior imagem principal do produto.
- Após a execução, recomenda-se conferir a aba `Fontes` e o arquivo `log_execucao.csv`.

## Como usar com o Codex

Depois de colocar os arquivos na pasta `dados`, peça ao Codex:

```text
Execute a ferramenta SearchImages usando o arquivo config.example.yaml. 
Se houver erro, corrija o script. 
Depois, confira o log_execucao.csv e ajuste os itens marcados como REVISAR quando possível.
```

## Resultado esperado

Uma planilha Excel pronta para uso, com os itens, descrições, marcas, imagens reais dos produtos, unidade, quantidade, valor unitário e valor total calculado.
