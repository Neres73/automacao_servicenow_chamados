# Automação de Fechamento de Chamados — ServiceNow

Automação (RPA) em Python que lê chamados de uma planilha Excel e executa o fluxo completo de fechamento no ServiceNow de forma autônoma: busca cada chamado, valida os dados, preenche os campos necessários e encerra o registro — gerando ainda um relatório do que foi concluído com sucesso e do que precisou de atenção.

## Sobre o projeto

O fechamento de chamados no ServiceNow era feito manualmente, um a um, a partir de uma planilha compartilhada com dezenas de registros. Cada chamado seguia sempre a mesma sequência de passos repetitivos. Este projeto automatiza esse processo de ponta a ponta, reduzindo o tempo gasto e a margem de erro.

## Tecnologias

- **Python**
- **Playwright** — automação de navegador (Microsoft Edge)
- **openpyxl** — leitura e escrita de planilhas Excel
- **python-dotenv** — gerenciamento de configuração

## Como funciona

1. Lê os dados dos chamados a partir de uma planilha (arquivo local ou download automático do Excel Online).
2. Para cada chamado, monta uma busca filtrada por URL no ServiceNow e abre o registro correspondente.
3. Valida se o técnico atribuído confere com o esperado antes de qualquer ação.
4. Executa a sequência de fechamento: aceitar, iniciar trabalho, preencher horários reais, informações de fechamento e encerrar.
5. Marca o status de volta na planilha (quando configurado) e gera um relatório `.csv` com o resultado de cada chamado.

A automação foi projetada para ser resiliente: pula etapas que já foram feitas por outra pessoa, lida com janelas de confirmação, registra capturas de tela em caso de erro e continua processando os demais chamados mesmo que um falhe.

## Instalação

```bash
# 1. Instalar as dependências
pip install -r requirements.txt

# 2. Instalar o navegador (Microsoft Edge)
playwright install msedge

# 3. Criar o arquivo de configuração a partir do modelo
cp .env.example .env
```

Depois, edite o arquivo `.env` com os valores do seu ambiente (URLs, caminho da planilha, etc.).

## Uso

**Primeiro login** (apenas uma vez, ou quando a sessão expirar):

```bash
python main.py --login
```

Isso abre o navegador para você fazer login manualmente. A sessão fica salva localmente, então as execuções seguintes já abrem autenticadas.

**Execução normal:**

```bash
python main.py
```

## Configuração (`.env`)

As variáveis de ambiente necessárias estão documentadas no arquivo `.env.example`. As principais:

- `SERVICENOW_URL` — URL base da sua instância do ServiceNow
- `EXCEL_LOCAL_PATH` — caminho da planilha local (recomendado; opcional)
- `EXCEL_ONLINE_URL` — link da planilha no Excel Online (usado se não houver arquivo local)

## Observações importantes

> Os seletores de interface (botões, campos, abas) são específicos da instância do ServiceNow em que a automação foi desenvolvida. Ao adaptar para outro ambiente, será necessário ajustar os seletores no código. O passo a passo de como localizar cada seletor está comentado no próprio `main.py` e no `MANUAL_INTERNO.md`.

> O arquivo `.env`, a pasta de sessão do navegador e as planilhas **não** são versionados (ver `.gitignore`), por conterem credenciais e dados sensíveis.

## Aviso

Este projeto foi desenvolvido para um contexto específico e é compartilhado para fins de portfólio e referência. Automações que interagem com sistemas corporativos devem sempre ser usadas em conformidade com as políticas da organização.
