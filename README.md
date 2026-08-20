# Automação de Fechamento de Chamados — ServiceNow

Automação (RPA) em Python + Playwright que lê chamados de uma planilha Excel e executa o fluxo completo de fechamento no ServiceNow: localiza cada registro, valida o responsável, preenche horários reais e informações de fechamento e encerra — gerando relatório auditável de cada execução.

> Processo antes executado manualmente, chamado a chamado, passou a rodar em lote sem intervenção humana — sobrando para a pessoa apenas os casos que exigem decisão.

---

## Contexto

O fechamento de ordens de serviço era feito um a um, a partir de uma planilha compartilhada com dezenas de registros por lote. A sequência de passos era sempre a mesma — abrir o chamado, aceitar, iniciar, preencher horários, fechar —, mas com dois agravantes que impedem uma automação ingênua:

- **os dados de entrada são sujos**: a planilha é preenchida por várias pessoas, com números em formatos diferentes e datas em padrões mistos;
- **o estado dos chamados varia**: parte deles já foi parcialmente processada por outra pessoa quando o lote roda.

O projeto automatiza o fluxo de ponta a ponta tratando essas duas realidades, em vez de assumir o caminho feliz.

---

## Como funciona

1. Obtém a planilha — arquivo local sincronizado (preferencial) ou download automático do Excel Online
2. Lê e normaliza os dados de cada chamado (número, técnico, validador, horários)
3. Localiza o registro no ServiceNow por busca filtrada via URL
4. **Valida o técnico atribuído antes de qualquer ação** — divergência descarta o chamado
5. Executa a sequência de fechamento: aceitar → iniciar trabalho → horários reais → informações de fechamento → encerrar
6. Grava o status de volta na planilha (quando configurado) e emite relatório `.csv` da execução

---

## Stack

| Camada | Tecnologia |
|---|---|
| Linguagem | Python 3.10+ |
| Automação de navegador | Playwright (Microsoft Edge, contexto persistente) |
| Planilhas | openpyxl |
| Configuração | Variáveis de ambiente (`python-dotenv`) |

---

## Decisões técnicas

- **Validação antes de ação.** O técnico atribuído no ServiceNow é conferido contra a planilha antes do primeiro clique. Num fluxo que encerra registros de trabalho, o erro seguro é não executar.
- **Etapas idempotentes.** Cada passo do fechamento verifica se o controle existe na tela; se não existe, assume que a etapa já foi concluída por outra pessoa e continua do estado real do registro. O lote pode rodar sobre chamados em estágios diferentes sem quebrar.
- **Normalização defensiva da entrada.** O número do chamado é reconstruído a partir do que a célula realmente contém — incluindo o caso clássico do Excel remover o zero à esquerda de células numéricas — e a estrutura da planilha é conferida pelos cabeçalhos, alertando se alguma coluna foi deslocada.
- **Falha isolada por item.** Erro em um chamado registra o motivo, salva screenshot da tela no momento da falha e segue para o próximo. O lote nunca é derrubado por um caso.
- **Diálogos de confirmação tratados por evento.** O popup de confirmação do ServiceNow pode surgir em qualquer ponto; o handler é registrado no evento de load da página, além de chamadas diretas nos pontos críticos.
- **Sessão persistente, sem credenciais no código.** O login é feito manualmente uma única vez num perfil de navegador persistido em disco. A automação nunca lê, armazena ou digita senhas.

---

## Instalação

```bash
git clone https://github.com/Neres73/automacao_servicenow_chamados.git
cd automacao_servicenow_chamados

python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install -r requirements.txt
playwright install msedge
```

---

## Configuração

Toda a configuração fica em variáveis de ambiente — **nenhum dado de ambiente corporativo é versionado**.

```bash
cp .env.example .env
```

| Variável | Descrição | Obrigatória |
|---|---|---|
| `SERVICENOW_URL` | URL base da instância do ServiceNow | Sim |
| `SERVICENOW_TASK_LIST_URL` | URL da listagem de tarefas usada na busca | Não (padrão: `SERVICENOW_URL`) |
| `EXCEL_LOCAL_PATH` | Caminho da planilha local sincronizada — recomendado; habilita a gravação do status de volta | Não |
| `EXCEL_ONLINE_URL` | Link de download da planilha no Excel Online — usado quando não há arquivo local | Sim, se não houver `EXCEL_LOCAL_PATH` |
| `DOWNLOAD_DIR` | Pasta de downloads temporários | Não (padrão: `./downloads`) |
| `USER_DATA_DIR` | Pasta do perfil de navegador com a sessão salva | Não (padrão: `./perfil_navegador`) |

---

## Uso

### Primeira execução — autenticação

```bash
python main.py --login
```

Abre o navegador para login manual na Microsoft e no ServiceNow. A sessão fica persistida e as execuções seguintes já abrem autenticadas.

### Execução normal

```bash
python main.py
```

---

## Saídas de cada execução

| Saída | Onde | Conteúdo |
|---|---|---|
| Relatório | `resultados/relatorio_<timestamp>.csv` | Uma linha por chamado: número, sucesso/falha e motivo |
| Lista de falhas | `resultados/sgms_com_falha_<timestamp>.txt` | Só os números que falharam, prontos para reprocessar |
| Evidência de erro | `erros/erro_<numero>.png` | Screenshot da tela no momento de cada falha |
| Status na planilha | Coluna "Status" (criada se não existir) | "Concluído" ou o motivo do erro — apenas com `EXCEL_LOCAL_PATH` |

---

## Limitações e próximos passos

- Os seletores de interface são específicos da instância em que a automação foi desenvolvida — adaptar para outro ambiente exige ajustá-los no código
- Execução sequencial e com navegador visível; modo headless ainda não foi validado
- Sem suíte de testes automatizados — as funções puras de normalização são as primeiras candidatas
- Próximas evoluções: `argparse` com modo `--dry-run` (simular o lote sem encerrar nada), `logging` estruturado em arquivo e separação em módulos (planilha / ServiceNow / relatório)

---

## Estrutura

```
.
├── main.py            # fluxo completo: leitura, validação, fechamento e relatório
├── requirements.txt
├── .env.example       # modelo de configuração
├── .gitignore         # ignora .env, perfil de navegador, planilhas e saídas
└── README.md
```

As pastas `downloads/`, `erros/`, `resultados/` e `perfil_navegador/` são criadas em tempo de execução e não são versionadas.

---

## Aviso

Projeto desenvolvido a partir de uma necessidade real de trabalho e compartilhado para fins de portfólio. Nenhuma URL, credencial ou dado corporativo está versionado neste repositório. Automações que interagem com sistemas corporativos devem sempre respeitar as políticas da organização.
