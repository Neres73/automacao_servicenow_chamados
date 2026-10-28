# ServiceNow — Encerramento de Ordens de Serviço em Lote

Script em Python + Playwright que lê os chamados atendidos numa planilha Excel e encerra cada ordem de serviço no ServiceNow, preenchendo horários reais e informações de fechamento.

Uso no meu trabalho, no back-office de Manutenção. O encerramento era dividido entre várias pessoas da equipe; com o script, uma pessoa só dá conta do lote.

---

## Por que existe

Cada encerramento segue sempre os mesmos passos (aceitar, iniciar, preencher horários, informar o fechamento, encerrar), mas dois detalhes impedem uma automação simples:

- **A planilha é preenchida por várias pessoas**: números de chamado com espaço ou sem o zero da frente, datas e horas em formatos diferentes
- **Os chamados não estão todos no mesmo estado**: alguns já foram aceitos, outros estão pausados

O script trata esses dois casos em vez de assumir que tudo está no formato e no estado esperados.

---

## O que o script faz

1. **Obtém a planilha**: usa o arquivo local sincronizado (recomendado) ou baixa do Excel Online
2. **Lê e corrige os dados**:
   - Confere se os cabeçalhos estão nas colunas esperadas e avisa se a planilha mudou de estrutura
   - Padroniza o número do chamado (`SGM` + 7 dígitos), corrigindo espaços e zeros à esquerda que o Excel remove
   - Junta data e hora do atendimento em qualquer formato que o Excel entregue (data, texto ou fração do dia)
3. **Para cada chamado**:
   - Busca a tarefa (MAN) vinculada ao chamado e abre o registro
   - **Confere se o técnico atribuído é o mesmo da planilha**; se não for, não mexe no chamado
   - Executa só as etapas que ainda faltam: Aceitar → Start Work → Retomar Trabalho → Não aplicar material
   - Preenche início e término reais na aba **Real**
   - Preenche código de fechamento, validador e matrícula na aba **Informações de Fechamento**
   - Clica em **Encerrado Concluído**
4. **Registra o resultado** na planilha, em CSV e com print da tela nos erros

---

## Cuidados embutidos

- **Confirma cada mudança de status**: depois de Start Work, Retomar Trabalho e Encerrar, espera o botão sumir da tela, que é o sinal de que o ServiceNow aceitou. Se não aceitou, lê a mensagem de erro da tela e tenta mais uma vez. Se o Start Work ou o Retomar Trabalho continuarem sem efeito, para aquele chamado com o motivo em vez de seguir com dado errado; no encerramento, avisa no terminal
- **Fecha o aviso de "chamados em andamento"** em qualquer ponto em que ele apareça, inclusive dentro do formulário
- **Espera elementos concretos da tela** em vez de esperar a rede ficar parada, que no ServiceNow quase nunca acontece
- **Erro em um chamado não derruba o lote**: o motivo é registrado e o script segue para o próximo
- **Detecta sessão expirada** e orienta a refazer o login

---

## Saídas de cada execução

| Saída | Onde | Conteúdo |
|---|---|---|
| Relatório | `resultados/relatorio_<data_hora>.csv` | Um chamado por linha: sucesso ou falha, com o motivo |
| Lista de falhas | `resultados/sgms_com_falha_<data_hora>.txt` | Só os chamados que falharam, prontos para rodar de novo |
| Prints de erro | `erros/erro_<chamado>.png` | Tela no momento de cada falha |
| Status na planilha | Coluna "Status" (criada se não existir) | "Concluído" ou o motivo do erro (só com planilha local) |

---

## Instalação

Requisitos: Python 3.10+ e Microsoft Edge.

```bash
git clone https://github.com/Neres73/automacao_servicenow_chamados.git
cd automacao_servicenow_chamados

python -m venv .venv
.venv\Scripts\activate            # Linux/Mac: source .venv/bin/activate

pip install -r requirements.txt
playwright install msedge
```

## Configuração

Copie o modelo e preencha com os dados do seu ambiente:

```bash
copy .env.example .env            # Linux/Mac: cp .env.example .env
```

| Variável | O que é | Padrão |
|---|---|---|
| `SERVICENOW_URL` | URL da sua instância do ServiceNow | obrigatória |
| `SERVICENOW_TASK_LIST_URL` | URL da lista de tarefas usada na busca | igual a `SERVICENOW_URL` |
| `EXCEL_LOCAL_PATH` | Planilha local sincronizada (permite gravar o status) | — |
| `EXCEL_ONLINE_URL` | Link de download da planilha no Excel Online | usado se não houver arquivo local |
| `DOWNLOAD_DIR` | Pasta de downloads temporários | `./downloads` |
| `USER_DATA_DIR` | Pasta do perfil do navegador com a sessão | `./perfil_navegador` |

**Colunas lidas da planilha** (cabeçalho na linha 1):

| Coluna | Conteúdo |
|---|---|
| E | Número Chamado SGM |
| G | Técnico (comparado com o técnico atribuído no ServiceNow) |
| I | Data de atendimento do chamado |
| J / K | Horário de início / fim do atendimento |
| L | Coordenador (validador) |
| M | Matrícula do validador |

Os seletores que podem variar de uma instância do ServiceNow para outra estão marcados no código com `🔧 AJUSTE AQUI`.

## Uso

```bash
python main.py --login    # 1ª vez: abre o Edge para login na Microsoft e no ServiceNow; a sessão fica salva
python main.py            # execuções seguintes
```

O script **não lê, guarda nem digita senha**.

---

## Limitações

- Seletores e campos personalizados são os da instância onde o script foi feito; outra instância exige ajustes
- O código de fechamento é sempre "Outros"
- Roda um chamado por vez, com o navegador visível
- Sem modo de simulação: rodar o script **encerra os chamados de verdade**

## Estrutura

```
├── main.py            # leitura da planilha, encerramento no ServiceNow e relatórios
├── requirements.txt
├── .env.example       # modelo de configuração
└── .gitignore         # ignora .env, perfil do navegador, planilhas e saídas
```

---

Desenvolvido por [Gustavo Neres](https://github.com/Neres73), com apoio do Claude (IA) para modelar a lógica, escrever e revisar o código. As regras de negócio e os critérios de validação vêm do processo real e foram verificados em execução. Projeto compartilhado como portfólio: nenhuma URL, credencial ou dado da empresa está no repositório. Licença MIT.
