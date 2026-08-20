"""
Automação de encerramento de ordens de serviço no ServiceNow.

Lê os números SGM de uma planilha (Excel Online ou arquivo local
sincronizado), localiza a tarefa MAN correspondente no ServiceNow e executa
o fluxo completo de encerramento: aceite, início do trabalho, preenchimento
dos horários reais, informações de fechamento e conclusão. Ao final, grava
o status de cada chamado na planilha (quando aplicável) e gera relatório
em CSV.

A autenticação usa um perfil de navegador persistente — o script nunca
manipula credenciais.

Uso:
    python main.py --login    # primeira execução: login manual, salva a sessão
    python main.py            # execução normal

Configuração via arquivo .env — ver README.md.
"""

import csv
import os
import re
import sys
import time
from datetime import datetime, date as date_cls, time as time_cls
from pathlib import Path

import openpyxl
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, Page, BrowserContext

load_dotenv()

EXCEL_ONLINE_URL = os.getenv("EXCEL_ONLINE_URL")
EXCEL_LOCAL_PATH = os.getenv("EXCEL_LOCAL_PATH", "").strip().strip('"').strip("'")
SERVICENOW_URL = os.getenv("SERVICENOW_URL")
SERVICENOW_TASK_LIST_URL = os.getenv("SERVICENOW_TASK_LIST_URL", SERVICENOW_URL)
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "./downloads")).resolve()
USER_DATA_DIR = Path(os.getenv("USER_DATA_DIR", "./perfil_navegador")).resolve()

DOWNLOAD_DIR.mkdir(exist_ok=True, parents=True)
USER_DATA_DIR.mkdir(exist_ok=True, parents=True)

ERROS_DIR = Path("./erros").resolve()
ERROS_DIR.mkdir(exist_ok=True, parents=True)

RESULTADOS_DIR = Path("./resultados").resolve()
RESULTADOS_DIR.mkdir(exist_ok=True, parents=True)

# Índices (0-based) das colunas da planilha de chamados.
COLUNA_SGM_INDEX = 4                  # E — "Número Chamado SGM"
COLUNA_TECNICO_INDEX = 5              # F — técnico (validado contra o ServiceNow)
COLUNA_VALIDADOR_INDEX = 6            # G — "Coordenador (a)" (validador)
COLUNA_MATRICULA_INDEX = 7            # H — matrícula do validador
COLUNA_DATA_ATENDIMENTO_INDEX = 8     # I — "Data de atendimento do chamado"
COLUNA_HORARIO_INICIO_INDEX = 9       # J — "Horário de inicio do atendimento"
COLUNA_HORARIO_TERMINO_INDEX = 10     # K — "Horário do fim do atendimento"

# Formato de data/hora aceito pelos campos de data do ServiceNow.
FORMATO_DATA_SERVICENOW = "%d/%m/%Y %H:%M:%S"


# ============================================================================
# ETAPA 1 — Obtenção da planilha
# ============================================================================
def baixar_planilha_excel_online(page: Page) -> Path:
    """
    Baixa a planilha a partir da URL do Excel Online.

    A URL configurada serve o arquivo como download direto. Nesse cenário o
    page.goto() levanta exceção — a navegação é interrompida quando vira
    download —, comportamento esperado do Playwright; o download em si é
    capturado pelo expect_download().
    """
    with page.expect_download() as download_info:
        try:
            page.goto(EXCEL_ONLINE_URL)
        except Exception:
            pass  # esperado: a navegação é abortada quando vira download
    download = download_info.value

    caminho_arquivo = DOWNLOAD_DIR / "planilha_chamados.xlsx"
    download.save_as(caminho_arquivo)
    return caminho_arquivo


# ============================================================================
# ETAPA 2 — Leitura e normalização dos dados
# ============================================================================
def combinar_data_hora(data_valor, hora_valor):
    """
    Combina a data de atendimento com um horário em um único datetime.

    Aceita os tipos que o openpyxl pode entregar para células de data/hora
    (datetime, date, time ou texto). Retorna None quando algum dos valores
    não pode ser interpretado — o chamador decide como reportar.
    """
    if isinstance(data_valor, datetime):
        data = data_valor.date()
    elif isinstance(data_valor, date_cls):
        data = data_valor
    elif isinstance(data_valor, str):
        data = None
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d/%m/%y", "%m/%d/%y"):
            try:
                data = datetime.strptime(data_valor.strip(), fmt).date()
                break
            except ValueError:
                continue
        if data is None:
            return None
    else:
        return None

    if isinstance(hora_valor, datetime):
        hora = hora_valor.time()
    elif isinstance(hora_valor, time_cls):
        hora = hora_valor
    elif isinstance(hora_valor, str):
        hora = None
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                hora = datetime.strptime(hora_valor.strip(), fmt).time()
                break
            except ValueError:
                continue
        if hora is None:
            return None
    else:
        return None

    return datetime.combine(data, hora)


def formatar_data_hora_servicenow(valor) -> str:
    """
    Formata um datetime no padrão aceito pelos campos de data do ServiceNow
    (DD/MM/AAAA HH:MM:SS), que difere do formato de exibição da planilha.

    Valores que chegarem como texto puro são repassados como estão, com
    aviso — a conversão só é garantida para objetos de data reais.
    """
    if valor is None:
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime(FORMATO_DATA_SERVICENOW)
    texto = str(valor).strip()
    print(f"Aviso: valor de data/hora veio como texto puro ('{texto}') e não foi "
          "reformatado — confirmar se está em DD/MM/AAAA HH:MM:SS.")
    return texto


def normalizar_sgm(valor_bruto) -> str:
    """
    Normaliza o número SGM para o formato padrão "SGM" + 7 dígitos.

    Trata os desvios comuns de digitação e de formatação do Excel:
    - espaços em qualquer posição ("SGM 0670464") são removidos;
    - células numéricas perdem o zero à esquerda ("0751687" é lido como
      "751687"), então valores só com dígitos são completados com zfill(7)
      antes de receber o prefixo.

    Valores fora do padrão são usados como estão, com aviso para
    conferência manual.
    """
    texto = str(valor_bruto).strip().replace(" ", "")

    if texto.isdigit():
        texto = "SGM" + texto.zfill(7)

    if not re.fullmatch(r"SGM\d{7}", texto):
        print(f"Aviso: SGM '{texto}' (original: '{valor_bruto}') fora do formato "
              "esperado SGM + 7 dígitos — conferir manualmente.")

    return texto


def ler_itens_da_planilha(caminho_arquivo: Path) -> list[dict]:
    """
    Lê a planilha de chamados e devolve a lista de itens a processar.

    Antes da leitura, os cabeçalhos das colunas mapeadas são conferidos
    contra os nomes esperados: se a estrutura da planilha mudar (uma coluna
    inserida desloca todas as seguintes), o desvio é reportado em vez de
    processar dados errados em silêncio.
    """
    wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
    ws = wb.active

    colunas_esperadas = {
        COLUNA_SGM_INDEX: "Número Chamado SGM",
        COLUNA_TECNICO_INDEX: "Técnico",
        COLUNA_VALIDADOR_INDEX: "Coordenador (a)",
        COLUNA_MATRICULA_INDEX: "Técnico",
        COLUNA_DATA_ATENDIMENTO_INDEX: "Data de atendimento do chamado",
        COLUNA_HORARIO_INICIO_INDEX: "Horário de inicio do atendimento",
        COLUNA_HORARIO_TERMINO_INDEX: "Horário do fim do atendimento",
    }
    for indice, cabecalho_esperado in colunas_esperadas.items():
        cabecalho_real = ws.cell(row=1, column=indice + 1).value  # openpyxl é 1-based
        if cabecalho_real != cabecalho_esperado:
            print(
                f"Aviso: esperado o cabeçalho '{cabecalho_esperado}' na coluna "
                f"{indice + 1}, encontrado '{cabecalho_real}'. A planilha pode "
                "ter mudado de estrutura."
            )

    itens = []
    for row in ws.iter_rows(min_row=2):  # pula o cabeçalho
        celula_sgm = row[COLUNA_SGM_INDEX]
        if not celula_sgm.value:
            continue
        tecnico = row[COLUNA_TECNICO_INDEX].value
        validador = row[COLUNA_VALIDADOR_INDEX].value
        matricula = row[COLUNA_MATRICULA_INDEX].value

        data_atendimento = row[COLUNA_DATA_ATENDIMENTO_INDEX].value
        hora_inicio = row[COLUNA_HORARIO_INICIO_INDEX].value
        hora_termino = row[COLUNA_HORARIO_TERMINO_INDEX].value

        sgm_atual = normalizar_sgm(celula_sgm.value)

        inicio_combinado = combinar_data_hora(data_atendimento, hora_inicio)
        if inicio_combinado is None:
            print(f"Aviso: não foi possível combinar data+hora de início do SGM "
                  f"{sgm_atual} (data='{data_atendimento}', hora='{hora_inicio}').")

        termino_combinado = combinar_data_hora(data_atendimento, hora_termino)
        if termino_combinado is None:
            print(f"Aviso: não foi possível combinar data+hora de término do SGM "
                  f"{sgm_atual} (data='{data_atendimento}', hora='{hora_termino}').")

        itens.append({
            "linha": celula_sgm.row,  # linha original, usada para gravar o status
            "sgm": sgm_atual,
            "inicio_real": formatar_data_hora_servicenow(inicio_combinado) if inicio_combinado else "",
            "termino_real": formatar_data_hora_servicenow(termino_combinado) if termino_combinado else "",
            "tecnico": str(tecnico).strip() if tecnico is not None else "",
            "validador": str(validador).strip() if validador is not None else "",
            "matricula": str(matricula).strip() if matricula is not None else "",
        })
    return itens


# ============================================================================
# ETAPA 3 — Localização e encerramento do chamado no ServiceNow
# ============================================================================
def obter_alvo_servicenow(page: Page):
    """
    Retorna o objeto de interação correto para a tela atual: o iframe
    'gsft_main' (interface clássica) quando existir, ou a própria página
    (Next Experience).
    """
    frame = page.frame(name="gsft_main")
    return frame if frame else page


def montar_url_busca_por_sgm(sgm: str) -> str:
    """
    Monta a URL de busca da tarefa a partir do número SGM.

    A tarefa MAN (tabela wm_task) é um registro filho do chamado de origem,
    então a consulta filtra por 'parent.number' com STARTSWITH — buscar o
    número do pai é o caminho direto até a tarefa correspondente.
    """
    base = SERVICENOW_TASK_LIST_URL.split("?")[0]
    return f"{base}?sysparm_query=parent.numberSTARTSWITH{sgm}"


def buscar_chamado_e_clicar_man(page: Page, sgm: str):
    """
    Pesquisa o SGM na listagem de tarefas e abre o registro MAN.

    O número MAN não é um campo separado: é a numeração padrão dos
    registros da tabela wm_task. Por isso o alvo do clique é qualquer link
    cujo texto comece com "MAN" seguido de dígitos, independentemente da
    posição na lista.
    """
    if "login" in page.url.lower():
        raise RuntimeError(
            "Sessão do ServiceNow expirou ou não está logada. "
            "Rode: python main.py --login"
        )

    page.goto(montar_url_busca_por_sgm(sgm))
    page.wait_for_load_state("networkidle")

    alvo = obter_alvo_servicenow(page)

    link_man = alvo.get_by_text(re.compile(r"^MAN\d+"))

    # Sem resultados, falha imediatamente em vez de aguardar o timeout
    # padrão de clique — acelera o avanço para o próximo chamado.
    if link_man.count() == 0:
        raise RuntimeError(f"SGM {sgm} não encontrado no ServiceNow.")

    link_man.first.click(timeout=5000)
    alvo.wait_for_load_state("networkidle")

    print(f"Chamado {sgm}: registro MAN aberto.")

    # O reload garante que o formulário carregue no estado atual do
    # registro; como o iframe é recriado, o alvo é obtido novamente.
    page.reload()
    page.wait_for_load_state("networkidle")
    fechar_popup_chamados_pausados(page)
    return obter_alvo_servicenow(page)


def obter_ou_criar_coluna_status(ws) -> int:
    """
    Localiza a coluna "Status" na planilha; se não existir, cria uma nova
    após a última coluna usada. Retorna o índice 1-based.
    """
    for cell in ws[1]:
        if cell.value == "Status":
            return cell.column
    nova_coluna = ws.max_column + 1
    ws.cell(row=1, column=nova_coluna, value="Status")
    return nova_coluna


def fechar_popup_chamados_pausados(page: Page) -> None:
    """
    Fecha o diálogo de confirmação de chamados em andamento ("Tem
    certeza? ... Chamado(s): MANxxxxxxx") clicando em OK (id "ok_button").

    O diálogo pode surgir em qualquer ponto da execução, então esta função
    é registrada no evento de load da página e também chamada diretamente
    nos pontos críticos do fluxo. Falhas na checagem são ignoradas — a
    ausência do diálogo é o caso normal.
    """
    try:
        botao_ok = page.locator("#ok_button")
        if botao_ok.count() > 0 and botao_ok.first.is_visible():
            botao_ok.first.click(timeout=2000)
            print("  (diálogo de chamados em andamento fechado)")
    except Exception:
        pass


def clicar_se_existir(alvo, localizador, descricao: str) -> bool:
    """
    Clica no elemento se ele estiver presente; caso contrário, considera a
    etapa já concluída (ou não aplicável) e segue sem erro. Retorna True se
    clicou.

    O force=True é necessário porque os botões de ação do ServiceNow podem
    ficar recolhidos em um menu overflow quando a tela é estreita — o clique
    é disparado mesmo sem a checagem de visibilidade.
    """
    if localizador.count() == 0:
        print(f"  (etapa \"{descricao}\" ignorada: não presente na tela — já concluída ou não se aplica)")
        return False
    localizador.first.click(timeout=5000, force=True)
    alvo.wait_for_load_state("networkidle")
    return True


def clicar_esperando(alvo, localizador, descricao: str, segundos: int = 8) -> bool:
    """
    Variante de clicar_se_existir que aguarda o elemento surgir por até o
    tempo indicado antes de desistir — para controles que aparecem alguns
    instantes após a ação anterior. Retorna True se clicou.
    """
    try:
        localizador.first.wait_for(state="attached", timeout=segundos * 1000)
    except Exception:
        print(f"  (etapa \"{descricao}\" ignorada: não surgiu após {segundos}s — não se aplica)")
        return False
    localizador.first.click(timeout=5000, force=True)
    alvo.wait_for_load_state("networkidle")
    return True


def validar_tecnico_assignado(alvo, tecnico_planilha: str) -> None:
    """
    Confere se o técnico atribuído no ServiceNow corresponde ao informado
    na planilha antes de qualquer ação sobre o chamado.

    A comparação usa o primeiro nome e o operador "contém", porque o campo
    do ServiceNow pode trazer matrícula ou outras informações junto ao
    nome. Divergência interrompe o processamento deste chamado — ele é
    registrado como falha no relatório.
    """
    if not tecnico_planilha:
        return  # sem nome na planilha, não há o que validar

    primeiro_nome = tecnico_planilha.strip().split()[0]
    valor_servicenow = alvo.input_value("input[id='sys_display.wm_task.assigned_to']")

    if primeiro_nome.lower() not in valor_servicenow.lower():
        raise RuntimeError(
            f"Técnico não confere: esperado '{primeiro_nome}' no campo "
            f"'Atribuído a', encontrado '{valor_servicenow.strip()}'."
        )


def preencher_etapas_man(page: Page, alvo, inicio_real: str, termino_real: str,
                          validador: str, matricula: str, tecnico: str) -> None:
    """
    Executa o fluxo de encerramento dentro do registro MAN:
    Aceitar -> Start Work -> Não Aplicar Material -> horários reais ->
    informações de fechamento -> Encerrado Concluído.

    As três primeiras etapas são opcionais por design: se o chamado já
    avançou (outro usuário clicou em Aceitar, por exemplo), o botão
    correspondente não existe e o fluxo continua do estado real do
    registro. O chamado só é descartado se nenhuma das três etapas
    iniciais estiver disponível.

    Escolha de seletores: ids fixos ("accept", "close_complete") têm
    prioridade sobre texto, que varia em caixa e acentuação. O padrão
    "^Não Aplicar Material$" é ancorado para não colidir com o botão
    vizinho "Aplicar Material".
    """
    # A validação de técnico acontece antes de qualquer clique — impede
    # aceitar ou processar chamado atribuído a outra pessoa.
    validar_tecnico_assignado(alvo, tecnico)

    loc_aceitar = alvo.locator("#accept")
    loc_start_work = alvo.get_by_text("Start Work", exact=True)
    loc_nao_aplicar = alvo.get_by_text(re.compile(r"^Não Aplicar Material$", re.IGNORECASE))

    if loc_aceitar.count() == 0 and loc_start_work.count() == 0 and loc_nao_aplicar.count() == 0:
        raise RuntimeError(
            "Nenhuma das etapas iniciais (Aceitar / Start Work / Não Aplicar "
            "Material) foi encontrada — chamado descartado."
        )

    clicar_se_existir(alvo, loc_aceitar, "Aceitar")
    clicar_se_existir(alvo, loc_start_work, "Start Work")
    # Este controle surge alguns instantes após o Start Work, por isso a
    # variante com espera.
    clicar_esperando(alvo, loc_nao_aplicar, "Não Aplicar Material", segundos=8)

    # Aba "Real" — horários efetivos do atendimento.
    alvo.locator("span.tab_caption_text:text-is('Real')").first.click(force=True)
    alvo.wait_for_load_state("networkidle")

    # Campos de data do ServiceNow só confirmam o valor quando o campo
    # perde o foco — o Tab após cada preenchimento simula esse comportamento.
    alvo.fill("input[id='wm_task.work_start']", inicio_real)
    alvo.press("input[id='wm_task.work_start']", "Tab")

    alvo.fill("input[id='wm_task.work_end']", termino_real)
    alvo.press("input[id='wm_task.work_end']", "Tab")

    # Aba "Informações de Fechamento".
    alvo.locator("span.tab_caption_text:text-is('Informações de Fechamento')").first.click(force=True)
    alvo.wait_for_load_state("networkidle")

    alvo.select_option("select[id='wm_task.u_fsm_close_code']", "Outros")
    alvo.fill("input[id='wm_task.u_validated_by']", validador)
    alvo.fill("input[id='wm_task.u_validated_id_by']", matricula)

    fechar_popup_chamados_pausados(page)

    alvo.locator("#close_complete").click(force=True)
    alvo.wait_for_load_state("networkidle")

    print(f"  Início real: {inicio_real} | Término real: {termino_real} — chamado encerrado.")


# ============================================================================
# Login manual (modo --login)
# ============================================================================
def fazer_login_manual(context: BrowserContext) -> None:
    """
    Abre as páginas de autenticação para login manual e persiste a sessão
    no perfil de navegador. As credenciais nunca passam pelo script.
    """
    page = context.new_page()
    page.goto(EXCEL_ONLINE_URL)
    print("\n>>> Faça login na Microsoft nesta janela (Excel Online).")

    page2 = context.new_page()
    page2.goto(SERVICENOW_URL)
    print(">>> Faça login no ServiceNow nesta outra aba.")

    input("\nApós logar nos dois, volte aqui e pressione Enter para salvar a sessão...")
    print("Sessão salva em:", USER_DATA_DIR)


# ============================================================================
# Execução principal
# ============================================================================
def main() -> None:
    modo_login = "--login" in sys.argv

    with sync_playwright() as p:
        # O contexto persistente mantém cookies e sessão entre execuções,
        # como um perfil de navegador comum. channel="msedge" usa o Edge
        # instalado no sistema em vez do Chromium empacotado.
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            channel="msedge",
            headless=False,
            accept_downloads=True,
        )

        if modo_login:
            fazer_login_manual(context)
            context.close()
            return

        page = context.pages[0] if context.pages else context.new_page()

        # Diálogos nativos do navegador são aceitos automaticamente; o
        # diálogo relevante do ServiceNow é HTML comum e é tratado por
        # fechar_popup_chamados_pausados, registrada no evento de load.
        page.on("dialog", lambda dialog: dialog.accept())
        page.on("load", lambda *_args: fechar_popup_chamados_pausados(page))

        if EXCEL_LOCAL_PATH and Path(EXCEL_LOCAL_PATH).exists():
            print(f"Lendo planilha local em: {EXCEL_LOCAL_PATH}")
            caminho_planilha = Path(EXCEL_LOCAL_PATH)
        else:
            if EXCEL_LOCAL_PATH:
                print(
                    f"Aviso: EXCEL_LOCAL_PATH configurado ('{EXCEL_LOCAL_PATH}'), "
                    "mas o arquivo não foi encontrado. Usando o download via "
                    "Excel Online."
                )
            print("Baixando a planilha do Excel Online...")
            caminho_planilha = baixar_planilha_excel_online(page)

        itens = ler_itens_da_planilha(caminho_planilha)
        print(f"{len(itens)} chamado(s) encontrado(s) na planilha.")

        # O status só é gravado de volta quando a origem é o arquivo local
        # sincronizado — escrever na cópia temporária baixada do Excel
        # Online não refletiria na planilha usada pela equipe.
        permite_escrever_status = bool(EXCEL_LOCAL_PATH) and caminho_planilha == Path(EXCEL_LOCAL_PATH)

        if permite_escrever_status:
            wb_escrita = openpyxl.load_workbook(caminho_planilha)  # sem data_only: preserva fórmulas
            ws_escrita = wb_escrita.active
            coluna_status = obter_ou_criar_coluna_status(ws_escrita)
            wb_escrita.save(caminho_planilha)
            print(f"Status será gravado na planilha local (coluna {coluna_status}).")
        else:
            print(
                "Aviso: EXCEL_LOCAL_PATH não configurado (ou não é o arquivo em "
                "uso) — o status não será gravado na planilha; a cópia baixada "
                "do Excel Online é temporária."
            )

        print("Acessando o ServiceNow...")
        page.goto(SERVICENOW_URL)
        page.wait_for_load_state("networkidle")

        sucesso = []
        falha = []

        for item in itens:
            sgm = item["sgm"]
            try:
                alvo = buscar_chamado_e_clicar_man(page, sgm)
                preencher_etapas_man(
                    page, alvo,
                    item["inicio_real"], item["termino_real"],
                    item["validador"], item["matricula"],
                    item["tecnico"],
                )
                sucesso.append(sgm)
                if permite_escrever_status:
                    ws_escrita.cell(row=item["linha"], column=coluna_status, value="Concluído")
                    wb_escrita.save(caminho_planilha)
            except Exception as erro:
                print(f"Erro ao processar o chamado {sgm}: {erro}")
                falha.append((sgm, str(erro)))
                if permite_escrever_status:
                    ws_escrita.cell(row=item["linha"], column=coluna_status, value=str(erro))
                    wb_escrita.save(caminho_planilha)
                # Screenshot no momento da falha — evita ter que reproduzir
                # o erro para investigar.
                try:
                    caminho_print = ERROS_DIR / f"erro_{sgm}.png"
                    page.screenshot(path=str(caminho_print))
                    print(f"  Screenshot do erro salva em: {caminho_print}")
                except Exception:
                    pass  # a falha do screenshot não interrompe o lote

            time.sleep(1)  # pausa entre chamados para não sobrecarregar o ServiceNow

        # Relatório persistente da execução — o terminal não é a única
        # fonte para revisar o que falhou.
        agora = datetime.now().strftime("%Y%m%d_%H%M%S")

        caminho_relatorio = RESULTADOS_DIR / f"relatorio_{agora}.csv"
        with open(caminho_relatorio, "w", newline="", encoding="utf-8-sig") as f:
            escritor = csv.writer(f)
            escritor.writerow(["SGM", "Status", "Motivo"])
            for sgm in sucesso:
                escritor.writerow([sgm, "Sucesso", ""])
            for sgm, motivo in falha:
                escritor.writerow([sgm, "Falha", motivo])

        caminho_falhas = RESULTADOS_DIR / f"sgms_com_falha_{agora}.txt"
        with open(caminho_falhas, "w", encoding="utf-8") as f:
            for sgm, _ in falha:
                f.write(sgm + "\n")

        print("\n--- Resumo da execução ---")
        print(f"Sucesso: {len(sucesso)}/{len(itens)}")
        if falha:
            print(f"Falha: {len(falha)}/{len(itens)}")
            for sgm, motivo in falha:
                print(f"  - {sgm}: {motivo}")
        print(f"\nRelatório completo salvo em: {caminho_relatorio}")
        print(f"Lista dos SGMs com falha salva em: {caminho_falhas}")

        context.close()


if __name__ == "__main__":
    main()
