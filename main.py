"""
Automação de encerramento de ordens de serviço no ServiceNow.

Lê os chamados de uma planilha (arquivo local ou Excel Online), localiza a
tarefa MAN de cada um, confere o técnico atribuído e executa o fluxo de
encerramento: aceite, início/retomada do trabalho, horários reais,
informações de fechamento e conclusão. Grava o status na planilha (quando
local) e gera relatório CSV da execução.

LOGIN: usa um perfil de navegador persistente — o script nunca lê, guarda
ou digita usuário e senha.

  Primeiro uso (login manual):   python main.py --login
  Uso normal (já logado):        python main.py

Configuração via arquivo .env (veja .env.example). Os seletores que podem
variar entre instâncias do ServiceNow estão marcados com:  🔧 AJUSTE AQUI
"""

import csv
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import (
    sync_playwright,
    Page,
    BrowserContext,
    TimeoutError as PlaywrightTimeoutError,
)

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
COLUNA_SGM_INDEX = 4                  # Coluna E => "Número Chamado SGM"
COLUNA_TECNICO_INDEX = 6              # Coluna G => nome do técnico (validado contra o ServiceNow)
COLUNA_DATA_ATENDIMENTO_INDEX = 8     # Coluna I => "Data de atendimento do chamado"
COLUNA_HORARIO_INICIO_INDEX = 9       # Coluna J => "Horário de inicio do atendimento"
COLUNA_HORARIO_TERMINO_INDEX = 10     # Coluna K => "Horário do fim do atendimento"
COLUNA_VALIDADOR_INDEX = 11           # Coluna L => "Coordenador (a)" (nome do validador)
COLUNA_MATRICULA_INDEX = 12           # Coluna M => matrícula do validador


# ============================================================================
# ESPERAS — substituem o antigo wait_for_load_state("networkidle")
# ============================================================================
# O ServiceNow mantém requisições em segundo plano o tempo todo (polling,
# notificações, websocket), então "networkidle" quase nunca acontece e o
# Playwright estoura o timeout de 30s. Em vez disso: esperamos o HTML
# carregar (domcontentloaded) e depois um elemento CONCRETO da tela.

# 🔧 AJUSTE AQUI se algum seletor não bater com o seu ServiceNow
SELETOR_SERVICENOW_CARREGADO = "iframe#gsft_main, #gsft_main, sn-polaris-layout, table.data_list_table, form"
SELETOR_LISTA_CARREGADA = "table.data_list_table, .list2_no_records, tr.list2_no_records"
SELETOR_FORM_MAN = "input[id='sys_display.wm_task.assigned_to']"
SELETOR_CAMPO_INICIO_REAL = "input[id='wm_task.work_start']"
SELETOR_CODIGO_FECHAMENTO = "select[id='wm_task.u_fsm_close_code']"


def aguardar_pagina(alvo, seletor: str = None, timeout: int = 30000, estado: str = "visible") -> None:
    """
    Espera o documento carregar (domcontentloaded) e, se informado, espera
    o `seletor` ficar no `estado` pedido ("visible" ou "attached" — use
    "attached" para <input type="hidden">).
    """
    alvo.wait_for_load_state("domcontentloaded", timeout=timeout)
    if seletor:
        alvo.locator(seletor).first.wait_for(state=estado, timeout=timeout)
    else:
        alvo.wait_for_timeout(1000)  # folga mínima quando não há seletor específico


# ============================================================================
# ETAPA 1 — Baixar a planilha atualizada do Excel Online
# ============================================================================
def baixar_planilha_excel_online(page: Page) -> Path:
    try:
        with page.expect_download(timeout=30000) as download_info:
            try:
                page.goto(EXCEL_ONLINE_URL, timeout=30000)
            except Exception:
                pass  # esperado: a navegação "falha" porque virou download direto
        download = download_info.value
    except PlaywrightTimeoutError:
        raise RuntimeError(
            "Nenhum download começou em 30s ao acessar o Excel Online. "
            "Provavelmente a sessão expirou (a página carregou uma tela de "
            "login em vez de baixar o arquivo). Rode: python main.py --login "
            "e faça login novamente. Também confira se não há outra janela "
            "do Edge aberta usando o mesmo perfil (./perfil_navegador)."
        )

    caminho_arquivo = DOWNLOAD_DIR / "planilha_chamados.xlsx"
    download.save_as(caminho_arquivo)

    falha = download.failure()
    if falha:
        raise RuntimeError(f"Download falhou: {falha}")

    return caminho_arquivo


# ============================================================================
# ETAPA 2 — Ler os dados da planilha (SGM, início real, término real)
# ============================================================================
def combinar_data_hora(data_valor, hora_valor):
    from datetime import datetime as dt, date as date_cls, time as time_cls

    if isinstance(data_valor, dt):
        data = data_valor.date()
    elif isinstance(data_valor, date_cls):
        data = data_valor
    elif isinstance(data_valor, str):
        data = None
        # Inclui o formato ISO ("2026-10-22 00:00:00"), que aparece quando a
        # célula de data está formatada como TEXTO na planilha.
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d/%m/%y", "%m/%d/%y",
                    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                data = dt.strptime(data_valor.strip(), fmt).date()
                break
            except ValueError:
                continue
        if data is None:
            return None
    else:
        return None

    if isinstance(hora_valor, dt):
        hora = hora_valor.time()
    elif isinstance(hora_valor, time_cls):
        hora = hora_valor
    elif isinstance(hora_valor, str):
        hora = None
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                hora = dt.strptime(hora_valor.strip(), fmt).time()
                break
            except ValueError:
                continue
        if hora is None:
            return None
    elif isinstance(hora_valor, (int, float)):
        # Excel guarda hora como fração do dia (ex: 0.4375 = 10:30)
        segundos = round(float(hora_valor) % 1 * 86400)
        hora = time_cls(segundos // 3600 % 24, segundos % 3600 // 60, segundos % 60)
    else:
        return None

    return dt.combine(data, hora)


def formatar_data_hora_servicenow(valor) -> str:
    if valor is None:
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y %H:%M:%S")
    texto = str(valor).strip()
    print(f"⚠️  Aviso: valor de data/hora veio como texto puro ('{texto}'), não consegui reformatar — confirme se está no formato certo (DD/MM/AAAA HH:MM:SS).")
    return texto


def normalizar_sgm(valor_bruto) -> str:
    texto = str(valor_bruto).strip()
    texto = texto.replace(" ", "")

    if texto.isdigit():
        texto = "SGM" + texto.zfill(7)

    if not re.fullmatch(r"SGM\d{7}", texto):
        print(f"⚠️  Aviso: SGM '{texto}' (original: '{valor_bruto}') não está no "
              "formato esperado SGM0999999 (SGM + 7 dígitos) — confira manualmente.")

    return texto


def ler_itens_da_planilha(caminho_arquivo: Path) -> list[dict]:
    import openpyxl

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
        cabecalho_real = ws.cell(row=1, column=indice + 1).value
        if cabecalho_real != cabecalho_esperado:
            print(
                f"⚠️  Aviso: esperava o cabeçalho '{cabecalho_esperado}' na coluna "
                f"{indice + 1}, mas encontrei '{cabecalho_real}'. A planilha pode "
                "ter mudado de estrutura."
            )

    itens = []
    for row in ws.iter_rows(min_row=2):
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
            print(f"⚠️  Aviso: não consegui combinar data+hora de início pro SGM "
                  f"{sgm_atual} (data='{data_atendimento}', hora='{hora_inicio}').")

        termino_combinado = combinar_data_hora(data_atendimento, hora_termino)
        if termino_combinado is None:
            print(f"⚠️  Aviso: não consegui combinar data+hora de término pro SGM "
                  f"{sgm_atual} (data='{data_atendimento}', hora='{hora_termino}').")

        itens.append({
            "linha": celula_sgm.row,
            "sgm": sgm_atual,
            "inicio_real": formatar_data_hora_servicenow(inicio_combinado) if inicio_combinado else "",
            "termino_real": formatar_data_hora_servicenow(termino_combinado) if termino_combinado else "",
            "tecnico": str(tecnico).strip() if tecnico is not None else "",
            "validador": str(validador).strip() if validador is not None else "",
            "matricula": str(matricula).strip() if matricula is not None else "",
        })
    return itens


# ============================================================================
# ETAPA 3 — Pesquisar o SGM no ServiceNow, clicar no MAN e preencher as etapas
# ============================================================================
def obter_alvo_servicenow(page: Page):
    frame = page.frame(name="gsft_main")
    return frame if frame else page


def montar_url_busca_por_sgm(sgm: str) -> str:
    base = SERVICENOW_TASK_LIST_URL.split("?")[0]
    return f"{base}?sysparm_query=parent.numberSTARTSWITH{sgm}"


def buscar_chamado_e_clicar_man(page: Page, sgm: str):
    if "login" in page.url.lower():
        raise RuntimeError(
            "Sessão do ServiceNow expirou ou não está logada. "
            "Rode: python main.py --login"
        )

    page.goto(montar_url_busca_por_sgm(sgm), wait_until="domcontentloaded")
    alvo = obter_alvo_servicenow(page)
    aguardar_pagina(alvo, SELETOR_LISTA_CARREGADA)

    link_man = alvo.get_by_text(re.compile(r"^MAN\d+"))

    if link_man.count() == 0:
        raise RuntimeError(f"SGM {sgm} não encontrado no ServiceNow.")

    link_man.first.click(timeout=5000)
    aguardar_pagina(obter_alvo_servicenow(page), SELETOR_FORM_MAN)

    print(f"Chamado {sgm}: entrou no registro MAN com sucesso.")

    page.reload(wait_until="domcontentloaded")
    alvo = obter_alvo_servicenow(page)
    aguardar_pagina(alvo, SELETOR_FORM_MAN)
    fechar_popup_chamados_pausados(page)
    return obter_alvo_servicenow(page)


def obter_ou_criar_coluna_status(ws) -> int:
    for cell in ws[1]:
        if cell.value == "Status":
            return cell.column
    nova_coluna = ws.max_column + 1
    ws.cell(row=1, column=nova_coluna, value="Status")
    return nova_coluna


def fechar_popup_chamados_pausados(page: Page) -> bool:
    """
    Procura o botão OK do popup "chamados em andamento" na página E em todos
    os iframes (inclusive o gsft_main, onde fica o formulário). Antes só a
    página de fora era verificada — se o popup abria dentro do iframe, ele
    ficava aberto bloqueando o Start Work.
    """
    fechou = False
    for frame in page.frames:
        try:
            botao_ok = frame.locator("#ok_button")
            if botao_ok.count() > 0 and botao_ok.first.is_visible():
                botao_ok.first.click(timeout=2000)
                print("  (popup 'chamados em andamento' fechado automaticamente)")
                fechou = True
        except Exception:
            pass
    return fechou


def ler_mensagens_erro_servicenow(alvo) -> str:
    """Lê mensagens de erro/aviso exibidas no topo do formulário ou nos campos."""
    try:
        textos = alvo.locator(
            ".outputmsg_error, .outputmsg_warning, .fieldmsg.notification-error, "
            "div.notification-error, .alert-danger"
        ).all_inner_texts()
        textos = [t.strip() for t in textos if t.strip()]
        return " | ".join(textos)
    except Exception:
        return ""


def clicar_se_existir(alvo, localizador, descricao: str, page: Page = None) -> bool:
    if localizador.count() == 0:
        print(f"  (pulando \"{descricao}\": não encontrado na tela — provavelmente já foi feito)")
        return False

    # Fecha o popup "chamados em andamento" (se estiver aberto) ANTES do
    # clique — ele pode interceptar o clique no botão certo.
    if page is not None:
        fechar_popup_chamados_pausados(page)

    try:
        localizador.first.click(timeout=5000)
    except Exception as erro:
        # Último recurso: clique forçado (ignora overlays por cima).
        print(f"  ⚠️  Clique normal em \"{descricao}\" falhou ({erro}); tentando clique forçado...")
        if page is not None:
            fechar_popup_chamados_pausados(page)
        localizador.first.click(timeout=5000, force=True)

    alvo.wait_for_timeout(500)  # dá tempo do ServiceNow iniciar o recarregamento
    aguardar_pagina(alvo, SELETOR_FORM_MAN)

    # Fecha o popup também DEPOIS do clique (ele pode surgir como reação ao
    # Start Work, por exemplo).
    if page is not None:
        fechar_popup_chamados_pausados(page)

    return True


def clicar_esperando(alvo, localizador, descricao: str, segundos: int = 8, page: Page = None) -> bool:
    try:
        localizador.first.wait_for(state="visible", timeout=segundos * 1000)
    except Exception:
        print(f"  (pulando \"{descricao}\": não apareceu após {segundos}s — provavelmente não se aplica)")
        return False

    if page is not None:
        fechar_popup_chamados_pausados(page)

    localizador.first.click(timeout=5000)
    alvo.wait_for_timeout(500)
    aguardar_pagina(alvo, SELETOR_FORM_MAN)

    if page is not None:
        fechar_popup_chamados_pausados(page)

    return True


def preencher_campo_oculto_servicenow(alvo, seletor: str, valor: str) -> None:
    """
    Preenche um <input type="hidden"> do ServiceNow (ex: "wm_task.work_start"/
    "wm_task.work_end") setando o .value via JS e disparando "change", o que
    aciona o onchange do próprio ServiceNow.
    """
    alvo.locator(seletor).evaluate(
        """(el, valor) => {
            el.value = valor;
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }""",
        valor,
    )
    alvo.wait_for_timeout(300)  # pequena folga pra o onChange do ServiceNow processar


def validar_tecnico_assignado(alvo, tecnico_planilha: str) -> None:
    if not tecnico_planilha:
        return

    primeiro_nome = tecnico_planilha.strip().split()[0]
    valor_servicenow = alvo.input_value("input[id='sys_display.wm_task.assigned_to']")

    if primeiro_nome.lower() not in valor_servicenow.lower():
        raise RuntimeError(
            f"Técnico não confere: esperava encontrar '{primeiro_nome}' no campo "
            f"'Atribuído a', mas o valor lá é '{valor_servicenow.strip()}'."
        )


def aguardar_transicao_de_estado(alvo, localizador_botao_clicado, descricao: str,
                                 segundos: int = 10, page: Page = None,
                                 obrigatorio: bool = False) -> bool:
    """
    Depois de clicar num botão que muda o ESTADO do chamado (Start Work,
    Retomar Trabalho), espera o próprio botão sumir da tela — sinal de que
    o ServiceNow confirmou a transição.

    Se não sumir: fecha popups, lê mensagens de erro da tela e tenta clicar
    de novo UMA vez. Se continuar sem efeito e `obrigatorio=True`, aborta o
    chamado (com o motivo) em vez de seguir e falhar mais adiante.
    """
    try:
        localizador_botao_clicado.first.wait_for(state="hidden", timeout=segundos * 1000)
        return True
    except Exception:
        pass

    # 1ª tentativa não funcionou — investiga e tenta de novo
    popup_fechado = fechar_popup_chamados_pausados(page) if page is not None else False
    msg = ler_mensagens_erro_servicenow(alvo)
    print(f"  ⚠️  \"{descricao}\" não surtiu efeito em {segundos}s"
          + (" (havia um popup aberto, já fechado)" if popup_fechado else "")
          + (f". Mensagem do ServiceNow: {msg}" if msg else "")
          + ". Tentando clicar de novo...")

    try:
        if localizador_botao_clicado.count() > 0 and localizador_botao_clicado.first.is_visible():
            localizador_botao_clicado.first.click(timeout=5000)
            alvo.wait_for_timeout(500)
            if page is not None:
                fechar_popup_chamados_pausados(page)
        localizador_botao_clicado.first.wait_for(state="hidden", timeout=segundos * 1000)
        print(f"  ✅ \"{descricao}\" confirmado na 2ª tentativa.")
        return True
    except Exception:
        pass

    msg = ler_mensagens_erro_servicenow(alvo)
    texto = (f"\"{descricao}\" não mudou o estado do chamado após 2 tentativas"
             + (f" — ServiceNow diz: {msg}" if msg else ""))
    if obrigatorio:
        raise RuntimeError(texto)
    print(f"  ⚠️  {texto}. Seguindo mesmo assim.")
    return False


def preencher_etapas_man(page: Page, alvo, inicio_real: str, termino_real: str,
                          validador: str, matricula: str, tecnico: str) -> None:
    """
    Sequência dentro do chamado MAN:
    Aceitar -> Start Work -> Retomar Trabalho -> Não aplicar material ->
    aba "Real" -> preencher Início/Término real -> Encerrar Concluído.
    Etapas iniciais são opcionais (pula as que já foram feitas).
    """
    validar_tecnico_assignado(alvo, tecnico)

    # 🔧 AJUSTE AQUI (1/5 a 4/5) — Aceitar / Start Work / Retomar Trabalho / Não aplicar material
    loc_aceitar = alvo.get_by_text("Aceitar", exact=True)
    loc_start_work = alvo.locator("//*[@id='start_work']")
    loc_retomar = alvo.locator("//*[@id='resume_work']")
    loc_nao_aplicar = alvo.get_by_text(re.compile(r"^Não Aplicar Material$", re.IGNORECASE))

    if (loc_aceitar.count() == 0 and loc_start_work.count() == 0
            and loc_retomar.count() == 0 and loc_nao_aplicar.count() == 0):
        raise RuntimeError(
            "Nenhuma das etapas iniciais (Aceitar / Start Work / Retomar "
            "Trabalho / Não aplicar material) foi encontrada — pulando esse "
            "chamado."
        )

    clicar_se_existir(alvo, loc_aceitar, "Aceitar", page=page)

    clicou_start_work = clicar_se_existir(alvo, loc_start_work, "Start Work", page=page)
    if clicou_start_work:
        aguardar_transicao_de_estado(alvo, loc_start_work, "Start Work",
                                     page=page, obrigatorio=True)

    clicou_retomar = clicar_se_existir(alvo, loc_retomar, "Retomar Trabalho", page=page)
    if clicou_retomar:
        aguardar_transicao_de_estado(alvo, loc_retomar, "Retomar Trabalho",
                                     page=page, obrigatorio=True)

    clicar_esperando(alvo, loc_nao_aplicar, "Não aplicar material", segundos=8, page=page)

    # --- Aba "Real" ---
    fechar_popup_chamados_pausados(page)
    alvo.locator("span.tab_caption_text:text-is('Real')").first.click()
    # O campo é <input type="hidden">, por isso estado "attached" (não "visible")
    aguardar_pagina(alvo, SELETOR_CAMPO_INICIO_REAL, estado="attached")

    preencher_campo_oculto_servicenow(alvo, "input[id='wm_task.work_start']", inicio_real)
    preencher_campo_oculto_servicenow(alvo, "input[id='wm_task.work_end']", termino_real)

    # --- Aba "Informações de Fechamento" ---
    fechar_popup_chamados_pausados(page)
    alvo.locator("span.tab_caption_text:text-is('Informações de Fechamento')").first.click()
    aguardar_pagina(alvo, SELETOR_CODIGO_FECHAMENTO)

    alvo.select_option("select[id='wm_task.u_fsm_close_code']", "Outros")

    alvo.fill("input[id='wm_task.u_validated_by']", validador)
    alvo.fill("input[id='wm_task.u_validated_id_by']", matricula)

    fechar_popup_chamados_pausados(page)

    loc_close_complete = alvo.locator("#close_complete")
    try:
        loc_close_complete.wait_for(state="visible", timeout=15000)
    except Exception:
        raise RuntimeError(
            "Botão 'Encerrado Concluído' (#close_complete) não apareceu depois "
            "de preencher os campos. Provavelmente uma etapa anterior (Start "
            "Work / Retomar Trabalho / preenchimento de Início-Término real) "
            "não foi confirmada de verdade pelo ServiceNow — confira a "
            "screenshot do erro pra ver em que estado o chamado ficou."
        )
    loc_close_complete.click()
    # Espera o botão sumir = ServiceNow processou o encerramento
    aguardar_transicao_de_estado(alvo, loc_close_complete, "Encerrado Concluído", segundos=15)
    aguardar_pagina(alvo)

    print(f"  Início real: {inicio_real} | Término real: {termino_real} preenchidos e encerrado.")


# ============================================================================
# LOGIN MANUAL (modo --login)
# ============================================================================
def fazer_login_manual(context: BrowserContext) -> None:
    page = context.new_page()
    page.goto(EXCEL_ONLINE_URL)
    print("\n>>> Faça login manualmente na Microsoft nesta janela (Excel Online).")

    page2 = context.new_page()
    page2.goto(SERVICENOW_URL)
    print(">>> Faça login manualmente no ServiceNow nesta outra aba.")

    input("\nDepois de logar nos dois, volte aqui e aperte Enter para salvar a sessão...")
    print("Sessão salva em:", USER_DATA_DIR)


# ============================================================================
# EXECUÇÃO PRINCIPAL
# ============================================================================
def main() -> None:
    modo_login = "--login" in sys.argv

    if not SERVICENOW_URL:
        sys.exit("SERVICENOW_URL não configurada. Copie o .env.example para .env e preencha.")
    if not EXCEL_LOCAL_PATH and not EXCEL_ONLINE_URL:
        sys.exit("Configure EXCEL_LOCAL_PATH ou EXCEL_ONLINE_URL no .env.")

    with sync_playwright() as p:
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

        page.on("dialog", lambda dialog: dialog.accept())
        page.on("load", lambda *_args: fechar_popup_chamados_pausados(page))

        if EXCEL_LOCAL_PATH and Path(EXCEL_LOCAL_PATH).exists():
            print(f"Lendo planilha local em: {EXCEL_LOCAL_PATH}")
            caminho_planilha = Path(EXCEL_LOCAL_PATH)
        else:
            if EXCEL_LOCAL_PATH:
                print(
                    f"⚠️  EXCEL_LOCAL_PATH foi configurado ('{EXCEL_LOCAL_PATH}') mas "
                    "o arquivo não foi encontrado nesse caminho. Caindo para o "
                    "método via Excel Online."
                )
            print("Baixando a planilha do Excel Online...")
            caminho_planilha = baixar_planilha_excel_online(page)

        itens = ler_itens_da_planilha(caminho_planilha)
        print(f"{len(itens)} chamado(s) encontrado(s) na planilha.")

        import openpyxl
        permite_escrever_status = bool(EXCEL_LOCAL_PATH) and caminho_planilha == Path(EXCEL_LOCAL_PATH)

        if permite_escrever_status:
            wb_escrita = openpyxl.load_workbook(caminho_planilha)
            ws_escrita = wb_escrita.active
            coluna_status = obter_ou_criar_coluna_status(ws_escrita)
            wb_escrita.save(caminho_planilha)
            print(f"Status será escrito na planilha local (coluna {coluna_status}).")
        else:
            print(
                "⚠️  Aviso: EXCEL_LOCAL_PATH não está configurado (ou não é o arquivo "
                "em uso agora) — o Status NÃO será escrito de volta na planilha, já "
                "que a cópia baixada do Excel Online é só temporária."
            )

        print("Acessando o ServiceNow...")
        page.goto(SERVICENOW_URL, wait_until="domcontentloaded")
        try:
            # Até 2 min: dá tempo de concluir SSO/MFA se a sessão tiver expirado
            aguardar_pagina(page, SELETOR_SERVICENOW_CARREGADO, timeout=120000)
        except PlaywrightTimeoutError:
            raise RuntimeError(
                "O ServiceNow não carregou em 2 minutos. Verifique se a sessão "
                "expirou (tela de login/SSO) — se sim, rode: python main.py --login"
            )

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
                try:
                    caminho_print = ERROS_DIR / f"erro_{sgm}.png"
                    page.screenshot(path=str(caminho_print))
                    print(f"  Screenshot do erro salva em: {caminho_print}")
                except Exception:
                    pass

            time.sleep(1)

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
        print(f"Lista só dos SGMs com falha salva em: {caminho_falhas}")

        context.close()


if __name__ == "__main__":
    main()