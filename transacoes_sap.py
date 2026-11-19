import os
import sys
import threading
import time

import pandas as pd
import pyperclip

from conexao import fechar_sessoes_por_id, garantir_quantidade_sessoes
from funcoes import assassinar_excel, esperar_download_concluir, selecionar_layout_por_nome


def ck5n(session, msd, nome_arquivo, path, transacao="/nck51n"):

    try:
        # session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = transacao
        session.findById("wnd[0]").sendVKey(0)

        time.sleep(0.2)

        session.findById("wnd[0]/usr/subALL:SAPLCKDI:4611/subKOPF:SAPLCKDI:4620/ctxtCKI64A-VBELN").text = msd
        session.findById("wnd[0]/usr/subALL:SAPLCKDI:4611/subKOPF:SAPLCKDI:4620/txtCKI64A-POSNR").text = "10"

        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.3)

        session.findById("wnd[0]/tbar[1]/btn[20]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[1]").select()

        session.findById(
            "wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME"
        ).text = nome_arquivo
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

    except Exception as e:
        # Pega a linha exata onde o código explodiu
        _, _, tb = sys.exc_info()
        linha = tb.tb_lineno

        print(f"❌ [ERRO LINHA {linha}] MSD '{msd}': {e}")

        raise

    return {"msd": msd, "status": "Concluído com sucesso"}


def extrair_ck5n(session, msd, nome_arquivo, path, transacao="/nck51n"):

    ck5n(session=session, msd=msd, nome_arquivo=nome_arquivo, path=path, transacao=transacao)

    esperar_download_concluir(caminho_arquivo=os.path.join(path, nome_arquivo))
    assassinar_excel(nome_parcial_planilha=nome_arquivo)

    # Lê e retorna o DataFrame prontinho
    return {"msd": msd, "status": "Concluído com sucesso"}


# ---------


def export_se16n(session, diretorio, nome_arquivo):
    """Descrição
    Extrator Nativo: Executa a rotina padrão do SAP para acionar o botão de exportação
    XXL (Planilha Eletrônica), definindo o nome e o caminho de destino no disco do usuário.
    """

    session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById(
        "wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME"
    ).text = nome_arquivo
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]").sendVKey(4)
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = diretorio
    session.findById("wnd[2]").sendVKey(0)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


def se16n_aplicar_filtros(
    session, nome_tabela, dicionario_filtros, nome_layout, col_nome=6, col_input=2, col_botao=4, aplicar_layout=True
):

    grid = "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell"
    id_grid_str = "wnd[0]/usr/subTAB_SUB:SAPLSE16N:0121/tblSAPLSE16NSELFIELDS_TC"

    print(f"\n--- Iniciando automação SE16N para tabela {nome_tabela} ---")

    # Configuração Inicial do Ambiente SAP
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16N"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = nome_tabela
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
    session.findById("wnd[0]").sendVKey(0)

    # Prepara o scroll lendo a capacidade da tela atual
    linhas_visiveis = session.findById(id_grid_str).VisibleRowCount
    limite_maximo = session.findById(id_grid_str).VerticalScrollbar.Maximum

    session.findById(id_grid_str).VerticalScrollbar.Position = 0
    posicao_atual = 0

    campos_pendentes = list(dicionario_filtros.keys())

    while True:
        if not campos_pendentes:
            print("Todos os campos foram encontrados!")
            break

        # Itera sobre as linhas visíveis na janela atual
        for i in range(linhas_visiveis):
            try:
                cell_nome = session.findById(id_grid_str).GetCell(i, col_nome)
                texto_lido = cell_nome.Text

                if texto_lido in campos_pendentes:
                    dados = dicionario_filtros[texto_lido]

                    # Desmonta a estrutura de dados para entender se há condições avançadas
                    valor_final = dados
                    opcao_row = None

                    if isinstance(dados, tuple):
                        valor_final = dados[0]
                        opcao_row = dados[1]

                    # Avalia se a inserção será direta no grid ou se exige a janela múltipla
                    usar_popup = isinstance(valor_final, list) or (opcao_row is not None)

                    if usar_popup:
                        id_botao_push = f"{id_grid_str}/btnPUSH[{col_botao},{i}]"
                        session.findById(id_botao_push).setfocus()
                        session.findById(id_botao_push).press()

                        # Inserção em Massa via Área de Transferência
                        if str(valor_final).strip() != "":
                            lista_valores = valor_final if isinstance(valor_final, list) else [valor_final]
                            texto_clip = "\r\n".join([str(v) for v in lista_valores])
                            pyperclip.copy(texto_clip)
                            session.findById("wnd[1]/tbar[0]/btn[24]").press()

                        # Aplicação de regras matemáticas lógicas (ex: diferente, maior, menor)
                        if opcao_row is not None and str(opcao_row).strip() != "":
                            print(f"   -> Aplicando Opção Index {opcao_row}...")
                            session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/btnOPTION[0,0]").press()
                            session.findById("wnd[2]/usr/cntlGRID/shellcont/shell").setCurrentCell(opcao_row, "TEXT")
                            session.findById("wnd[2]/usr/cntlGRID/shellcont/shell").selectedRows = str(opcao_row)
                            session.findById("wnd[2]").sendVKey(0)

                        # Confirma e retorna para a tela principal
                        session.findById("wnd[1]/tbar[0]/btn[8]").press()

                    else:
                        # Inserção de valor único diretamente na célula
                        cell_input = session.findById(id_grid_str).GetCell(i, col_input)
                        cell_input.SetFocus()
                        cell_input.Text = str(valor_final)

                    campos_pendentes.remove(texto_lido)
                    if not campos_pendentes:
                        break

            except Exception as e:
                # Silencia o erro para manter o loop, mas registra no console caso precise de debug futuro
                print(f"Alerta silencioso ao processar linha visual: {e}")
                continue

        # Avalia se a missão foi cumprida antes de descer a tela
        if not campos_pendentes:
            print("Todos os filtros preenchidos. Executando busca...")
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            if aplicar_layout:
                session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
                session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("&LOAD")
                selecionar_layout_por_nome(session, nome_layout, grid, "VARIANT")
            return True

        # Gerenciamento dinâmico da rolagem para a próxima página de filtros
        if posicao_atual >= limite_maximo:
            print(f"Fim da tabela alcançado. Faltou preencher: {campos_pendentes}")
            break

        proxima_posicao_ideal = posicao_atual + linhas_visiveis
        nova_posicao_real = min(proxima_posicao_ideal, limite_maximo)
        diferenca = nova_posicao_real - posicao_atual

        if diferenca == 0:
            break

        session.findById("wnd[0]").maximize()
        session.findById(id_grid_str).VerticalScrollbar.Position = nova_posicao_real
        posicao_atual = nova_posicao_real

    return False


def extrair_se16n_autonoma(config, diretorio, aplica_layout=True):
    """Descrição
    Envelope Moderno Singular: Gerencia todo o ciclo do dado. Verifica idempotência no disco (lidando com arquivos trancados)
    abre o SAP sob demanda, preenche e exporta a SE16N, abate o Excel fantasma e entrega o DataFrame pronto, fechando a conexão ao finalizar.
    """

    tabela = config["tabela"]
    nome_arquivo = config["nome_arquivo"]
    caminho_completo = os.path.join(diretorio, nome_arquivo)

    # Verifica a idempotência (Se já tem, foge e nem abre o SAP)
    if os.path.exists(caminho_completo):
        try:
            print(f"\n ⏩ Arquivo {nome_arquivo} já existe na base, lendo para memoria...")
            # Tenta ler o arquivo pacificamente
            return pd.read_excel(caminho_completo)

        except PermissionError:
            # O arquivo existe, mas alguém (ou o próprio robô antes) deixou aberto!
            print(f"⚠️ O arquivo {nome_arquivo} está aberto e trancado. Fechando...")
            assassinar_excel(nome_parcial_planilha=nome_arquivo)
            time.sleep(2)  # Respiro para o Windows soltar o arquivo de vez

            # Agora lê com o caminho livre
            return pd.read_excel(caminho_completo)

    # Garante a Sessão
    ids_matar, sessoes = garantir_quantidade_sessoes(quantidade_desejada=1, session_janela=True)
    session = sessoes[0]

    semaforo_seguro = threading.Lock()

    try:
        with semaforo_seguro:
            se16n_aplicar_filtros(
                session=session,
                nome_tabela=tabela,
                dicionario_filtros=config.get("filtros", []),
                nome_layout=config.get("layout", ""),
                aplicar_layout=aplica_layout,
            )

            export_se16n(session, diretorio, nome_arquivo)
            time.sleep(3)

        # Assassina (Garantindo que o download terminou antes)
        esperar_download_concluir(caminho_completo)
        assassinar_excel(nome_parcial_planilha=nome_arquivo)

        # Lê e retorna
        return pd.read_excel(caminho_completo)

    finally:
        # Limpa a sujeira do SAP
        if ids_matar:
            fechar_sessoes_por_id(ids_matar)
