import os
import time
from datetime import date, datetime

import openpyxl
import pandas as pd
import win32con
import win32gui
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


def montar_config_sap(tabela, filtros, layout=""):
    """
    Fábrica que gera o dicionário padrão de extração para o motor SAP - se16n.
    """
    return {
        "tabela": tabela,
        "filtros": filtros,
        "nome_arquivo": f"Base_{tabela}.xlsx",  # Gera o nome automaticamente!
        "layout": layout,
    }


def limpar_arquivos_antigos(diretorio):
    """
    Varre a pasta e deleta apenas os arquivos que NÃO foram gerados hoje.
    Garante que arquivos antigos sumam, mas preserva a idempotência se o robô
    cair no meio do dia e precisar ser reiniciado.
    """
    print(f"\n🧹 Iniciando faxina na pasta: {diretorio}")
    hoje = date.today()
    arquivos_apagados = 0

    # Verifica se a pasta existe antes de tentar limpar
    if not os.path.exists(diretorio):
        os.makedirs(diretorio)
        return

    for nome_arquivo in os.listdir(diretorio):
        caminho = os.path.join(diretorio, nome_arquivo)

        # Ignora pastas, olha só para arquivos
        if os.path.isfile(caminho):
            # Pega a data da última modificação do arquivo
            timestamp = os.path.getmtime(caminho)
            data_arquivo = date.fromtimestamp(timestamp)

            # Se o arquivo for mais velho que hoje, FOGO NELE!
            if data_arquivo < hoje:
                try:
                    os.remove(caminho)
                    arquivos_apagados += 1
                except Exception as e:
                    print(f"⚠️ Não consegui apagar {nome_arquivo}: {e}")

    if arquivos_apagados > 0:
        print(f"🗑️ Faxina concluída: {arquivos_apagados} arquivos de dias anteriores foram fumados.")
    else:
        print("✨ A pasta já está limpa (sem arquivos velhos).")


def esperar_download_concluir(caminho_arquivo, tempo_maximo=180):
    """
    Monitora a pasta até o SAP terminar de gerar o arquivo físico.
    """
    inicio = time.time()
    while not os.path.exists(caminho_arquivo):
        if time.time() - inicio > tempo_maximo:
            raise TimeoutError(f"O arquivo demorou mais de {tempo_maximo}s para baixar.")
        time.sleep(1)

    time.sleep(3)  # Respiro para o Excel aparecer na tela do Windows


def assassinar_excel(nome_parcial_planilha):
    """
    Fecha uma janela específica do Excel usando a API nativa do Windows (win32gui).
    Ignora se o arquivo está no Modo Protegido ou não registrado na tabela COM.
    """
    # O Excel muitas vezes esconde a extensão '.xlsx' no título da janela superior
    nome_limpo = nome_parcial_planilha.replace(".xlsx", "").replace(".xls", "")

    # Variável de controle para sabermos se ele achou ou não
    janela_encontrada = []

    def enumerar_janelas(hwnd, _):
        titulo_janela = win32gui.GetWindowText(hwnd)

        # Procura por uma janela que tenha o nome da planilha E a palavra Excel no título
        if nome_limpo in titulo_janela and "Excel" in titulo_janela:
            print(f"🔪 Janela fantasma localizada e abatida: {titulo_janela}")
            # Envia o comando nativo do Windows de "Clicar no X" da janela
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            janela_encontrada.append(True)

    # Varre todas as janelas abertas no Windows (visíveis ou não)
    win32gui.EnumWindows(enumerar_janelas, None)
    time.sleep(0.2)

    if not janela_encontrada:
        print(f"🔍 Nenhuma janela do Excel aberta correspondente a '{nome_limpo}'.")


def arquivo_foi_baixado_hoje(caminho_completo):
    """
    Retorna True se o arquivo existir E tiver sido modificado na data de hoje.
    """
    if not os.path.exists(caminho_completo):
        return False

    # Pega o timestamp de modificação do arquivo e converte para data
    data_modificacao = datetime.fromtimestamp(os.path.getmtime(caminho_completo)).date()
    data_hoje = datetime.now().date()

    return data_modificacao == data_hoje


def filtrar_fila_idempotente(lista_tarefas, var_nome_arquivo="nome_arquivo", var_diretorio="path"):
    """
    Filtra uma lista de tarefas do SAP verificando a idempotência no disco.

    Regras:
    - Se o arquivo existe e é de hoje: Ignora a tarefa (já está pronto).
    - Se o arquivo é antigo: Deleta (mata o Excel se estiver trancado) e reprocessa.
    - Se não existe: Adiciona à fila normalmente.

    :param lista_tarefas: Lista contendo os dicionários de tarefas.
    :return: Lista filtrada contendo apenas as tarefas que realmente precisam rodar.
    """
    print("\n🔍 Verificando bases locais e filtrando fila de extração...")
    tarefas_efetivas = []

    for tarefa in lista_tarefas:
        kwargs = tarefa.get("kwargs", {})
        nome_arq = kwargs.get(var_nome_arquivo)
        caminho_destino = kwargs.get(var_diretorio)

        # Trava de segurança: Se a tarefa estiver mal formatada, joga pra fila por precaução
        if not nome_arq or not caminho_destino:
            tarefas_efetivas.append(tarefa)
            continue

        caminho_completo = os.path.join(caminho_destino, nome_arq)

        if os.path.exists(caminho_completo):
            if arquivo_foi_baixado_hoje(caminho_completo):
                print(f" ✅ Atualizado, pulando: {nome_arq}")
            else:
                print(f" 🗑️ Base desatualizada. Tentando excluir: {nome_arq}")
                try:
                    os.remove(caminho_completo)
                    tarefas_efetivas.append(tarefa)
                except PermissionError:
                    print(f" ⚠️ Arquivo trancado pelo Windows. Acionando atirador de elite para: {nome_arq}")
                    # assassinar_excel(nome_parcial_planilha=nome_arq)
                    time.sleep(1)  # Respiro para o sistema operacional liberar o arquivo

                    try:
                        os.remove(caminho_completo)
                        tarefas_efetivas.append(tarefa)
                    except Exception as e:
                        print(f" ❌ Erro crítico ao deletar {nome_arq} mesmo após abater Excel: {e}")
                        # Adiciona na fila mesmo assim, pois o SAP pode conseguir sobrescrever
                        tarefas_efetivas.append(tarefa)
        else:
            # Não existe
            tarefas_efetivas.append(tarefa)

    print(f"\n📊 Resumo da Fila: {len(tarefas_efetivas)} de {len(lista_tarefas)} materiais precisam ser extraídos.")
    return tarefas_efetivas


def consolidar_bases(lista_caminhos, caminho_saida, n_linhas=0, colunas_alvo=None):
    """
    Consolida arquivos do SAP com segurança total e idempotência:
    - Retoma de onde parou caso o robô tenha quebrado no meio.
    - Lê apenas os arquivos informados na lista e valida o cabeçalho.
    - Pula metadados iniciais do excel com 'skiprows'.
    """
    if not lista_caminhos:
        print("⚠️ Lista de arquivos vazia. Nada para consolidar.")
        return None

    arquivos_faltantes = lista_caminhos  # Por padrão, assume que precisa de todos
    resultados = []
    cabecalho_referencia = None

    # =================================================================
    # -------- IDEMPOTÊNCIA: Verifica o que já foi consolidado --------

    if os.path.exists(caminho_saida):
        print(f"\n🔍 Arquivo consolidado encontrado em '{os.path.basename(caminho_saida)}'. Verificando conteúdo...")
        try:
            df_existente = pd.read_excel(caminho_saida)

            if "arquivo_origem" in df_existente.columns:
                # Força para texto (str) e arranca possíveis decimais (.0) fantasma do Pandas
                arquivos_ja_consolidados = set(
                    df_existente["arquivo_origem"].astype(str).str.replace(r"\.0$", "", regex=True).unique()
                )

                # Filtra a lista deixando APENAS os arquivos que ainda não estão no consolidado
                arquivos_faltantes = [
                    p
                    for p in lista_caminhos
                    if os.path.splitext(os.path.basename(p))[0] not in arquivos_ja_consolidados
                ]

                if not arquivos_faltantes:
                    print("✅ Idempotência: Todos os arquivos já estão no consolidado! Nenhuma ação necessária.")
                    return df_existente
                else:
                    print(f"🧩 Faltam {len(arquivos_faltantes)} arquivos no consolidado. Retomando de onde parou...")
                    # Coloca a base que já existe como o primeiro item da nossa lista de resultados
                    resultados.append(df_existente)
                    # Define a régua de cabeçalho (ignorando a coluna de origem)
                    cabecalho_referencia = [col for col in df_existente.columns if col != "arquivo_origem"]
            else:
                print("⚠️ O consolidado existente não possui a coluna 'arquivo_origem'. Refazendo do zero.")

        except Exception as e:
            print(f"⚠️ Erro ao ler o consolidado antigo (pode estar corrompido). Refazendo do zero. Erro: {e}")
            # Se deu erro, a lista de faltantes volta a ser a lista completa
            arquivos_faltantes = lista_caminhos

    # ================================================================
    # ---- LEITURA E VALIDAÇÃO DOS ARQUIVOS (Apenas os faltantes) ----

    for caminho_arquivo in arquivos_faltantes:
        nome_arquivo = os.path.basename(caminho_arquivo)

        if not os.path.exists(caminho_arquivo):
            print(f"⚠️ Arquivo ausente no disco, pulando: {nome_arquivo}")
            continue

        try:
            df = pd.read_excel(
                caminho_arquivo,
                skiprows=n_linhas,
                usecols=colunas_alvo,
            )
            df_cols = df.columns.tolist()

            # Define a régua de cabeçalho se ainda não existir
            if cabecalho_referencia is None:
                cabecalho_referencia = df_cols

            # Valida se o layout não mudou
            if df_cols != cabecalho_referencia:
                print(f"❌ Layout diferente detectado no arquivo {nome_arquivo}. Ignorando.")
                continue

            # Rastreabilidade O splitext separa o nome da extensão e o [0] pega só o nome)
            nome_sem_extensao = os.path.splitext(nome_arquivo)[0]
            df["arquivo_origem"] = nome_sem_extensao
            resultados.append(df)

        except Exception as e:
            print(f"❌ Falha ao ler o arquivo {nome_arquivo}: {e}")
            continue

    # ================================================================
    # ------------------------ MONTAGEM FINAL ------------------------

    if len(resultados) > (1 if "df_existente" in locals() and df_existente is not None else 0):
        print("\n🔄 Juntando peças e salvando consolidação...")

        df_final = pd.concat(resultados, ignore_index=True, sort=False)

        try:
            df_final.to_excel(caminho_saida, index=False)
            print(f"✅ Arquivo mestre salvo com sucesso em: {caminho_saida}")
            return df_final
        except PermissionError:
            print(
                f"⚠️ Erro crítico: O arquivo consolidado ({os.path.basename(caminho_saida)}) está aberto no Excel. Feche-o e rode novamente."
            )
            return None
    else:
        # Se ele chegou aqui, é porque a base existente já era suficiente ou nenhum arquivo novo foi válido
        if "df_existente" in locals() and df_existente is not None:
            return df_existente
        print("⚠️ Nenhum dado válido foi processado.")
        return None


def cols_to_string(dataframe, colunas=None):
    """
    Converte colunas para o tipo string de forma segura.
    Se 'colunas' não for informado, aplica em todo o DataFrame.
    """

    df = dataframe.copy()

    # Define o escopo de atuação (Todas as colunas ou apenas as selecionadas)
    alvo = colunas if colunas is not None else df.columns

    for col in alvo:
        if col in df.columns:
            df[col] = df[col].astype("object").fillna("").astype(str).str.strip()
            df[col] = df[col].str.replace(r"\.0$", "", regex=True)

    return df


def formatar_layout_excel(caminho_arquivo: str, nome_aba: str = "Sheet1"):
    """
    Aplica formatação visual na aba: Negrito no cabeçalho, Filtros,
    Autoajuste de largura, Centralização e Remoção das linhas de grade.
    """
    print(f"🎨 Aplicando formatação visual na aba '{nome_aba}'...")

    # Abre o arquivo existente e seleciona a aba
    wb = openpyxl.load_workbook(caminho_arquivo)

    if nome_aba not in wb.sheetnames:
        raise ValueError(f"🚨 Aba '{nome_aba}' não encontrada no arquivo.")

    ws = wb[nome_aba]

    # Cabeçalho em Negrito
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Adicionar Filtro (Pega as dimensões exatas da tabela)
    ws.auto_filter.ref = ws.dimensions

    # Autoajuste de Colunas (ALT + C + O + T) e Centralização (ALT + C + A + C)
    for col in ws.columns:
        tamanho_maximo = 0
        letra_coluna = get_column_letter(col[0].column)

        for cell in col:
            # Aplica a centralização em todas as células
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # Descobre o tamanho do texto para o autoajuste
            try:
                if len(str(cell.value)) > tamanho_maximo:
                    tamanho_maximo = len(str(cell.value))
            except:
                pass

        # Define a largura da coluna (+3 de margem para o ícone da setinha do filtro não espremer o texto)
        ws.column_dimensions[letra_coluna].width = tamanho_maximo + 3

    # Tirar as Linhas de Grade
    ws.sheet_view.showGridLines = False

    # Salva e fecha
    wb.save(caminho_arquivo)
    wb.close()
    print("✅ Formatação concluída!")


def control_exists(session, id_str):
    session.findById(id_str)
    return True


def selecionar_layout_por_nome(session, layout_nome, layout_id, col_name="VARIANT"):
    """Seleciona um layout no SAP fazendo o scroll dinâmico das linhas."""

    if not control_exists(session, layout_id):
        print("Janela de seleção de Layout não apareceu. Pulando etapa!")
        return False

    shell = session.findById(layout_id)

    total_linhas = shell.RowCount
    linhas_visiveis = shell.VisibleRowCount

    shell.firstVisibleRow = 0

    for i in range(total_linhas):
        if i % linhas_visiveis == 0:
            shell.firstVisibleRow = i
        valor_celula = shell.GetCellValue(i, col_name)

        if valor_celula.strip().upper() == layout_nome.strip().upper():
            shell.currentCellRow = i
            shell.selectedRows = str(i)

            try:
                shell.clickCurrentCell()

            except Exception:
                shell.doubleClickCurrentCell()

            print(f"Layout '{layout_nome}' selecionado!")
            return True

    print(f"Layout '{layout_nome}' não encontrado.")
    return False
