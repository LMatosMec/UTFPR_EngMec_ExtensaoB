import time

import win32com.client


def get_sap_sessions():
    """
    Retorna uma lista de todas as sessões SAP GUI abertas.
    """
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    sessions = []
    for c in range(application.Children.Count):
        connection = application.Children(c)
        for s in range(connection.Children.Count):
            session = connection.Children(s)
            sessions.append(session)
    return sessions


def garantir_quantidade_sessoes(quantidade_desejada: int, session_janela=False):
    """
    Cria novas sessões de forma otimizada.
    Impede o congelamento verificando se a sessão mestre está livre antes de usá-la.
    """
    sessions = get_sap_sessions()
    janelas_para_abrir = quantidade_desejada
    ids_antes = [s.Id for s in sessions]

    ids_criados_agora = []
    lista_sessoes = []

    if janelas_para_abrir > 0:
        print(f"⚡ Solicitando {janelas_para_abrir} nova(s) janela(s)...")

        # ===================================================================
        # --- O PULO DO GATO: Achar uma sessão LIVRE para evitar Deadlock ---

        sessao_mestre = None
        for s in sessions:
            if not s.Busy:  # Verifica se a sessão está livre de processamento
                sessao_mestre = s
                break

        if sessao_mestre is None:
            raise RuntimeError(
                "🚨 Todas as sessões do SAP estão ocupadas ou travadas. "
                "O código foi interrompido para evitar o congelamento do terminal."
            )

        # PEDE TODAS DE UMA VEZ
        for _ in range(janelas_para_abrir):
            try:
                sessao_mestre.createSession()
                time.sleep(0.5)
            except Exception as e:
                print(f"⚠️ SAP recusou comando de janela: {e}")

        # O RADAR GERAL
        meta_total = len(ids_antes) + janelas_para_abrir
        tempo_inicial = time.time()
        timeout_global = 30

        sessoes_atuais = get_sap_sessions()

        while len(sessoes_atuais) < meta_total:
            if (time.time() - tempo_inicial) > timeout_global:
                raise TimeoutError(
                    "🚨 ERRO CRÍTICO: SAP falhou ao abrir as janelas a tempo. "
                    "Isso geralmente acontece após exportações pesadas da SE16N."
                )
            time.sleep(0.5)
            sessoes_atuais = get_sap_sessions()

        print("✅ Todas as janelas abriram com sucesso!")

        # MAPEIA OS IDs
        for s in sessoes_atuais:
            if s.Id not in ids_antes:
                ids_criados_agora.append(s.Id)
                if session_janela:
                    lista_sessoes.append(s)

    time.sleep(1)
    return ids_criados_agora, lista_sessoes


def get_session_by_index(index):
    """Retorna uma sessão SAP GUI específica pelo índice da lista."""
    sessions = get_sap_sessions()
    if not sessions:
        raise Exception("Nenhuma sessão SAP encontrada!")
    try:
        return sessions[index]
    except IndexError:
        raise Exception(
            f"Erro: A sessão no índice {index} não existe. Atualmente existem apenas {len(sessions)} janelas abertas."
        )


def fechar_sessoes_por_id(lista_ids_para_fechar):
    """
    Mata cirurgicamente apenas as sessões cujos IDs estão na lista.
    """
    if not lista_ids_para_fechar:
        return

    print(
        f"\n🧹 Limpando as {len(lista_ids_para_fechar)} janelas geradas pela automação..."
    )
    sessoes_atuais = get_sap_sessions()

    for sessao in sessoes_atuais:
        if sessao.Id in lista_ids_para_fechar:
            try:
                sessao.parent.CloseSession(sessao.Id)
            except Exception as e:
                print(f" ↳ Erro ao fechar a janela {sessao.Id}: {e}")
