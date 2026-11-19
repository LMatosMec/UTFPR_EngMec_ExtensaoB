import concurrent.futures
import time

import pythoncom

from conexao import (
    fechar_sessoes_por_id,
    garantir_quantidade_sessoes,
    get_sap_sessions,
    get_session_by_index,
)


def _worker_sap(funcao_alvo, index_sessao, kwargs_funcao, nome_param_sessao="ss"):
    """
    Worker interno que inicializa a thread, pega a sessão correta e roda a função.
    """
    time.sleep(0.1)
    pythoncom.CoInitialize()

    # Prepara um identificador caso dê erro (tenta pegar o nome do arquivo no payload)
    pyl = kwargs_funcao.get("pyl")
    id_tarefa = (
        pyl.nome_arquivo
        if pyl and hasattr(pyl, "nome_arquivo")
        else "Tarefa Desconhecida"
    )

    sessao_thread = None

    try:
        sessao_thread = get_session_by_index(index_sessao)
        kwargs_funcao[nome_param_sessao] = sessao_thread

        # Executa a função core
        resultado = funcao_alvo(**kwargs_funcao)

        # APRIMORAMENTO 1: Evita dicionários aninhados ("Inception")
        if isinstance(resultado, dict):
            return resultado  # Se o operário já montou o dicionário, repassa direto
        else:
            return {"status": "ok", "retorno": resultado, "arquivo": id_tarefa}

    except Exception as e:
        # APRIMORAMENTO 2: Rastreabilidade. Agora sabemos QUEM falhou.
        return {
            "status": "erro",
            "erro": f"Falha crítica no worker: {str(e)}",
            "arquivo": id_tarefa,
        }

    finally:
        # APRIMORAMENTO 3: Limpeza da janela para o próximo lote
        if sessao_thread:
            try:
                # Tenta forçar a saída de qualquer transação presa (comando /n)
                sessao_thread.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                sessao_thread.findById("wnd[0]").sendVKey(0)
            except Exception:
                pass  # Se a janela já estiver limpa ou não aceitar o comando, ignora em silêncio

        pythoncom.CoUninitialize()


def executar_tarefas_em_lotes_sap(
    lista_tarefas,
    tamanho_lote_desejado=4,
    limite_maximo_sap=12,
    param_sessao_padrao="ss",
):
    """
    Executa tarefas no SAP em paralelo, divididas em lotes.
    Abre as janelas necessárias APENAS UMA VEZ e as reaproveita em todos os lotes,
    fechando-as apenas no final da fila (Pool de Conexões).
    """
    if not lista_tarefas:
        return []

    # Cálculo de segurança
    sessoes_abertas_agora = len(get_sap_sessions())
    janelas_livres = limite_maximo_sap - sessoes_abertas_agora

    if janelas_livres <= 0:
        raise RuntimeError(
            f"O SAP já atingiu o limite de {limite_maximo_sap} janelas. Feche algumas manualmente para rodar o robô."
        )

    # Otimização: Se tem 2 tarefas, não precisa abrir 4 janelas. Pega o menor valor.
    lote_seguro = min(tamanho_lote_desejado, janelas_livres, len(lista_tarefas))

    print(
        f"\n📊 SAP Status: {sessoes_abertas_agora} janelas já abertas. Limite: {limite_maximo_sap}."
    )
    print(
        f"🔒 Motor ajustado para rodar na esteira com {lote_seguro} janelas simultâneas."
    )

    resultados = []
    ids_matar = []

    try:
        # SETUP DA FÁBRICA: Abre as janelas UMA ÚNICA VEZ
        print(f"⏳ Ligando as máquinas... Abrindo {lote_seguro} janelas no SAP.")
        ids_matar, _ = garantir_quantidade_sessoes(
            quantidade_desejada=lote_seguro, session_janela=True
        )

        # Mapeia o índice fixo dessas novas janelas
        total_sessoes_abertas = len(get_sap_sessions())
        indice_inicial = total_sessoes_abertas - lote_seguro

        # ESTEIRA DE PRODUÇÃO: Roda os lotes reaproveitando as janelas
        for i in range(0, len(lista_tarefas), lote_seguro):
            lote_atual = lista_tarefas[i : i + lote_seguro]
            numero_lote = (i // lote_seguro) + 1
            total_lotes = (len(lista_tarefas) + lote_seguro - 1) // lote_seguro

            print(
                f"\n🚀 Iniciando Lote {numero_lote}/{total_lotes} com {len(lote_atual)} tarefas..."
            )

            with concurrent.futures.ThreadPoolExecutor(
                max_workers=len(lote_atual)
            ) as executor:
                futures = []

                for j, tarefa in enumerate(lote_atual):
                    # Reaproveita sempre as mesmas janelas (Ex: índices 1, 2, 3 e 4)
                    index_janela = indice_inicial + j

                    funcao_alvo = tarefa["funcao"]
                    kwargs_funcao = tarefa["kwargs"]
                    nome_param = tarefa.get("nome_param_sessao") or param_sessao_padrao

                    future = executor.submit(
                        _worker_sap,
                        funcao_alvo,
                        index_janela,
                        kwargs_funcao,
                        nome_param,
                    )
                    futures.append(future)

                # Aguarda o lote atual terminar para liberar as janelas pro próximo
                for future in concurrent.futures.as_completed(futures):
                    resultados.append(future.result())

    finally:
        # FAXINA FINAL: Só fecha as portas quando todo o trabalho acabar

        if ids_matar:
            print(
                f"\n🧹 Fila concluída! Encerrando as {len(ids_matar)} janelas auxiliares..."
            )
            fechar_sessoes_por_id(ids_matar)
            time.sleep(1)

    return resultados
