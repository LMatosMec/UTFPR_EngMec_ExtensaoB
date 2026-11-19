import os
import time
from datetime import datetime

from config import MSD
from funcoes import (
    assassinar_excel,
    consolidar_bases,
    filtrar_fila_idempotente,
    formatar_layout_excel,
    limpar_arquivos_antigos,
)
from motor import executar_tarefas_em_lotes_sap
from transacoes_sap import (
    extrair_ck5n,
)

if __name__ == "__main__":
    #
    # ====================================================================================
    # region ----------------------------------- SETUP -----------------------------------

    start = time.perf_counter()

    data_hoje = datetime.today().strftime("%d_%m_%Y")

    CAMINHO = MSD.PATH

    caminho_msd = os.path.join(MSD.PATH, (MSD.NOME_MSD + "_" + data_hoje + ".xlsx"))

    lista_msd = MSD.LISTA_MSD

    todas_tarefas = []

    # endregion

    # =================================================================
    # region ------------------------- CK53N --------------------------

    limpar_arquivos_antigos(diretorio=MSD.PATH)

    for msd_atual in lista_msd:
        todas_tarefas.append(
            {
                "funcao": extrair_ck5n,
                "kwargs": {
                    "msd": msd_atual,
                    "nome_arquivo": f"{msd_atual}.xlsx",
                    "path": CAMINHO,
                },
            }
        )

    tarefas_para_rodas = filtrar_fila_idempotente(todas_tarefas)

    if tarefas_para_rodas:
        resultados_ck53n = executar_tarefas_em_lotes_sap(
            lista_tarefas=tarefas_para_rodas,
            tamanho_lote_desejado=10,
            param_sessao_padrao="session",
        )

        for res in resultados_ck53n:
            if res.get("status") == "erro":
                print(f"🚨 ERRO FATAL NA THREAD: {res.get('erro')}")

    else:
        print("🎉 Nenhuma extração pendente para a CK53N!")

    # --------------------------- TRANSFORM ---------------------------

    print("\n- 📦 Iniciando etapa de Consolidação CK5_N - (TRANSFORM)...")

    # Montar a lista com os caminhos completos de TODOS os materiais da MSD
    # Como a nossa função é blindada, passamos a lista completa. Ela mesma vai checar o que já está no consolidado e ler apenas o que falta.
    lista_caminhos_ck53n = [os.path.join(CAMINHO, f"{msd_atual}.xlsx") for msd_atual in lista_msd]

    for arquivos in lista_msd:
        assassinar_excel(nome_parcial_planilha=arquivos)

    time.sleep(2)

    # Acionar a função de consolidação
    df_ck53n_consolidado = consolidar_bases(
        lista_caminhos=lista_caminhos_ck53n,
        caminho_saida=caminho_msd,
        n_linhas=6,
    )

    if df_ck53n_consolidado is not None:
        print(f"\n📊 Resumo Final: A base consolidada possui {len(df_ck53n_consolidado)} linhas.")

    # endregion

    # =========================================================================
    # region ------------------------------ LOAD ------------------------------

    formatar_layout_excel(caminho_arquivo=caminho_msd)

    end = time.perf_counter()
    print("\n✅ Rotina finalizada com sucesso!")
    print(f"⏱️ Tempo total: {end - start:.2f} segundos")
