# ======================================================================================
# PARCEIRO: PARCEIRO F PJ
# Arquivo próprio do parceiro. No momento, usa o fluxo geral de processamento.
# ======================================================================================

from interface.acoes import selecionar_e_processar_geral


def selecionar_e_processar_parceiro_f(event=None):
    """
    Executa o processamento da PARCEIRO F PJ usando o fluxo geral.
    """
    return selecionar_e_processar_geral("PARCEIRO F PJ", "247", event)