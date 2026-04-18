# ======================================================================================
# PARCEIRO: PARCEIRO G TRAK
# Arquivo próprio do parceiro. No momento, usa o fluxo geral de processamento.
# ======================================================================================

from interface.acoes import selecionar_e_processar_geral


def selecionar_e_processar_parceiro_g(event=None):
    """
    Executa o processamento do PARCEIRO G TRAK usando o fluxo geral.
    """
    return selecionar_e_processar_geral("PARCEIRO G TRAK", "286", event)