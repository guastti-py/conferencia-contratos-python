# ======================================================================================
# AÇÕES DA INTERFACE
# Funções chamadas pelos botões e fluxos gerais do sistema
# ======================================================================================

import os
from tkinter import messagebox, filedialog

from openpyxl import Workbook

from utils.arquivos import (
    validar_data,
    carregar_arquivo,
    padronizar_colunas,
    converter_colunas_numericas,
)
from utils.excel import obter_nome_arquivo_saida
from parceiros.geral import (
    calcular_resumo,
    criar_aba_resumo,
    criar_aba_dados_originais,
    gerar_arquivo_exportacao,
)
from config.parceiros import CODIGOS_GRUPO_D


def selecionar_e_processar_geral(root, entry_data, progress_bar, nome_parceiro, codigo_parceiro, event=None):
    data_digitada = entry_data.get().strip()

    if not validar_data(data_digitada):
        messagebox.showwarning("Atenção", "Por favor, digite uma data válida no formato dd.mm.aaaa.")
        return

    caminho_arquivo = filedialog.askopenfilename(
        title=f"Selecionar - {nome_parceiro}",
        filetypes=[("Excel ou CSV", "*.xlsx *.xls *.csv")]
    )
    if not caminho_arquivo:
        return

    try:
        progress_bar.pack(pady=10)
        progress_bar.set(0.20)
        root.update_idletasks()

        df = carregar_arquivo(caminho_arquivo)
        df = padronizar_colunas(df)
        df = converter_colunas_numericas(df)

        progress_bar.set(0.45)
        root.update_idletasks()

        dados_resumo, df = calcular_resumo(df, codigo_parceiro)
        is_torra = codigo_parceiro in CODIGOS_GRUPO_D

        pasta_saida = os.path.dirname(caminho_arquivo)
        nome_arquivo_resumo = obter_nome_arquivo_saida(
            nome_parceiro,
            codigo_parceiro,
            data_digitada,
            is_torra
        )

        caminho_resumo = os.path.join(pasta_saida, nome_arquivo_resumo)

        wb_resumo = Workbook()
        criar_aba_resumo(wb_resumo, dados_resumo)
        criar_aba_dados_originais(wb_resumo, df)
        wb_resumo.save(caminho_resumo)

        progress_bar.set(0.80)
        root.update_idletasks()

        if not is_torra:
            gerar_arquivo_exportacao(
                df,
                pasta_saida,
                nome_parceiro,
                codigo_parceiro,
                data_digitada
            )

        progress_bar.set(1.0)
        root.update_idletasks()
        progress_bar.pack_forget()

        os.startfile(pasta_saida)
        messagebox.showinfo("Sucesso", "Processamento concluído!")

    except Exception as erro:
        progress_bar.pack_forget()
        messagebox.showerror("Erro", str(erro))