import sys
import os
import customtkinter as ctk
from tkinter import messagebox, filedialog
from datetime import datetime


def resource_path(relative_path):
    """Retorna o caminho correto tanto no ambiente Python quanto no executável PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)


from config.parceiros import TABELA_PARCEIROS, CODIGOS_ATIVOS, CODIGOS_GRUPO_D, CODIGOS_GRUPO_B
from interface.acoes import selecionar_e_processar_geral
from parceiros.parceiro_a import selecionar_e_processar_parceiro_a, fazer_conferencia_parceiro_a
from parceiros.parceiro_b import (
    selecionar_e_processar_parceiro_b,
    fazer_conferencia_parceiro_b,
    fazer_conferencia_parceiro_b_endosso
)
from parceiros.parceiro_e import fazer_conferencia
from parceiros.parceiro_d import fazer_conferencia_parceiro_d
from parceiros.parceiro_c import processar_parceiro_c, conferencia_parceiro_c, conferencia_pdf_parceiro_c


def limpar_interface_resultado():
    progress_bar.pack_forget()
    btn_conf.pack_forget()
    btn_conf_parceiro_a.pack_forget()
    btn_conf_parceiro_d.pack_forget()
    btn_conf_parceiro_b.pack_forget()
    btn_conf_parceiro_b_endosso.pack_forget()
    btn_conf_parceiro_c_pdf.pack_forget()
    btn_anexar.pack_forget()
    label_instrucao_data.pack_forget()
    entry_data.pack_forget()
    label_aviso.configure(text="")
    label_resultado.configure(text="")


def formatar_data(event):
    if event.keysym in ("BackSpace", "Left", "Right", "Delete"):
        return

    conteudo = entry_data.get()
    apenas_numeros = "".join(filter(str.isdigit, conteudo))
    novo_texto = ""

    for i, char in enumerate(apenas_numeros):
        if i == 2 or i == 4:
            novo_texto += "."
        novo_texto += char

    novo_texto = novo_texto[:10]

    if conteudo != novo_texto:
        entry_data.delete(0, "end")
        entry_data.insert(0, novo_texto)


def buscar_parceiro(event=None):
    codigo = entry_codigo.get().strip()

    if codigo in ["127", "128"]:
        codigo = "127/128"
    if codigo in ["125", "126"]:
        codigo = "125/126"

    limpar_interface_resultado()

    if codigo not in TABELA_PARCEIROS:
        label_resultado.configure(text="Código não encontrado!", text_color="#e74c3c")
        return

    nome_parceiro = TABELA_PARCEIROS[codigo]
    label_resultado.configure(text=f"Parceiro: {nome_parceiro}", text_color="#1a73e8")

    if codigo not in CODIGOS_ATIVOS and codigo != "127/128":
        label_aviso.configure(
            text=f"Regras para {nome_parceiro} em desenvolvimento.",
            text_color="#e67e22"
        )
        return

    if codigo == "127/128":
        label_aviso.configure(text="Fluxo Parceiro Z pronto para uso.", text_color="#2ecc71")

    label_instrucao_data.pack(pady=(10, 0))
    entry_data.pack(pady=5)

    if codigo == "101":
        comando_processar = lambda e=None: selecionar_e_processar_parceiro_a(
            root, entry_data, progress_bar
        )
    elif codigo in CODIGOS_GRUPO_B:
        comando_processar = lambda e=None: selecionar_e_processar_parceiro_b(
            root, entry_data, progress_bar, nome_parceiro, codigo
        )
    elif codigo == "127/128":
        comando_processar = lambda e=None: processar_parceiro_c(
            root, entry_data, progress_bar
        )
    else:
        comando_processar = lambda e=None: selecionar_e_processar_geral(
            root, entry_data, progress_bar, nome_parceiro, codigo
        )

    entry_data.bind("<Return>", comando_processar)
    entry_data.bind("<KeyRelease>", formatar_data)

    btn_anexar.configure(
        text=f"Anexar e Gerar para {nome_parceiro}",
        command=comando_processar
    )
    btn_anexar.pack(pady=10, padx=30, fill="x")

    if codigo == "102":
        btn_conf.pack(pady=5, padx=30, fill="x")

    elif codigo == "101":
        btn_conf_parceiro_a.pack(pady=5, padx=30, fill="x")

    elif codigo in CODIGOS_GRUPO_D:
        btn_conf_parceiro_d.pack(pady=5, padx=30, fill="x")

    elif codigo in CODIGOS_GRUPO_B:
        btn_conf_parceiro_b.pack(pady=5, padx=30, fill="x")
        btn_conf_parceiro_b_endosso.pack(pady=5, padx=30, fill="x")

    elif codigo == "127/128":
        btn_conf.configure(
            text="Conferência Tesouraria - Parceiro Z",
            command=lambda: conferencia_parceiro_c(root, entry_data, progress_bar)
        )
        btn_conf.pack(pady=5, padx=30, fill="x")
        btn_conf_parceiro_c_pdf.pack(pady=5, padx=30, fill="x")


# ======================================================================================
# CRIAÇÃO DA JANELA PRINCIPAL
# ======================================================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

COR_PRIMARIA      = "#1C3F6E"
COR_PRIMARIA_HOVER = "#16335A"
COR_SECUNDARIA     = "#E8581A"
COR_SECUNDARIA_HVR = "#C44A12"
COR_HEADER_BG      = "#1C3F6E"
COR_BRANCA         = "#FFFFFF"

root = ctk.CTk()
root.title("ConferenciApp - Conferência de Contratos")
root.geometry("460x660")
root.resizable(False, False)

# ── Header ──────────────────────────────────────────────────────────────────
frame_header = ctk.CTkFrame(root, fg_color=COR_BRANCA, corner_radius=0)
frame_header.pack(fill="x")

frame_header_inner = ctk.CTkFrame(frame_header, fg_color="transparent")
frame_header_inner.pack(pady=14, padx=20)

frame_header_texts = ctk.CTkFrame(frame_header_inner, fg_color="transparent")
frame_header_texts.pack(side="left")

ctk.CTkLabel(
    frame_header_texts,
    text="ConferenciApp",
    font=ctk.CTkFont(family="Arial", size=24, weight="bold"),
    text_color="black"
).pack(anchor="w")

ctk.CTkLabel(
    frame_header_texts,
    text="Sistema de Conferência de Contratos",
    font=ctk.CTkFont(family="Arial", size=12),
    text_color="black"
).pack(anchor="w")

ctk.CTkFrame(root, height=4, fg_color=COR_SECUNDARIA, corner_radius=0).pack(fill="x")

# ── Corpo principal ──────────────────────────────────────────────────────────
frame_main = ctk.CTkFrame(root, fg_color="transparent")
frame_main.pack(fill="both", expand=True, padx=30, pady=20)

ctk.CTkLabel(
    frame_main,
    text="Código do Parceiro:",
    font=ctk.CTkFont(family="Arial", size=12, weight="bold")
).pack(pady=(0, 6))

entry_codigo = ctk.CTkEntry(
    frame_main,
    font=ctk.CTkFont(family="Arial", size=16),
    justify="center",
    width=200,
    height=40,
    placeholder_text="Ex: 101",
    border_color=COR_PRIMARIA,
    border_width=2
)
entry_codigo.pack(pady=(0, 8))
entry_codigo.bind("<Return>", buscar_parceiro)

ctk.CTkButton(
    frame_main,
    text="Confirmar Código",
    command=buscar_parceiro,
    width=200,
    height=38,
    font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    fg_color=COR_PRIMARIA,
    hover_color=COR_PRIMARIA_HOVER,
    corner_radius=8
).pack(pady=(0, 12))

ctk.CTkFrame(frame_main, height=2, fg_color=COR_SECUNDARIA, corner_radius=0).pack(fill="x", pady=(0, 12))

label_resultado = ctk.CTkLabel(
    frame_main,
    text="",
    font=ctk.CTkFont(family="Arial", size=13, weight="bold")
)
label_resultado.pack(pady=(0, 4))

label_aviso = ctk.CTkLabel(
    frame_main,
    text="",
    font=ctk.CTkFont(family="Arial", size=11),
    text_color=COR_SECUNDARIA
)
label_aviso.pack(pady=(0, 4))

label_instrucao_data = ctk.CTkLabel(
    frame_main,
    text="Digite a data (dd.mm.aaaa):",
    font=ctk.CTkFont(family="Arial", size=11)
)

entry_data = ctk.CTkEntry(
    frame_main,
    font=ctk.CTkFont(family="Arial", size=14),
    justify="center",
    width=160,
    height=36,
    border_color=COR_PRIMARIA,
    border_width=2
)
entry_data.insert(0, datetime.now().strftime("%d.%m.%Y"))

btn_anexar = ctk.CTkButton(
    frame_main,
    text="",
    fg_color=COR_SECUNDARIA,
    hover_color=COR_SECUNDARIA_HVR,
    font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    height=44,
    corner_radius=8
)

btn_conf = ctk.CTkButton(
    frame_main,
    text="Fazer Conferência (Excel x PDF)",
    fg_color=COR_PRIMARIA,
    hover_color=COR_PRIMARIA_HOVER,
    font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    height=44,
    corner_radius=8,
    command=fazer_conferencia
)

btn_conf_parceiro_a = ctk.CTkButton(
    frame_main,
    text="Fazer Conferência Parceiro A (Excel x CSV)",
    fg_color="#17a2b8",
    hover_color="#117a8b",
    font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    height=44,
    corner_radius=8,
    command=fazer_conferencia_parceiro_a
)

btn_conf_parceiro_d = ctk.CTkButton(
    frame_main,
    text="Fazer Conferência Parceiro D (Excel x PDF)",
    fg_color=COR_SECUNDARIA,
    hover_color=COR_SECUNDARIA_HVR,
    font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    height=44,
    corner_radius=8,
    command=fazer_conferencia_parceiro_d
)

btn_conf_parceiro_b = ctk.CTkButton(
    frame_main,
    text="Fazer Conferência Parceiro B (Excel x CSV)",
    fg_color="#198754",
    hover_color="#146c43",
    font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    height=44,
    corner_radius=8,
    command=fazer_conferencia_parceiro_b
)

btn_conf_parceiro_b_endosso = ctk.CTkButton(
    frame_main,
    text="Fazer Conferência Parceiro B - Endosso (Excel x PDF)",
    fg_color="#6f42c1",
    hover_color="#59359c",
    font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    height=44,
    corner_radius=8,
    command=fazer_conferencia_parceiro_b_endosso
)

btn_conf_parceiro_c_pdf = ctk.CTkButton(
    frame_main,
    text="Conferência Parceiro Z - PDF (Endosso)",
    fg_color="#2e7d32",
    hover_color="#1b5e20",
    font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    height=44,
    corner_radius=8,
    command=lambda: conferencia_pdf_parceiro_c(root, entry_data, progress_bar)
)

progress_bar = ctk.CTkProgressBar(
    frame_main,
    width=360,
    height=14,
    corner_radius=7,
    progress_color=COR_SECUNDARIA
)
progress_bar.set(0)


def criar_janela_principal():
    return root
