import pandas as pd
import requests
from tqdm import tqdm
import time
import re
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading


ARQUIVO_ENTRADA = "entrada.xlsx"
ARQUIVO_SAIDA = "saida.xlsx"

DELAY_BASE = 3        # delay inicial
DELAY_MAX = 10        # delay máximo permitido

def padronizar_cnpj(cnpj):
    if pd.isna(cnpj):
        return None

    cnpj = re.sub(r"\D", "", str(cnpj))

    if len(cnpj) == 13:
        cnpj = "0" + cnpj

    if len(cnpj) != 14:
        return None

    return cnpj


def bool_para_sim_nao(valor):
    if valor is True:
        return "SIM"
    elif valor is False:
        return "NÃO"
    return ""

def encontrar_coluna_cnpj(colunas):
    for col in colunas:
        if "cnpj" in col.strip().lower():
            return col
    return None

def consultar_cnpj_com_backoff(cnpj):
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"
    delay_atual = DELAY_BASE

    while True:
        try:
            response = requests.get(url, timeout=20)

            if response.status_code != 200:
                time.sleep(delay_atual)
                delay_atual = min(delay_atual + 2, DELAY_MAX)
                continue

            dados = response.json()

            if dados.get("status") != "OK":
                time.sleep(delay_atual)
                delay_atual = min(delay_atual + 2, DELAY_MAX)
                continue

            simples = dados.get("simples", {})
            simei = dados.get("simei", {})

            return {
                "nome": dados.get("nome"),
                "cnpj": cnpj,
                "simples_nacional": bool_para_sim_nao(simples.get("optante")),
                "mei": bool_para_sim_nao(simei.get("optante"))
            }

        except requests.exceptions.RequestException as e:
            time.sleep(delay_atual)
            delay_atual = min(delay_atual + 2, DELAY_MAX)

def log(msg):
    root.after(0, lambda: _log(msg))

def _log(msg):
    log_area.config(state="normal")
    log_area.insert(tk.END, msg + "\n")
    log_area.see(tk.END)
    log_area.config(state="disabled")

def processar():
    try:
        df = pd.read_excel(ARQUIVO_ENTRADA)
        coluna_cnpj = encontrar_coluna_cnpj(df.columns)
        
        if not coluna_cnpj:
            log("❌ Coluna de CNPJ não encontrada.")
            return

        log(f"✅ Coluna detectada: {coluna_cnpj}")

        total = len(df)
        progress["maximum"] = total

        inicio = time.time()
        resultados = []

        for i, (_, row) in enumerate(df.iterrows()):
            cnpj_original = row[coluna_cnpj]
            nome = row.get("nome", "")

            cnpj = padronizar_cnpj(cnpj_original)
            
            exibicao = nome if pd.notna(nome) and nome != "" else cnpj
            
            if not cnpj:
                log(f"⚠ CNPJ inválido: {cnpj_original}")
                continue

            log(f"Consultando {exibicao}...")

            resultado = consultar_cnpj_com_backoff(cnpj)
            resultados.append(resultado)

            percent = int((i + 1) / total * 100)
            progress["value"] = i + 1
            percent_label.config(text=f"{percent}%")

            decorrido = time.time() - inicio
            media = decorrido / (i + 1)
            restante = media * (total - (i + 1))

            m = int(restante // 60)
            s = int(restante % 60)
            eta_label.config(text=f"Tempo restante: {m}m {s}s")

            root.update_idletasks()
            time.sleep(DELAY_BASE)

        df_saida = pd.DataFrame(resultados)
        df_saida.to_excel(ARQUIVO_SAIDA, index=False)

        eta_label.config(text="Concluído ✔")
        log(f"✅ Arquivo salvo: {ARQUIVO_SAIDA}")

        root.after(0, lambda: messagebox.showinfo("Concluído", "Processamento finalizado!"))

    except Exception as e:
        log(f"❌ ERRO: {e}")

def iniciar():
    threading.Thread(target=processar).start()

root = tk.Tk()
root.title("Consulta CNPJ Pro")
root.geometry("560x520")
root.configure(bg="#f5f6fa")

# ===== Frame principal =====
main_frame = tk.Frame(root, bg="#f5f6fa", padx=20, pady=20)
main_frame.pack(fill="both", expand=True)

# ===== Título =====
title = tk.Label(
    main_frame,
    text="Consulta CNPJ Pro",
    font=("Segoe UI", 16, "bold"),
    bg="#f5f6fa"
)
title.pack(pady=(0, 15))

# ===== Barra de progresso =====
progress = ttk.Progressbar(main_frame, length=450)
progress.pack(pady=10)

# ===== Porcentagem =====
percent_label = tk.Label(
    main_frame,
    text="0%",
    font=("Segoe UI", 18, "bold"),
    bg="#f5f6fa"
)
percent_label.pack()

# ===== ETA =====
eta_label = tk.Label(
    main_frame,
    text="Tempo restante: --",
    font=("Segoe UI", 10),
    fg="#555",
    bg="#f5f6fa"
)
eta_label.pack(pady=(0, 10))

# ===== Botão =====
btn = tk.Button(
    main_frame,
    text="Iniciar",
    font=("Segoe UI", 11),
    width=20,
    height=1,
    bg="#2ecc71",
    fg="white",
    activebackground="#27ae60",
    relief="flat",
    command=iniciar
)
btn.pack(pady=10)

# ===== Console =====
log_frame = tk.Frame(main_frame, bg="#dcdde1", padx=1, pady=1)
log_frame.pack(fill="both", expand=True, pady=10)

log_area = ScrolledText(
    log_frame,
    height=12,
    font=("Consolas", 10),
    bg="#1e1e1e",
    fg="#dcdcdc",
    insertbackground="white",
    state="disabled"
)
log_area.pack(fill="both", expand=True)

root.mainloop()


