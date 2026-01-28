import pandas as pd
import requests
import time
import re


ARQUIVO_ENTRADA = "entrada.xlsx"
ARQUIVO_SAIDA = "saida.xlsx"

DELAY_BASE = 3 
DELAY_MAX = 10


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
                print(f"[HTTP {response.status_code}] Erro. Aguardando {delay_atual}s...")
                time.sleep(delay_atual)
                delay_atual = min(delay_atual + 2, DELAY_MAX)
                continue

            dados = response.json()

            if dados.get("status") != "OK":
                print(f"[API ERRO] {cnpj}: {dados.get('message')} | Esperando {delay_atual}s...")
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
            print(f"[ERRO DE REDE] {cnpj}: {e} | Esperando {delay_atual}s...")
            time.sleep(delay_atual)
            delay_atual = min(delay_atual + 2, DELAY_MAX)

def main():
    df = pd.read_excel(ARQUIVO_ENTRADA)

    coluna_cnpj = encontrar_coluna_cnpj(df.columns)

    if not coluna_cnpj:
        raise Exception(
            f"Nenhuma coluna de CNPJ encontrada. Colunas disponíveis: {list(df.columns)}"
        )

    print(f"✅ Coluna de CNPJ detectada: '{coluna_cnpj}'\n")

    resultados = []

    for _, row in df.iterrows():
        cnpj_original = row[coluna_cnpj]
        cnpj = padronizar_cnpj(cnpj_original)

        print(f"Consultando: {cnpj_original} → {cnpj}")

        if not cnpj:
            print("  [CNPJ INVÁLIDO]\n")
            continue

        resultado = consultar_cnpj_com_backoff(cnpj)
        resultados.append(resultado)
        
        time.sleep(DELAY_BASE)

    df_saida = pd.DataFrame(resultados)
    df_saida.to_excel(ARQUIVO_SAIDA, index=False)

    print(f"\n✅ Arquivo '{ARQUIVO_SAIDA}' criado com {len(df_saida)} registros.")


if __name__ == "__main__":
    main()


