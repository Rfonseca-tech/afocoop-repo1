"""
Enriquece o Relatório de Veículos Ativos com Valor Equipamento FIPE.
Etapa 1: cruzamento com base consolidada (por placa)
Etapa 2: PlacaFipe API — paralelo com N workers
"""
import pandas as pd
import urllib.request
import json
import time
import shutil
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# ==========================================
RELATORIO  = r"c:\Users\Sairon\OneDrive - R FONSECA ADVOGADOS\01 - PROJETOS\AFOOCOP\Analise dos rateios\01 - RELATÓRIO DE VEÍCULOS ATIVOS - 21-01-2026.xlsx"
BASE       = r"c:\Users\Sairon\OneDrive - R FONSECA ADVOGADOS\01 - PROJETOS\AFOOCOP\AFOOCOP_Rateios_Consolidado.xlsx"
SHEET      = "Relatorio-relVeiculos (15)"
API_TOKEN  = "9AEC24604B8026EA30E45703E4D6F8BCB11BEA31ADBB8DC6A2949536CB6960A9"
COL_VALOR  = "Valor Equipamento"
COL_PLACA  = "Placa"
WORKERS    = 10   # requisições paralelas
SAVE_EVERY = 500  # salvar progresso a cada N placas
# ==========================================

print_lock = threading.Lock()
counter    = {"ok": 0, "fail": 0, "done": 0}


def get_fipe_por_placa(placa):
    url = f"https://api.placafipe.com.br/getplacafipe/{placa}/{API_TOKEN}"
    req = urllib.request.Request(url, headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                return json.loads(r.read().decode())
        except urllib.error.HTTPError as e:
            if e.code == 429:
                time.sleep(5)
                continue
            return {"erro": f"HTTP {e.code}"}
        except Exception as e:
            if attempt == 2:
                return {"erro": str(e)}
            time.sleep(1)
    return {"erro": "Max retries"}


def consultar(args):
    idx_list, placa_orig, placa_norm = args
    res = get_fipe_por_placa(placa_orig)

    if "fipe" in res and len(res["fipe"]) > 0:
        fipe = res["fipe"][0]
        try:
            valor = float(str(fipe.get("valor", "0")).replace(",", "."))
        except:
            valor = 0
        if valor > 0:
            return (idx_list, placa_norm, valor, True, None)
        return (idx_list, placa_norm, 0, False, "valor R$0")
    else:
        msg = res.get("msg", res.get("erro", "?"))
        return (idx_list, placa_norm, 0, False, msg)


def salvar(df, path, sheet):
    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df.to_excel(w, sheet_name=sheet, index=False)


def main():
    print("=" * 60)
    print(" ENRIQUECIMENTO — RELATÓRIO DE VEÍCULOS ATIVOS (paralelo)")
    print("=" * 60)

    ts = time.strftime("%Y%m%d_%H%M%S")
    backup = Path(RELATORIO).with_suffix(f".backup_{ts}.xlsx")
    shutil.copy2(RELATORIO, backup)
    print(f"\n1. Backup: {backup.name}")

    print("2. Lendo dados...")
    df = pd.read_excel(RELATORIO, sheet_name=SHEET)
    total = len(df)
    print(f"   {total} veículos | sem valor: {df[COL_VALOR].isna().sum()}")

    # ----------------------------------------------------------------
    # ETAPA 1 — base interna
    # ----------------------------------------------------------------
    print("\n3. Etapa 1 — Base interna...")
    base = pd.read_excel(BASE, sheet_name="MASTER_DATA")
    col_val_base = next(c for c in base.columns if "valor equip" in c.lower())
    mapa = (
        base.drop_duplicates(subset=["PLACA"])
        .pipe(lambda x: x[x[col_val_base] > 0])
        .assign(N=lambda x: x["PLACA"].str.replace("-", "").str.upper())
        .set_index("N")[col_val_base]
        .to_dict()
    )

    df["_NORM"] = df[COL_PLACA].astype(str).str.replace("-", "").str.upper()
    etapa1 = 0
    for idx, row in df.iterrows():
        if pd.notna(row[COL_VALOR]):
            continue
        v = mapa.get(row["_NORM"])
        if v:
            df.at[idx, COL_VALOR] = v
            etapa1 += 1
    print(f"   [OK] {etapa1} preenchidos")

    # ----------------------------------------------------------------
    # ETAPA 2 — API paralela
    # ----------------------------------------------------------------
    sem = df[df[COL_VALOR].isna()]
    # agrupar por placa (evitar chamar mesma placa N vezes)
    placa_to_idx = {}
    for idx, row in sem.iterrows():
        n = row["_NORM"]
        placa_to_idx.setdefault(n, {"orig": row[COL_PLACA], "idxs": []})
        placa_to_idx[n]["idxs"].append(idx)

    placas = [(v["idxs"], v["orig"], n) for n, v in placa_to_idx.items()]
    total_api = len(placas)
    print(f"\n4. Etapa 2 — API ({total_api} placas únicas | {WORKERS} workers paralelos)...")

    api_cache = {}
    limite_atingido = False
    done = 0
    ok = 0
    fail = 0
    t0 = time.time()

    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(consultar, p): p for p in placas}
        for fut in as_completed(futures):
            idx_list, placa_norm, valor, sucesso, msg = fut.result()
            done += 1

            if sucesso:
                api_cache[placa_norm] = valor
                ok += 1
            else:
                fail += 1
                if msg and "limite" in str(msg).lower():
                    limite_atingido = True
                    pool.shutdown(wait=False, cancel_futures=True)
                    print(f"\n   [!] LIMITE DA API atingido após {done} consultas. Salvando...")
                    break

            if done % 100 == 0 or done == total_api:
                elapsed = time.time() - t0
                rate = done / elapsed if elapsed > 0 else 0
                eta = (total_api - done) / rate if rate > 0 else 0
                print(f"   [{done}/{total_api}] OK={ok} FALHA={fail} | {rate:.1f} req/s | ETA {eta/60:.1f}min")

            # Salvar progresso intermediário
            if done % SAVE_EVERY == 0:
                for pn, v in api_cache.items():
                    for i in placa_to_idx.get(pn, {}).get("idxs", []):
                        if pd.isna(df.at[i, COL_VALOR]):
                            df.at[i, COL_VALOR] = v
                salvar(df, RELATORIO, SHEET)
                print(f"   [SAVE] Progresso salvo ({done} processadas)")

    # Injetar resultados finais
    for pn, v in api_cache.items():
        for i in placa_to_idx.get(pn, {}).get("idxs", []):
            if pd.isna(df.at[i, COL_VALOR]):
                df.at[i, COL_VALOR] = v

    df.drop(columns=["_NORM"], inplace=True)

    ainda_sem = df[COL_VALOR].isna().sum()
    print(f"\n5. Resumo:")
    print(f"   Base interna:    {etapa1}")
    print(f"   API (OK):        {ok}")
    print(f"   API (falha):     {fail}")
    print(f"   Ainda sem valor: {ainda_sem} de {total}")
    if limite_atingido:
        print(f"   ⚠ Limite da API atingido — rode novamente amanhã para completar")

    print("\n6. Salvando arquivo final...")
    salvar(df, RELATORIO, SHEET)
    print(f"[SUCESSO] {Path(RELATORIO).name} atualizado!")


if __name__ == "__main__":
    main()
