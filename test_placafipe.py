import pandas as pd
import urllib.request
import json
import time
import shutil
from pathlib import Path

# ==========================================
# CONFIGURAÇÕES
# ==========================================
ARQUIVO_MASTER = 'AFOOCOP_Rateios_Consolidado.xlsx'
API_TOKEN = "9AEC24604B8026EA30E45703E4D6F8BCB11BEA31ADBB8DC6A2949536CB6960A9"
# ==========================================

def get_fipe_por_placa(placa):
    url = f"https://api.placafipe.com.br/getplacafipe/{placa}/{API_TOKEN}"
    req = urllib.request.Request(url, headers={'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'})
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            with urllib.request.urlopen(req, timeout=10) as response:
                return json.loads(response.read().decode())
        except urllib.error.HTTPError as e:
            if e.code == 429:
                print("   [!] Limite de taxa (429). Aguardando 10s...")
                time.sleep(10)
                continue
            return {"erro": f"HTTP {e.code}"}
        except Exception as e:
            if attempt == max_retries - 1:
                return {"erro": str(e)}
            time.sleep(2)
    return {"erro": "Max retries excedido"}

def main():
    print("========================================")
    print(" ENRIQUECIMENTO DE DADOS - PLACAFIPE.COM ")
    print("========================================\n")

    if not Path(ARQUIVO_MASTER).exists():
        print(f"[ERRO] Arquivo {ARQUIVO_MASTER} não encontrado.")
        return

    # Backup de segurança
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    backup_file = f"{ARQUIVO_MASTER}.backup_antes_placafipe_{timestamp}.xlsx"
    shutil.copy2(ARQUIVO_MASTER, backup_file)
    print(f"1. Backup criado: {backup_file}")

    print("2. Carregando dados...")
    df = pd.read_excel(ARQUIVO_MASTER, sheet_name='MASTER_DATA')
    
    # Isolar linhas que ainda não possuem Valor Equipamento
    faltantes = df[df['Valor Equipamento'].isna() | (df['Valor Equipamento'] == 0)]
    placas_unicas = faltantes['PLACA'].dropna().unique()
    
    total_placas = len(placas_unicas)
    print(f"   - Total de Placas SECA (Sem Fipe): {total_placas}")
    
    if total_placas == 0:
        print("\n[OK] Todas as placas já possuem Valor do Equipamento!")
        return

    print("\n3. Iniciando disparos à API (Buscando...)\n")
    
    sucessos = 0
    erros = 0
    updates_cache = {} # placa -> (valor, modelo, marca, cavalo_carreta)
    
    for i, placa in enumerate(placas_unicas, 1):
        print(f"   [{i}/{total_placas}] Consultando {placa}...", end=" ")
        
        res = get_fipe_por_placa(placa)
        
        if 'fipe' in res and len(res['fipe']) > 0:
            fipe = res['fipe'][0]
            info = res.get('informacoes_veiculo', {})
            
            valor_str = str(fipe.get('valor', '0')).replace(',', '.')
            try:
                 valor_float = float(valor_str)
            except:
                 valor_float = 0
                 
            if valor_float > 0:
                modelo_fipe = fipe.get('modelo', 'Desconhecido')
                marca_detran = info.get('marca', '')
                modelo_detran = info.get('modelo', '')
                segmento = info.get('sub_segmento', '')
                
                # Se a API nos disser que é reboque, a gente guarda
                tipo = 'Carreta' if 'REBOQUE' in str(segmento).upper() else 'Cavalo'
                
                updates_cache[placa] = {
                    'valor': valor_float,
                    'modelo': f"{marca_detran} {modelo_detran}".strip() or modelo_fipe,
                    'marca': marca_detran or fipe.get('marca', ''),
                    'tipo': tipo
                }
                
                print(f"[OK] R$ {valor_float:,.2f}")
                sucessos += 1
            else:
                print(f"[ERRO] Retornou R$ 0.00")
                erros += 1
        else:
            erro_msg = res.get('msg', res.get('erro', 'Desconhecido'))
            print(f"[FALHA] {erro_msg}")
            erros += 1
            
        # Respeitar APIs - 1 requisição a cada 1 segundo (aumentar se a chave chorar)
        time.sleep(1)

    print(f"\n4. Resumo da Busca:")
    print(f"   - Sucessos: {sucessos}")
    print(f"   - Falhas  : {erros}")
    
    if sucessos > 0:
        print("\n5. Injetando dados recuperados na planilha original...")
        linhas_atualizadas = 0
        
        for idx, row in df.iterrows():
            placa = row['PLACA']
            if pd.isna(row['Valor Equipamento']) or row['Valor Equipamento'] == 0:
                if placa in updates_cache:
                    novo_dado = updates_cache[placa]
                    df.at[idx, 'Valor Equipamento'] = novo_dado['valor']
                    
                    # Se faltar Categoria (Marca/Modelo/Tipo) nessa placa, preenchemos tambem!
                    if pd.isna(row.get('Marca')) and novo_dado['marca']:
                         df.at[idx, 'Marca'] = novo_dado['marca']
                    if pd.isna(row.get('Modelo')) and novo_dado['modelo']:
                         df.at[idx, 'Modelo'] = novo_dado['modelo']
                    if pd.isna(row.get('Cavalo/Carreta')) and novo_dado['tipo']:
                         df.at[idx, 'Cavalo/Carreta'] = novo_dado['tipo']

                    # Annotar em notas o milagre
                    nota_antiga = str(row['NOTES']) if pd.notna(row.get('NOTES')) else ""
                    if "PlacaFipe API:" not in nota_antiga:
                         nota_nova = f"PlacaFipe API: R$ {novo_dado['valor']} | "
                         df.at[idx, 'NOTES'] = nota_nova + nota_antiga
                    
                    linhas_atualizadas += 1
                    
        print(f"   [OK] {linhas_atualizadas} transações tarifadas foram atualizadas com sucesso!")
        
        print("\n6. Salvando o arquivo...")
        
        # Salvar apenas se teve att
        with pd.ExcelWriter(ARQUIVO_MASTER, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='MASTER_DATA', index=False)
            
        print(f"[SUCESSO] Arquivo {ARQUIVO_MASTER} atualizado e pronto para o Dashboard.")

if __name__ == '__main__':
    main()



