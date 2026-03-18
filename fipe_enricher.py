import pandas as pd
import urllib.request
import urllib.parse
import json
import time
from pathlib import Path
from difflib import SequenceMatcher

def similar(a, b):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def get_json(url, retries=3):
    headers = {
        'Accept': 'application/json',
        'User-Agent': 'AFOOCOP-Integration/1.0 (sistemas@rfonseca.adv.br)'
    }
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=10) as resp:
                if resp.status == 200:
                    return json.loads(resp.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            if e.code == 429: # Too Many Requests
                print(f"      [!] HTTP 429 Too Many Requests. Aguardando 2s...")
                time.sleep(2)
            else:
                print(f"      [!] HTTP Error {e.code}: {e.reason}")
        except Exception as e:
            print(f"      [!] Erro de conexão: {e}")
        time.sleep(1)
    return None

def main():
    BASE_DIR = Path(r"C:\Users\Sairon\OneDrive - R FONSECA ADVOGADOS\01 - PROJETOS\AFOOCOP")
    file_path = BASE_DIR / "AFOOCOP_Rateios_Consolidado.xlsx"
    
    if not file_path.exists():
        print(f"Arquivo não encontrado: {file_path}")
        return

    print("1. Lendo arquivo consolidado...")
    df = pd.read_excel(file_path, sheet_name='MASTER_DATA')
    
    # Normalizando nomes possivelmente estranhos
    col_cavalo = next((c for c in df.columns if 'cavalo' in str(c).lower()), None)
    col_marca = next((c for c in df.columns if 'marca' in str(c).lower()), None)
    col_modelo = next((c for c in df.columns if 'modelo' in str(c).lower() and 'ano' not in str(c).lower()), None)
    col_ano = next((c for c in df.columns if 'ano' in str(c).lower() and 'mod' in str(c).lower()), None)
    col_valor = next((c for c in df.columns if 'valor equip' in str(c).lower()), None)
    
    if not all([col_marca, col_modelo, col_ano, col_valor]):
        print("Erro: Colunas necessárias não encontradas no arquivo.")
        print("Colunas atuais:", df.columns.tolist())
        return

    if 'NOTES' not in df.columns:
        df['NOTES'] = ""

    # Filtrar placas únicas que têm categoria mas estão SEM valor
    df_placas = df.drop_duplicates(subset=['PLACA']).copy()
    mask_needs_fipe = (df_placas[col_valor].isna() | (df_placas[col_valor] == 0)) & df_placas[col_marca].notna() & df_placas[col_modelo].notna() & df_placas[col_ano].notna()
    placas_para_buscar = df_placas[mask_needs_fipe]
    
    print(f"2. Identificadas {len(placas_para_buscar)} placas elegíveis para busca na FipeX.")
    if len(placas_para_buscar) == 0:
        print("Nenhuma placa precisa de atualização. Encerrando.")
        return

    # Cache de Marcas (Make UUIDs)
    print("\n3. Carregando Cache de Marcas (Makes) da FipeX...")
    make_cache = {}
    
    # Coletar marcas únicas da nossa base elegível
    marcas_unicas = placas_para_buscar[col_marca].dropna().unique()
    for marca in marcas_unicas:
        marca_str = str(marca).strip()
        
        # Correções de marcas muito comuns que falham no FipeX se exatas da nossa base
        busca_marca = marca_str
        # Correções de alias — marcas com nomes diferentes na FipeX
        alias_map = {
            'MERCEDES BENZ': 'MERCEDES-BENZ',
            'VW': 'VW - VolksWagen',
            'VOLKSWAGEN': 'VW - VolksWagen',
        }
        busca_marca = alias_map.get(marca_str, marca_str)

        url_make = f"https://api.fipex.com.br/v1/makes?filters[0].field=name&filters[0].op=CONTAINS&filters[0].value={urllib.parse.quote(busca_marca)}"
        res = get_json(url_make)
        if res and res.get('data'):
            # Para alias conhecidos, aceitar direto sem checar similaridade com o nome original
            if marca_str in alias_map:
                best_make = res['data'][0]
                make_cache[marca_str] = best_make['id']
                print(f"   [+] Marca mapeada (alias): '{marca_str}' -> '{best_make['name']}' ({best_make['id']})")
            else:
                best_make = max(res['data'], key=lambda x: similar(marca_str, x['name']))
                if similar(marca_str, best_make['name']) > 0.6:
                    make_cache[marca_str] = best_make['id']
                    print(f"   [+] Marca mapeada: '{marca_str}' -> '{best_make['name']}' ({best_make['id']})")
                else:
                    print(f"   [-] Marca '{marca_str}' rejeitada (Similaridade baixa com '{best_make['name']}')")
        else:
             print(f"   [-] Marca '{marca_str}' não encontrada na FipeX.")
        time.sleep(0.1) # Respect rate limits

    print("\n4. Buscando Modelos na FipeX...")
    
    # Cache por (MakeID, Year) -> list de dicionários {model_name, price}
    model_cache = {}
    atualizados = 0
    erros = 0
    cache_hits = 0

    updates = {} # dict placa -> (valor, score, nome_fipe)

    for idx, row in placas_para_buscar.iterrows():
        placa = row['PLACA']
        marca = str(row[col_marca]).strip()
        modelo = str(row[col_modelo]).strip()
        
        # O Ano Modelo pode vir como float '2013.0' ou string '2013/2014'. Limpando:
        ano_str = str(row[col_ano]).split('.')[0].strip()
        try:
            ano = int(ano_str)
            if ano < 1980 or ano > 2030:
                continue
        except:
             continue # Ano inválido

        make_id = make_cache.get(marca)
        if not make_id:
            continue
            
        cache_key = (make_id, ano)
        
        if cache_key not in model_cache:
            # Buscar todos os veículos dessa Marca e Ano
            url = f"https://api.fipex.com.br/v1/search?filters[0].field=make&filters[0].op=%3D&filters[0].value={make_id}&filters[1].field=year&filters[1].op=%3D&filters[1].value={ano}&limit=200"
            res = get_json(url)
            if res and res.get('data'):
                model_cache[cache_key] = [
                    {
                        'name': item.get('model_name', ''),
                        'price': item.get('latest_market_price_cents', 0) / 100
                    }
                    for item in res['data']
                ]
            else:
                 model_cache[cache_key] = []
            time.sleep(0.1)

        fipe_models = model_cache.get(cache_key, [])
        if not fipe_models:
             continue
             
        # Fuzzy Matching
        best_match = None
        best_score = 0
        
        # Limpar o nosso nome de modelo para melhorar o match
        # Remover textos comuns como "2p", "caminhão", e sufixos técnicos que a FIPE costuma ignorar
        # ou grafar diferente
        import re
        modelo_limpo = modelo.lower()
        
        # Remove coisas como "6X2T", "6X4", "8X2", "4X2" etc
        modelo_limpo = re.sub(r'\d+x\d+[a-z]*', '', modelo_limpo)
        
        # Remove coisas como "E5", "E6" (padrões de emissão)
        modelo_limpo = re.sub(r'e[56]', '', modelo_limpo)
        
        modelo_limpo = modelo_limpo.replace("especial", "").replace("basico", "")
        modelo_limpo = modelo_limpo.replace("caminhao", "").replace("trator", "")
        modelo_limpo = modelo_limpo.strip()

        for fm in fipe_models:
            fm_nom = fm['name'].lower()
            # Limpa o nome da Fipe também pra comparar semente com semente
            fm_nom_limpo = re.sub(r'\d+x\d+[a-z]*', '', fm_nom)
            fm_nom_limpo = re.sub(r'e[56]', '', fm_nom_limpo)
            fm_nom_limpo = fm_nom_limpo.replace("caminhão", "").replace("caminhao", "").replace("2p", "").replace("3p", "").replace("4p", "").strip()
            
            score = similar(modelo_limpo, fm_nom_limpo)
            
            # Boost no score se uma string está totalmente contida na outra
            if modelo_limpo and (modelo_limpo in fm_nom_limpo or fm_nom_limpo in modelo_limpo):
                 score += 0.25
            
            if score > best_score:
                best_score = score
                best_match = fm

        # Limitar uma similaridade mínima de 0.60 (60%)
        # a pedido do usuário para maximizar a cobertura
        if best_match and best_score >= 0.60:
            valor_fipe = best_match['price']
            if valor_fipe > 0:
                updates[placa] = (valor_fipe, best_score, best_match['name'])
                atualizados += 1
                if atualizados % 50 == 0:
                     print(f"   Progresso: {atualizados} valores encontrados...")
        else:
             erros += 1

    print(f"\n5. Busca Finalizada!")
    print(f"   - {atualizados} valores encontrados (Fuzzy Match >= 60%).")
    print(f"   - {erros} não encontrados/sem match seguro.")
    
    if atualizados > 0:
        print("\n6. Atualizando MASTER_DATA...")
        # Atualizar a planilha principal
        mask_placas = df['PLACA'].isin(updates.keys())
        linhas_afetadas = df[mask_placas].index
        
        for idx in linhas_afetadas:
            placa = df.at[idx, 'PLACA']
            valor, score, nome_fipe = updates[placa]
            df.at[idx, col_valor] = valor
            
            nota_antiga = str(df.at[idx, 'NOTES']) if pd.notna(df.at[idx, 'NOTES']) and df.at[idx, 'NOTES'] != "" else ""
            nota_nova = f"FipeX Auto-Match ({nome_fipe}) | "
            df.at[idx, 'NOTES'] = nota_nova + nota_antiga

        # Salvar backup e em seguida sobrescrever
        print("   Salvando arquivo consolidado atualizado...")
        
        backup_path = BASE_DIR / "AFOOCOP_Rateios_Consolidado_Backup_PreFipe.xlsx"
        if not backup_path.exists():
            import shutil
            shutil.copy2(file_path, backup_path)
            
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
             df.to_excel(writer, sheet_name='MASTER_DATA', index=False)
             
        print(f"   [OK] {len(linhas_afetadas)} linhas de rateio atualizadas com sucesso!")

if __name__ == "__main__":
    main()
