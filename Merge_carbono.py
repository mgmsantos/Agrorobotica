# %%
import os
from pathlib import Path
import pandas as pd

# %%
# ===================== CONFIGURAÇÕES =====================

BASE_DIR = r"C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\04_XLSX_OP_Formatado\04_Projetos_2022\OS_39"

OUTPUT_DIR_SUST = r"C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\ENTRADAS\ENTRADA_TESTE"
OUTPUT_DIR_COMBINED = r"C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\ENTRADAS\ENTRADA_TESTE"

OUTPUT_NAME_SUST = "F2022{codigo_os}-SUST.xlsx"
OUTPUT_NAME_COMBINED = "Fazenda_{codigo_os}_amostras_simples.xlsx"

SHEET_NAME_SUST = "SUST_ALL"
SHEET_NAME_COMBINED = "SUST6_LIBSrest"

JOIN_COL = "QR-Code"


# ===================== FUNÇÕES AUXILIARES =====================

def ignorar_pasta(folder: Path):
    return folder.name.upper().endswith(("_R2", "_DUPLICATA", "_DUPLICATR2"))

def extrair_codigo_os_de_nome(arquivo_sust: Path) -> str:
    parts = arquivo_sust.name.split('_')
    if len(parts) >= 2:
        return parts[1]
    else:
        return "OSDESCONHECIDA"

def extrair_timestamp(nome, tipo):
    try:
        parte = nome.split(f"_{tipo}_")[-1].replace(".xlsx", "")
        return pd.to_datetime(parte, format="%Y-%m-%d_%Hh%Mm%Ss")
    except Exception:
        return pd.NaT

def extrair_prefixo_id(nome, tipo):
    """Extrai a identificação da amostra (ex: OS_71_01) antes do tipo."""
    try:
        return nome.split(f"_{tipo}_")[0]
    except Exception:
        return nome

def selecionar_ultimas_versoes_por_id(lista_nomes, tipo):
    """Retorna um dicionário {ID: nome} com a versão mais recente de cada ID."""
    if not lista_nomes:
        return {}
    df = pd.DataFrame({
        "nome": lista_nomes,
        "ts": [extrair_timestamp(n, tipo) for n in lista_nomes],
        "prefixo": [extrair_prefixo_id(n, tipo) for n in lista_nomes]
    }).dropna(subset=["ts"])
    if df.empty:
        return {}
    # Pega o índice do maior timestamp para cada prefixo
    idx_max = df.groupby("prefixo")["ts"].idxmax()
    df_recentes = df.loc[idx_max]
    return dict(zip(df_recentes["prefixo"], df_recentes["nome"]))

def iter_all_subfolders(base_dir):
    base_path = Path(base_dir)
    for root, dirs, files in os.walk(base_path):
        folder = Path(root)
        if ignorar_pasta(folder):
            continue
        yield folder, files


# ===================== LÓGICA PRINCIPAL =====================

def main():
    sust_frames_corrigidos = []
    combined_frames = []
    first_sust_path = None

    # Garante que as pastas de saída existam
    Path(OUTPUT_DIR_SUST).mkdir(parents=True, exist_ok=True)
    Path(OUTPUT_DIR_COMBINED).mkdir(parents=True, exist_ok=True)

    for folder, files in iter_all_subfolders(BASE_DIR):
        all_sust = [f for f in files if f.lower().endswith(".xlsx") and "_sust_" in f.lower()]
        all_libs = [f for f in files if f.lower().endswith(".xlsx") and "_libs_" in f.lower()]

        dict_sust = selecionar_ultimas_versoes_por_id(all_sust, "SUST")
        dict_libs = selecionar_ultimas_versoes_por_id(all_libs, "LIBS")

        # IDs que vamos processar nesta pasta
        todos_ids = set(dict_sust.keys()) | set(dict_libs.keys())
        
        for prefixo in todos_ids:
            sust_file = dict_sust.get(prefixo)
            libs_file = dict_libs.get(prefixo)

            if not sust_file:
                continue # Se não tem SUST, não temos coordenadas/base

            try:
                sust_path = folder / sust_file
                if first_sust_path is None: first_sust_path = sust_path
                
                # Leitura
                df_sust = pd.read_excel(sust_path, dtype=str)
                df_sust.columns = [str(c).strip() for c in df_sust.columns]

                # Tenta ler o LIBS se ele existir para buscar Ponto/Profundidade
                df_libs = None
                if libs_file:
                    df_libs = pd.read_excel(folder / libs_file, dtype=str)
                    df_libs.columns = [str(c).strip() for c in df_libs.columns]

                # --- 1. RECUPERAÇÃO DE PONTO ---
                cands_ponto = ['Ponto']
                col_ponto = next((c for c in df_sust.columns if c in cands_ponto or 'ponto' in c.lower()), None)
                
                if col_ponto:
                    df_sust.rename(columns={col_ponto: 'Ponto'}, inplace=True)
                elif df_libs is not None:
                    col_p_libs = next((c for c in df_libs.columns if c in cands_ponto or 'ponto' in c.lower()), None)
                    if col_p_libs:
                        df_sust = df_sust.merge(df_libs[[JOIN_COL, col_p_libs]], on=JOIN_COL, how='left')
                        df_sust.rename(columns={col_p_libs: 'Ponto'}, inplace=True)
                
                if 'Ponto' not in df_sust.columns: df_sust['Ponto'] = ""

                # --- 2. RECUPERAÇÃO DE PROFUNDIDADE ---
                cands_prof = ['Profundidade']
                col_prof = next((c for c in df_sust.columns if c in cands_prof or 'prof' in c.lower()), None)
                
                if col_prof:
                    df_sust.rename(columns={col_prof: 'Profundidade'}, inplace=True)
                elif df_libs is not None:
                    col_pr_libs = next((c for c in df_libs.columns if c in cands_prof or 'prof' in c.lower()), None)
                    if col_pr_libs:
                        df_sust = df_sust.merge(df_libs[[JOIN_COL, col_pr_libs]], on=JOIN_COL, how='left')
                        df_sust.rename(columns={col_pr_libs: 'Profundidade'}, inplace=True)
                
                if 'Profundidade' not in df_sust.columns: df_sust['Profundidade'] = ""

                # --- 3. ORGANIZAÇÃO DO LAYOUT SUST (A-I com QR na I) ---
                cols_base = list(df_sust.columns[:6])
                ordem_fixa = cols_base + ['Ponto', 'Profundidade', JOIN_COL]
                
                # Criamos a versão do SUST corrigida para o mestre
                df_sust_corrigido = df_sust.reindex(columns=ordem_fixa + [c for c in df_sust.columns if c not in ordem_fixa])
                sust_frames_corrigidos.append(df_sust_corrigido)

                # --- 4. SE TIVER LIBS, GERA O COMBINADO ---
                if df_libs is not None:
                    libs_dados = [c for c in df_libs.columns if c not in [JOIN_COL, 'Ponto', 'Profundidade']]
                    df_comb = df_sust_corrigido.merge(df_libs[[JOIN_COL] + libs_dados], on=JOIN_COL, how='left')
                    combined_frames.append(df_comb)

            except Exception as e:
                print(f"[ERRO] Falha no ID {prefixo}: {e}")

    # --- SALVAMENTO ---
    if first_sust_path:
        codigo_os = extrair_codigo_os_de_nome(first_sust_path)
        out_sust = Path(OUTPUT_DIR_SUST) / OUTPUT_NAME_SUST.format(codigo_os=codigo_os)
        out_comb = Path(OUTPUT_DIR_COMBINED) / OUTPUT_NAME_COMBINED.format(codigo_os=codigo_os)

        if sust_frames_corrigidos:
            pd.concat(sust_frames_corrigidos, ignore_index=True).to_excel(out_sust, index=False)
            print(f"✔ Mestre SUST corrigido (com Ponto/Prof): {out_sust}")

        if combined_frames:
            pd.concat(combined_frames, ignore_index=True).to_excel(out_comb, index=False)
            print(f"✔ Combinado corrigido: {out_comb}")

if __name__ == "__main__":
    main()