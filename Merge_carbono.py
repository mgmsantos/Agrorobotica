# %%
import os
from pathlib import Path
import pandas as pd

# ===================== CONFIGURAÇÕES =====================

BASE_DIR = r"\\Agroserver\processos\04_XLSX_OP_Formatado\05_Projetos_2023\OS_140"

OUTPUT_DIR_SUST = r"C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\ENTRADAS\TESTE"
OUTPUT_DIR_COMBINED = r"C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\ENTRADAS\TESTE"

OUTPUT_NAME_SUST = "F2025{codigo_os}-SUST.xlsx"
OUTPUT_NAME_COMBINED = "Fazenda_{codigo_os}_amostras_simples.xlsx"

SHEET_NAME_SUST = "SUST_ALL"
SHEET_NAME_COMBINED = "SUST6_LIBSrest"

JOIN_COL = "QR-Code"


# ===================== FUNÇÕES AUXILIARES =====================

def ignorar_pasta(folder: Path):
    return folder.name.endswith(("_R2", "_DUPLICATA", "_DUPLICATR2"))


def extrair_codigo_os_de_nome(arquivo_sust: Path) -> str:
    parts = arquivo_sust.name.split('_')
    if len(parts) >= 2:
        return parts[1]
    else:
        return "OSDESCONHECIDA"


def extrair_timestamp(nome, tipo):
    """
    Extrai timestamp do final do nome, exemplo:
    *_SUST_2025-11-25_11h25m46s.xlsx
    *_LIBS_2025-11-25_11h25m46s.xlsx
    """
    try:
        parte = nome.split(f"_{tipo}_")[-1].replace(".xlsx", "")
        return pd.to_datetime(parte, format="%Y-%m-%d_%Hh%Mm%Ss")
    except Exception:
        return pd.NaT


def selecionar_ultima_versao(lista_nomes, tipo):
    """Retorna o arquivo mais recente baseado no timestamp."""
    if not lista_nomes:
        return None
    df = pd.DataFrame({
        "nome": lista_nomes,
        "ts": [extrair_timestamp(n, tipo) for n in lista_nomes]
    }).dropna(subset=["ts"])
    if df.empty:
        return None
    return df.sort_values("ts").iloc[-1]["nome"]


def iter_all_subfolders(base_dir):
    base_path = Path(base_dir)
    for root, dirs, files in os.walk(base_path):
        folder = Path(root)
        if ignorar_pasta(folder):
            print(f"→ Ignorando pasta (duplicata/R2): {folder}")
            continue
        yield folder, files


# ===================== LÓGICA PRINCIPAL =====================

def main():
    sust_frames = []
    combined_frames = []
    first_sust_path = None

    for folder, files in iter_all_subfolders(BASE_DIR):

        all_sust = [f for f in files if f.lower().endswith(".xlsx") and "_sust_" in f.lower()]
        all_libs = [f for f in files if f.lower().endswith(".xlsx") and "_libs_" in f.lower()]

        # Selecionar somente a última versão em cada pasta
        sust_file = selecionar_ultima_versao(all_sust, "SUST")
        libs_file = selecionar_ultima_versao(all_libs, "LIBS")

        # ---------- CONCAT (JUSTAPOSIÇÃO) de TODOS os SUST (SEM MODIFICAR NADA) ----------
        if sust_file:
            fpath = folder / sust_file
            if first_sust_path is None:
                first_sust_path = fpath

            try:
                df_sust = pd.read_excel(fpath)
                sust_frames.append(df_sust)
            except Exception as e:
                print(f"[ERRO SUST] Não consegui ler {fpath}: {e}")

        # ---------- COMBINADO: SUST6 + LIBSrest (ligando por QR-Code) ----------
        if sust_file and libs_file:
            sust_path = folder / sust_file
            libs_path = folder / libs_file

            try:
                df_sust = pd.read_excel(sust_path)
                df_libs = pd.read_excel(libs_path)
            except Exception as e:
                print(f"[ERRO COMBINADO] {folder}: {e}")
                continue

            # Checar coluna de junção
            if JOIN_COL not in df_sust.columns:
                print(f"[AVISO] {folder} tem SUST+LIBS, mas SUST não tem coluna '{JOIN_COL}'. Pulando combinado.")
                continue
            if JOIN_COL not in df_libs.columns:
                print(f"[AVISO] {folder} tem SUST+LIBS, mas LIBS não tem coluna '{JOIN_COL}'. Pulando combinado.")
                continue

            # Pegar as 6 primeiras colunas do SUST (como antes)
            sust_cols6 = list(df_sust.columns[:6])

            # Garantir que JOIN_COL exista no lado esquerdo do merge (mesmo que não esteja nas 6 primeiras)
            add_join_temp = JOIN_COL not in sust_cols6
            if add_join_temp:
                sust_left = df_sust[sust_cols6 + [JOIN_COL]].copy()
            else:
                sust_left = df_sust[sust_cols6].copy()

            # Merge por QR-Code (mantém todas as linhas do SUST; traz colunas do LIBS)
            df_merged = sust_left.merge(
                df_libs,
                on=JOIN_COL,
                how="left",
                suffixes=("", "_LIBS")
            )

            # Montar output: 6 colunas da SUST + todas do LIBS (exceto QR-Code duplicado)
            libs_cols_out = [c for c in df_libs.columns if c != JOIN_COL]
            out_cols = sust_cols6 + libs_cols_out

            # Se o JOIN_COL foi adicionado só para merge e não faz parte das 6 colunas, remove do output
            # (out_cols já não inclui JOIN_COL nesse caso)
            df_combined = df_merged.reindex(columns=out_cols)

            combined_frames.append(df_combined)

        else:
            if sust_file and not libs_file:
                print(f"[AVISO] Pasta {folder} tem SUST mas NÃO tem LIBS.")
            elif libs_file and not sust_file:
                print(f"[AVISO] Pasta {folder} tem LIBS mas NÃO tem SUST.")

    # ===================== NOMES DOS OUTPUTS =====================

    if first_sust_path is not None:
        codigo_os = extrair_codigo_os_de_nome(first_sust_path)
    else:
        codigo_os = "OSDESCONHECIDA"

    Path(OUTPUT_DIR_SUST).mkdir(parents=True, exist_ok=True)
    Path(OUTPUT_DIR_COMBINED).mkdir(parents=True, exist_ok=True)

    output_sust_all = Path(OUTPUT_DIR_SUST) / OUTPUT_NAME_SUST.format(codigo_os=codigo_os)
    output_combined_all = Path(OUTPUT_DIR_COMBINED) / OUTPUT_NAME_COMBINED.format(codigo_os=codigo_os)

    # ===================== SALVANDO =====================

    if sust_frames:
        # Justaposição/concat SEM ordenar, SEM cortar, SEM mexer: só empilha
        df_sust_all = pd.concat(sust_frames, ignore_index=True, sort=False)

        df_sust_all.to_excel(
            output_sust_all, index=False, sheet_name=SHEET_NAME_SUST
        )
        print("✔ MERGE SUST (justaposição, sem modificar nada) salvo em:", output_sust_all)
    else:
        print("Nenhum arquivo SUST encontrado.")

    if combined_frames:
        # Sem ordenação também
        df_combined_all = pd.concat(combined_frames, ignore_index=True, sort=False)

        df_combined_all.to_excel(
            output_combined_all, index=False, sheet_name=SHEET_NAME_COMBINED
        )
        print("✔ MERGE SUST6+LIBSrest (join por QR-Code) salvo em:", output_combined_all)
    else:
        print("Nenhum par SUST+LIBS encontrado para gerar o combinado.")


if __name__ == "__main__":
    main()
