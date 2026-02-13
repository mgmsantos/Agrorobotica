# %%

import pandas as pd
import os
from pathlib import Path
from natsort import natsorted

# %%

## EXTRAI AS INFORMAÇÕES DOS CLIENTES DE CADA OS E MESCLA EM UM SÓ EXCEL

caminho = Path(r'C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\fazer_entradas\2025')

lista = []

for arquivo in sorted(caminho.iterdir()):

    if arquivo.is_file() and arquivo.suffix == '.xlsx':

        caminho_completo = os.path.join(caminho, arquivo)

        try:
            df = pd.read_excel(caminho_completo, header = None)

            dados = {
                "ID": Path(arquivo).stem.split('-')[0].strip(),
                "Fazenda": str(df.iloc[8, -1]).strip(),
                "Razão Social": str(df.iloc[8, 1]).strip(),
                "Endereço": (
                    f"{str(df.iloc[9, 1]).strip()}, "
                    f"{str(df.iloc[10, 1]).strip()}, "
                    f"{str(df.iloc[11, 1]).strip()}, "
                    f"{str(df.iloc[12, 1]).strip()}"
                ),
                "CNPJ": str(df.iloc[9, -1]).strip()
            }

            lista.append(dados)

        except Exception as e:
            print(f'Erro: {e}')

df_final = pd.DataFrame(lista)

print(f'Relatório gerado:')
df_final["ID"] = natsorted(df_final["ID"])
df_final = df_final.sort_index()
df_final

# %%

## SALVAR DF COM TODOS OS ENDEREÇOS POR OS

df_final.to_excel(r'C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\enderecos\2025.xlsx',
                  index = False)

# %%

## CONCATENA TODOS OS DFS GERADOS NA PASTA ENDEREÇOS

caminho = Path(r'C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\enderecos')

dfs = []

for arquivo in sorted(caminho.iterdir()):

    if arquivo.is_file() and arquivo.suffix == '.xlsx':

        df = pd.read_excel(os.path.join(caminho, arquivo))
        dfs.append(df)

df_uniao = pd.concat(dfs, ignore_index = True)

df_uniao.to_excel(r'C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\enderecos\DADOS_COMPLETO.xlsx',
                  index = False)

