
### AUTOMAÇÃO - AJUSTAR PLANILHAS

# %%

import pyautogui
import os
import time
import pyperclip

# %%
ORIGEM = r"C:\Users\migue\OneDrive - Agrorobotica Fotonica Em Certificacoes Agroambientais\_Fertilidade\ENTRADAS\ENTRADA_TESTE"

# Aumentando o intervalo para garantir que o Excel processe cada comando
pyautogui.PAUSE = 0.8

arquivos = [f for f in os.listdir(ORIGEM) if f.endswith(".xlsx")]

def ajustar(caminho):
    print(f"Abrindo: {os.path.basename(caminho)}")
    os.startfile(caminho)

    print(f"ATENÇÃO: NÃO MECHA NO TECLADO NEM NO MOUSE")
    
    # Tempo generoso para o Office 365 carregar e sincronizar com OneDrive
    time.sleep(15) 

    # 1. NAVEGAR PARA A SEGUNDA ABA
    # Garantir que começa da primeira
    for _ in range(5):
        pyautogui.hotkey("ctrl", "pgup")
    # Vai para a segunda
    pyautogui.hotkey("ctrl", "pgdn")
    time.sleep(1.5)

    # 2. CRIAR NOVA COLUNA (via B8 para evitar títulos mesclados)
    pyautogui.press('f5')
    pyautogui.write('B8')
    pyautogui.press('enter')
    time.sleep(0.5)
    
    pyautogui.hotkey('ctrl', 'shift', '+')
    time.sleep(0.8)
    pyautogui.press('c') # 'c' seleciona "Coluna Inteira" no Excel PT-BR
    pyautogui.press('enter')
    time.sleep(2) # Espera o Excel processar o deslocamento de colunas

    # 3. CABEÇALHO B8: "Talhao Agro"
    pyautogui.press('f5')
    pyautogui.write('B8')
    pyautogui.press('enter')
    pyperclip.copy("Talhao Agro")
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    # 4. INICIAR O ASTERISCO NA B9
    pyautogui.press('f5')
    pyautogui.write('B9')
    pyautogui.press('enter')
    pyperclip.copy("*")
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # 5. SELEÇÃO PRECISA E PREENCHIMENTO
    # Passo A: Ir para a coluna C (que tem os dados) e achar a última linha
    pyautogui.press('f5')
    pyautogui.write('C9')
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'down') # Pula para a última linha da tabela
    time.sleep(0.5)
    
    # Passo B: Voltar para a coluna B (estamos agora na última linha da B)
    pyautogui.press('left') 
    time.sleep(0.5)
    
    # Passo C: O PULO DO GATO - Selecionar até a B9 usando F5 + Shift
    pyautogui.press('f5')
    pyautogui.write('B9')
    # Shift + Enter no menu "Ir Para" seleciona do ponto atual até o destino
    pyautogui.hotkey('shift', 'enter') 
    time.sleep(0.8)

    # Passo D: Preencher a seleção com o asterisco que está na B9
    pyautogui.hotkey('ctrl', 'd') 
    time.sleep(1)

    # 6. CABEÇALHO C8: "Talhao Comercial"
    pyautogui.press('f5')
    pyautogui.write('C8')
    pyautogui.press('enter')
    pyperclip.copy("Talhao Comercial")
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # 7. SALVAR E FECHAR
    pyautogui.hotkey('ctrl', 'b') # Salvar
    print(f"Salvando {os.path.basename(caminho)}...")
    time.sleep(5) # Tempo extra para o OneDrive processar o salvamento
    pyautogui.hotkey('alt', 'f4') 
    time.sleep(3)

# --- EXECUÇÃO ---
if not arquivos:
    print("Nenhum arquivo encontrado na pasta especificada.")
else:
    for arquivo in arquivos:
        try:
            ajustar(os.path.join(ORIGEM, arquivo))
        except Exception as e:
            print(f"Erro ao processar {arquivo}: {e}")

print("Automação concluída!")