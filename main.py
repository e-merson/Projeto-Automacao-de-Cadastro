# -Entrar na Planilha
import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']


# -copiar informação de um campo e passar para o correspondente

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1163,325, duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1176,445, duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1194,598, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1171,700, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1167,811, duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1174,915, duration=1)
    pyautogui.hotkey('ctrl','v')

    # Próxima página

    pyautogui.click(1157,994, duration=1)
    sleep(3)


    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1151,357, duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(1157,453, duration=1)
    pyautogui.hotkey('ctrl','v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(1168,567, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1171,673, duration=1)
    pyautogui.hotkey('ctrl','v')

    # botão seleção

    tamanho = linha[10].value
    pyautogui.click(1216,783, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1178,819, duration=1)        
    elif tamanho == 'Médio':
        pyautogui.click(1171,848, duration=1)        
    else:
        pyautogui.click(1165,873, duration=1)
        
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1220,890, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1162,961, duration=1)
    sleep(3)
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1161,377, duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_de_origem = linha[13].value
    pyperclip.copy(pais_de_origem)
    pyautogui.click(1182,479, duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1205,604, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1200,755, duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1197,867, duration=1)
    pyautogui.hotkey('ctrl','v')

    #concluir
    pyautogui.click(1171,942, duration=1)
    #produto salvo
    pyautogui.click(1649,239, duration=1)
    # adicionar mais um produto
    pyautogui.click(1446,656, duration=1)
    sleep(3)
   

# -Repetir até preencher campos da página

# -Clicar em proxima

# -repetir até a pagina 2

# -repetir passos e finalizar

# -clicar em ok na mensagem de confirmação de salvamento

# -clicar em adicionar mais um e repetir o processo até finalizar

