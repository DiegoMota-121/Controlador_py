import openpyxl
import pyperclip
import pyautogui
# importarei abaixo um intervalo conforme comentário 014
from time import sleep

# https://www.youtube.com/watch?v=UtkPIpov6h8&list=PLnNURxKyyLIJ5ftIIYFLNNLyCmBx5uXYM&index=2 local do curso
# https://cadastro-produtos-devaprender.netlify.app/index.html
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(429, 345, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(298, 420, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Categoria = linha[2].value
    pyperclip.copy(Categoria)
    pyautogui.click(428, 557, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Codigo_produto = linha[3].value
    pyperclip.copy(Codigo_produto)
    pyautogui.click(302, 641, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Peso = linha[4].value
    pyperclip.copy(Peso)
    pyautogui.click(302, 726, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Dimensoes = linha[5].value
    pyperclip.copy(Dimensoes)
    pyautogui.click(306, 808, duration=1)
    pyautogui.hotkey('ctrl', 'v')

   # COMMITE 014 após realizar todos os click ELE vai para outra aba da automação, sendo assim ele tem um intervalo para realizar o ajuste da próxima ETAPA
    pyautogui.click(320, 872, duration=1)
    sleep(3)

    Preco = linha[6].value
    pyperclip.copy(Preco)
    pyautogui.click(300, 358, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Qtd_em_estoque = linha[7].value
    pyperclip.copy(Qtd_em_estoque)
    pyautogui.click(306, 444, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.click(302, 530, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Cor = linha[9].value
    pyperclip.copy(Cor)
    pyautogui.click(324, 622, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # NO TAMANHA IREI COLOCAR UMA LOGICA DE PROGRAMAÇÃO, POIS NA PLAN DO CLIENTE TEM UMA CAIXA DE SELEÇÃO
    # aplicando uma condicional
    Tamanho = linha[10].value
    pyautogui.click(291, 706, duration=1)

    if Tamanho == 'Pequeno':
        pyautogui.click(325, 732, duration=1)
    elif Tamanho == 'médio':
        pyautogui.click(331, 754, duration=1)
    else:
        pyautogui.click(305, 782, duration=1)
    # LER INFORMAÇÃO DA PLAN
    # SE FOR P,M,G, TRES LAÇO)

    Material = linha[11].value
    pyperclip.copy(Material)
    pyautogui.click(291, 793, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(300, 849, duration=1)
    sleep(3)

    Fabricante = linha[12].value
    pyperclip.copy(Fabricante)
    pyautogui.click(367, 376, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    origem = linha[13].value
    pyperclip.copy(origem)
    pyautogui.click(360, 465, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Observacoes = linha[14].value
    pyperclip.copy(Observacoes)
    pyautogui.click(364, 545, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Cod_de_barras = linha[15].value
    pyperclip.copy(Cod_de_barras)
    pyautogui.click(367, 681, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    Localizacao = linha[16].value
    pyperclip.copy(Localizacao)
    pyautogui.click(325, 765, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(329, 820, duration=1)
    sleep(3)

    pyautogui.click(320, 834, duration=2)
    pyautogui.click(916, 187, duration=2)
    pyautogui.click(909, 191, duration=2)
    pyautogui.click(698, 606, duration=2)
