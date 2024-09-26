import openpyxl
import pyperclip
import pyautogui
#entra na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produto = workbook['Produtos']
#copiar informação de um campo e colar no seu campo correspondente 
for linha in sheet_produto.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(809,247,duration=1)
    pyautogui.hotkey('ctrl','v')
    Descricao = linha[1].value
    pyperclip.copy(Descricao)
    pyautogui.click(815,345,duration=1)
    pyautogui.hotkey('ctrl','v')
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(818,470,duration=1)
    pyautogui.hotkey('ctrl','v')
    codigo = linha[3].value
    pyperclip.copy(codigo)
    pyautogui.click(822,552,duration=1)
    pyautogui.hotkey('ctrl','v')
    peso_kg_ = linha[4].value
    pyperclip.copy(peso_kg_)
    pyautogui.click(824,644,duration=1)
    pyautogui.hotkey('ctrl','v')
    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(799,721,duration=1)
    pyautogui.hotkey('ctrl','v')
    #Clique para proximo
    pyautogui.click(819,784,duration=1)
    #copiar informação de um campo e colar no seu campo correspondente
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(800,271,duration=1)
    pyautogui.hotkey('ctrl','v')
    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.click(799,358,duration=1)
    pyautogui.hotkey('ctrl','v')
    datavalidade = linha[8].value
    pyperclip.copy(datavalidade)
    pyautogui.click(810,449,duration=1)
    pyautogui.hotkey('ctrl','v')
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(822,524,duration=1)
    pyautogui.hotkey('ctrl','v')
    tamanho = linha[10].value
    if tamanho == 'Pequeno':
         pyautogui.click(831,617,duration=1)
         pyautogui.click(826,640,duration=1)
         
    elif tamanho == 'Médio':
       pyautogui.click(820,617,duration=1)
       pyautogui.click(816,676,duration=1)
        
    else:
       pyautogui.click(804,617,duration=1)
       pyautogui.click(801,71160,duration=1)
    
    material= linha[11].value
    pyperclip.copy(material)
    pyautogui.click(828,697,duration=1)
    pyautogui.hotkey('ctrl','v')
     #Clique para proximo
    pyautogui.click(826,764,duration=1)
#copiar informação de um campo e colar no seu campo correspondente
    fabricante= linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(815,287,duration=1)
    pyautogui.hotkey('ctrl','v')
    pais_origem= linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(810,374,duration=1)
    pyautogui.hotkey('ctrl','v')
    observa= linha[14].value
    pyperclip.copy(observa)
    pyautogui.click(791,486,duration=1)
    pyautogui.hotkey('ctrl','v')
    codigo_barra= linha[15].value
    pyperclip.copy(codigo_barra)
    pyautogui.click(822,597,duration=1)
    pyautogui.hotkey('ctrl','v')
    local_armazem= linha[16].value
    pyperclip.copy(local_armazem)
    pyautogui.click(820,684,duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.click(822,740,duration=1)
    pyautogui.click(1196,192,duration=1)
    pyautogui.click(1016,520,duration=1)

    
#repetir esses passos para outros campos ate preencher campor daquela pagina
#clicar em proximo
#repetir os mesmo passos e ir para a proxima pagina (pagina2)
#repetir os mesmos passos e finalizar o cadastro daquele produtor e clicar em conclur
#clicar em ok, para finalizar o processo
#clicar no ok mais um vez na mensagem de configuração e salvamento no banco de dados
#clicar em adicionar mais um e repetir o processo ate finalizar a planilha

#Pyautogui ( automação de clicks e teclado)Super Bolsa
#Openpyxl (leitura e automação de planilhas)

