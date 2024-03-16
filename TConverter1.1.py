import PySimpleGUI as sg
import time
import pyautogui as py
import pygetwindow as gw

def set_theme():
    theme_dict = {
        'BACKGROUND': '#FFFFFF',
        'TEXT': '#4D148C',
        'INPUT': '#F2EFE8',
        'TEXT_INPUT': '#000000',
        'SCROLL': '#F2EFE8',
        'BUTTON': ('#000000', '#C2D4D8'),
        'PROGRESS': ('#FFFFFF', '#C7D5E0'),
        'BORDER': 1, 'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 0
    }
    
    sg.LOOK_AND_FEEL_TABLE['Conversor de Terminais'] = theme_dict
    sg.theme('Conversor de Terminais')

def create_layout():
    
    BPAD_INPUTS =((40,0), (0,0))
    BPAD_SIGNATURE = ((100,0), (20,0))

    layout = [
        [sg.Image(filename="image\Logo_with_robot.png")],
        [
            [sg.Text("Insira as informações abaixo:", justification='c', font="Any 10")],
            [sg.Text("Part_Number Origem:      ", pad=BPAD_INPUTS),
             sg.InputText(pad=BPAD_INPUTS, size=(30, 11))],
            [sg.Text("Part_Number Destino:     ", pad=BPAD_INPUTS),
             sg.InputText(pad=BPAD_INPUTS, size=(30, 10))],
            [sg.Text("Quantidade de Terminais:", pad=BPAD_INPUTS),
             sg.Input(pad=BPAD_INPUTS, size=(30, 10))],
            [sg.Button("Execute", key="execute", pad=BPAD_INPUTS),
             sg.Button("Cancelar", key="cancel", pad=BPAD_INPUTS)]
        ],
        [
            [sg.Text("Progresso:", justification='c', font="Any 10")],
            [sg.Text("====>"), sg.Text("", key="percent")],

            [sg.Text("Analise de Tempo:", justification='c', font="Any 10")],
            [sg.Text("Total Terminais:"), sg.Text("", key="Total_Terminais")],
            [sg.Text("Tempo_por_Terminal:"), sg.Text("", key="Tempo_por_Terminal")],
            [sg.Text("Tempo Restante:"), sg.Text("", key="Tempo_Restante")],
            [sg.Text("------------Copyright © Robson Carlos - Developed on 2023------------", font=("Helvetica", 8), pad=BPAD_SIGNATURE)]
        ],
    ]

    return layout

def execute_conversion(repeticoes, de, para, window):
    EXCEL_WINDOW = 'template_conversor_python.xlsx - Excel'
    RLOG_DISPLAY = 'Atualizar ESN - Google Chrome'
    CONTADOR = 1
    start_time = time.time()
    loop = repeticoes-1

    try:
        left = gw.getWindowsWithTitle(EXCEL_WINDOW)[0]
    except IndexError:
        sg.popup("Error:", "Janela do Excel não foi localizada, abra e reinicie o programa")
        return
    
    left.activate()
    left.maximize()
    left.resizeTo(800, 600)
    left.moveTo(10,10)
    time.sleep(2)
    py.click(40, 195)
    time.sleep(1)
    py.write('c2')
    py.press('enter')

    try:
        right = gw.getWindowsWithTitle(RLOG_DISPLAY)[0]
    except IndexError:
        sg.popup("Error:", "O Google Chrome ou a Aba do RLOG Atualizar ESN não foi localizada, abra e reinicie o programa")
        return
    right.activate()
    right.maximize()
   
    py.hotkey('win', 'right')
    time.sleep(1)
    py.press('esc')
    py.click(1416, 357)
    py.press('f', presses=2)
    py.press('enter')
    time.sleep(2)
    py.click(1209, 379)
    time.sleep(1)
    py.write(de)
    #time.sleep(1)
    py.press("Enter")
    py.click(771, 122)
    #time.sleep(1)
    py.hotkey("ctrl", "c")
    #time.sleep(1)
    py.click(1191, 401)
    py.hotkey("ctrl", "v")
    #time.sleep(1)
    py.click(1209, 422)
    py.write(para)
    #time.sleep(1)
    py.press("Enter")
    py.click(1184, 465)
    time.sleep(1)
    py.press("Enter")
    time.sleep(1)

    tempo_execucao = time.time() - start_time
    repeticoes_restante = loop - (CONTADOR-1)
    tempo_restante = f"{(tempo_execucao * repeticoes_restante):.2f}segundos"
    percentual = (repeticoes_restante / (loop+1)) * 100
    percentual_final = f"{100 - percentual:.2f}%"
    total_terminais = f"{CONTADOR}/{(loop+1)}"
    tempo_por_terminal = f"{tempo_execucao:.2f} segundos"

    window["percent"].update(percentual_final)
    window["Total_Terminais"].update(total_terminais)
    window["Tempo_por_Terminal"].update(tempo_por_terminal)
    window["Tempo_Restante"].update(tempo_restante)
    window.refresh()
    

    while CONTADOR <= loop:
        
        time.sleep(2)
        py.click(1209, 379)
        py.write(de)
        #time.sleep(1)
        py.press("Enter")
        py.click(1191, 401)
        py.hotkey("ctrl", "a")
        py.press("Del")
        py.click(771, 122)
        py.press("Esc")
        py.press("Down")
        py.hotkey("ctrl", "c")
        #time.sleep(1)
        py.click(1191, 401)
        #time.sleep(1)
        py.hotkey("ctrl", "v")
        #time.sleep(1)
        py.click(1209, 422)
        py.write(para)
        #time.sleep(1)
        py.press("Enter")
        py.click(1184, 465)
        time.sleep(1)
        py.press("Enter")
        time.sleep(1)

        CONTADOR +=1
        
        repeticoes_restante = loop - (CONTADOR-1)
        percentual = (repeticoes_restante / (loop+1)) * 100
        percentual_final = f"{100 - percentual:.2f}%"
        total_terminais = f"{CONTADOR}/{(loop+1)}"
        tempo_por_terminal = f"{tempo_execucao:.2f} segundos"

        window["percent"].update(percentual_final)
        window["Total_Terminais"].update(total_terminais)
        window.refresh()

    sg.popup("Alert:", "   O Processo foi finalizado")

def main():

    set_theme()

    layout = create_layout()
    window = sg.Window("TConverter", layout, location=(-10, 390))


    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "cancel":
            break

        if event == "execute":

            de = values[1].upper()
            para = values[2].upper()
            repeticoes = int(values[3])
            execute_conversion(repeticoes, de, para, window)

    

if __name__ == "__main__":
    main()
