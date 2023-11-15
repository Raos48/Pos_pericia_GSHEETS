import pyautogui
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import time
import subprocess
#teste GIT PC DESKTOP
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
# SAMPLE_SPREADSHEET_ID = "1kBX7EUe9SnVP98_dvT8ixB-mh3jlbF2ziVpBWITcfwM"
SAMPLE_SPREADSHEET_ID = "17vIYdlF36Uc54Lgza4ilJFMhNbb7IIc3BMV3weNVj2Y"
SAMPLE_RANGE_NAME = "Página1!A2:V8381"

creds = None

if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
# If there are no (valid) credentials available, let the user log in.

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            "client_secret.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
        token.write(creds.to_json())

service = build("sheets", "v4", credentials=creds)
sheet = service.spreadsheets()

caminho_programa = r"C:\Program Files\SABI\Controle Operacional\Controle.exe"
processo_controle = None
nr_requerimento = None
ID = None
dut = None
status = None
arquivo = "Pendencias"
tratar_dut = True

# Abrir SABI
subprocess.call(['start', '', caminho_programa], shell=True)
time.sleep(5)
try:
    result = pyautogui.locateOnScreen("1/imagem28_sabi_ativo.png", grayscale=True, confidence=0.9)
    time.sleep(2)
    print("SABI já encontra-se aberto!")
    x, y = pyautogui.locateCenterOnScreen('1/imagem28_sabi_ativo.png', grayscale=False, confidence=0.8)
    pyautogui.moveTo(x, y, duration=0.5)
    pyautogui.click(x, y)
except:
    pyautogui.write('go2035843')
    pyautogui.press('tab')
    pyautogui.write('tema4512')
    pyautogui.press('enter')
    time.sleep(15)
    # ====fechar===================
    x, y = pyautogui.locateCenterOnScreen('1/imagem1_fechar_.png', grayscale=False, confidence=0.8)
    pyautogui.moveTo(x, y, duration=0.5)
    pyautogui.click(x, y)
    # ==============================
    time.sleep(3)
    pyautogui.keyDown('alt')
    pyautogui.press('m')
    pyautogui.keyUp('alt')
    pyautogui.press('enter')
    time.sleep(3)

print()
print("   _____         ____ _____  ")
print("  / ____|  /\   |  _ \_   _| ")
print(" | (___   /  \  | |_) || |   ")
print("  \___ \ / /\ \ |  _ < | |   ")
print("  ____) / ____ \| |_) || |_  ")
print(" |_____/_/    \_\____/_____| ")
print()

linha = int(input("Informe o numero da linha manualmente e ENTER para continuar: "))

while True:
    print()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!B{linha}").execute()
    values = result.get('values', [])
    nr_requerimento = values[0][0]
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}").execute()
    values = result.get('values', [])
    sit = values
    print(linha, nr_requerimento, sit)

    # Clicar em Pesquiar ( binóculus )
    x, y = pyautogui.locateCenterOnScreen('1/imagem2_pesquisar.png', grayscale=False, confidence=0.8)
    pyautogui.moveTo(x, y, duration=0.1)
    pyautogui.click(x, y)
    print("Clica em pesquisar >>")
    time.sleep(2)
    # Clicar em Pesquiar ( binóculus )
    x, y = pyautogui.locateCenterOnScreen('1/imagem2_pesquisar.png', grayscale=False, confidence=0.8)
    pyautogui.moveTo(x, y, duration=0.1)
    pyautogui.click(x, y)
    print("Clica em pesquisar >>")
    time.sleep(2)

    tentativas = 0  # aguardando===========================================
    while True:
        try:
            result = pyautogui.locateOnScreen("1/imagem3_sim.png", grayscale=True, confidence=0.9)
            break
        except Exception as e:
            tentativas += 1
            print('wait 1...')
            if tentativas >= 50:  # Verifica se atingiu 50 tentativas
                print('Limite de tentativas atingido. >>')
                break
            time.sleep(0.5)
            continue
    # aguardando===========================================

    tentativas = 30
    encontrado = False
    for _ in range(tentativas):
        element_location = pyautogui.locateCenterOnScreen('1/imagem3_sim.png', grayscale=True, confidence=0.9)
        if element_location:
            x, y = pyautogui.locateCenterOnScreen('1/imagem3_sim.png', grayscale=False, confidence=0.8)
            pyautogui.moveTo(x, y, duration=0.1)
            pyautogui.click(x, y)
            print("Clica Sim >>")
            encontrado = True
            break
        else:
            print("Botão Sim não localizado >>")
            time.sleep(1)
    if not encontrado:
        print("Botão 'Sim' não encontrado após 30 tentativas.")
        break

    # digitar=====================
    time.sleep(2)
    x, y = pyautogui.locateCenterOnScreen('1/imagem4_campo_requerimento.png', grayscale=False, confidence=0.8)
    pyautogui.moveTo(x, y + 8, duration=0.1)
    pyautogui.click(x + 10, y)
    pyautogui.press('backspace', presses=20)
    pyautogui.write(str(nr_requerimento))
    print("Digita NR requerimento >>")

    # pesquisar e aguardar===========================================
    x, y = pyautogui.locateCenterOnScreen('1/imagem2_pesquisar.png', grayscale=False, confidence=0.8)
    pyautogui.moveTo(x, y, duration=0.5)
    pyautogui.click(x, y)
    print("Clica pesquisar >>")

    while True:  # aguardando===========================================
        try:
            result = pyautogui.locateOnScreen("1/imagem25_aguarda.png", grayscale=True, confidence=0.9)
            break
        except:
            print("wait 2... >>")
            time.sleep(0.5)
            continue
    time.sleep(5)  # aguardando===========================================

    try:  # verificar situação=================================================
        result = pyautogui.locateOnScreen("1/imagem27_critica2.png", grayscale=True, confidence=0.9)
        if result:
            valores_adicionar = [["Crítica 2"]]
            body = {"values": valores_adicionar}
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                           valueInputOption="USER_ENTERED", body=body).execute()
            print("Crítica 2 >>")
            linha += 1
            continue
    except:
        pass

    try:  # verificar situação=================================================
        print("Analisando situação... >>")
        result = pyautogui.locateOnScreen("1/imagem5_req_deferido.png", grayscale=True, confidence=0.9)
        if result:
            print("Analisando situação - Deferido!")
            print("Fim!")
            valores_adicionar = [["Deferido"]]
            body = {"values": valores_adicionar}
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                           valueInputOption="USER_ENTERED", body=body).execute()
            print("Deferido")
            linha += 1
            continue
    except:
        pass  ################################################################

    try:  # verificar situação===============================================
        result = pyautogui.locateOnScreen("1/imagem7_req_indeferido.png", grayscale=True, confidence=0.9)
        if result:
            print("Analisando situação - Indeferido!")
            print("Fim!")
            valores_adicionar = [["Indeferido"]]
            body = {"values": valores_adicionar}
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                           valueInputOption="USER_ENTERED", body=body).execute()
            print("Indeferido")
            linha += 1
            continue
    except:
        pass  ################################################################

    try:  # verificar situação=============================================
        result = pyautogui.locateOnScreen("1/imagem23_req_cancelado.png", grayscale=True, confidence=0.9)
        if result:
            print("Analisando situação - Cancelado!")
            print("Fim!")
            valores_adicionar = [["Requerimento Cancelado"]]
            body = {"values": valores_adicionar}
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                           valueInputOption="USER_ENTERED", body=body).execute()
            print("Requerimento Cancelado")
            linha += 1
            continue
    except:
        pass  ################################################################


    time.sleep(10)
    # clicar em serviço
    x, y = pyautogui.locateCenterOnScreen('1/imagem10_servico.png', grayscale=True, confidence=0.7)
    pyautogui.moveTo(x, y, duration=0.5)
    time.sleep(3)

    tentativas = 30  ################################################################
    encontrado = False
    for _ in range(tentativas):
        print("Move mouse até elemento serviço >>")
        x, y = pyautogui.locateCenterOnScreen('1/imagem10_servico.png', grayscale=True, confidence=0.7)
        pyautogui.moveTo(x, y, duration=0.5)
        time.sleep(3)
        element_location = pyautogui.locateCenterOnScreen('1/imagem10_servico.png', grayscale=True, confidence=0.7)
        if element_location:
            time.sleep(3)
            print("Elemento Serviço encontrado.")
            x, y = pyautogui.locateCenterOnScreen('1/imagem24_servico_menor.png', grayscale=True, confidence=0.9)
            if (x, y) != (None, None):
                pyautogui.moveTo(x, y, duration=0.5)
                pyautogui.click(x, y)
                print("Elemento Serviço Menor encontrado e clicado.")
                encontrado = True
                break
            else:
                print("Elemento Serviço Menor não encontrado após encontrar Serviço.")
                break
        else:
            print("Elemento Serviço não encontrado. Tentando novamente. >>")
            time.sleep(1)  # Aguarda 1 segundo antes da próxima tentativa

    if not encontrado:
        print("Elemento Serviço não encontrado após 30 tentativas.")
        break

    # # clicar em serviço
    # x, y = pyautogui.locateCenterOnScreen('1/imagem24_servico_menor.png', grayscale=True, confidence=0.9)
    # pyautogui.moveTo(x, y, duration=0.5)
    # pyautogui.click(x, y)  ################################################################
    # print("Clica elemento Serviço")

    time.sleep(2)
    try:  # clicar em tratar DUT====================================================
        x, y = pyautogui.locateCenterOnScreen('1/imagem11_tratar_dut.png', grayscale=False, confidence=0.8)
        print("Clica menu Tratar DUT >>")
        pyautogui.moveTo(x, y, duration=0.1)
        pyautogui.click(x, y)
    except:
        print("Pendente de Tratamento - DUT inibida!")
        print("Fim!")
        valores_adicionar = [["Pendente de Tratamento - DUT inibida >>"]]  # DUT INIBIDA====================
        body = {"values": valores_adicionar}
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                       valueInputOption="USER_ENTERED", body=body).execute()
        linha += 1
        continue  ################################################################

    try:  # ERRO=====================
        print("Verifica existência da Tela: Erro DUT >>")
        result = pyautogui.locateOnScreen("1/imagem12_erro.png", grayscale=True, confidence=0.9)
        if result:
            print("Erro.. DUT")
            x, y = pyautogui.locateCenterOnScreen('1/imagem12_erro.png', grayscale=False, confidence=0.8)
            pyautogui.moveTo(x, y, duration=0.2)
            pyautogui.click(x, y)
    except:
        pass  ################################################################

    time.sleep(3)
    try:
        print("Verifica janela DUT aberta >>")
        result = pyautogui.locateOnScreen("1/imagem_26_campo_dut.png", grayscale=True, confidence=0.9)
        if result:
            try:  # verifica se deu erro============
                x, y = pyautogui.locateCenterOnScreen('1/imagem13_ok_erro.png', grayscale=False, confidence=0.8)
                pyautogui.moveTo(x, y, duration=0.5)
                pyautogui.click(x, y)
                print("Erro ao lançar a DUT!")
                print("Fim")
                valores_adicionar = [["Erro - consta exigencia"]]
                body = {"values": valores_adicionar}
                result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                               valueInputOption="USER_ENTERED", body=body).execute()
                linha += 1
                continue
            except:
                print("Avançar >>")
                pass
            # verifica se deu erro============
    except:
        print("Campo inibido")
        print("Fim")
        valores_adicionar = [["Campo inibido"]]
        body = {"values": valores_adicionar}
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                       valueInputOption="USER_ENTERED", body=body).execute()
        linha += 1
        continue

    # lançamento da DUT===================================
    element_location = pyautogui.locateCenterOnScreen('1/imagem_14_campo_dut.png', grayscale=True, confidence=0.9)
    if element_location:

        print("Verificando se possui DUT informada >>")
        try:  # não possui DUT===========================================
            result = pyautogui.locateOnScreen("1/imagem15_sem_dut.png", grayscale=True, confidence=0.9)
            if result:
                print("Campo DUT Vazio!")
                print("Fim")
                x, y = pyautogui.locateCenterOnScreen('1/imagem16_cancelar.png', grayscale=False, confidence=0.8)
                pyautogui.moveTo(x, y, duration=0.1)
                pyautogui.click(x, y)
                valores_adicionar = [["Não possui DUT"]]
                body = {"values": valores_adicionar}
                result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                               valueInputOption="USER_ENTERED", body=body).execute()
                print(valores_adicionar)
                linha += 1
                continue
        except:
            pass
        # não possui DUT===========================================

        print("Radius não >>")
        x, y = pyautogui.locateCenterOnScreen('1/imagem17_radius_nao.png', grayscale=False, confidence=0.8)
        pyautogui.moveTo(x, y, duration=0.2)
        pyautogui.click(x, y)

        print("Confirma >>")
        x, y = pyautogui.locateCenterOnScreen('1/imagem19_confirmar.png', grayscale=False, confidence=0.8)
        pyautogui.moveTo(x, y, duration=0.2)
        pyautogui.click(x, y)

        time.sleep(3)
        print("Sim Confirma >>")
        x, y = pyautogui.locateCenterOnScreen('1/imagem_26_sim_confirmar.png', grayscale=False, confidence=0.8)
        pyautogui.moveTo(x, y, duration=0.2)
        pyautogui.click(x, y)
        time.sleep(5)

        try:  # DUT MAIOR DER==============
            print("Verifica Erro - DUT maior DER >>")
            result = pyautogui.locateOnScreen("1/campo20_dut_maior_der.png", grayscale=True, confidence=0.9)
            if result:
                print("DUT maior DER!")
                print("Fim")
                valores_adicionar = [["DUT maior que a DER"]]
                body = {"values": valores_adicionar}
                result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                               valueInputOption="USER_ENTERED", body=body).execute()
                x, y = pyautogui.locateCenterOnScreen('1/campo20_dut_maior_der.png', grayscale=False, confidence=0.8)
                pyautogui.moveTo(x, y, duration=0.2)
                pyautogui.click(x, y)
                time.sleep(3)
                x, y = pyautogui.locateCenterOnScreen('1/imagem21_cancelar.png', grayscale=False, confidence=0.8)
                pyautogui.moveTo(x, y, duration=0.2)
                pyautogui.click(x, y)
                time.sleep(3)
        except:
            pass

        # atualizar 3 vezes==================================================
        #         for i in range(3):
        #             time.sleep(5)
        #             try:
        #                 print("clicar em atualizar")
        #                 x, y = pyautogui.locateCenterOnScreen('1/imagem22_atualizar.png', grayscale=False, confidence=0.8)
        #                 pyautogui.moveTo(x, y, duration=0.5)
        #                 pyautogui.click(x, y)
        #             except:
        #                 time.sleep(2)

        print("Aguarda atualizar status >>")
        time.sleep(5)

        try:  # verificar situação=============================================
            result = pyautogui.locateOnScreen("1/imagem5_req_deferido.png", grayscale=True, confidence=0.9)
            if result:
                valores_adicionar = [["Deferido"]]
                body = {"values": valores_adicionar}
                result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                               valueInputOption="USER_ENTERED", body=body).execute()
                print("Deferido!")
        except:
            pass

        print("Verifica retorno STATUS. >>")
        try:  # ===============================================================
            result = pyautogui.locateOnScreen("1/imagem7_req_indeferido.png", grayscale=True, confidence=0.9)
            if result:
                valores_adicionar = [["Indeferido"]]
                body = {"values": valores_adicionar}
                result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                               valueInputOption="USER_ENTERED", body=body).execute()
                print("Indeferido!")
        except:
            pass

        try:  # ===============================================================
            result = pyautogui.locateOnScreen("1/imagem8_req_pendente_tratamento.png", grayscale=True, confidence=0.9)
            if result:
                valores_adicionar = [["Pendente Tratamento"]]
                body = {"values": valores_adicionar}
                result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                               valueInputOption="USER_ENTERED", body=body).execute()
                print("Pendente Tratamento!")
        except:
            pass
    else:
        print("Janela DUT não localizada!")
        print("Campo inibido")
        print("Fim")
        valores_adicionar = [["Pendente Tratamento - Campo DUT inibido"]]
        body = {"values": valores_adicionar}
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=f"Página1!V{linha}",
                                       valueInputOption="USER_ENTERED", body=body).execute()
        linha += 1
        continue

    print("Final!")
    linha += 1
