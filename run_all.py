import threading
import subprocess
import os
import time
from tqdm import tqdm
import shutil
import json
import tkinter as tk
from tkinter import messagebox
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import re

is_cupom = ""
# def com_cupom():
#     messagebox.showinfo("Cupom", "Com cupom!")
#     is_cupom = "True"
#     janela.destroy() 

# def sem_cupom():
#     messagebox.showinfo("Cupom", "Sem Cupom!")
#     is_cupom = ""
#     janela.destroy()  

# janela = tk.Tk()
# janela.title("Caixa de Pagamento")

# janela.geometry("300x150")

# botao_com_cupom = tk.Button(janela, text="Com Cupom", command=com_cupom)
# botao_com_cupom.pack(pady=20)

# botao_sem_cupom = tk.Button(janela, text="Sem Cupom", command=sem_cupom)
# botao_sem_cupom.pack(pady=20)

# janela.mainloop()

cookie = "_fbp=fb.1.1719583917979.99650778628702461; template=template-light; ai_user=PZ4M|2024-06-28T14:12:01.744Z; lang_nubi=pt_br; hubspotutk=2f2c25210fc5f4647166a6d443465c4a; intercom-id-re8wm274=8eebef63-1134-43bf-8e26-ed2750ec21a4; intercom-device-id-re8wm274=b2141e40-c318-4bd0-963a-cabaa4bda6d6; _ga=GA1.3.1410409649.1719583924; _gid=GA1.2.251311202.1721730974; _ce.clock_event=1; _ce.clock_data=2857%2C177.91.74.213%2C1%2C362d7fe3d8b2581bffa359f0eeda7106%2CChrome%2CBR; _clck=70t169%7C2%7Cfnq%7C0%7C1640; _hp2_ses_props.2056355555=%7B%22ts%22%3A1721820064786%2C%22d%22%3A%22app.nubimetrics.com%22%2C%22h%22%3A%22%2Faccount%2Flogin%22%2C%22q%22%3A%22%3FReturnUrl%3D%252fopportunity%252fcategoryDetail%22%2C%22g%22%3A%22%23%3Fcategory%3DMLB5672-MLB1747-MLB114675-MLB45905%22%7D; ARRAffinity=0a022efd41c53065d08c7e4e9af60a16a775eb774d83da8c0e7a061e9ac207c2; ARRAffinitySameSite=0a022efd41c53065d08c7e4e9af60a16a775eb774d83da8c0e7a061e9ac207c2; i18next=pt_br; _ce.irv=returning; cebs=1; __hstc=154116135.2f2c25210fc5f4647166a6d443465c4a.1719583931327.1721820071590.1721821502493.10; __hssrc=1; TiPMix=74.14588657906556; x-ms-routing-name=self; ASP.NET_SessionId=knixgfss2x5mtps3rrrc3nsl; _gcl_au=1.1.428859039.1719583922.1474109683.1721820092.1721826085; .ASPXAUTH=ABE18D75BD83080AE7843458456E1BB52B84CD6B7C010DA3CE2BD0CDFD25944884013A28B0C00184EDD8EDFD507154BE948BE7136BF0CE466158D3F1C2B0195CC72E9A524166A7144380061F96A97FE6A6AA843BEA266FEA6E322D9EF029BE44A80BBFF475963EFB1F34FDA3036302CFA751379AD3A2596B435898D1C051C655D25973872275A9F5CF9B9013385367CB1A3DD2D53175523A6F0CDA6565C56736C440E0DDF29D81B284DB81877AC9A6BC658EF47A0CA841EEA71FCF007F7627DFF88BFCB8E10D15E23E3E19D5B64BA799B6408DAB6A96DD47D8C30C3F67B834CAC1D849B7E04CFE1E9313A488903043AD311514F891F9818EA2D8E7B1B4D0AC65; _ga=GA1.1.1410409649.1719583924; _uetsid=6317c99048df11ef8cb4a1e37632fad0; _uetvid=d4f1f2b01ce711efb06041a93bf318c4; cebsp_=16; _clsk=7qkc9u%7C1721826210329%7C18%7C1%7Cp.clarity.ms%2Fcollect; __hssc=154116135.11.1721821502493; intercom-session-re8wm274=d2JQQmRuRWh2Nmt1c1NtWGQ5VGxnVlBiRHN3RXJ0MlBGbGJ4RUM1OWdhdzA2azVWNVUybGJiRCtzbVZYaEJCOC0tL0VQa09SeUtGS2QycVloc3JoUjY1dz09--92727e44354e2a92069709065b399c88c8aecffe; _hp2_id.2056355555=%7B%22userId%22%3A%22967345408458646%22%2C%22pageviewId%22%3A%228045139369766454%22%2C%22sessionId%22%3A%221508023483604914%22%2C%22identity%22%3Anull%2C%22trackerVersion%22%3A%224.0%22%7D; _ga_26N5SV28FF=GS1.1.1721825606.98.1.1721826263.60.0.0; _ga_X9JW5VPF68=GS1.1.1721817527.37.1.1721826295.60.0.0; _ga_1BD6V1LPWP=GS1.1.1721817526.10.1.1721826295.60.0.0; _ce.s=lcw~1721826267390~gtrk.la~lxyrv6su~v~14ab46a1534964ae49bb107549fe3a6703c9bf93~lva~1721821449958~vpv~7~v11.fhb~1721826207409~v11.lhb~1721826267388~v11slnt~1720611997725~v11.cs~229172~v11.s~b8ba64f0-49bb-11ef-b737-51a0d327d132~v11.sla~1721826295326~v11.send~1721826259596~lcw~1721826295327; sc_is_visitor_unique=rx12923916.1721826297.B2D94BDE701C4FAF4F4B3CFE0A2A9D73.9.8.8.7.6.6.5.5.1; ai_session=vRNpN|1721817526508|1721826297624.2"


pasta = r"./dados"

# Verificar se a pasta existe
if os.path.exists(pasta):
    # Iterar sobre todos os itens na pasta
    for item in os.listdir(pasta):
        item_path = os.path.join(pasta, item)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
                print(f"Pasta removida: {item_path}")
        except Exception as e:
            print(f"Erro ao remover {item_path}: {e}")
else:
    print(f"A pasta {pasta} não existe.")
    
if os.path.exists(r"./dados_extraidos.docx"):
    os.remove(r"./dados_extraidos.docx")
    
def run_command(args):
    subprocess.run(args)


commands = [
    ["python", "rodar.py", "EQUALIZADOR PARA BANCO DE BATERIAS", is_cupom],
    # ["python", "rodar.py", "FONTE NOBREAK 12V/8A", is_cupom],
    # ["python", "rodar.py", "FONTE NOBREAK 24V/6A", is_cupom],
    # ["python", "rodar.py", "FONTE NOBREAK -48V 15A 15A", is_cupom],
    # ["python", "rodar.py", "FONTE NOBREAK -48V 30A 15A", is_cupom],
    # ["python", "rodar.py", "FONTE NOBREAK -48V 40A 10A", is_cupom],
    # ["python", "rodar.py", "INVERSOR OFF GRID SENOIDAL PURA JFA 1000W 48V/220V RACK", is_cupom],
    # ["python", "rodar.py", "INVERSOR OFF GRID SENOIDAL PURA JFA 3000W 48/220V C/ GER RACK", is_cupom],
    # ["python", "rodar.py", "INVERSOR OFF GRID SENOIDAL PURA JFA 5000W 48/220V C/ GER RACK", is_cupom],
]


threads = []


for cmd in commands:
    thread = threading.Thread(target=run_command, args=(cmd,))
    thread.start()
    threads.append(thread)


for thread in tqdm(threads):
    thread.join()

subprocess.run(["python", r"./ordenar.py"])


#https://app.nubimetrics.com/api/search/items?site_id=MLB&to_search=fonte%20120a%20bob%20jfa&buying_mode=buy_it_now&limit=50&offset=0&attributes=results,available_filters,paging,filters&seller_id=1242763049&order=relevance&typeSearch=q&search_filters=condition=new@&exportData=true&language=pt_BR&pvp=499&isControlPrice=true