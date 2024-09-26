import keyboard
from lxml import html
from collections import defaultdict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from datetime import datetime
import pygetwindow as gw
import pyautogui
import time
import re
import json
import requests
import pandas
from tqdm import tqdm

df = pandas.read_excel("PoliticadePrecos_Telecom.xlsx", engine='openpyxl', sheet_name="Sugestão de Preços", skiprows=1)

for index, i in df.iterrows():
    if i['CONSTRUÇÃO DE CUSTO '] == "EQUALIZADOR PARA BANCO DE BATERIAS":
        if i['Anuncio'] == "Classico":
            equalizadorClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == "Premium":
            equalizadorPremium = round(i['PREÇO ML '], 2) - 0.01
    elif i['CONSTRUÇÃO DE CUSTO '] == "FONTE NOBREAK 12V/8A":
        if i['Anuncio'] == "Classico":
            fonte12v8aClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == 'Premium':
            fonte12v8aPremium = round(i['PREÇO ML '], 2) - 0.01
        
    elif i['CONSTRUÇÃO DE CUSTO '] == "FONTE NOBREAK 24V/6A":
        if i['Anuncio'] == "Classico":
            fonte24v6aClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == 'Premium':
            fonte24v6aPremium = round(i['PREÇO ML '], 2) - 0.01
        
    elif i['CONSTRUÇÃO DE CUSTO '] == "FONTE NOBREAK -48V 15A 15A":
        if i['Anuncio'] == "Classico":
            fonte48v15aClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == 'Premium':
            fonte48v15aPremium = round(i['PREÇO ML '], 2) - 0.01
        
    elif i['CONSTRUÇÃO DE CUSTO '] == "FONTE NOBREAK -48V 30A 15A":
        if i['Anuncio'] == "Classico":
            fonte48v30aClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == 'Premium':
            fonte48v30aPremium = round(i['PREÇO ML '], 2) - 0.01
        
    elif i['CONSTRUÇÃO DE CUSTO '] == "FONTE NOBREAK -48V 40A 10A":
        if i['Anuncio'] == "Classico":
            fonte48v40aClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == 'Premium':
            fonte48v40aPremium = round(i['PREÇO ML '], 2) - 0.01
        
    elif i['CONSTRUÇÃO DE CUSTO '] == "INVERSOR OFF GRID SENOIDAL PURA JFA 1000W 48V/220V  RACK ":
        if i['Anuncio'] == "Classico":
            inversor1000wClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == 'Premium':
            inversor1000wPremium = round(i['PREÇO ML '], 2) - 0.01
        
    elif i['CONSTRUÇÃO DE CUSTO '] == "INVERSOR OFF GRID SENOIDAL PURA JFA 3000W 48/220V C/  GER RACK ":
        if i['Anuncio'] == "Classico":
            inversor3000wClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == 'Premium':
            inversor3000wPremium = round(i['PREÇO ML '], 2) - 0.01
        
    elif i['CONSTRUÇÃO DE CUSTO '] == "INVERSOR OFF GRID SENOIDAL PURA JFA 5000W 48/220V C/  GER RACK ":
        if i['Anuncio'] == "Classico":
            inversor5000wClassico = round(i['PREÇO ML '], 2) - 0.01
        elif i['Anuncio'] == 'Premium':
            inversor5000wPremium = round(i['PREÇO ML '], 2) - 0.01
        
#"search_filters": "BRAND=2466336@category=MLB3381@", #MLB2227, 22292586
        
options_req = [
    "EQUALIZADOR PARA BANCO DE BATERIAS",
    "FONTE NOBREAK 12V/8A",
    "FONTE NOBREAK 24V/6A",
    "FONTE NOBREAK -48V 15A 15A",
    "FONTE NOBREAK -48V 30A 15A",
    "FONTE NOBREAK -48V 40A 10A",
    "INVERSOR OFF GRID SENOIDAL PURA JFA 1000W 48V/220V RACK",
    "INVERSOR OFF GRID SENOIDAL PURA JFA 3000W 48/220V C/ GER RACK",
    "INVERSOR OFF GRID SENOIDAL PURA JFA 5000W 48/220V C/ GER RACK",
]
        
url = "https://app.nubimetrics.com/api/search/items"



service = Service()
options = webdriver.ChromeOptions()
titulo_arquivo = ""
# options.add_argument("--headless=new")

options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)


driver = webdriver.Chrome(service=service, options=options)
driver.get("https://www.google.com.br/?hl=pt-BR")
time.sleep(3)
try:
    driver.get("https://app.nubimetrics.com/account/login?ReturnUrl=%2fopportunity%2fcategoryDetail#?category=MLB5672")#https://app.nubimetrics.com/opportunity/categoryDetail#?category=MLB263532
    counter = 0
    while True:
        test = driver.find_elements(By.XPATH, '//*[@id="content"]/div[1]/div/form/div/div[1]/fieldset/section[1]/label/input')
        if test:
            break
        else:
            counter += 1
            if counter > 20:
                break;
            time.sleep(0.5)
    driver.find_element(By.XPATH, '//*[@id="content"]/div[1]/div/form/div/div[1]/fieldset/section[1]/label/input').send_keys("carlosbartojr@yahoo.com")
    driver.find_element(By.XPATH, '//*[@id="content"]/div[1]/div/form/div/div[1]/fieldset/section[2]/label/input').send_keys("JFA2004")
    driver.find_element(By.XPATH, '//*[@id="content"]/div[1]/div/form/div/footer/button').click()
except TimeoutException as e:
    print(f"Timeout ao tentar carregar a página ou encontrar um elemento: {e}")
except NoSuchElementException as e:
    print(f"Elemento não encontrado na página: {e}")
except WebDriverException as e:
    print(f"Erro no WebDriver: {e}")

driver.get("https://app.nubimetrics.com/search/layout#?op1=q-searchTypeOption3-icPubliActivas&op2=fonte%2060a%20jfa&category=")

time.sleep(5)
cookies_list = []

cookies = driver.get_cookies()
for cookie in cookies:
    objeto = cookie['name']
    value = cookie['value']
    cookies_list.append(f"{objeto}={value};")

cookies = "".join(cookies_list)
driver.quit()

headers = {
    "Cookie": cookies
}

base_params = {
    "site_id": "MLB",
    "buying_mode": "buy_it_now",
    "limit": 50,
    "offset": 0,
    "attributes": "results,available_filters,paging,filters",
    "seller_id": 1242763049,
    "order": "relevance",
    "typeSearch": "q",
    "exportData": "false",
    "language": "pt_BR",
    "isControlPrice": "true"
}

# Parâmetros específicos
params_list = [
    {"search_filters": "BRAND=2466336@"},
    {"search_filters": "BRAND=2466336@"},
    {"search_filters": "BRAND=22292586@"},
    {"search_filters": "BRAND=22292586@"}
]

# Lista para armazenar todos os resultados filtrados
all_filtered_results = []

# Loop para cada opção e para cada conjunto de parâmetros
for option in tqdm(options_req):
    for params in params_list:
        # Atualizar o campo 'to_search' com a opção atual
        params.update(base_params)
        params['to_search'] = option

        # Inicializar offset para paginação
        offset = 0
        while True:
            params['offset'] = offset

            # Fazer a requisição GET
            response = requests.get(url, params=params, headers=headers)

            # Verificar se a requisição foi bem-sucedida
            if response.status_code != 200:
                print(f"Erro ao fazer a requisição para {option} com {params['search_filters']}: {response.status_code}")
                break

            data = response.json()
            results = data.get('data', {}).get('results', [])
            total = data.get('data', {}).get('paging', {}).get('total', 0)

            # Filtrar os resultados
            for item in results:
                title = item.get('title', '').lower()
                price = item.get('price', float('inf'))
                real_price = item.get('original_price', float('inf'))
                link = item.get('permalink', '')
                sellernickname = item.get('sellernickname', '')
                listing_type_id = item.get('listing_type_id', '')
                if real_price:
                    real_price = float(real_price)
                if option == "EQUALIZADOR PARA BANCO DE BATERIAS":
                    item['modelo'] = "EQUALIZADOR PARA BANCO DE BATERIAS"
                    if "equalizador" in title and ("bateria" in title or "baterias" in title) and "jfa" in title:
                        if "kit" in title:
                            kit_match = re.search(r'kit\s+(\d+)', title)
                            if kit_match:
                                num_kits = int(kit_match.group(1))
                                if num_kits > 1 and price:
                                    price = round(price / num_kits, 2)
                                    cupom = f"KIT: {num_kits} UNIDADES"
                        if listing_type_id == "gold_pro" and price < equalizadorPremium:
                            item['price_previsto'] = equalizadorPremium
                            all_filtered_results.append(item) 

                        elif price < equalizadorClassico:
                            item['price_previsto'] = equalizadorClassico
                            all_filtered_results.append(item)
                elif option == "FONTE NOBREAK 12V/8A":
                    item['modelo'] = "FONTE NOBREAK 12V/8A"
                    if "fonte" in title and "nobreak" in title and "12v" in title and "8a" in title:
                        if listing_type_id == "gold_pro" and price < fonte12v8aPremium:
                            item['price_previsto'] = fonte12v8aPremium
                            all_filtered_results.append(item) 

                        elif price < fonte12v8aClassico:
                            item['price_previsto'] = fonte12v8aClassico                            
                            all_filtered_results.append(item)
                elif option == "FONTE NOBREAK 24V/6A":
                    item['modelo'] = "FONTE NOBREAK 24V/6A"
                    if "fonte" in title and "nobreak" in title and "24v" in title and "6a" in title:
                        if listing_type_id == "gold_pro" and price < fonte24v6aPremium:
                            item['price_previsto'] = fonte24v6aPremium
                            all_filtered_results.append(item) 

                        elif price < fonte24v6aClassico:
                            item['price_previsto'] = fonte24v6aClassico                            
                            all_filtered_results.append(item)
                elif option == "FONTE NOBREAK -48V 15A 15A":
                    item['modelo'] = "FONTE NOBREAK -48V 15A 15A"
                    if "fonte" in title and "nobreak" in title and "48v" in title and "15a" in title:
                        if listing_type_id == "gold_pro" and price < fonte48v15aPremium:
                            item['price_previsto'] = fonte48v15aPremium
                            all_filtered_results.append(item) 

                        elif price < fonte48v15aClassico:
                            item['price_previsto'] = fonte48v15aClassico                            
                            all_filtered_results.append(item)
                elif option == "FONTE NOBREAK -48V 30A 15A":
                    item['modelo'] = "FONTE NOBREAK -48V 30A 15A"
                    if "fonte" in title and "nobreak" in title and "48v" in title and "30a" in title:
                        if listing_type_id == "gold_pro" and price < fonte48v30aPremium:
                            item['price_previsto'] = fonte48v30aPremium
                            all_filtered_results.append(item) 

                        elif price < fonte48v30aClassico:
                            item['price_previsto'] = fonte48v30aClassico                            
                            all_filtered_results.append(item)
                elif option == "FONTE NOBREAK -48V 40A 10A":
                    item['modelo'] = "FONTE NOBREAK -48V 40A 10A"
                    if "fonte" in title and "nobreak" in title and ("48v" in title or "48" in title) and ("40a" in title or "40" in title):
                        if listing_type_id == "gold_pro" and price < fonte48v40aPremium:
                            item['price_previsto'] = fonte48v40aPremium
                            all_filtered_results.append(item) 

                        elif price < fonte48v40aClassico:
                            item['price_previsto'] = fonte48v40aClassico                            
                            all_filtered_results.append(item)
                elif option == "INVERSOR OFF GRID SENOIDAL PURA JFA 1000W 48V/220V RACK":
                    item['modelo'] = "INVERSOR OFF GRID SENOIDAL PURA JFA 1000W 48V/220V RACK"
                    if "inversor" in title and "senoidal" in title and "1000w" in title and ("48v" in title or "48" in title) and ("220v" in title or "220" in title) and "ger" in title and "rack" in title:
                        if listing_type_id == "gold_pro" and price < inversor1000wPremium:
                            item['price_previsto'] = inversor1000wPremium
                            all_filtered_results.append(item) 

                        elif price < inversor1000wClassico:
                            item['price_previsto'] = inversor1000wClassico                            
                            all_filtered_results.append(item)
                elif option == "INVERSOR OFF GRID SENOIDAL PURA JFA 3000W 48/220V C/ GER RACK":
                    item['modelo'] = "INVERSOR OFF GRID SENOIDAL PURA JFA 3000W 48/220V C/ GER RACK"
                    if "inversor" in title and "senoidal" in title and "3000w" in title and ("48v" in title or "48" in title) and ("220v" in title or "220" in title) and "ger" in title and "rack" in title:
                        if listing_type_id == "gold_pro" and price < inversor3000wPremium:
                            item['price_previsto'] = inversor3000wPremium
                            all_filtered_results.append(item) 

                        elif price < inversor3000wClassico:
                            item['price_previsto'] = inversor3000wClassico                            
                            all_filtered_results.append(item)
                elif option == "INVERSOR OFF GRID SENOIDAL PURA JFA 5000W 48/220V C/ GER RACK":
                    item['modelo'] = "INVERSOR OFF GRID SENOIDAL PURA JFA 5000W 48/220V C/ GER RACK"
                    if "inversor" in title and "senoidal" in title and "5000w" in title and ("48v" in title or "48" in title) and ("220v" in title or "220" in title) and "ger" in title and "rack" in title:
                        if listing_type_id == "gold_pro" and price < inversor5000wPremium:
                            item['price_previsto'] = inversor5000wPremium
                            all_filtered_results.append(item) 

                        elif price < inversor5000wClassico:
                            item['price_previsto'] = inversor5000wClassico                            
                            all_filtered_results.append(item)
                    

            # Atualizar o offset para a próxima página
            offset += params['limit']

            # Verificar se todos os itens foram processados
            if offset >= total:
                break

def get_loja(loja):
    # Formatar a URL com o nome da loja
    location_url = f'https://www.mercadolivre.com.br/perfil/{loja.replace(" ", "+")}'
    
    # Fazer a requisição HTTP
    response = requests.get(location_url)
    
    if response.status_code == 200:
        # Parsear o conteúdo HTML da resposta
        tree = html.fromstring(response.content)
        
        # Extrair o texto do elemento especificado pelo XPath
        loja_info = tree.xpath('//*[@id="profile"]/div/div[2]/div[1]/div[3]/p/text()')
        
        if loja_info:
            return loja_info[0].strip() 
        else:
            return "Informação não encontrada"
    else:
        return f"Erro ao acessar a página: {response.status_cod}"
    

def get_greeting():
    current_hour = datetime.now().hour
    if 5 <= current_hour < 12:
        return "Bom dia!"
    elif 12 <= current_hour < 18:
        return "Boa tarde!"
    else:
        return "Boa noite!"

def enviar(grouped_by_seller):
    whatsapp_window = None
    for window in gw.getAllTitles():
        if 'WhatsApp' in window:
            whatsapp_window = gw.getWindowsWithTitle(window)[0]
            break

    # Se a janela foi encontrada, traz para o foco
    if whatsapp_window is not None:
        try:
            whatsapp_window.activate()
            time.sleep(1)  # Espera um pouco para garantir que a janela está em foco

            # Pressiona Ctrl+F para abrir a busca
            pyautogui.hotkey('ctrl', 'f')
            time.sleep(1)
            # Digita o texto desejado
            pyautogui.typewrite('politica telecom ml')
            time.sleep(1)
            
            # Pressiona Tab e Enter
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(1)
        except Exception as e:
            print(f"Erro ao interagir com a janela do WhatsApp: {e}")
            return
    else:
        print("Janela do WhatsApp não encontrada.")
        return

    try:
        keyboard.write(get_greeting())
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        keyboard.write("Segue anúncios fora da política")
        time.sleep(1)
        pyautogui.press('enter')
        for seller, items in grouped_by_seller.items():
            time.sleep(1)
            keyboard.write(f"*{seller}*")
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'enter')
            time.sleep(1)
            loja_info = get_loja(item['Seller'])
            for item in items:
                if item['Listing Type'] == "gold_special":
                    item['Listing Type'] = "Clássico"
                else:
                    item['Listing Type'] = "Premium"
                
                keyboard.write(f"{item['modelo']} - {item['Seller']} - {loja_info} - Preço Anúncio: {item['Price']} - Preço Política: {item['price_previsto']} ({item['Listing Type']})")
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'enter')
                time.sleep(1)
                keyboard.write(f"{item['Link']}")
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'enter')
                time.sleep(1)
            pyautogui.press('enter')
    except Exception as e:
        print(f"Erro ao enviar mensagens: {e}")


formatted_results = [
    {
        "modelo": result['modelo'],
        "Seller": result['sellernickname'],
        "Title": result['title'],
        "Price": result['price'],
        "price_previsto": result['price_previsto'],
        "Listing Type": result['listing_type_id'],
        "Link": result['permalink'],
        "attributes": result['attributes'],
    }
    for result in all_filtered_results
]

grouped_by_seller = defaultdict(list)

for item in formatted_results:
    seller = item['Seller']
    grouped_by_seller[seller].append(item)
    
grouped_by_seller = dict(grouped_by_seller)
    
enviar(grouped_by_seller)
# Salva os dados em um arquivo JSON

# with open('filtered_results.json', 'w', encoding='utf-8') as json_file:
#     json.dump(formatted_results, json_file, ensure_ascii=False, indent=4)

# print("Dados salvos em 'filtered_results.json'")

