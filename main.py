import re
import requests
import keyboard
from lxml import html
from collections import defaultdict
from datetime import datetime
import schedule
import time
import subprocess
from unidecode import unidecode
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
      
      

def get_loja(loja):
    response = requests.get(f"https://api.mercadolibre.com/sites/MLB/search?nickname={loja}")
    user_id = response.json()['results'][0]['seller']['id']
    user_response = requests.get(f"https://api.mercadolibre.com/users/{user_id}")
    address = user_response.json()['address']['city']
    state = user_response.json()['address']['state']
    return address + " - " + state
    
    

def get_greeting():
    current_hour = datetime.now().hour
    if 5 <= current_hour < 12:
        return "Bom dia!"
    elif 12 <= current_hour < 18:
        return "Boa tarde!"
    else:
        return "Boa noite!"

def enviar(grouped_by_seller):
    requests.post("http://localhost:3000/api/sendText", {
        "chatId": "120363337104474327@g.us",
        "text": f"{get_greeting()} \n Segue anúncios fora da política",
        "session": "default"
    })
    try:
        for seller, items in grouped_by_seller.items():
            dados = f"*{seller}* \n"
            time.sleep(1)
            for item in items:
                if item['listing_type'] == "gold_special":
                    item['listing_type'] = "Clássico"
                else:
                    item['listing_type'] = "Premium"
                
                loja_info = get_loja(item['seller'])
                dados =  dados + f"{item['model']} - {item['seller']} - {loja_info} - Preço Anúncio: {item['price']} - Preço Política: {round(item['predicted_price'], 2)} ({item['listing_type']}) \n {item['link']} \n"
            requests.post("http://localhost:3000/api/sendText", {
            "chatId": "120363337104474327@g.us",
            "text": dados,
            "session": "default"
            })  
    except Exception as e:
        print(f"Erro ao enviar mensagens: {e}")


def get_diferenca(price, previsto):
    return (price / previsto) * 100;

        
options_req = [
    "equalizador de baterias",
    "FONTE NOBREAK 12V/8A",
    "FONTE NOBREAK 24V/6A",
    "FONTE NOBREAK -48V 15A 15A",
    "FONTE NOBREAK -48V 30A 15A",
    "FONTE NOBREAK -48V 40A 10A",
    "INVERSOR SENOIDAL JFA 1000W",
    "INVERSOR SENOIDAL JFA 3000W",
    "INVERSOR SENOIDAL JFA 5000W",
]
        

def politica():
    urls = [
    "https://api.mercadolibre.com/sites/MLB/search?sort=price_asc&BRAND=2466336",
    "https://api.mercadolibre.com/sites/MLB/search?sort=price_asc&BRAND=22292586",
    ]
    # Loop para cada opção e para cada conjunto de parâmetros
    all_filtered_results = []
    for option in tqdm(options_req):
        for url in urls:
            # Atualizar parâmetros com a opção atual
            params = {"q": option}

            # Inicializar offset para paginação
            offset = 0
            while True:
                # Atualizar parâmetros com o offset atual
                params['offset'] = offset

                # Fazer a requisição GET
                try:
                    response = requests.get(url, params=params)
                except Exception as e:
                    # Tentar novamente após 10 segundos em caso de erro
                    time.sleep(10)
                    response = requests.get(url, params=params)
                    print(f"Erro ao fazer a requisição para {option}: {e}")

                # Verificar se a requisição foi bem-sucedida
                if response.status_code != 200:
                    print(f"Erro ao fazer a requisição para {option} com {params}: {response.status_code}")
                    break

                data = response.json()
                results = data.get('results', [])
                total = data.get('paging', {}).get('total', 0)
                # Filtrar os resultados
                for item in results:
                    title = unidecode(item.get('title', '').lower())
                    price = item.get('price', float('inf'))
                    real_price = item.get('original_price', float('inf'))
                    link = item.get('permalink', '')
                    sellernickname = item.get('sellernickname', '')
                    listing_type_id = item.get('listing_type_id', '')
                    if real_price:
                        real_price = float(real_price)
                    if option == "equalizador de baterias":
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
                                item['diferenca'] = get_diferenca(price, equalizadorPremium)
                                item['price_previsto'] = equalizadorPremium + 0.01
                                all_filtered_results.append(item) 

                            elif price < equalizadorClassico:
                                item['diferenca'] = get_diferenca(price, equalizadorClassico)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = equalizadorClassico + 0.01
                                    all_filtered_results.append(item)
                    elif option == "FONTE NOBREAK 12V/8A":
                        item['modelo'] = "FONTE NOBREAK 12V/8A"
                        if "fonte" in title and ("nobreak" in title or "gerenciavel" in title) and "12v" in title and "8a" in title:
                            if listing_type_id == "gold_pro" and price < fonte12v8aPremium:
                                item['diferenca'] = get_diferenca(price, fonte12v8aPremium)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte12v8aPremium + 0.01
                                    all_filtered_results.append(item) 

                            elif price < fonte12v8aClassico:
                                item['diferenca'] = get_diferenca(price, fonte12v8aClassico)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte12v8aClassico                             + 0.01
                                    all_filtered_results.append(item)
                    elif option == "FONTE NOBREAK 24V/6A":
                        item['modelo'] = "FONTE NOBREAK 24V/6A"
                        if "fonte" in title and ("nobreak" in title or "gerenciavel" in title) and "24v" in title and "6a" in title:
                            if listing_type_id == "gold_pro" and price < fonte24v6aPremium:
                                item['diferenca'] = get_diferenca(price, fonte24v6aPremium)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte24v6aPremium + 0.01
                                    all_filtered_results.append(item) 

                            elif price < fonte24v6aClassico:
                                item['diferenca'] = get_diferenca(price, fonte24v6aClassico)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte24v6aClassico                             + 0.01
                                    all_filtered_results.append(item)
                    elif option == "FONTE NOBREAK -48V 15A 15A":
                        item['modelo'] = "FONTE NOBREAK -48V 15A 15A"
                        if "fonte" in title and ("nobreak" in title or "gerenciavel" in title) and "48v" in title and "15a" in title:
                            if listing_type_id == "gold_pro" and price < fonte48v15aPremium:
                                item['diferenca'] = get_diferenca(price, fonte48v15aPremium)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte48v15aPremium + 0.01
                                    all_filtered_results.append(item) 

                            elif price < fonte48v15aClassico:
                                item['diferenca'] = get_diferenca(price, fonte48v15aClassico)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte48v15aClassico                             + 0.01
                                    all_filtered_results.append(item)
                    elif option == "FONTE NOBREAK -48V 30A 15A":
                        item['modelo'] = "FONTE NOBREAK -48V 30A 15A"
                        if "fonte" in title and ("nobreak" in title or "gerenciavel" in title) and "48v" in title and "30a" in title:
                            if listing_type_id == "gold_pro" and price < fonte48v30aPremium:
                                item['diferenca'] = get_diferenca(price, fonte48v30aPremium)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte48v30aPremium + 0.01
                                    all_filtered_results.append(item) 

                            elif price < fonte48v30aClassico:
                                item['diferenca'] = get_diferenca(price, fonte48v30aClassico)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte48v30aClassico                             + 0.01
                                    all_filtered_results.append(item)
                    elif option == "FONTE NOBREAK -48V 40A 10A":
                        item['modelo'] = "FONTE NOBREAK -48V 40A 10A"
                        if "fonte" in title and ("nobreak" in title or "gerenciavel" in title) and ("48v" in title or "48" in title) and ("40a" in title or "40" in title):
                            if listing_type_id == "gold_pro" and price < fonte48v40aPremium:
                                item['diferenca'] = get_diferenca(price, fonte48v40aPremium)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte48v40aPremium + 0.01
                                    all_filtered_results.append(item) 

                            elif price < fonte48v40aClassico:
                                item['diferenca'] = get_diferenca(price, fonte48v40aClassico)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = fonte48v40aClassico                             + 0.01
                                    all_filtered_results.append(item)
                    elif option == "INVERSOR SENOIDAL JFA 1000W":
                        item['modelo'] = "INVERSOR OFF GRID SENOIDAL PURA JFA 1000W 48V/220V RACK"
                        if "inversor" in title and "senoidal" in title and "1000w" in title and ("48v" in title or "48" in title) and ("220v" in title or "220" in title) and ("ger" in title or "rack" in title):
                            if listing_type_id == "gold_pro" and price < inversor1000wPremium:
                                item['diferenca'] = get_diferenca(price, inversor1000wPremium)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = inversor1000wPremium + 0.01
                                    all_filtered_results.append(item) 

                            elif price < inversor1000wClassico:
                                item['diferenca'] = get_diferenca(price, inversor1000wClassico)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = inversor1000wClassico                             + 0.01
                                    all_filtered_results.append(item)
                    elif option == "INVERSOR SENOIDAL JFA 3000W":
                        item['modelo'] = "INVERSOR OFF GRID SENOIDAL PURA JFA 3000W 48/220V C/ GER RACK"
                        if "inversor" in title and "senoidal" in title and "3000w" in title and ("48v" in title or "48" in title) and ("220v" in title or "220" in title):#and "ger" in title and "rack" in title
                            if listing_type_id == "gold_pro" and price < inversor3000wPremium:
                                item['diferenca'] = get_diferenca(price, inversor3000wPremium)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = inversor3000wPremium + 0.01
                                    all_filtered_results.append(item) 

                            elif price < inversor3000wClassico:
                                item['diferenca'] = get_diferenca(price, inversor3000wClassico)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = inversor3000wClassico                             + 0.01
                                    all_filtered_results.append(item)
                    elif option == "INVERSOR SENOIDAL JFA 5000W":
                        item['modelo'] = "INVERSOR OFF GRID SENOIDAL PURA JFA 5000W 48/220V C/ GER RACK"
                        if "inversor" in title and "senoidal" in title and "5000w" in title and ("48v" in title or "48" in title) and ("220v" in title or "220" in title):#and "ger" in title and "rack" in title
                            if listing_type_id == "gold_pro" and price < inversor5000wPremium:
                                item['diferenca'] = get_diferenca(price, inversor5000wPremium)
                                if item['diferenca'] > 70:
                                    item['price_previsto'] = inversor5000wPremium + 0.01
                                    all_filtered_results.append(item) 

                            elif price < inversor5000wClassico:
                                item['price_previsto'] = inversor5000wClassico  + 0.01
                                if item['diferenca'] > 70:
                                    item['diferenca'] = get_diferenca(price, inversor5000wClassico)                           
                                    all_filtered_results.append(item)
                        

                # Atualizar o offset para a próxima página
                offset += 50
                # Verificar se todos os itens foram processados
                if offset >= total:
                    break
    formatted_results = [
        {
            "image": result['thumbnail'],
            "model": result['modelo'],
            "seller": result['seller']['nickname'],
            "title": result['title'],
            "price": result['price'],   
            "predicted_price": result['price_previsto'],
            "listing_type": result['listing_type_id'],
            "link": result['permalink'],
        }
        for result in all_filtered_results
    ]

    print(formatted_results)
    # grouped_by_seller = defaultdict(list)

    # for item in formatted_results:
    #     seller = item['seller']
    #     grouped_by_seller[seller].append(item)
        
    # grouped_by_seller = dict(grouped_by_seller)
        
    # enviar(grouped_by_seller)


def executar_codigo():
    politica()

executar_codigo()
# Agendar a execução nos horários especificados
# schedule.every().day.at("08:00").do(executar_codigo)
# schedule.every().day.at("11:00").do(executar_codigo)
# schedule.every().day.at("14:00").do(executar_codigo)
# schedule.every().day.at("16:00").do(executar_codigo)
# schedule.every().day.at("20:00").do(executar_codigo)
# schedule.every().day.at("00:00").do(executar_codigo)

# while True:
#     schedule.run_pending()
#     time.sleep(60) 