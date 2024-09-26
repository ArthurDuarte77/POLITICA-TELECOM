from datetime import datetime
import schedule
import time
import subprocess

def executar_codigo():
    subprocess.run(['python', 'main.py'])

# Agendar a execução nos horários especificados
schedule.every().day.at("08:00").do(executar_codigo)
schedule.every().day.at("11:00").do(executar_codigo)
schedule.every().day.at("14:00").do(executar_codigo)
schedule.every().day.at("16:00").do(executar_codigo)
schedule.every().day.at("20:00").do(executar_codigo)
schedule.every().day.at("00:00").do(executar_codigo)

while True:
    schedule.run_pending()
    time.sleep(60) 