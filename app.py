import tkinter as tk
from tkinter import messagebox
import requests
import subprocess
from datetime import datetime
from threading import Timer
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side

# Substitua pela sua chave de API
API_KEY = 'bfbccb1c6bc4841a8f1d4e7632c1bafc'

def buscar_previsao():
    cidade = entrada_cidade.get()
    if not cidade:
        messagebox.showerror("Erro", "Por favor, insira o nome de uma cidade.")
        return
    
    # Construindo a URL para buscar a previsão com base no nome da cidade
    url = f'http://api.openweathermap.org/data/2.5/weather?q={cidade}&appid={API_KEY}&units=metric'

    try:
        response = requests.get(url)
        response.raise_for_status()  # Verifica se a requisição foi bem-sucedida
        data = response.json()
        temperatura = round(data['main']['temp'], 1)  # Formata a temperatura para 1 casa decimal
        umidade = data['main']['humidity']

        # Abrindo a URL no navegador Chrome
        abrir_url(f'https://www.google.com/search?q=weather+{cidade}', cidade, temperatura, umidade)
        
        limpar_campos()  # Limpa o campo após a busca
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao obter os dados: {e}")

def abrir_url(url, cidade, temperatura, umidade):
    # Caminho para o navegador Chrome
    chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe'
    try:
        # Abrir o navegador como um processo separado
        subprocess.Popen([chrome_path, url])
        
        # Fechar automaticamente após 5 segundos e salvar os dados
        Timer(5.0, fechar_navegador, [cidade, temperatura, umidade]).start()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir o navegador: {e}")

def fechar_navegador(cidade, temperatura, umidade):
    try:
        # Encerra todos os processos do navegador Chrome abertos
        subprocess.call("taskkill /F /IM chrome.exe", shell=True)
        salvar_dados(cidade, temperatura, umidade)
        messagebox.showinfo("Sucesso", f"Dados salvos para {cidade}.\nTemperatura: {temperatura:.1f}°C\nUmidade: {umidade}%")
    except Exception as e:
        print(f"Erro ao fechar o navegador: {e}")

def salvar_dados(cidade, temperatura, umidade):
    try:
        workbook = load_workbook('historico_clima.xlsx')
        sheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Histórico Climático'
        headers = ['Data', 'Hora', 'Cidade', 'Temperatura (°C)', 'Umidade (%)']
        for col_num, header in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))

    data_hora = datetime.now()
    data = data_hora.strftime('%Y-%m-%d')
    hora = data_hora.strftime('%H:%M:%S')
    nova_linha = [data, hora, cidade, temperatura, umidade]
    sheet.append(nova_linha)

    ultima_linha = sheet.max_row

    # Definindo cores de acordo com os níveis de umidade
    cell_umidade = sheet.cell(row=ultima_linha, column=5)
    
    # Aplicando cores baseadas na umidade
    if umidade < 12:
        cell_umidade.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Vermelho
    elif 12 <= umidade < 20:
        cell_umidade.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # Laranja
    elif 20 <= umidade <= 30:
        cell_umidade.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Amarelo
    else:
        cell_umidade.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # Verde para umidade acima de 30%

    # Estilizando a célula da umidade
    cell_umidade.font = Font(bold=False)

    workbook.save('historico_clima.xlsx')
    print(f'Dados salvos: {data}, {hora}, Cidade: {cidade}, Temp: {temperatura:.0f}°C, Umidade: {umidade}%')

def limpar_campos():
    entrada_cidade.delete(0, tk.END)

# Configurando a interface com Tkinter
janela = tk.Tk()
janela.title("Previsão do Tempo")

label_cidade = tk.Label(janela, text="Digite o nome da cidade para a previsão do tempo:")
label_cidade.pack(pady=5)

entrada_cidade = tk.Entry(janela, width=30)
entrada_cidade.pack(pady=5)

botao_buscar = tk.Button(janela, text="Buscar previsão", command=buscar_previsao)
botao_buscar.pack(pady=10)

janela.mainloop()
