import pandas as pd
import os
import pywhatkit as pwk
import keyboard
import time
import pyautogui
import sys


def exibir_mensagem_inicial():
    print("="*50)
    print("Bem-vindo ao Programa de Envio de Mensagens de Vacinação!")
    print("Versão: 1.0.0")
    print("-Este programa envia mensagens automáticas via WhatsApp informando os clientes sobre vacinas e produtos de seus pets que estão vencendo.")
    print("Instruções:")
    print("-- Baixe o relatório de vacinas programadas;")
    print("-- Converta o arquivo .xls para .xlsx (vários sites podem fazer isso facilmente);")
    print("-- O arquivo deve ser renomeado para 'vacinacao.xlsx';")
    print("-- O Whatsapp Web da clínica deve estar conectado no computador. Não precisa estar aberto;")
    print("-- Certifique-se de que o arquivo 'vacinacao.xlsx' está no mesmo diretório que este executável;")
    print("-- Execute esse arquivo. Você será questionado sobre o modo simulação. Se 's', apenas escreverá as mensagens simuladas aqui, para verificação. Rode simulação; 'n' para enviar as mensagens.")
    print("-- Não mexa nas abas enquanto o programa estiver rodando. Ele vai automaticamente enviar as mensagens e fechar as abas;")
    print("="*50)
    print("Iniciando o processo...\n")


# Get the path to the directory where the script or executable is located
if getattr(sys, 'frozen', False):
    # If the script is running as a bundled executable
    current_dir = os.path.dirname(sys.executable)
else:
    # If running in a normal Python environment
    current_dir = os.path.dirname(os.path.abspath(__file__))

# Path to the XLSX file
arquivo_xlsx = os.path.join(current_dir, 'vacinacao.xlsx')

# Conversão xlsx para csv
df = pd.read_excel(arquivo_xlsx, skiprows=3)
df.to_csv('vacinacao.csv', encoding='utf-8')

# Mapeamento para substituir os nomes das vacinas/produtos na mensagem final
nome_vacina_produto_mapeamento = {
    'Antirrábica': 'de raiva',
    'Múltipla canina (Inicio 60 dias)': 'múltipla',
    'Múltipla canina (Início 45 dias)': 'múltipla',
    'TRAQUEOBRONQUITE': "de gripe",
    'Quádrupla': 'quádrupla',
    'CARDIOLOGIA REAVALIAÇÃO': 'avaliação cardiológica',
    'COLEIRA ANTIPARASITÁRIA': 'coleira',
    'COLEIRA ANTIPARASITÁRIA SERESTO': 'coleira',
    'DRONTAL SO GATOS TRANSDERMAL': 'drontal'
}

# Função para converter nomes de vacinas/produtos
def substituir_nome_produto(vacina_produto):
    return nome_vacina_produto_mapeamento.get(vacina_produto, vacina_produto)

# Função para enviar mensagens
def enviar_mensagens(clientes, simular=False):
    
    # Lista de produtos específicos
    produtos_especificos = ['MILBEMAX', 'LIBRELA', 'SOLENCIA', 'Tópico Revolution', 'CREDELI', 'BRAVECTO', 'Vermífugo', 'SIMPARIC', 'COMFORTIS', 'COLEIRA ANTIPARASITÁRIA', 'COLEIRA ANTIPARASITÁRIA SERESTO']

    # Agrupar vacinas por cliente
    for nome_cliente, dados_cliente in clientes.groupby('Cliente'):
        nome_cliente_formatado = nome_cliente.split()[0].capitalize()  # Só o primeiro nome, capitalizado

        # Extrair telefones e preferir celular
        telefones = dados_cliente['Telefones'].iloc[0].split(',')
        numero_cliente = None
        for telefone in telefones:
            telefone = telefone.strip()
            if len(telefone) >= 5 and telefone[4] == ' ':  # Verifica se há um espaço após o DDD
                numero = telefone[5:]
                if numero.startswith('9'):  # Preferir números de celular
                    numero_cliente = telefone
                    break

        if not numero_cliente:
            print(f"[Atenção] Nenhum número de celular disponível para: {nome_cliente}. Ignorando cliente.")
            continue
        
        # Construir a mensagem
        mensagem = f"Bom dia, {nome_cliente_formatado}, tudo bem?\n"
        mensagem_produtos = ""
        mensagem_vacinas = ""

        for animal, vacinas in dados_cliente.groupby('Animal')['Vacina']:
            vacinas_lista = [substituir_nome_produto(vacina) for vacina in vacinas]
            
            # Separar vacinas específicas e vacinas normais
            produtos_animal = [v for v in vacinas_lista if v in produtos_especificos]
            vacinas_normais_animal = [v for v in vacinas_lista if v not in produtos_especificos]
            
            # Adicionar produtos específicos na mensagem
            if produtos_animal:
                mensagem_produtos += f"O {', '.join(produtos_animal).capitalize()} do(a) {animal.capitalize()} vence essa semana.\n"
            
            # Adicionar vacinas normais na mensagem
            if vacinas_normais_animal:
                mensagem_vacinas += f"A vacina {', '.join(vacinas_normais_animal)} do(a) {animal.capitalize()} vence essa semana.\n"
        
        # Adicionar mensagens finais para produtos específicos e vacinas normais
        if mensagem_produtos:
            mensagem += mensagem_produtos + "Caso já tenha feito em casa, nos avise para atualizarmos o sistema! 😊\n"
        
        if mensagem_vacinas:
            mensagem += mensagem_vacinas + "Gostaria de agendar um horário para atualizarmos? 🐶 🐾 "

        # Enviar ou simular envio
        if simular:
            print("Simulação -> ON.")
            print(f"[Simulação] Mensagem para: {numero_cliente}\n{mensagem}\n")
        else:
            print("Simulação -> OFF.")
            pwk.sendwhatmsg_instantly(f"+{numero_cliente}", mensagem)

            # Aguardar um pouco para garantir que a janela do WhatsApp Web abra
            time.sleep(10)  # Aguardar 10 segundos (ajustar conforme a velocidade da internet)

            # Simular a tecla "Enter"
            pyautogui.press('enter')

            # Fechar a janela do WhatsApp Web (Ctrl + W fecha a aba no navegador)
            time.sleep(1)  # Aguardar um pouco antes de fechar a aba
            pyautogui.hotkey('ctrl', 'w')

            print(f"Mensagem enviada automaticamente para: {nome_cliente} ({numero_cliente})")

            print(f"Mensagem enviada para: {nome_cliente_formatado} ({numero_cliente})")

if __name__ == '__main__':
    while True:
        exibir_mensagem_inicial()
        user_input = input("Deseja rodar o programa em modo simulação? (s/n): ").strip().lower()
        if user_input == 's':
            simular = True
            break
        elif user_input == 'n':
            simular = False
            break
        else:
            print("Entrada inválida, por favor responda com 's' (sim) ou 'n' (não).")

    enviar_mensagens(df, simular)