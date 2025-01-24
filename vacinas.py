import pandas as pd
import os
import pywhatkit as pwk
import time
import re
import sys

# Diretório do script
if getattr(sys, 'frozen', False):
    current_dir = os.path.dirname(sys.executable)
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))

# Arquivo XLSX
arquivo_xlsx = os.path.join(current_dir, 'vacinacao.xlsx')

# Conversão xlsx para dataframe
df = pd.read_excel(arquivo_xlsx, skiprows=3)

# Excluir vacinas de giárdia
df = df[~df['Vacina'].str.contains('Giárdia')]

# Mapeamento para substituir os nomes das vacinas/produtos na mensagem final
NOME_VACINA_PRODUTO_MAPEAMENTO = {
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

PRODUTOS_ESPECIFICOS = ['MILBEMAX', 'SOLENCIA', 'Tópico Revolution', 'CREDELI', 'BRAVECTO', 
                        'Vermífugo', 'SIMPARIC', 'COMFORTIS', 'COLEIRA ANTIPARASITÁRIA', 
                        'COLEIRA ANTIPARASITÁRIA SERESTO']

# Função para converter nomes de vacinas/produtos
def substituir_nome_produto(vacina_produto):
    return NOME_VACINA_PRODUTO_MAPEAMENTO.get(vacina_produto, vacina_produto)

# Função para extrair o número de celular
def extrair_numero_celular(telefones):
    for telefone in telefones:
        telefone = telefone.strip()
        if len(telefone) >= 5 and telefone[4] == ' ':
            numero = telefone[5:]
            if numero.startswith('9'):  # Preferir números de celular
                return telefone
    return None

# Função para enviar mensagens
def enviar_mensagens(clientes, simular=False):
    for nome_cliente, dados_cliente in clientes.groupby('Cliente'):
        nome_cliente_formatado = nome_cliente.split()[0].capitalize()

        # Extrair telefones e preferir celular
        telefones = dados_cliente['Telefones'].iloc[0].split(',')
        numero_cliente = extrair_numero_celular(telefones)

        if not numero_cliente:
            print(f"[Atenção] Nenhum número de celular disponível para: {nome_cliente}. Ignorando cliente.")
            continue
        
        # Construir a mensagem
        mensagem = f"Bom dia, tudo bem?\n"
        mensagem_produtos = ""
        mensagem_vacinas = ""

        for animal, vacinas in dados_cliente.groupby('Animal')['Vacina']:
            vacinas_lista = [substituir_nome_produto(vacina) for vacina in vacinas]

            # Separar vacinas específicas e vacinas normais
            produtos_animal = [v for v in vacinas_lista if v in PRODUTOS_ESPECIFICOS]
            vacinas_normais_animal = [v for v in vacinas_lista if v not in PRODUTOS_ESPECIFICOS]
            
            # Adicionar produtos específicos na mensagem
            if produtos_animal:
                mensagem_produtos += f"O(s) {', '.join(produtos_animal).capitalize()} do(a) {animal.capitalize()} vence(m) essa semana.\n"
            
            # Adicionar vacinas normais na mensagem
            if vacinas_normais_animal:
                mensagem_vacinas += f"A vacina(s) {', '.join(vacinas_normais_animal)} do(a) {animal.capitalize()} vence(m) essa semana.\n"
        
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
            pwk.sendwhatmsg_instantly(f"+55 {numero_cliente}", mensagem, wait_time=10, tab_close=True)

            # Aguardar um pouco para garantir que a janela do WhatsApp Web abra
            time.sleep(10)

            print(f"Mensagem enviada automaticamente para: {nome_cliente} ({numero_cliente})")

if __name__ == '__main__':
    while True:
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
