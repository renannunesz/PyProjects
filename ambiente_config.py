import subprocess
import sys
import os

def instalar_bibliotecas(requirements_file):
    try:
        with open(requirements_file, 'r') as file:
            bibliotecas = file.readlines()
            bibliotecas = [biblioteca.strip() for biblioteca in bibliotecas]

            for biblioteca in bibliotecas:
                try:
                    __import__(biblioteca)
                    print(f"A biblioteca '{biblioteca}' já está instalada.")
                except ImportError:
                    print(f"A biblioteca '{biblioteca}' não está instalada. Instalando...")
                    subprocess.check_call([sys.executable, '-m', 'pip', 'install', biblioteca])
                    print(f"A biblioteca '{biblioteca}' foi instalada com sucesso.")
                    
    except FileNotFoundError:
        print(f"Erro: O arquivo '{requirements_file}' não foi encontrado.")

# Defina o arquivo de requisitos
requirements_file = os.path.abspath('requirements.txt')

# Chame a função para instalar as bibliotecas
instalar_bibliotecas(requirements_file)
