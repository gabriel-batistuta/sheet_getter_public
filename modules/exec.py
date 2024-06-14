import subprocess
# to generate .exe file
# pyinstaller --onefile --console modules/exec.py

# Instalar dependências
subprocess.run(["pip", "install", "-r", "requirements.txt"])

# Executar o script principal
try:
    subprocess.run(["python", "main.py"])
except Exception as e:
    with open('log.txt', 'w') as file:
        file.write(f'Error: {e}')