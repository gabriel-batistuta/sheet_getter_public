import os

# Diretório onde estão os arquivos Excel
diretorio = 'planilhas'

# Percorre todos os arquivos no diretório
for arquivo in os.listdir(diretorio):
    # Verifica se o arquivo é um arquivo Excel (.xlsx)
    if arquivo.endswith('.xlsx'):
        # Cria o caminho completo para o arquivo
        caminho_arquivo = os.path.join(diretorio, arquivo)
        # Remove o arquivo
        os.remove(caminho_arquivo)
        print(f'Arquivo {arquivo} excluído com sucesso.')
