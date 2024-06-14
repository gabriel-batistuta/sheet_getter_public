#!/bin/bash

/usr/bin/ruby -e “$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)

chmod +x run_app.sh

brew install create-dmg

appify run_app.sh "Relatórios Salão Jovem"

# Verifica se o usuário possui permissões de superusuário (root)
if [ "$(id -u)" != "0" ]; then
  echo "Este script precisa ser executado como root." 1>&2
  exit 1
fi

# Define o diretório de instalação do Chrome
install_dir="/Applications"

# Caminho para o arquivo ZIP contendo o Chrome
chrome_zip="chrome-mac-arm64.zip"

# Extrai o arquivo ZIP para o diretório de instalação
unzip -o "$chrome_zip" -d "$install_dir"

# Verifica se a extração foi bem-sucedida
if [ $? -eq 0 ]; then
  echo "Chrome instalado com sucesso."
else
  echo "Erro ao instalar o Chrome." 1>&2
  exit 1
fi

# Executa o Chrome
open -a "Google Chrome"

exit 0
