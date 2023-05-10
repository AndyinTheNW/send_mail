# Automação para envios de arquivos XML e PDF por E-mail

Este script Python foi desenvolvido para enviar automaticamente arquivos XML e PDF por e-mail para a empresa e, quando aplicável, para os clientes.

## Funcionalidades

- Enviar arquivos XML e PDF para um endereço de e-mail padrão (empresa)
- Enviar arquivos XML e PDF para clientes com base no CFOP encontrado no arquivo XML
- Anexar uma imagem no corpo do e-mail para clientes
- Mover arquivos enviados para pastas específicas

## Requisitos

- Python 3.x
- Bibliotecas Python:
    - smtplib
    - os
    - shutil
    - xml.etree.ElementTree
    - email.mime.text
    - email.mime.multipart
    - email.mime.application
    - email.mime.image

## Como usar

1. Configure as informações da conta de e-mail de origem e destino nas variáveis `email_de`, `senha` e `email_para`.
2. Ajuste o diretório onde estão os arquivos XML e PDF na variável `diretorio`.
3. (Opcional) Modifique o caminho da imagem na variável `imagem_caminho`.
4. Execute o script Python.

Os arquivos XML e PDF serão enviados por e-mail para o endereço de e-mail padrão (empresa) e, quando aplicável, para os clientes com base no CFOP encontrado no arquivo XML. Além disso, os arquivos enviados serão movidos para as pastas "enviados" e "enviados_cliente", conforme apropriado.

## Observações

- Certifique-se de ter as bibliotecas Python necessárias instaladas.

- O script foi projetado para uso com o servidor SMTP do Outlook. Ajuste as configurações do servidor SMTP conforme necessário.

- Ao executar o script pela primeira vez, as pastas "enviados" e "enviados_cliente" serão criadas automaticamente no diretório especificado.

- O arquivo .py deve obrigatóriamente estar na mesma pasta dos arquivos XML e PDF para a correta execução do código.
