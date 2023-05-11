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

## Descrição dos processos 

### Este programa realiza as seguintes ações:

1. Importação de bibliotecas: O programa importa as bibliotecas necessárias, como `datetime`, `os`, `smtplib`, `shutil`, `xml.etree.ElementTree` e classes específicas dessas bibliotecas para uso posterior.

2. Configuração do servidor SMTP: São definidas as configurações do servidor SMTP para o Outlook.com, especificando o host SMTP e a porta a ser utilizada.

3. Definição das contas de e-mail: São definidos o endereço de e-mail e a senha da conta de origem (remetente) e o endereço de e-mail de destino.

4. Definição de diretórios: É definido um diretório onde os arquivos serão localizados e armazenados. São criados diretórios específicos para arquivos enviados e arquivos enviados para clientes.

5. Loop principal: O programa entra em um loop infinito que verifica periodicamente se existem arquivos XML no diretório especificado. Se houver, realiza as seguintes etapas:

   . Analisa o arquivo XML: O programa abre o arquivo XML, lida com o namespace e busca por elementos específicos dentro do XML, como 'cfop' e 'ncm'.

   . Verificação dos elementos encontrados: Verifica se o valor de 'cfop' é igual a '6102' e se o valor de 'ncm' começa com '8703'. Se as condições forem atendidas, o programa procura pelo e-mail do cliente no arquivo XML.

   . Envio de e-mails: O programa utiliza o servidor SMTP para enviar e-mails contendo os arquivos XML e PDF anexados. Os e-mails são enviados tanto para a empresa quanto, se necessário, para o cliente. O corpo do e-mail varia dependendo do destinatário.

   . Movimentação de arquivos: Após o envio bem-sucedido, os arquivos XML e PDF são movidos para o diretório de arquivos enviados. Se o arquivo XML foi enviado para o cliente, ele também é copiado para o diretório de arquivos enviados para clientes.

   . Aguarda e repete: O programa aguarda por um determinado período de tempo (5 segundos) antes de verificar novamente se existem arquivos XML no diretório.
