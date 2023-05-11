from datetime import datetime
import os
import smtplib
import shutil
import xml.etree.ElementTree as ET
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from time import sleep


# Configurações do servidor SMTP para Outlook.com
smtp_host = 'smtp.office365.com'
smtp_port = 587

print(f'configurando o servidor SMTP para {smtp_host}:{smtp_port}')

# Dados da conta de e-mail de origem

email_de = 'DestinoEmpresaXML@outlook.com.br'
senha = '@Roma!@#23'

# Dados da conta de e-mail de destino

email_para = 'andersonalmeida1008@outlook.com'

print(f' Email padrão para recebimento: {email_para}')

# Diretório 'home/pyfiles
diretorio = os.path.join(os.path.expanduser('~'), 'pyfiles')

imagem_caminho = os.path.join(os.path.expanduser('~'), 'pyfiles', 'car.jpeg')

print(f'localizando os componentes no diretório {diretorio}')

# Diretório onde os componentes enviados serão armazenados
todos_enviados_dir = os.path.join(diretorio, 'enviados')
enviados_cliente_dir = os.path.join(todos_enviados_dir, 'enviados_cliente')

# verifica se existem os diretórios 'todos_enviados' e 'enviados_cliente' e cria ambas caso não existam, no diretório 'mailfiles'

if not os.path.exists(todos_enviados_dir):
    os.makedirs(todos_enviados_dir)

if not os.path.exists(enviados_cliente_dir):
    os.makedirs(enviados_cliente_dir)

print(f'criando os diretórios {todos_enviados_dir} e {enviados_cliente_dir}')

# Declara a variavél cfop como None para que possa ser sobrescrita posteriormente pelo valor encontrado no arquivo XML

cfop = None

# Percorre a pasta em busca de componentes XML, se houver algum, envia para o e-mail de destino, se não, aguarda 5 segundos e verifica novamente

while True:
    cliente_email = None
    xml_files = [f for f in os.listdir(diretorio) if f.endswith('.xml')]

    # Verifica se existem arquivos XML no diretório 'pyfiles'
    if xml_files:

        for arquivo in xml_files:

            # Abre o arquivo XML e procura pela tag 'cfop' pelo caminho nfeProc → NFe → infNFe → det → prod → CFOP
            caminho_arquivo = os.path.join(diretorio, arquivo)
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()

            # Lidando com o namespace
            namespaces = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
            nfe_element = root.find('ns:NFe', namespaces)

            # Encontre a tag CFOP
            if nfe_element is not None:
                infNFe_element = nfe_element.find('ns:infNFe', namespaces)

                if infNFe_element is not None:
                    det_element = infNFe_element.find('ns:det', namespaces)

                    if det_element is not None:
                        prod_element = det_element.find('ns:prod', namespaces)

                        if prod_element is not None:
                            cfop_element = prod_element.find(
                                'ns:CFOP', namespaces)

                            if cfop_element is not None:
                                cfop = cfop_element.text
                            else:
                                cfop = None
                        else:
                            cfop = None
                    else:
                        cfop = None
                else:
                    cfop = None
            else:
                cfop = None
            print(f"Encontrado o CFOP {cfop} no arquivo {arquivo}")

            # Se o campo 'cfop' = 6102, procurar pela tag 'email' e enviar o e-mail para cliente
            if cfop == '6102':
                if infNFe_element is not None:
                    entrega_element = infNFe_element.find(
                        'ns:entrega', namespaces)

                    if entrega_element is not None:
                        email_element = entrega_element.find(
                            'ns:email', namespaces)

                        if email_element is not None:
                            cliente_email = email_element.text
                        else:
                            cliente_email = None
                    else:
                        cliente_email = None
                else:
                    cliente_email = None
                print(
                    f"Encontrado o e-mail {cliente_email} no arquivo {arquivo}")

            # Nome do arquivo PDF que deve ser enviado junto com o XML
            arquivo_sem_sufixo = arquivo.replace('-nfe', '')
            nome_pdf = arquivo_sem_sufixo.replace('.xml', '.pdf')
            caminho_pdf = os.path.join(diretorio, nome_pdf)

            print(f'localizando o arquivo PDF {caminho_pdf}')

            # Se o arquivo PDF correspondente não existir, pular para o próximo arquivo XML
            if not os.path.exists(caminho_pdf):
                print(
                    f'O arquivo PDF correspondente {nome_pdf} não foi encontrado, pulando o arquivo XML {arquivo}')
                continue

        # Enviar e-mail para a empresa e, se necessário, para o cliente
        destinatarios = [email_para]

        if cliente_email is not None:
            destinatarios.append(cliente_email)

        server = smtplib.SMTP(smtp_host, smtp_port)
        server.starttls()
        server.login(email_de, senha)  # Faz login no servidor

        for email_destino in destinatarios:
            mensagem = f"Enviando o arquivo {arquivo} para {email_destino}"
            if email_destino == cliente_email:
                mensagem += " (cliente)"
            print(mensagem)

            # Cria a mensagem de e-mail
            msg = MIMEMultipart()
            msg['From'] = email_de
            msg['To'] = email_destino
            msg['Subject'] = 'Arquivos XML e PDF enviados por e-mail'

            # Adicione o corpo do e-mail
            if email_destino == cliente_email:
                with open(imagem_caminho, 'rb') as img_file:
                    img_data = img_file.read()
                    image = MIMEImage(img_data)
                    image.add_header('Content-ID', '<cars_image>')
                    msg.attach(image)

                corpo_email = f"""
        <html>
        <body>
            <p>Olá, cliente! espero que esse e-mail o encontre bem</p>
            <p>Segue em anexo os arquivos referentes á compra do seu mais novo carro! Parabéns!</p>
            <p><img src="cid:cars_image" alt="cars image"></p>
            <p>Atenciosamente,</p>
            <p>Anderson</p>
        </body>
        </html>
        """
                msg.attach(MIMEText(corpo_email, 'html'))
            else:
                corpo_email = f'Olá Empresa,\n\nSegue em anexo o arquivo XML e o arquivo PDF.\n\nAtenciosamente,\nAnderson'
                msg.attach(MIMEText(corpo_email, 'plain'))

            with open(caminho_arquivo, 'rb') as xml_file:
                anexo = MIMEApplication(xml_file.read(), _subtype='xml')
                anexo.add_header('Content-Disposition',
                                 'attachment', filename=arquivo)
                msg.attach(anexo)

            if os.path.exists(caminho_pdf):
                with open(caminho_pdf, 'rb') as pdf_file:
                    anexo = MIMEApplication(pdf_file.read(), _subtype='pdf')
                    anexo.add_header('Content-Disposition',
                                     'attachment', filename=nome_pdf)
                    msg.attach(anexo)
            else:
                print(f'Arquivo {nome_pdf} não encontrado')

            # Envie o e-mail
            server.sendmail(email_de, email_destino, msg.as_string())
            print(
                f'Arquivo {arquivo} enviado com sucesso para {email_destino}')

        server.quit()

        print(f'Arquivo {arquivo} enviado com sucesso para {email_destino}')

        # Move o arquivo enviado para a pasta 'todos_enviados'

        shutil.move(caminho_arquivo, todos_enviados_dir)
        if os.path.exists(caminho_pdf):
            shutil.move(caminho_pdf, todos_enviados_dir)
            print(f'Arquivo {arquivo} movido para {todos_enviados_dir}')

            # Se o arquivo XML foi enviado para o cliente, copiar xml e pdf e enviar para a pasta 'enviados_cliente'

    if cfop == '6102':
        shutil.copy(os.path.join(todos_enviados_dir,
                    arquivo), enviados_cliente_dir)
        if os.path.exists(caminho_pdf):
            shutil.copy(os.path.join(todos_enviados_dir,
                        nome_pdf), enviados_cliente_dir)

        print(f'Arquivo {arquivo} copiado para {enviados_cliente_dir}')

        for arquivo in (f for f in os.listdir(todos_enviados_dir) if f.endswith('.pdf')):
            shutil.copy(os.path.join(todos_enviados_dir,
                        arquivo), enviados_cliente_dir)
            print(f'Arquivo {arquivo} copiado para {enviados_cliente_dir}')

    sleep(5)
    # fim do for
