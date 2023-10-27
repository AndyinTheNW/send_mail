from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from time import sleep
import os
import smtplib
import shutil
import xml.etree.ElementTree as ET
import logging


# Configurações de log

logging.basicConfig(filename='mailtoclient.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Loga um evento de início do programa

logging.info('Iniciando o programa')

try:

    # Configurações do servidor SMTP para Outlook.com
    smtp_host = 'smtp.office365.com'
    smtp_port = 587

    print(f'configurando o servidor SMTP para {smtp_host}:{smtp_port}')
    logging.info(
        f'configurando o servidor SMTP para {smtp_host}:{smtp_port}')

    # Dados da conta de e-mail de origem

    email_de = 'destino@outlook.com'
    senha = 'senha'

    print(f' Definido E-mail de origem: {email_de}')
    logging.info(f' Definido E-mail de origem: {email_de}')

    # Dados da conta de e-mail de destino

    email_para = 'DestinoEmpresaXML@outlook.com'

    print(f' Email para recebimento: {email_para}')
    logging.info(f' Email para recebimento: {email_para}')

    # Diretorio 'home/pyfiles
    diretorio_inicial = os.path.join(os.path.expanduser('~'), 'pyfiles')

    imagem_caminho = os.path.join(
        os.path.expanduser('~'), 'pyfiles', 'car.jpeg')

    print(f'localizando os componentes no diretorio {diretorio_inicial}')
    logging.info(f'localizando os componentes no diretorio {diretorio_inicial}')

    # Diretorio onde os componentes enviados serão armazenados
    todos_enviados_dir = os.path.join(diretorio_inicial, 'enviados')
    enviados_cliente_dir = os.path.join(
        todos_enviados_dir, 'enviados_cliente')

    # verifica se existem os diretorios 'todos_enviados' e 'enviados_cliente' e cria ambas caso não existam, no diretorio 'mailfiles'

    if not os.path.exists(todos_enviados_dir):
        os.makedirs(todos_enviados_dir)

    if not os.path.exists(enviados_cliente_dir):
        os.makedirs(enviados_cliente_dir)

    print(
        f'criando os diretorios {todos_enviados_dir} e {enviados_cliente_dir}')
    logging.info(
        f'criando os diretorios {todos_enviados_dir} e {enviados_cliente_dir}')

    # Declara a variavel cfop e ncm como None para que possa ser sobrescrita posteriormente pelo valor encontrado no arquivo XML

    cfop = None
    ncm = None

    # Percorre a pasta em busca de componentes XML, se houver algum, envia para o e-mail de destino, se não, aguarda 5 segundos e verifica novamente

    while True:
        cliente_email = None
        xml_files = [f for f in os.listdir(
            diretorio_inicial) if f.endswith('.xml')]

        # Verifica se existem arquivos XML no diretorio 'pyfiles'
        
        if xml_files:
            for arquivo in xml_files:
                
                # Abre o arquivo XML e procura pela tag 'cfop' pelo caminho nfeProc → NFe → infNFe → det → prod → CFOP
                
                caminho_arquivo = os.path.join(diretorio_inicial, arquivo)
                tree = ET.parse(caminho_arquivo)
                root = tree.getroot()

                # Lidando com o namespace
                namespaces = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
                nfe_element = root.find('ns:NFe', namespaces)

                # Encontre a tag CFOP
                if nfe_element is not None:
                    infNFe_element = nfe_element.find(
                        'ns:infNFe', namespaces)

                    if infNFe_element is not None:
                        det_element = infNFe_element.find(
                            'ns:det', namespaces)

                        if det_element is not None:
                            prod_element = det_element.find(
                                'ns:prod', namespaces)

                            if prod_element is not None:
                                cfop_element = prod_element.find(
                                    'ns:CFOP', namespaces)
                                ncm_element = prod_element.find(
                                    'ns:NCM', namespaces)

                                if cfop_element is not None:
                                    cfop = cfop_element.text
                                else:
                                    cfop = None
                                if ncm_element is not None:
                                    ncm = ncm_element.text
                                else:
                                    ncm = None
                            else:
                                cfop = None
                                ncm = None
                        else:
                            cfop = None
                            ncm = None
                    else:
                        cfop = None
                        ncm = None
                else:
                    cfop = None
                    ncm = None

                print(f"Encontrado o CFOP {cfop} no arquivo {arquivo}")
                print(f"Encontrado o NCM {ncm} no arquivo {arquivo}")

                # Se o campo 'cfop' = 6102 e o campo 'ncm' começar com 8703, procura pelo e-mail do cliente no arquivo XML
                if cfop == '6102' and ncm.startswith('8703'):
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
                    logging.info(
                        f'Encontrado o e-mail {cliente_email} no arquivo {arquivo}')

                # Nome do arquivo PDF que deve ser enviado junto com o XML
                arquivo_sem_sufixo = arquivo.replace('-nfe', '')
                nome_pdf = arquivo_sem_sufixo.replace('.xml', '.pdf')
                caminho_pdf = os.path.join(diretorio_inicial, nome_pdf)

                print(f'localizando o arquivo PDF {caminho_pdf}')

                # Se o arquivo PDF correspondente não existir, pular para o proximo arquivo XML
                if not os.path.exists(caminho_pdf):
                    print(
                        f'O arquivo PDF correspondente {nome_pdf} nao foi encontrado, pulando o arquivo XML {arquivo}')
                    logging.info(
                        f'O arquivo PDF correspondente {nome_pdf} nao foi encontrado, pulando o arquivo XML {arquivo}')
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
                    anexo = MIMEApplication(
                        xml_file.read(), _subtype='xml')
                    anexo.add_header('Content-Disposition',
                                     'attachment', filename=arquivo)
                    msg.attach(anexo)

                if os.path.exists(caminho_pdf):
                    with open(caminho_pdf, 'rb') as pdf_file:
                        anexo = MIMEApplication(
                            pdf_file.read(), _subtype='pdf')
                        anexo.add_header('Content-Disposition',
                                         'attachment', filename=nome_pdf)
                        msg.attach(anexo)
                else:
                    print(f'Arquivo {nome_pdf}Not found')

                # Envie o e-mail
                server.sendmail(email_de, email_destino, msg.as_string())
                print(
                    f'Arquivo {arquivo} enviado com sucesso para {email_destino}')
                logging.info(
                    f'Arquivo {arquivo} enviado com sucesso para {email_destino}')

            server.quit()

            print(
                f'Arquivo {arquivo} enviado com sucesso para {email_destino}')
            logging.info(
                f'Arquivo {arquivo} enviado com sucesso para {email_destino}')

            # Move o arquivo enviado para a pasta 'todos_enviados'

            shutil.move(caminho_arquivo, todos_enviados_dir)
            if os.path.exists(caminho_pdf):
                shutil.move(caminho_pdf, todos_enviados_dir)
                print(
                    f'Arquivo {arquivo} movido para {todos_enviados_dir}')

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
                print(
                    f'Arquivo {arquivo} copiado para {enviados_cliente_dir}')
        logging.info('Ciclo reiniciado')
        sleep(5)

except Exception as e:
    # Loga um evento de erro
    logging.error('Ocorreu um erro: %s', e)

    smtp_host = 'smtp-mail.outlook.com'
    smtp_port = 587

    # Dados da conta de e-mail dos administradores do sistema
    email_admins = ['exampleAdmin@CorpMail.com.ru',
                    # '',
                    # '',
                    # ''
                    ]

    for email_admin in email_admins:
        msg = MIMEMultipart()
        msg['From'] = email_de
        msg['To'] = email_admin
        msg['Subject'] = 'Erro na Aplicação PyFiles'

        body = 'Alerta: ocorreu um erro na aplicação Pyfiles - GWM. Corrigir assim que possível. Erro:\n{} \n\n\n'.format(
            e)
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP(smtp_host, smtp_port)
        server.starttls()
        server.login(email_de, senha)
        text = msg.as_string()
        server.sendmail(email_de, email_admin, text)
        server.quit()

finally:
    # loga um evento de fim de execução
    logging.info('Fim do programa')
