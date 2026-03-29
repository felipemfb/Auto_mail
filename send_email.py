# bibliotecas necessárias
import pandas as pd
from email.message import EmailMessage
import smtplib
import ssl
import os

# Coloque aqui a senha que você criou no gmail após a autenticação de dois fatores. Para mais informações, consulte o read.me
email_senha = ''

# email de quem está enviando
email_origem = ''

# Lê a planilha de onde os dados serão extraídos. Basta inserir o nome da planilha desejada dentro das aspas, com a extensão csv.
planilha = pd.read_csv(r"")

# Coletar os dados e colocar numa lista. Entre [''] temos o nome das colunas. Basta trocar para o nome das colunas convenientes a você
emails_de_envio = planilha['column1'].tolist()
textos = planilha['column2'].tolist()
cursos = planilha['column3'].tolist()
nome = 'nome'
cargo = 'cargo'

# Lembre de trocar o range para acabar na quantidade de linhas que você quer
for i, email_destino in enumerate(emails_de_envio):
    # Verifica se o email_destino é válido
    if not email_destino or not isinstance(email_destino, str) or "@" not in email_destino:
        print(f"Endereço de e-mail inválido ou vazio na linha {i + 1}. Pulando...")
        continue
    texto = textos[i]
    assunto = 'subject'
    curso = cursos[i]

    # Mensagem que vai no corpo do email. Pode-se usar uma string ou um arquivo.
    corpo = f"""
    <html>
        <body>
        </body>
    </html>
    """
    # Configuração da mensagem
    mensagem = EmailMessage()
    mensagem["From"] = email_origem
    mensagem["To"] = email_destino
    mensagem["Subject"] = assunto
    mensagem.add_alternative(corpo, subtype='html')  # Define o corpo como HTML

    # Adiciona a imagem como anexo
    with open(r"path_image", "rb") as img:
        mensagem.add_attachment(
            img.read(),
            maintype="image",
            subtype="png",
            filename="filename.png",
            cid="filename"  # Define o Content-ID para referenciar no HTML
        )

    # Garante a segurança da mensagem
    safe = ssl.create_default_context()

    # Parte que realmente envia
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=safe) as servidor:
        servidor.login(email_origem, email_senha)
        servidor.send_message(mensagem)
        print(f"E-mail enviado com sucesso para {email_destino}.")