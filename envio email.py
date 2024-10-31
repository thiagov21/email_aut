import os
import time
import win32com.client as win32
import openpyxl
import re

# Caminho da planilha de clientes (exemplo)
caminho_planilha = r"data/exemplo_email_clientes.xlsx"

# Caminho do documento em PDF a ser anexado (exemplo)
caminho_anexo = r"data/exemplo_manual_troca.pdf"

# Link para informações adicionais (exemplo)
link_info = 'https://exemplo.com/instrucoes'

# Inicializa o Outlook
outlook = win32.Dispatch('outlook.application')

# Endereços de e-mail para cópia oculta (BCC) (exemplo)
emails_cco = "exemplo@gmail.com"

# Função para validar endereço de e-mail
def email_valido(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

# Tamanho do lote (quantidade de e-mails por envio)
lote_tamanho = 100

# Pausa entre lotes (em segundos)
pausa_lotes = 5

# Abrir a planilha
try:
    workbook = openpyxl.load_workbook(caminho_planilha)
    sheet = workbook.active

    # Total de linhas na planilha
    total_linhas = sheet.max_row

    # Loop para envio por lotes
    for inicio in range(1, total_linhas, lote_tamanho):
        fim = min(inicio + lote_tamanho, total_linhas)  # Define o fim do lote
        
        for row in sheet.iter_rows(min_row=inicio, max_row=fim, max_col=8, values_only=True):
            cliente_planilha = row[6]  # Coluna G (Cliente)
            email_destinatario = row[7]  # Coluna H (Email)
            relogio_cliente = row[3]  # Coluna D (Relógio)  
            OS_cliente = row[0]  # Coluna B (Ordem de Serviço)
            OS_troca = row[1]
            
            # Verificar se o e-mail é válido
            if email_valido(email_destinatario):
                # Criar um novo e-mail
                email = outlook.CreateItem(0)
                email.Subject = f"Grupo Exemplo - Ordem de Serviço"
                
                # Corpo do e-mail em HTML com texto de exemplo
                corpo_email = f"""
                <p>Prezado(a) {cliente_planilha},</p>
                <p>Este é um e-mail de exemplo para demonstrar a automação de envio de mensagens.</p>
                <p>Em referência ao serviço do relógio marca Exemplo, modelo {relogio_cliente},<br>
                informamos que não será possível concluir o reparo do produto referente
                à Ordem de Serviço {OS_cliente} dentro do prazo legal de 30 (trinta) dias.<br>
                Gerando assim a OS de troca: {OS_troca}.</p>

                <p>Com o intuito de proporcionar uma melhor experiência, gostaríamos de oferecer
                a substituição do produto por um modelo novo. Por favor, verifique as opções disponíveis.</p>

                <p>É fundamental que essa escolha seja feita em até 05 (CINCO) dias úteis. Para auxiliá-lo(a),
                disponibilizamos um passo a passo em nosso site para que você possa realizar a seleção do novo produto.</p>

                <p>Para acessar as instruções, <a href="{link_info}">CLIQUE AQUI</a>.</p>
                
                <p>Em anexo, está o manual para acessar o site.</p>

                <p>Após a confirmação de sua escolha, o novo produto será enviado com um prazo estimado
                de até 20 (VINTE) dias para entrega. Para dúvidas, entre em contato conosco através do e-mail
                <a href="mailto:exemplo@exemplo.com">exemplo@exemplo.com</a>.</p>

                <p>Lamentamos os transtornos e agradecemos sua compreensão.</p>
                <p>Atenciosamente,<br>
                Grupo Exemplo<br>
                <a href="https://exemplo.com">www.exemplo.com</a></p>
                <b>Este é um email automático. Por favor, não responda essa mensagem.</b>
                """

                email.HTMLBody = corpo_email  # Define o corpo do e-mail como HTML
                email.To = email_destinatario
                email.BCC = emails_cco 

                # Anexar o documento PDF
                email.Attachments.Add(caminho_anexo)

                # Enviar o e-mail
                email.Send()

                print(f"E-mail enviado para {cliente_planilha} ({email_destinatario}).")

            else:
                print(f"E-mail inválido para {cliente_planilha}: {email_destinatario}")

        # Pausa entre lotes
        print(f"Esperando {pausa_lotes} segundos antes de enviar o próximo lote...")
        time.sleep(pausa_lotes)

except Exception as e:
    print(f"Erro ao processar a planilha: {e}")

print(f"Todos os e-mails foram enviados. Total de: {total_linhas}")
