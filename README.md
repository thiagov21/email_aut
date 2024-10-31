Descrição do Script de Envio de E-mails Automáticos
Este script em Python foi desenvolvido para automatizar o envio de e-mails de acompanhamento a clientes que estão aguardando o reparo de seus relógios. Utilizando a biblioteca win32com.client, o script se conecta ao Outlook e permite que os e-mails sejam enviados em lotes, garantindo eficiência e organização no processo de comunicação.

Funcionalidades:
Leitura de Dados: O script lê informações de clientes, incluindo nome, e-mail e detalhes do produto, a partir de uma planilha Excel.
Validação de E-mail: Antes de enviar, o script valida se os endereços de e-mail são corretos, evitando erros de envio.
Envio em Lotes: Os e-mails são enviados em grupos (lotes), com um intervalo configurável entre cada envio, permitindo um controle melhor do processo.
Conteúdo Personalizado: Cada e-mail é personalizado com o nome do cliente e detalhes do produto, além de opções de solução para o problema do reparo.
Anexos: Possibilidade de incluir documentos em PDF, como manuais ou instruções adicionais, para melhor suporte ao cliente.
Exemplo de E-mail Enviado:
O e-mail enviado ao cliente contém informações sobre o status do reparo, opções de troca e um link para um catálogo online. O tom é amigável e profissional, garantindo uma boa experiência de comunicação.

Pré-requisitos:
Python 3.x
Bibliotecas: openpyxl, win32com
Microsoft Outlook instalado e configurado
Instruções de Uso:
Configure o caminho para a planilha de clientes e o documento PDF no script.
Certifique-se de que o Outlook esteja aberto e configurado corretamente.
Execute o script para iniciar o processo de envio automático de e-mails.
