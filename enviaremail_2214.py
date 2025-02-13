import win32com.client as win32
import os
import pandas as pd
from time import gmtime, strftime


def gerar_corpo_email(dados):
    # Gerar a data atual formatada
    data_atual = strftime('%d/%m/%Y', gmtime())
    corpo = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20px;
                line-height: 1.6;
                font-size: 14px;
                color: #333;
                background-color: #f9f9f9;
            }}
            .container {{
                background: #fff;
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 20px;
                max-width: 800px;
                margin: auto;
                box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.1);
            }}
            h2 {{
                color: #007BFF;
                text-align: center;
                font-size: 22px;
                margin-bottom: 20px;
            }}
            h3 {{
                color: #555;
                font-size: 16px;
                margin-top: 10px;
                margin-bottom: 5px;
            }}
            table {{
                width: 80%;
                border-collapse: collapse;
                margin: 15px auto;
                font-size: 14px;
                background-color: #fff;
                color: #000;
                border: 1px solid #333;
            }}
            th, td {{
                border: 1px solid #333;
                text-align: left;
                padding: 8px;
            }}
            th {{
                background-color: #f4f4f4;
                color: #333;
                font-weight: bold;
            }}
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
            .footer {{
                text-align: center;
                margin-top: 20px;
                font-size: 12px;
                color: #888;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <p>Prezados,</p>
            <p>Encaminhamos informações atualizadas sobre os lançamentos em aberto até a data <strong>{data_atual}</strong>.</p>
            <p>Observa-se que o tipo de pagamento com maior valor é <strong></strong> totalizando <strong></strong>, na Faculdade Sebrae <strong></strong> tem o maior número de registros totalizando <strong></strong>, já a <strong></strong> tem o maior valor de parcelas em aberto totalizando <strong></strong>.</p>

            </div>
            <div class="container">
            <h2>LANÇAMENTOS EM ABERTO</h2>

    """

    #for loop preenche tabela 1
    for resumo in dados[:1]:
        print(type(resumo))  # Verifique se é um dicionário
        print(resumo)  # Verifique o conteúdo de resumo
        # Calcular a quantidade total de registros e o valor total corretamente
        valor_total = sum(detalhe.get("VALORBRUTO", 0) for detalhe in resumo['detalhes'])
        # Formatar o valor total com separador de milhar no formato brasileiro
        valor_total_formatado = f"{valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        corpo += f"""
       
            <h3>Valor Total: R$ {valor_total_formatado}</h3>
      
        """

        # Tabela para Tipo de Pagamento, Quantidade e Valor Total
        corpo += """
        <table>
            <thead>
                <tr>
                    <th>Tipo Pagamento</th>
                    <th>Quantidade</th>
                    <th>Valor Total</th>
                </tr>
            </thead>
            <tbody>
        """

        # Criação do dicionário para agrupar por "Tipo de Pagamento"
        tipo_contrato_resumo = {}
        for detalhe in resumo['detalhes']:
            valor_por_contrato = detalhe.get("VALORBRUTO", 0)

            if valor_por_contrato not in tipo_contrato_resumo:
                tipo_contrato_resumo[valor_por_contrato] = {'VALORBRUTO': 0}

            tipo_contrato_resumo[valor_por_contrato]['VALORBRUTO'] += valor_por_contrato

        for valor_por_contrato, resumo_valorbruto in tipo_contrato_resumo.items():
            valor_formatado = f"{resumo_valorbruto['VALORBRUTO']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            corpo += f"""
            <tr>
                <td>{resumo_valorbruto['VALORBRUTO']}</td>
                <td>R$ {valor_formatado}</td>
            </tr>
            """
        corpo += """
            </tbody>
        </table>
        <br>
        """


    corpo += """
        </tbody>
    </table>
    <br>
"""  
    corpo += """
        </div>
        <div class="footer">
            <p>Este é um e-mail automático. Por favor, não responda.</p>
        </div>
    </body>
    </html>
    """
    return corpo

 
 
def enviar_email(destinatario, copiar, assunto, corpo, caminho_anexo):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  
    mail.To = destinatario
    mail.cc = copiar
    mail.Subject = assunto
    mail.HTMLBody = corpo
    if caminho_anexo and os.path.exists(caminho_anexo):
        mail.Attachments.Add(caminho_anexo)
   
    mail.Send()
    print("E-mail enviado com sucesso!")