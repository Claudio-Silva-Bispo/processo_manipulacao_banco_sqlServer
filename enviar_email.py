# Inicializa o objeto do aplicativo Outlook
    outlook = win32.Dispatch('outlook.application')
    accounts = outlook.Session.Accounts

    # Verificar se existem pelo menos dois resultados
    if len(accounts) >= 2:
        account_segunda_opcao = accounts[1]
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = account_segunda_opcao

    # Define as informações do e-mail
    mail.To = ''
    mail.Subject = ''

    # Formatação da tabela em HTML
    html_table = tabelaAcionamento.to_html(index=False)
    html1_table = tabelaChegada.to_html(index=False)
    html2_table = resultadosAcionamento.to_html(index=False)
    html3_table = resultadoChegada.to_html(index=False)

    if valor_tabela_2 != '0' or valor_tabela_3 != '0':
        mail.HTMLBody =f""" 
        <html>
        <head>

        <title>Page Title</title>
        </head>
        <body>
        <h2>REPORT SERVIÇOS EM ANDAMENTO</h2>
        <p><h3>SERVIÇOS EM ACIONAMENTO</h3></p> 
        <p>Monitoramento - Consolidado</p> 
        {html2}
        <p> Monitoramento - Detalhado </p> 
        {html8}
        <p><font color="red"> Escalonamento - Coordenador - Serviços acima de 45 minutos e até 60 minutos de atuação</font></p> 
        {html}
        <p><font color="red"> Escalonamento - Gerente - Serviços acima de 60 minutos e até 90 minutos de atuação</font></p> 
        {html4}
        <p><font color="red"> Escalonamento - Superintendente - Serviços acima de 90 minutos de atuação</font></p> 
        {html6}

        <br> <br>
        <hr> <!-- Linha divisória -->

        <p><h3>SERVIÇOS EM TEMPO DE CHEGADA</h3></p> 
        <p>Monitoramento - Consolidado</p>
        {html3}
        <p> Monitoramento - Detalhado </p>
        {html9}

        <br> <br>
        <hr> <!-- Linha divisória -->

        <p><h2><font color="red">Tempo de Experiência do cliente - Visão dia:</font> {tempoExperienciaFinal}<h2></p>
        <p><h4> Obs.: Somente estamos considerando serviços aceitos e sem agendamentos </p><h4>
        </body>
        </html>"""

        # Envia o e-mail
        mail.Send()

        print('E-mail do primeiro processo enviado com sucesso!')

    else:
        print("Não enviar e-mail")
