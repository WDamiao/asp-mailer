<%
' Receber dados do formulário
nome = Request.Form("nome")
email = Request.Form("email")
mensagem = Request.Form("mensagem")

Set objMail = Server.CreateObject("CDO.Message")
Set objConfig = Server.CreateObject("CDO.Configuration")

' Configuração SMTP autenticada
With objConfig.Fields
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingPort
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "email-ssl.com.br"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 ' cdoBasic
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "admin@macedoff.com.br"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Teste@30303030"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .Update
End With

Set objMail.Configuration = objConfig

' Criar o e-mail
With objMail
    .From = "admin@macedoff.com.br"
    .To = "suportelocaweb09@gmail.com"
    .Subject = "Formulário de Contato"
    .HTMLBody = "<b>Nome:</b> " & nome & "<br>" & _
                "<b>Email:</b> " & email & "<br>" & _
                "<b>Mensagem:</b><br>" & Replace(mensagem, vbCrLf, "<br>")
    .Send
End With

' Limpar
Set objMail = Nothing
Set objConfig = Nothing

Response.Write "Mensagem enviada com sucesso!"
%>