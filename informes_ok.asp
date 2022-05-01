<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ /  Gerando com PageMaster   \ ------>
'<----- /  http:\\www.pagemaster.tk   \ ----->
'<----- \_____________________________/ ----->

'NOME DO ARQUIVO...: usu_receberInform_INC.ASP
'CRIADO EM.........: 15/7/2003 18:04:37
'---------------------------------------------

'On Error Resume Next
    mru01usuario = request("cp5SEUNOME")
    mru02email   = request("cp2SEUEMAIL")
    mru03assunto = request("cp3ASSUNTO")
    mru04msg     = request("pminfor")

    sBodyText = sBodyText & "<font face='Courier New' color='#003300'>" & vbCrLf
	sBodyText = sBodyText & "NOME......: " & mru01usuario & "<br>" & vbCrLf
	sBodyText = sBodyText & "EMAIL.....: " & mru02email & "<br>" & vbCrLf
	sBodyText = sBodyText & "ASSUNTO...: " & mru03assunto & "<br>" & vbCrLf
	sBodyText = sBodyText & "MENSAGEM..: " & mru04msg & "<br>" & vbCrLf
    sBodyText = sBodyText & "<br>" & vbCrLf
    sBodyText = sBodyText & "<hr><br>" & vbCrLf
    sBodyText = sBodyText & "<font face=Verdana size=1 color='#0000FF'><a href='http://www.pagemaster.tk'>Gentileza www.pagemaster.tk</a></font>" & vbCrLf
    sBodyText = sBodyText & "</font>"

    On Error Resume Next
    Set pmMail = Server.CreateObject("CDONTS.NewMail")
    pmMail.mailFormat = 0
    pmMail.bodyFormat = 0
    pmMail.From = mru02email
    pmMail.To   = "tecnica@tecnicabalancas.com.br"
    pmMail.Subject = "Contato do site Tecnica Balanças"
    pmMail.Body = sBodyText
    pmMail.Send
    Set pmMail = Nothing
If Err <> 0 Then
    sMsgErr = "Ocorreu o seguinte erro ao tentar enviar o e-mail:" & Err.Description
End If
    On Error GoTo 0

response.redirect "informes_ok.htm"
%>
