<%
'<------- /같같같같같같같같같같같같같\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_CATEG_cnx.asp
'CRIADO EM.........: 21/3/2006 11:06:23
'-------------------------------------------------------------------------------
   Set objConect = Server.CreateObject("ADODB.Connection")
   strConn = "Provider=MSDataShape;DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & server.mappath("TEC009IOKE\TEC00949IRKF843DKI.mdb") & ";UID=;PWD="
   objConect.Open strConn
   session.LCID=1046

Dim asPathImage
'caminho da pasta das imagens
asPathImage = "" 'digite o caminho aqui - Server.MapPath("")
%>

