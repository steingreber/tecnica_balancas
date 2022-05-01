<!--#include file="_cnx.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Tecnica Balanças</title>
<link type="text/css" rel="stylesheet" href="suport/suport.css"></head>

<body background="images/tb_18.jpg" bgcolor="#E6E6E6" bgproperties="fixed">
<%
Dim iCateg, iCod

iCateg = Request("categ")
iCod   = Request("cod")

Select case iCateg
Case 1
	iNome = "SISTEMA DE PESAGEM"
Case 2
	iNome = "INDÚSTRIA"
Case 3
	iNome = "CÉLULAS DE CARGAS E INDICADORES DE PESAGEM"
Case 4
	iNome = "COMÉRCIO"
Case 5
	iNome = "FARMÁCIA E LABORATÓRIO"
Case 6
	iNome = "AGRONEGÓCIO"
Case 7
	iNome = "ASSISTÊNCIA TÉCNICA"
Case Else
	Response.Redirect "principal.htm"
End Select %>

<p class="titulotext"><%= iNome %></p>
<%
SQL = "SELECT CATEG.A01CATEG, PRODUTOS.A02DESC, PRODUTOS.A00ID, PRODUTOS.A01NOME, PRODUTOS.A03FOTO, CATEG.A00ID"
SQL = SQL & " FROM CATEG INNER JOIN PRODUTOS ON CATEG.A00ID = PRODUTOS.A04CATEG"
SQL = SQL & " WHERE (((PRODUTOS.A00ID)=" & iCod & "))"
Set RS = objConect.Execute(SQL)
%>
<table width="459" cellspacing="0" cellpadding="0" align="center">
    <tr>
        <td width="449" align="left"><span class="bodytext"><%= UCase(RS("A01NOME")) %></span></td>
    </tr>
    <tr>
        <td width="449" align="center"><br><img src="config/fotos/<%= RS("A03FOTO") %>" border="0"></td>
    </tr>
    <tr>
        <td width="449" height="30"><br><p align="justify"><span class="bodytext"><strong><%= RS("A02DESC") %></strong></span></p></td>
    </tr>
</table>
<p align="center"><a href="#" onClick="history.go(-1)"><img src="images/bt_voltar.gif" alt="" width="80" height="16" border="0"></a></p>
<br><br>
<hr>
<%
RS.Close
Set RS = Nothing
objConect.Close
%>
</body>
</html>
