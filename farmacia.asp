<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Tecnica balanças</title>
<link type="text/css" rel="stylesheet" href="suport/suport.css"></head>

<body background="images/tb_18.jpg" bgcolor="#E6E6E6" bgproperties="fixed">
<%
dim iCateg, iNome
iCateg = Request"categ")

SQL = "SELECT CATEG.A01CATEG, PRODUTOS.A01NOME, PRODUTOS.A03FOTO, CATEG.A00ID"
SQL = SQL & " FROM CATEG INNER JOIN PRODUTOS ON CATEG.A00ID = PRODUTOS.A04CATEG"
SQL = SQL & " WHERE (((CATEG.A00ID)=" & iCateg & "))"

Select case iCateg
Case 1
	iNome = "SISTEMA DE PESAGEM"
Case 2
	iNome = "INDUSTRIA"
Case 3
	iNome = "CÉLULAS DE CARGAS E INDICADORES DE PESAGEM"
Case 4
	iNome = "COMÉRCIO"
Case 5
	iNome = "FARMÁCIA E LABORATÓRIO"
Case 6
	iNome = "AGRONEGÓCIO"
Case Else
	Response.Redirect "principal.htm"
End Select

%>

<p class="titulotext"><%= UCase(iNome) %></p>
<table width="497" height="93" cellpadding="0" cellspacing="0">
<%
Set RS = objConect.Execute(SQL)
Do While Not RS.EOF
%>
    <tr>
        <td width="92" height="87" align="center"><img src="config/fotos/<%= RS("A03FOTO") %>" width="69" height="69" border="0"></td>
        <td width="389" height="87"><%= RS("A01NOME") %></td>
    </tr>
</table>
<%
RS.MoveNext
Loop
RS.Close
Set RS = Nothing
%>

</body>
</html>
