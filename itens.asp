<!--#include file="_cnx.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Tecnica balanças</title>
<link type="text/css" rel="stylesheet" href="suport/suport.css"></head>

<body background="images/tb_18.jpg" bgcolor="#E6E6E6" bgproperties="fixed">
<%
dim iCateg, iNome
iCateg = Request("categ")

SQL = "SELECT PRODUTOS.A00ID as idProd, CATEG.A01CATEG, PRODUTOS.A01NOME, PRODUTOS.A04FOTOP, CATEG.A00ID"
SQL = SQL & " FROM CATEG INNER JOIN PRODUTOS ON CATEG.A00ID = PRODUTOS.A04CATEG"
SQL = SQL & " WHERE (((CATEG.A00ID)=" & iCateg & "))"

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
	Response.Redirect "principal.asp"
End Select

%>

<p class="titulotext"><%= UCase(iNome) %><br><span class="mintext">Clique na imagem para mais detalhes.</span></p>
<table width="477" height="93" cellpadding="0" cellspacing="0">
<%
Set RS = objConect.Execute(SQL)
Do While Not RS.EOF
%>
    <tr>
        <td width="92" height="123" align="center"><a href="itens_view.asp?cod=<%= RS("idProd") %>&categ=<%= iCateg %>"><img src="config/fotos/<%= RS("A04FOTOP") %>" width="121" height="121" border="0"></a></td>
        <td width="369" height="123">&nbsp;&nbsp;<span class="bodytext"><%= UCase(RS("A01NOME")) %></span></td>
    </tr>
<%
RS.MoveNext
Loop
RS.Close
Set RS = Nothing
objConect.Close
%>
</table>
</body>
</html>
