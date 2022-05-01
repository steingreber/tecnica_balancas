<%
Dim sUser, sPass, vUser, vPass
sUser = Request("username")
sPass = Request("password")
vUser = "admin"
vPass = "124578"
If sUser <> vUser or sPass <> vPass Then
   response.redirect "default.asp"
Else
   Session("xUser") = True
   Session("xPass") = True
   response.redirect "pm_PRODUTOS.asp"
End If
%>

