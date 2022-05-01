<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_PRODUTOS_opt.asp
'CRIADO EM.........: 21/3/2006 11:08:28
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "default.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "Novo Registro"
ElseIf acao = "2" Then
   sNome = "Editar Registro"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>.::: PageMaster 2.0 :::.</title>
<SCRIPT language=Javascript>

function go2url() {
window.location = "pm_PRODUTOS.asp";
}
function Validator(theForm)
{
  if (theForm.cp_A04CATEG.value == "")
  {
    alert("O campo CATEGORIA não pode ser vazio!");
    theForm.cp_A04CATEG.focus();
    return (false);
  }
  if (theForm.cp_A01NOME.value == "")
  {
    alert("O campo NOME não pode ser vazio!");
    theForm.cp_A01NOME.focus();
    return (false);
  }
  if (theForm.cp_A02DESC.value == "")
  {
    alert("O campo DESCRIÇÃO não pode ser vazio!");
    theForm.cp_A02DESC.focus();
    return (false);
  }
  return (true);
}
function fotoZoom(img){
  foto1= new Image();
  foto1.src=(img);
  Controlla(img);
}
function Controlla(img){
  if((foto1.width!=0)&&(foto1.height!=0)){
    viewFoto(img);
  }
  else{
    funzione="Controlla('"+img+"')";
    intervallo=setTimeout(funzione,20);
  }
}
function viewFoto(img){
  largh=foto1.width+20;
  altez=foto1.height+20;
  stringa="width="+largh+",height="+altez;
  finestra=window.open(img,"",stringa);
}
</SCRIPT>
<meta name="robots" content="none">
<link rel="stylesheet" href="pagemaster.css">
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000" leftmargin=0 topmargin=0>
<table cellpadding="0" cellspacing="0" width="779" align="center">
    <tr>
        <td " class="tabelatitulo" align="center" height="50">
            <p>-:- <%= sNome %> -:-</p>
        </td>
    </tr>
    <tr>
        <td height="5"" class="tabelaitem">
        </td>
    </tr>
    <tr>
        <td" align="center" valign="top">
            <table align="center" border="1" cellspacing="0" width="775" bordercolordark="white" bordercolorlight="black">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table align="center" border=1 cellspacing=0 width="775" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_PRODUTOS_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">CATEGORIA<br></span>
           <select name="cp_A04CATEG" size="1" class="ud_caixa">
<%sLin = "Select * From CATEG"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option selected>Selecione</option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("A00ID") %>"><%= objRs("A01CATEG") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
       </P></TD>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT class=ud_caixa style="WIDTH:360px" name=cp_A01NOME></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">DESCRIÇÃO<br></span><TEXTAREA class=ud_caixa style="WIDTH:360px" name=cp_A02DESC rows=8></textarea></P></TD>
       <TD><P><span class="aFontePT">FOTO<br></span><INPUT class=ud_caixa style="WIDTH:360px" name=cp_A03FOTO type="file"></P></TD>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then %>
<%
On Error Resume Next
sNot = request.querystring("sF")
sSel = "Select * From PRODUTOS Where A00ID =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_A04CATEG                = Trim(objRS("A04CATEG"))
va_A04CATEG                = Replace(va_A04CATEG,"<BR>", chr(13),1)
'-----------------------------------------------
va_A01NOME                 = Trim(objRS("A01NOME"))
va_A01NOME                 = Replace(va_A01NOME,"<BR>", chr(13),1)
'-----------------------------------------------
va_A02DESC                 = Trim(objRS("A02DESC"))
va_A02DESC                 = Replace(va_A02DESC,"<BR>", chr(13),1)
'-----------------------------------------------
va_A03FOTO                 = Trim(objRS("A03FOTO"))
va_A03FOTO                 = Replace(va_A03FOTO,"<BR>", chr(13),1)
%>
                            <table align="center" border=1 cellspacing=0 width="775" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_PRODUTOS_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">CATEGORIA<br></span>
           <select name="cp_A04CATEG" size="1" class="ud_caixa">
<%
sLin = "Select * From CATEG"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option value="<%=va_A04CATEG%>" selected><%=va_A04CATEG%></option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("A00ID") %>"><%= objRs("A00ID") & "-" & objRs("A01CATEG") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
       </P></TD>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT class=ud_caixa style="WIDTH:360px" value='<% = va_A01NOME %>' name=cp_A01NOME></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">DESCRIÇÃO<br></span><TEXTAREA class=ud_caixa style="WIDTH:360px" name=cp_A02DESC rows=8><% = va_A02DESC %></textarea></P></TD>
       <TD><P><span class="aFontePT">FOTO - <a href="javascript:fotoZoom('<%= "fotos/" & va_A03FOTO%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:360px" name=cp_A03FOTO type="file"><BR></P>
       <INPUT type="hidden" name="hcp_A03FOTO" value="<%=va_A03FOTO%>">
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then %>
<%
sNot = request.querystring("sF")
sSel = "Delete From PRODUTOS Where A00ID =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_PRODUTOS.asp"
%>
<%
ElseIf acao = "4" Then
sNot = request.querystring("sF")
Response.Expires = 0
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest RequestBin
Dim sTipo, sCmd
sTipo  = request.querystring("tp")
%>
<%
'-----------------------------------------------
va_A04CATEG                = Trim(UploadRequest.Item("cp_A04CATEG").Item("Value"))
va_A04CATEG                = Replace(va_A04CATEG,chr(13),"<BR>" ,1)
'-----------------------------------------------
va_A01NOME                 = Trim(UploadRequest.Item("cp_A01NOME").Item("Value"))
va_A01NOME                 = Replace(va_A01NOME,chr(13),"<BR>" ,1)
'-----------------------------------------------
va_A02DESC                 = Trim(UploadRequest.Item("cp_A02DESC").Item("Value"))
va_A02DESC                 = Replace(va_A02DESC,chr(13),"<BR>" ,1)
'-----------------------------------------------
va_A03FOTO                 = Trim(UploadRequest.Item("cp_A03FOTO").Item("Value"))
va_A03FOTO                 = Replace(va_A03FOTO,chr(13),"<BR>" ,1)

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_A03FOTO <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM PRODUTOS WHERE A03FOTO = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_A03FOTO").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_A03FOTO").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM PRODUTOS WHERE A03FOTO = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_A03FOTO").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("fotos") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_A03FOTO = FileName0
End If

   sSel = sSel & "INSERT INTO PRODUTOS(A04CATEG, A01NOME, A02DESC, A03FOTO)"
   sSel = sSel & "VALUES ('" & va_A04CATEG & "', '" & va_A01NOME & "', '" & va_A02DESC & "', '" & va_A03FOTO & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_PRODUTOS.asp"

ElseIf sTipo = "2" Then

'***********************************************
va_A03FOTOh                 = Trim(UploadRequest.Item("hcp_A03FOTO").Item("Value"))
va_A03FOTOh                 = Replace(va_A03FOTOh,chr(13),"<BR>" ,1)
'***********************************************
sNot = request.querystring("sF")
'===============================================
If va_A03FOTO <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM PRODUTOS WHERE A03FOTO = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_A03FOTO").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_A03FOTO").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM PRODUTOS WHERE A03FOTO = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_A03FOTO").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("fotos") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_A03FOTO = FileName0
Else
va_A03FOTO = va_A03FOTOh
End If

   sSel = sSel & "UPDATE PRODUTOS SET "
   sSel = sSel & "A04CATEG = '" & va_A04CATEG & "', A01NOME = '" & va_A01NOME & "', A02DESC = '" & va_A02DESC & "', A03FOTO = '" & va_A03FOTO & "' "
   sSel = sSel & "WHERE A00ID = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_PRODUTOS.asp"
Else
   sSel = sSel & "UPDATE PRODUTOS SET "
   sSel = sSel & "A04CATEG = '" & va_A04CATEG & "', A01NOME = '" & va_A01NOME & "', A02DESC = '" & va_A02DESC & "' "
   sSel = sSel & "WHERE A00ID = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_PRODUTOS.asp"
End If
End If

Sub BuildUploadRequest(RequestBin)
PosBeg = 1
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(13)))
boundary = MidB(RequestBin, PosBeg, PosEnd - PosBeg)
BoundaryPos = InStrB(1, RequestBin, boundary)
Do Until (BoundaryPos = InStrB(RequestBin, boundary & getByteString("--")))
Dim UploadControl
Set UploadControl = CreateObject("Scripting.Dictionary")
Pos = InStrB(BoundaryPos, RequestBin, getByteString("Content-Disposition"))
Pos = InStrB(Pos, RequestBin, getByteString("name="))
PosBeg = Pos + 6
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(34)))
Name = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
PosFile = InStrB(BoundaryPos, RequestBin, getByteString("filename="))
PosBound = InStrB(PosEnd, RequestBin, boundary)
If PosFile <> 0 And (PosFile < PosBound) Then
PosBeg = PosFile + 10
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(34)))
FileName = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
UploadControl.Add "FileName", FileName
Pos = InStrB(PosEnd, RequestBin, getByteString("Content-Type:"))
PosBeg = Pos + 14
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(13)))
ContentType = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
UploadControl.Add "ContentType", ContentType
PosBeg = PosEnd + 4
PosEnd = InStrB(PosBeg, RequestBin, boundary) - 2
Value = MidB(RequestBin, PosBeg, PosEnd - PosBeg)
Else
Pos = InStrB(Pos, RequestBin, getByteString(Chr(13)))
PosBeg = Pos + 4
PosEnd = InStrB(PosBeg, RequestBin, boundary) - 2
Value = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
End If
UploadControl.Add "Value", Value
UploadRequest.Add Name, UploadControl
BoundaryPos = InStrB(BoundaryPos + LenB(boundary), RequestBin, boundary)
Loop
End Sub
'===============================================
Function getByteString(StringStr)
For i = 1 To Len(StringStr)
Char = Mid(StringStr, i, 1)
getByteString = getByteString & ChrB(AscB(Char))
Next
End Function

Function getString(StringBin)
getString = ""
For intCount = 1 To LenB(StringBin)
getString = getString & Chr(AscB(MidB(StringBin, intCount, 1)))
Next
End Function
'===============================================
%>
            </td>
    </tr>
</table>
<table cellpadding="0" cellspacing="0" width="779" align="center">
    <tr>
        <td class="tabelaitem" height="5" colspan="2">
        </td>
    </tr>
    <tr>
        <td height="21" class="tabelatitulo" colspan="2">
        </td>
    </tr>
    <tr>
        <td></td>
        <td align="center" valign="top">
                    </td>
    </tr>
</table>
</body>
</html>
<%
Sub Errou
If err <> 0 Then
%>
<br>
<table cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
        <td align="center" valign="middle">
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" align="center" cellpadding="0" cellspacing="0" width="571" bordercolordark="black" bordercolorlight="black">
                <tr>
                    <td width="571" height="63" align="center">
                        <p style="margin-left:5pt;"><span class="titulo">Informação do sistema...</span></p>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="128" align="center">
           <span class="titulo">Erro nº <%=err.Number%> <br><%=err.Source%><br><%= err.Description%></span>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="38" align="center">
           <span class="titulo"><a href="javascript:self.history.go(-1)">&lt;&lt; Voltar</a></span>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<br>
<%
Response.End
End If
End Sub
%>

