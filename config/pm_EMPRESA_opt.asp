<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_EMPRESA_opt.asp
'CRIADO EM.........: 23/3/2006 14:57:05
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
window.location = "pm_EMPRESA.asp";
}
function Validator(theForm)
{
  if (theForm.cp_A01END.value == "")
  {
    alert("O campo ENDEREÇO não pode ser vazio!");
    theForm.cp_A01END.focus();
    return (false);
  }
  if (theForm.cp_A02FONE.value == "")
  {
    alert("O campo FONE não pode ser vazio!");
    theForm.cp_A02FONE.focus();
    return (false);
  }
  if (theForm.cp_A03CEP.value == "")
  {
    alert("O campo CEP não pode ser vazio!");
    theForm.cp_A03CEP.focus();
    return (false);
  }
  if (theForm.cp_A04CIDADE.value == "")
  {
    alert("O campo CIDADE/UF não pode ser vazio!");
    theForm.cp_A04CIDADE.focus();
    return (false);
  }
  if (theForm.cp_A05TEXTO.value == "")
  {
    alert("O campo TEXTO não pode ser vazio!");
    theForm.cp_A05TEXTO.focus();
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
                        <form action="pm_EMPRESA_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">ENDEREÇO<br></span><INPUT class=ud_caixa style="WIDTH:360px" name=cp_A01END></P></TD>
       <TD><P><span class="aFontePT">FONE<br></span><INPUT class=ud_caixa style="WIDTH:360px" name=cp_A02FONE></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CEP<br></span><INPUT class=ud_caixa style="WIDTH:360px" name=cp_A03CEP></P></TD>
       <TD><P><span class="aFontePT">CIDADE/UF<br></span><INPUT class=ud_caixa style="WIDTH:360px" name=cp_A04CIDADE></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">TEXTO<br></span><TEXTAREA class=ud_caixa style="WIDTH:360px" name=cp_A05TEXTO rows=8></textarea></P></TD>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then %>
<%
On Error Resume Next
sNot = request.querystring("sF")
sSel = "Select * From EMPRESA Where A00ID =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_A01END                  = Trim(objRS("A01END"))
va_A01END                  = Replace(va_A01END,"<BR>", chr(13),1)
'-----------------------------------------------
va_A02FONE                 = Trim(objRS("A02FONE"))
va_A02FONE                 = Replace(va_A02FONE,"<BR>", chr(13),1)
'-----------------------------------------------
va_A03CEP                  = Trim(objRS("A03CEP"))
va_A03CEP                  = Replace(va_A03CEP,"<BR>", chr(13),1)
'-----------------------------------------------
va_A04CIDADE               = Trim(objRS("A04CIDADE"))
va_A04CIDADE               = Replace(va_A04CIDADE,"<BR>", chr(13),1)
'-----------------------------------------------
va_A05TEXTO                = Trim(objRS("A05TEXTO"))
va_A05TEXTO                = Replace(va_A05TEXTO,"<BR>", chr(13),1)
%>
                            <table align="center" border=1 cellspacing=0 width="775" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_EMPRESA_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">ENDEREÇO<br></span><INPUT class=ud_caixa style="WIDTH:360px" value='<% = va_A01END %>' name=cp_A01END></P></TD>
       <TD><P><span class="aFontePT">FONE<br></span><INPUT class=ud_caixa style="WIDTH:360px" value='<% = va_A02FONE %>' name=cp_A02FONE></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CEP<br></span><INPUT class=ud_caixa style="WIDTH:360px" value='<% = va_A03CEP %>' name=cp_A03CEP></P></TD>
       <TD><P><span class="aFontePT">CIDADE/UF<br></span><INPUT class=ud_caixa style="WIDTH:360px" value='<% = va_A04CIDADE %>' name=cp_A04CIDADE></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">TEXTO<br></span><TEXTAREA class=ud_caixa style="WIDTH:360px" name=cp_A05TEXTO rows=8><% = va_A05TEXTO %></textarea></P></TD>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then %>
<%
sNot = request.querystring("sF")
sSel = "Delete From EMPRESA Where A00ID =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_EMPRESA.asp"
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
va_A01END                  = Trim(UploadRequest.Item("cp_A01END").Item("Value"))
va_A01END                  = Replace(va_A01END,chr(13),"<BR>" ,1)
'-----------------------------------------------
va_A02FONE                 = Trim(UploadRequest.Item("cp_A02FONE").Item("Value"))
va_A02FONE                 = Replace(va_A02FONE,chr(13),"<BR>" ,1)
'-----------------------------------------------
va_A03CEP                  = Trim(UploadRequest.Item("cp_A03CEP").Item("Value"))
va_A03CEP                  = Replace(va_A03CEP,chr(13),"<BR>" ,1)
'-----------------------------------------------
va_A04CIDADE               = Trim(UploadRequest.Item("cp_A04CIDADE").Item("Value"))
va_A04CIDADE               = Replace(va_A04CIDADE,chr(13),"<BR>" ,1)
'-----------------------------------------------
va_A05TEXTO                = Trim(UploadRequest.Item("cp_A05TEXTO").Item("Value"))
va_A05TEXTO                = Replace(va_A05TEXTO,chr(13),"<BR>" ,1)

If sTipo = "1" Then
sNot = request.querystring("sF")
   sSel = sSel & "INSERT INTO EMPRESA(A01END, A02FONE, A03CEP, A04CIDADE, A05TEXTO)"
   sSel = sSel & "VALUES ('" & va_A01END & "', '" & va_A02FONE & "', '" & va_A03CEP & "', '" & va_A04CIDADE & "', '" & va_A05TEXTO & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_EMPRESA.asp"

ElseIf sTipo = "2" Then

'***********************************************
sNot = request.querystring("sF")
   sSel = sSel & "UPDATE EMPRESA SET "
   sSel = sSel & "A01END = '" & va_A01END & "', A02FONE = '" & va_A02FONE & "', A03CEP = '" & va_A03CEP & "', A04CIDADE = '" & va_A04CIDADE & "', A05TEXTO = '" & va_A05TEXTO & "' "
   sSel = sSel & "WHERE A00ID = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_EMPRESA.asp"
Else
   sSel = sSel & "UPDATE EMPRESA SET "
   sSel = sSel & "A01END = '" & va_A01END & "', A02FONE = '" & va_A02FONE & "', A03CEP = '" & va_A03CEP & "', A04CIDADE = '" & va_A04CIDADE & "', A05TEXTO = '" & va_A05TEXTO & "' "
   sSel = sSel & "WHERE A00ID = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_EMPRESA.asp"
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

