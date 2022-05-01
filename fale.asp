<!--#include file="_cnx.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>FALE CONOSCO</title>
<link type="text/css" rel="stylesheet" href="suport/suport.css"></head>
<script Language=JavaScript>
function Validator(theForm)
{
  if (theForm.cp5SEUNOME.value == "")
  {
    alert("O campo SEU NOME não pode ser vazio!");
    theForm.cp5SEUNOME.focus();
    return (false);
  }
  if (theForm.cp3ASSUNTO.value == "")
  {
    alert("O campo ASSUNTO não pode ser vazio!");
    theForm.cp3ASSUNTO.focus();
    return (false);
  }
  if (theForm.cp2SEUEMAIL.value == "")
  {
    alert("O campo SEU EMAIL não pode ser vazio!");
    theForm.cp2SEUEMAIL.focus();
    return (false);
  }
  if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(theForm.cp2SEUEMAIL.value)){
  return (true)
  }
  {
    alert("O SEU EMAIL não é válido!");
    theForm.cp2SEUEMAIL.focus();
    return (false);
  }
  return (true);
}
</script>
<body bgcolor="#E6E6E6" background="images/tb_18.jpg">


<p class="titulotext">FALE CONOSCO</p>
<p class="bodytext"><STRONG><span class="navtext">Entre em contato conosco por e-mail, correspondência 
ou telefone.</span></STRONG><BR><span class="navtext">Endereço: <%= RSemp("A01END") %></span><BR><span class="navtext"><%= RSemp("A04CIDADE") %> - 
CEP <%= RSemp("A03CEP") %></span><BR><span class="navtext">Fone: </span><%= RSemp("A02FONE") %><BR><span class="navtext">Email: <a href="mailto:tecnica@tecnicabalancas.com.br">tecnica@tecnicabalancas.com.br</a></span></p>
<hr>
<form action="informes_ok.asp" method="post" name="contatop" id="contatop" onSubmit="return Validator(this);">
		<table border="0">
	<tr><td class="texto_personalizacao"><span class="navtext">Seu Nome:</span></td>
	<td><input type="text" name="cp5SEUNOME" class="ud_caixa"></td></tr>
	<tr><td class="texto_personalizacao"><span class="navtext">Seu Email:</span></td>
	<td><input type="text" name="cp2SEUEMAIL" class="ud_caixa"></td></tr>
	<tr><td class="texto_personalizacao"><span class="navtext">Assunto:</span></td>
	<td><input type="text" name="cp3ASSUNTO" class="ud_caixa"></td></tr>
	<tr><td class="texto_personalizacao" valign="top"><span class="navtext">Mensagem:</span></td>
	<td><textarea cols="40" rows="4" name="pminfor" class="ud_caixa"></textarea></td></tr>
	<tr><td colspan="2">
	</td></tr>
	<tr>
		<td align="center" colspan="2">
			<input type="submit" value="Enviar" class="ud_caixa"> <input type="reset" value="Limpar" class="ud_caixa">
		</td>
	</tr>
		</table>
</form>
		</td>
    </tr>
</table>
<%
RSemp.close
Set RSemp = Nothing
%>
</body>
</html>
