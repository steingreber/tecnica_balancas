<!--#include file="_cnx.asp"-->
<HTML>
<HEAD>
<TITLE>Técnica Balanças Eletrônicas </TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
<SCRIPT TYPE="text/javascript">
<!--

function newImage(arg) {
	if (document.images) {
		rslt = new Image();
		rslt.src = arg;
		return rslt;
	}
}

function changeImages() {
	if (document.images && (preloadFlag == true)) {
		for (var i=0; i<changeImages.arguments.length; i+=2) {
			document[changeImages.arguments[i]].src = changeImages.arguments[i+1];
		}
	}
}

var preloadFlag = false;
function preloadImages() {
	if (document.images) {
		tb_03_over = newImage("images/tb_03-over.jpg");
		tb_04_over = newImage("images/tb_04-over.jpg");
		tb_05_over = newImage("images/tb_05-over.jpg");
		tb_06_over = newImage("images/tb_06-over.jpg");
		tb_07_over = newImage("images/tb_07-over.jpg");
		tb_07_tb_03_over = newImage("images/tb_07-tb_03_over.jpg");
		tb_08_over = newImage("images/tb_08-over.jpg");
		tb_08_tb_07_over = newImage("images/tb_08-tb_07_over.jpg");
		tb_08_tb_03_over = newImage("images/tb_08-tb_03_over.jpg");
		tb_09_over = newImage("images/tb_09-over.jpg");
		tb_09_tb_08_over = newImage("images/tb_09-tb_08_over.jpg");
		tb_10_over = newImage("images/tb_10-over.jpg");
		tb_10_tb_09_over = newImage("images/tb_10-tb_09_over.jpg");
		tb_11_over = newImage("images/tb_11-over.jpg");
		preloadFlag = true;
	}
}

// -->
</SCRIPT>
<!-- End Preload Script -->
<link type="text/css" rel="stylesheet" href="suport/suport.css">
</HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 ONLOAD="preloadImages();">
<TABLE WIDTH=779 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD COLSPAN=3 ROWSPAN=3>
			<IMG SRC="images/tb_01.jpg" WIDTH=214 HEIGHT=95></TD>
		<TD COLSPAN=9 width="565" height="52" bgcolor="black" align="center">
			<font size="5" face="Verdana" color="white"><b>HÁ <%= Year(date) - 1994 %> ANOS OFERECENDO QUALIDADE!</b></font></TD>
	</TR>
	<TR>
		<TD>
			<A HREF="principal.asp" target="quadro"
				ONMOUSEOVER="changeImages('tb_03', 'images/tb_03-over.jpg', 'tb_07', 'images/tb_07-tb_03_over.jpg', 'tb_08', 'images/tb_08-tb_03_over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_03', 'images/tb_03.jpg', 'tb_07', 'images/tb_07.jpg', 'tb_08', 'images/tb_08.jpg'); return true;">
				<IMG NAME="tb_03" SRC="images/tb_03.jpg" WIDTH=117 HEIGHT=22 BORDER=0></A></TD>
		<TD COLSPAN=2>
			<A HREF="itens.asp?categ=1" target="quadro"
				ONMOUSEOVER="changeImages('tb_04', 'images/tb_04-over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_04', 'images/tb_04.jpg'); return true;">
				<IMG NAME="tb_04" SRC="images/tb_04.jpg" WIDTH=122 HEIGHT=22 BORDER=0></A></TD>
		<TD>
			<A HREF="itens.asp?categ=2" target="quadro"
				ONMOUSEOVER="changeImages('tb_05', 'images/tb_05-over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_05', 'images/tb_05.jpg'); return true;">
				<IMG NAME="tb_05" SRC="images/tb_05.jpg" WIDTH=70 HEIGHT=22 BORDER=0></A></TD>
		<TD COLSPAN=5>
			<A HREF="itens.asp?categ=3" target="quadro"
				ONMOUSEOVER="changeImages('tb_06', 'images/tb_06-over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_06', 'images/tb_06.jpg'); return true;">
				<IMG NAME="tb_06" SRC="images/tb_06.jpg" WIDTH=256 HEIGHT=22 BORDER=0></A></TD>
	</TR>
	<TR>
		<TD>
			<A HREF="itens.asp?categ=7" target="quadro"
				ONMOUSEOVER="changeImages('tb_07', 'images/tb_07-over.jpg', 'tb_08', 'images/tb_08-tb_07_over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_07', 'images/tb_07.jpg', 'tb_08', 'images/tb_08.jpg'); return true;">
				<IMG NAME="tb_07" SRC="images/tb_07.jpg" WIDTH=117 HEIGHT=21 BORDER=0></A></TD>
		<TD>
			<A HREF="itens.asp?categ=4" target="quadro"
				ONMOUSEOVER="changeImages('tb_08', 'images/tb_08-over.jpg', 'tb_09', 'images/tb_09-tb_08_over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_08', 'images/tb_08.jpg', 'tb_09', 'images/tb_09.jpg'); return true;">
				<IMG NAME="tb_08" SRC="images/tb_08.jpg" WIDTH=69 HEIGHT=21 BORDER=0></A></TD>
		<TD COLSPAN=3>
			<A HREF="itens.asp?categ=5" target="quadro"
				ONMOUSEOVER="changeImages('tb_09', 'images/tb_09-over.jpg', 'tb_10', 'images/tb_10-tb_09_over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_09', 'images/tb_09.jpg', 'tb_10', 'images/tb_10.jpg'); return true;">
				<IMG NAME="tb_09" SRC="images/tb_09.jpg" WIDTH=152 HEIGHT=21 BORDER=0></A></TD>
		<TD COLSPAN=2>
			<A HREF="itens.asp?categ=6" target="quadro"
				ONMOUSEOVER="changeImages('tb_10', 'images/tb_10-over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_10', 'images/tb_10.jpg'); return true;">
				<IMG NAME="tb_10" SRC="images/tb_10.jpg" WIDTH=115 HEIGHT=21 BORDER=0></A></TD>
		<TD COLSPAN=2>
			<A HREF="fale.asp" target="quadro"
				ONMOUSEOVER="changeImages('tb_11', 'images/tb_11-over.jpg'); return true;"
				ONMOUSEOUT="changeImages('tb_11', 'images/tb_11.jpg'); return true;">
				<IMG NAME="tb_11" SRC="images/tb_11.jpg" WIDTH=112 HEIGHT=21 BORDER=0></A></TD>
	</TR>
	<TR>
		<TD COLSPAN=12 width="779" height="14" bgcolor="#E6E6E6"></TD>
	</TR>
	<TR>
		<TD COLSPAN=9>
			<IMG SRC="images/tb_13.jpg" WIDTH=649 HEIGHT=2></TD>
		<TD COLSPAN=2 ROWSPAN=2>
			<a href="itens.asp?categ=6" target="quadro"><IMG SRC="images/tb_14.jpg" WIDTH=115 HEIGHT=117 border="0" alt="AGRONEGÓCIO"></a></TD>
		<TD ROWSPAN=5 width="15" height="474" bgcolor="#E6E6E6">
			&nbsp;</TD>
	</TR>
	<TR>
		<TD ROWSPAN=4 width="15" height="472" bgcolor="#E6E6E6">
			&nbsp;</TD>
		<TD>
			<a href="itens.asp?categ=7" target="quadro"><IMG SRC="images/tb_17.jpg" WIDTH=108 HEIGHT=115 border="0" alt="ASSISTÊNCIA TÉCNICA"></a></TD>
		<TD COLSPAN=7 ROWSPAN=4 width="526" height="472" background="images/tb_18.jpg">
<iframe src="principal.asp" name="quadro" id="quadro" width="525" height="471" frameborder="0"></iframe>
		</TD>
	</TR>
	<TR>
		<TD>
			<a href="itens.asp?categ=4" target="quadro"><IMG SRC="images/tb_19.jpg" WIDTH=108 HEIGHT=120 border="0" alt="COMÉRCIO"></a></TD>
		<TD COLSPAN=2>
			<a href="itens.asp?categ=1" target="quadro"><IMG SRC="images/tb_20.jpg" WIDTH=115 HEIGHT=120 border="0" alt="SISTEMA DE PESAGEM"></a></TD>
	</TR>
	<TR>
		<TD>
			<a href="itens.asp?categ=2" target="quadro"><IMG SRC="images/tb_21.jpg" WIDTH=108 HEIGHT=117 border="0" alt="Industria"></a></TD>
		<TD COLSPAN=2>
			<a href="itens.asp?categ=5" target="quadro"><IMG SRC="images/tb_22.jpg" WIDTH=115 HEIGHT=117 border="0" alt="FARMÁCIA"></a></TD>
	</TR>
	<TR>
		<TD>
			<a href="itens.asp?categ=3" target="quadro"><IMG SRC="images/tb_23.jpg" WIDTH=108 HEIGHT=120 border="0" alt="CÉLULAS DE CARGAS E INDICADORES DE PESAGEM"></a></TD>
		<TD COLSPAN=2>
			<a href="fale.asp" target="quadro"><IMG SRC="images/tb_24.jpg" WIDTH=115 HEIGHT=120 border="0" alt="FALE CONOSCO"></a></TD>
	</TR>
    <TR>
        <TD COLSPAN=12 width="779" align="center" bgcolor="#E6E6E6">
            <p><font size="1" face="Arial"><%= RSemp("A01END") %> CEP <%= RSemp("A03CEP") %> - TELEFONE <%= RSemp("A02FONE") %><br>
<a href="mailto:tecnica@tecnicabalancas.com.br">tecnica@tecnicabalancas.com.br</a></font></p>
        </TD>
    </TR>
	<TR>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=15 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=108 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=91 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=117 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=69 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=53 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=70 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=29 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=97 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=18 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=97 HEIGHT=1></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=15 HEIGHT=1></TD>
	</TR>
</TABLE>
<span class="mintext">CopyRight 2006 - Técnica Balanças Eletrônicas&nbsp;/ Designer e programação <a href="http://www.gnove.com.br/" target="_blank">GNove WebStudio</a></span>
</BODY>
</HTML>
<%
RSemp.Close
Set RSemp = Nothing
objConect.Close
Set objConect = Nothing
%>