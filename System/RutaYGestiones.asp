<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
<% ' Capa 1 ' %>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	IntId= session("ses_codcli")
	IntId= 1100
%>

	<TITLE>Mantenedor de Clientes</TITLE>
	<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">

	<link href="../css/style.css" rel="Stylesheet">
</HEAD>
<BODY BGCOLOR='FFFFFF'>

<table width="500" border="0">
  <tr>
    <td colspan =2 bordercolor="#999999"  bgcolor="#380ACD" class="Estilo13" ALIGN="CENTER">ARCHIVOS DE GESTIONES</td>
  </tr>

<%
	intCorrelativo = 1




	AbrirScg()
		strSql = "SELECT * FROM ARCHIVO_GESTIONES WHERE ACTIVO = 1 ORDER BY FECHA DESC"
		set rsGest = Conn.execute(strSql )
		Do While not rsGest.eof
		%>
		<tr bgcolor="#FFFFFF" class="Estilo8">
			<td>
				Archivo Nro. <%=intCorrelativo%>
			</td>
			<td colspan =2>
				<a href="../Archivo/Otros/<%=IntId%>/Gestiones/<%=Trim(rsGest("NOMBRE_ARCHIVO"))%>"><%=Trim(rsGest("NOMBRE_ARCHIVO"))%></a>
			</td>
		</tr>


		<%
			intCorrelativo = intCorrelativo + 1
			rsGest.movenext
		Loop
	CerrarScg()
%>

<tr>
	<td colspan =2 bordercolor="#999999"  bgcolor="#380ACD" class="Estilo13" ALIGN="CENTER">&nbsp;</td>
</tr>

</table>

</FORM>
  </TD>
    </TR>
</TABLE>

</BODY>
</HTML>




