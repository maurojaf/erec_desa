<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<% ' Capa 1 ' %>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	IntId= session("ses_codcli")
	IntId= 1100
%>

<HTML>
<HEAD><TITLE>Mantenedor de Clientes</TITLE>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
</HEAD>
<link href="../css/style.css" rel="Stylesheet">

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




