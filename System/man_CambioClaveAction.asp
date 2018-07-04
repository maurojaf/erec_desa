<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>

<% ' Capa 1 ' %>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/odbc/insertUpdateA.inc"-->
<!--#include file="../lib/asp/comunes/odbc/ObtenerRecordset.inc"-->

<% ' Capa 2 ' %>
<!--#include file="../lib/asp/comunes/insert/Usuario.inc"-->
<!--#include file="../lib/asp/comunes/recordset/Usuario.inc"-->
<!--#include file="../lib/asp/comunes/general/funciones.inc"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

strUsuario = session("session_idusuario")
AbrirSCG()

	strClave = Trim(Request("CLAVE_NEW1"))
	strSql="UPDATE USUARIO SET CLAVE = '" & strClave &"' WHERE ID_USUARIO = " & strUsuario
	Conn.execute(strSql)


CerrarSCG()
%>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<FORM name=mantenedorForm method=post>
</FORM>
<SCRIPT Language=JavaScript>
	alert ('Cambio de clave exitoso');
	mantenedorForm.action = "man_CambioClave.asp";
	mantenedorForm.submit();
</SCRIPT>

