<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"


accion_ajax 		=request.querystring("accion_ajax")

abrirscg()

if trim(accion_ajax)="actualiza_CB_EMAIL_GESTION" then
	rut		=request.querystring("rut")

%>

<select name="CB_EMAIL_GESTION" id="CB_EMAIL_GESTION" onchange="set_CB_CONTACTO_ASOCIADO_EMAIL(this.value); return false;"  onchange="this.style.width=260">


	<option value="0">SELECCIONE</option>
  <%
	AbrirSCG1()
	ssql_ = "SELECT ID_EMAIL, UPPER(EMAIL) AS EMAIL FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & rut & "' AND ESTADO <> 2"
	set rsEmail=Conn1.execute(ssql_)
	Do until rsEmail.eof
		strEmailCB = rsEmail("EMAIL")
		strSel=""
		if strEmailCB = strEmailAgestionar Then strSel = "SELECTED"
		%>
		<option value="<%=rsEmail("ID_EMAIL")%>" <%=strSel%>><%=rsEmail("EMAIL")%></option>
		<%
			rsEmail.movenext
	Loop
	rsEmail.close
	set rsEmail=nothing
	CerrarSCG1()
 %>
</select>


<%
end if
cerrarscg()

%>

