<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>

<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"
 
strOrigen 	= Request("strOrigen")
MAIL 		= UCASE(LTRIM(RTRIM(request("EMAIL"))))
rut 		= request("rut")
strContacto = request("TX_CONTACTO")
strApellido = request("TX_APELLIDO")
strCargo 	= request("TX_CARGO")
strDpto 	= request("TX_DPTO")
strFuente 	= request("CB_FUENTE")
IF(request("strTipoContacto") = "") then strTipoContacto = "null" else strTipoContacto = request("strTipoContacto") end if
strUsuarioIngresa = session("session_login")

If strContacto <> "" and strApellido <> "" and strCargo <> "" and strDpto <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strCargo & " /" & strDpto
ElseIf strContacto <> "" and strApellido <> "" and strCargo <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strCargo
ElseIf strContacto <> "" and strApellido <> "" and strDpto <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strDpto
ElseIf strContacto <> "" and strApellido <> "" Then
	strContactoCargo = strContacto & " /"& strApellido
Else strContactoCargo = strContacto
END IF

If Trim(strFuente) = "" Then
	strFuente = "DEUDOR - TERCERO"
End If

abrirscg()

strSql="SELECT DISTINCT ID_EMAIL FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & rut & "' AND EMAIL = '" & MAIL & "'"
set rsTel= Conn.execute(strSql)
'REsponse.write "strSql=" & strSql
If Not rsTel.eof Then
%>
	<script>
		alert('Email ya existe, no puede ser ingresado');
		carga_funcion_email();
	</script>

<%
	Response.End
End If

ssql="execute SCG_WEB_NUEVO_COR '" & rut & "','" & MAIL & "','" &session("session_login")  & "','" & UCASE(strContactoCargo) & "','" & strFuente & "'," & strTipoContacto &""
Conn.execute(ssql)

'REsponse.write "strSql=" & ssql
cerrarscg()

If strOrigen = "" Then
	'strEnlace="mas_correos.asp?rut=" & rut
Else
	'strEnlace="deudor_email.asp?strRUT_DEUDOR=" & rut & "&strOrigen=" & strOrigen
End If

%>


