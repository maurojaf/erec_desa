<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->

<%



Response.CodePage 	=65001
Response.charset	="utf-8"

strOrigen		=Request.querystring("strOrigen")
calle 			=UCASE(LTRIM(RTRIM(request.querystring("calle"))))
numero 			=LTRIM(RTRIM(request.querystring("numero")))
resto 			=UCASE(LTRIM(RTRIM(request.querystring("resto"))))
strContacto 	= request("TX_CONTACTO")
strApellido		= request("TX_APELLIDO")
strCargo 		= request("TX_CARGO")
strDpto 		= request("TX_DPTO")
comuna 			=request.querystring("comuna")
rut 			=request.querystring("rut")
strDesde    	=Trim(request.querystring("TX_DESDE"))
strHasta 		=Trim(request.querystring("TX_HASTA"))
strDiasAtencion =Trim(request.querystring("strDiasAtencion"))
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

if resto="" then
resto=" "
end if


abrirscg()

strSql="SELECT DISTINCT ID_DIRECCION FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & rut & "' AND CALLE = '" & calle & "' AND NUMERO = '" & numero & "' AND RESTO = '" & resto & "' AND COMUNA = '" & comuna & "'"
set rsTel= Conn.execute(strSql)
'REsponse.write "strSql=" & strSql


ssql="execute SCG_WEB_NUEVA_DIR '" & rut & "','" & calle & "','" & numero & "','" & resto & "','" & comuna & "','" & session("session_login") & "','" & UCASE(strContactoCargo) & "','" & strDesde & "','" & strHasta & "','" & strDiasAtencion & "'," & strTipoContacto &""
'Response.write "ssql=" & ssql
'Response.End
Conn.execute(ssql)

cerrarscg()

If strOrigen = "" Then
	'strEnlace="mas_direcciones.asp?rut=" & rut
Else
	'strEnlace="deudor_direcciones.asp?strRUT_DEUDOR=" & rut & "&strOrigen=" & strOrigen
End If

%>