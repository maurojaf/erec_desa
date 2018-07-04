<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->


<%
Response.CodePage = 65001
Response.charset="utf-8"

strOrigen 		= Request("strOrigen")
COD_AREA 		= UCASE(Trim(request("COD_AREA")))
numero 			= Trim(request("numero"))
rut 			= request("rut")
strContacto 	= request("TX_CONTACTO")
strApellido		= request("TX_APELLIDO")
strCargo 		= request("TX_CARGO")
strDpto 		= request("TX_DPTO")
strFuente 		= request("CB_FUENTE")
strDesde 		= Trim(request("TX_DESDE"))
strHasta 		= Trim(request("TX_HASTA"))
strDiasAtencion = Trim(request("CH_DIAS"))
strAnexo 		= request("TX_ANEXO")
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

strSql="SELECT DISTINCT ID_TELEFONO FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" & rut & "' AND COD_AREA = '" & COD_AREA & "' AND TELEFONO = '" & numero & "'"
set rsTel= Conn.execute(strSql)
'REsponse.write "strSql=" & strSql&"<br>"
If Not rsTel.eof Then
%>
	<script>
		alert('Telefono ya existe, no puede ser ingresado o si no se logra visualizar, favor revisar telefonos no validos');
		history.back()
	</script>

<%
	Response.End
End If



strSql="EXEC SCG_WEB_NUEVO_TEL '" & rut & "','" & COD_AREA & "','" & numero & "','" & session("session_login") & "','" & strAnexo & "','" & strFuente & "','" & UCASE(strContactoCargo) & "','" & strDesde & "','" & strHasta & "','" & strDiasAtencion & "',"& strTipoContacto &""
'Response.write strSql
Conn.execute(strSql)
cerrarscg()



%>
