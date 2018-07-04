<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"


strOrigen 			=request("strOrigen")
accion_ajax 		=request("accion_ajax")
intIdUsuario        =session("session_idusuario")
IF(request("strTipoContacto") = "") then strTipoContacto = "null" else strTipoContacto = request("strTipoContacto") end if


abrirscg()

if trim(accion_ajax)="auditar_email" then
	
	rut 				=request("rut")
	CORRELATIVO 		=request("CORRELATIVO")
	strAnexo 			=request("strAnexo")
	estado_correlativo 	=request("estado_correlativo")

	If Trim(estado_correlativo) <> "" and Not IsNull(estado_correlativo) Then
		ssql2="UPDATE DEUDOR_EMAIL"
		ssql2 = ssql2 & " SET ESTADO=" & Trim(estado_correlativo) & ", ANEXO = '" & strAnexo & "', IdTipoContacto = "& strTipoContacto &", FECHA_REVISION = GETDATE(), USR_REVISION = " & intIdUsuario
		ssql2 = ssql2 & " WHERE RUT_DEUDOR='"& rut &"' and CORRELATIVO=" & CORRELATIVO
		
		'Response.write "<br>ssql2=" & ssql2
		
		Conn.execute(ssql2)

	End If

end if


cerrarscg()
%>

