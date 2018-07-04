<% @LCID = 1034 %>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"

rut 				=request("rut")
strOrigen 			=request("strOrigen")

strAnexo  			=request("strAnexo")
IF(request("strTipoContacto") = "") then strTipoContacto = "null" else strTipoContacto = request("strTipoContacto") end if
estado_correlativo 	=request("estado_correlativo")
CORRELATIVO 		=request("CORRELATIVO")
TX_HASTA 			=request("TX_HASTA")
TX_DESDE 			=request("TX_DESDE")
strDiasAtencion 	=request("strDiasAtencion")

intIdUsuario        =session("session_idusuario")

abrirscg()


	ssql2="UPDATE DEUDOR_DIRECCION SET ESTADO='"&cint(estado_correlativo)&"', RESTO = '" & strAnexo & "', IdTipoContacto = "& strTipoContacto &",HORA_DESDE = '" & TX_DESDE & "', HORA_HASTA = '" & TX_HASTA & "', DIAS_PAGO = '" & strDiasAtencion & "', FECHA_REVISION = GETDATE(), USR_REVISION = " & intIdUsuario 
	ssql2 = ssql2 & " WHERE RUT_DEUDOR='"&rut&"' and CORRELATIVO='"&CORRELATIVO&"'"
	
	'Response.write "<br>ssql2=" & ssql2 &"<br>"
	'Response.end
	 
	Conn.execute(ssql2)



cerrarscg()

%>

