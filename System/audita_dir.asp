<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>

<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->

<%

Response.CodePage=65001
Response.charset ="utf-8"
	
rut=request("rut")
cliente=session("ses_codcli")
strOrigen=request("strOrigen")


abrirscg()
ssql="SELECT CORRELATIVO FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR='"&rut&"'"
set rsDIR=Conn.execute(ssql)
do until rsDIR.eof
	estado_correlativo=request.Form("radiodir"+cstr(rsDIR("CORRELATIVO")))
	CORRELATIVO=rsDIR("CORRELATIVO")

	''Response.write "<br>estado_correlativo=" & estado_correlativo
	If Trim(estado_correlativo) <> "" and Not IsNull(estado_correlativo) Then

	Resto=Trim(request("TX_ANEXO_" & Trim(rsDIR("CORRELATIVO"))))
	strDesde=Trim(request("TX_DESDE_" & Trim(rsDIR("CORRELATIVO"))))
	strHasta=Trim(request("TX_HASTA_" & Trim(rsDIR("CORRELATIVO"))))
	strDiasAtencion=Trim(request("CH_DIAS_" & Trim(rsDIR("CORRELATIVO"))))


	ssql2="UPDATE DEUDOR_DIRECCION SET ESTADO='"&cint(estado_correlativo)&"' WHERE RUT_DEUDOR='"&rut&"' and CORRELATIVO='"&CORRELATIVO&"'"
	'Response.write "<br>ssql2=" & ssql2
	Conn.execute(ssql2)

	strSql = "UPDATE DEUDOR_DIRECCION SET RESTO = '" & resto & "',HORA_DESDE = '" & strDesde & "', HORA_HASTA = '" & strHasta & "', DIAS_PAGO = '" & strDiasAtencion & "' WHERE RUT_DEUDOR = '" & rut & "' AND CORRELATIVO = '" & CORRELATIVO & "'"
	'Response.write "<br>strSql=" & strSql
	Conn.execute(strSql)

	End If

rsDIR.movenext
loop

rsDIR.close
set rsDIR=nothing
cerrarscg()

'Response.write "<br>ssql=" & ssql
'Response.End
%>


<script language="JavaScript" type="text/JavaScript">
	alert('Datos de direcciones actualizadas');
	<%If strOrigen="deudor_direcciones" Then%>
		window.navigate('deudor_direcciones.asp?strRUT_DEUDOR=<%=rut%>&strFonoAgestionar=<%=strFonoAgestionar%>');
	<%Else%>
		window.navigate('mas_direcciones.asp?rut=<%=rut%>');
	<%End if%>



		//window.navigate('principal.asp?rut=<%=rut%>');
</script>
