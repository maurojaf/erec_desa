<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->

<%

	Response.CodePage=65001
	Response.charset ="utf-8"

	
rut=request("rut")
strOrigen=request("strOrigen")


abrirscg()
ssql="SELECT CORRELATIVO, ANEXO FROM DEUDOR_EMAIL WHERE RUT_DEUDOR='"&rut&"'"
set rsDIR=Conn.execute(ssql)
do until rsDIR.eof
	estado_correlativo=request("radiomail"+cstr(rsDIR("CORRELATIVO")))

	strAnexo=Trim(request("TX_ANEXO_" & Trim(rsDIR("CORRELATIVO"))))

	CORRELATIVO=rsDIR("CORRELATIVO")

	If Trim(estado_correlativo) <> "" and Not IsNull(estado_correlativo) Then
		ssql2="UPDATE DEUDOR_EMAIL SET ESTADO=" & Trim(estado_correlativo) & ", ANEXO = '" & strAnexo & "' WHERE RUT_DEUDOR='"& rut &"' and CORRELATIVO=" & CORRELATIVO
		''Response.write "<br>ssql2=" & ssql2
		Conn.execute(ssql2)
	End If

rsDIR.movenext
loop

'Response.End

rsDIR.close
set rsDIR=nothing
cerrarscg()
%>


<script language="JavaScript" type="text/JavaScript">
	alert('Datos de email actualizados');
	<%If strOrigen="deudor_email" Then%>
		window.navigate('deudor_email.asp?strRUT_DEUDOR=<%=rut%>&strFonoAgestionar=<%=strFonoAgestionar%>');
	<%Else%>
		window.navigate('mas_correos.asp?rut=<%=rut%>');
	<%End if%>

</script>
