<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/lib.asp"-->
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>
<script language="JavaScript" type="text/JavaScript">
function AbreArchivo(nombre){
window.open(nombre,"INFORMACION","width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes");
}
</script>
<html xmlns="http:www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
<title>CRM RSA</title>
<style type="text/css">
<!--body {	background-color: #cccccc;}-->
</style>
</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<%

	If Trim(Request("Limpiar"))="1" Then
		session("session_RUT_DEUDOR") = ""
		rut = ""
	End if

	If Trim(Request("TX_RUT")) = "" Then
		strRUT_DEUDOR = session("session_RUT_DEUDOR")
	Else
		strRUT_DEUDOR = Trim(Request("TX_RUT"))
		session("session_RUT_DEUDOR") = strRUT_DEUDOR
	End If

	intOrigen = Request("intOrigen")

	intTipoPP = Request("CB_TIPO")

	If Trim(intTipoPP) = "RP" or Trim(intTipoPP) = "RC" or Trim(intTipoPP) = "RL" Then
		intTipoPP = "CONV"
	End If

	''Response.write "intTipoPP=" & intTipoPP

	intCliente = session("ses_codcli")

AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

	strSql = "SELECT CONVERT(VARCHAR(10),GETDATE(),108) AS HORA"

	set rsHra=Conn.execute(strSql)

	strHora= rsHra("HORA")


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivo = "export_plan_de_pago " & strRUT_DEUDOR &" "& Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivo
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivo

	''Response.write "terceroCSV=" & terceroCSV


	''terceroCSV = "F:\" & strNomArchivo

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero=""

	strTextoTercero = "RUT_DEUDOR;NOMBRE_DEUDOR;RUT_CLIENTE;NOMBRE_CLIENTE;NRO_FACTURA;CUOTA;VENCIMIENTO;DIAS_MORA;TIPO_DOC;ABONADO;MONTO_PENDIENTE;                 ;RECEPCION FACTURA(S/N);RECEPCION NOTIFICACION(S/N);FECHA_PAGO_FACTURA;MONTO_PAGO;OBSERVACION (MENCIONAR NOTAS DE CREDITO ASOCIADAS, ETC.)"

	fichCA.writeline(strTextoTercero)


		strSql = "SELECT ID_CUOTA, NRO_DOC, NRO_CUOTA, NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO, GASTOS_PROTESTOS, CUENTA, FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,SALDO,"
		strSql = strSql & " CUOTA.RUT_DEUDOR,DEUDOR.NOMBRE_DEUDOR,RUT_SUBCLIENTE,NOMBRE_SUBCLIENTE, NRO_CUOTA,(VALOR_CUOTA-SALDO) ABONO,"
		strSql = strSql & " DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, CUOTA.CUSTODIO, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS "
		strSql = strSql & " FROM CUOTA, TIPO_DOCUMENTO,DEUDOR WHERE CUOTA.RUT_DEUDOR='" & strRUT_DEUDOR & "' AND CUOTA.COD_CLIENTE='" & intCliente & "' AND SALDO > 0 AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO AND DEUDOR.RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND DEUDOR.COD_CLIENTE = '" & intCliente & "'"

		strSql = strSql & " ORDER BY RUT_SUBCLIENTE, FECHA_VENC ASC"

		set rsDet=Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0

	Do While Not rsDet.Eof

	strObjeto = "CH_" & rsDet("ID_CUOTA")
	strObjeto1 = "TX_SALDO_" & rsDet("ID_CUOTA")

	If UCASE(Request(strObjeto)) = "ON" Then

		strTextoTercero = rsDet("RUT_DEUDOR")& ";" &rsDet("NOMBRE_DEUDOR")& ";" &rsDet("RUT_SUBCLIENTE")& ";" &rsDet("NOMBRE_SUBCLIENTE")& ";" &rsDet("NRO_DOC")& ";" &rsDet("NRO_CUOTA")& ";" &rsDet("FECHA_VENC")& ";" &rsDet("ANTIGUEDAD")& ";" &rsDet("TIPO_DOCUMENTO")& ";" &rsDet("SALDO")

		cantSiniestroC = cantSiniestroC + 1

		fichCA.writeline(strTextoTercero)

	End If

		rsDet.movenext

	Loop



	%>
	<table>
	<tr><td>Cantidad de registros generados : <%= cantSiniestroC %></td></tr>
	<tr><td>
	<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivo%>')">Descargar</a>
	&nbsp;
	<a href="#" onClick="history.back()">Volver</a>


	</td></tr>


	</table>
 <%


	'conectamos con el FSO
	set confile = createObject("scripting.filesystemobject")
	'creamos el objeto TextStream

	'response.write "terceroCSV=" & terceroCSV
	'response.End

	''set fichCA = confile.CreateTextFile(terceroCSV)
	''fichCA.write(strTextoTercero)
	fichCA.close()


%>







				</td>
			  </tr>
			</table>


		</td>

	</tr>

</table>

</body>
</html>

