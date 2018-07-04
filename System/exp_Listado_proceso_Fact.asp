<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/lib.asp"-->
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>
<script language="JavaScript" type="text/JavaScript">
function AbreArchivo(nombre){
window.open(nombre,"INFORMACION","width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes");
}
</script>
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

	strCliente = REQUEST("CB_CLIENTE")
	strEstado = REQUEST("CB_ESTADO")
	dtmFechaProc = REQUEST("CB_FECHA")

	'Response.write "strCliente=" & strCliente
	'Response.write "strEstado=" & strEstado
	'Response.write "dtmFechaProc=" & dtmFechaProc

	AbrirSCG()

	strSql="SELECT FORMULA_HONORARIOS_FACT,FORMULA_HONORARIOS,FORMULA_INTERESES FROM CLIENTE WHERE COD_CLIENTE = '" & strCliente & "'"
	set rsDET=Conn.execute(strSql)
	if Not rsDET.eof Then
		strNomFormHonFact = ValNulo(rsDET("FORMULA_HONORARIOS_FACT"),"C")
		strNomFormHon = ValNulo(rsDET("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsDET("FORMULA_INTERESES"),"C")
	Else
		strNomFormHon = "NO_DEFINIDA"
		strNomFormInt = "NO_DEFINIDA"
	end if






'Server.ScriptTimeout = 9000

'Conn.ConnectionTimeout = 9000

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivo = "export_Proceso_Facturacion_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivo
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivo

	''Response.write "terceroCSV=" & terceroCSV


	''terceroCSV = "F:\" & strNomArchivo

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero=""

	strTextoTercero = "IID_CUOTA;NRO_SAP;INTERLOCUTOR;DOCUMENTO;CLIENTE;TIPO_DOCUMENTO;FECHA_ASIG;VENCIMIENTO;RUT_DEUDOR;NOMBRE_DEUDOR;SUCURSAL;CAPITAL;ESTADO_DEUDA;FECHA_ESTADO;COMPROBANTE;HONORARIO;RUT_CLIENTE;RAZON_SOCIAL"

	fichCA.writeline(strTextoTercero)


	strSql = "SELECT 	ID_CUOTA, CUOTA.COD_CLIENTE, UPPER(LOGIN) AS USUARIO, CUOTA.RUT_DEUDOR AS RUT_DEUDOR,"
	strSql = strSql & "	NRO_DOC,FECHA_ENVIO_VISAR, MONTO_VISACION, FECHA_ENVIO_FACTURAR, MONTO_FACTURACION, NUMERO_FACTURA,"
	strSql = strSql & "	ESTADO_FACTURA, USUARIO_ESTADO_FACT, OBSERVACION_FACTURACION, CONVERT(VARCHAR(10),CUOTA.FECHA_VENC,103) AS FECHA_VENC,"
	strSql = strSql & "	SUCURSAL, DEUDOR.NOMBRE_DEUDOR AS NOMDEUDOR, ESTADO_DEUDA.DESCRIPCION AS DESCRIPT,"
	strSql = strSql & "	HONORARIOS_FACT = (CASE WHEN '" & strEstado & "' IN (1) THEN dbo." & strNomFormHonFact & "(ID_CUOTA)"
	strSql = strSql & "							WHEN '" & strEstado & "' IN (2) THEN CUOTA.MONTO_VISACION"
	strSql = strSql & "							WHEN '" & strEstado & "' IN (3) THEN CUOTA.MONTO_FACTURACION"
	strSql = strSql & "					   END),"
	strSql = strSql & "	VALOR_CUOTA,"
	strSql = strSql & "	ISNULL(CUOTA.NRO_CLIENTE_DEUDOR,'NO INGRESADO') AS INTERLOCUTOR,"
	strSql = strSql & "	CLIENTE.DESCRIPCION,"
	strSql = strSql & "	NOM_TIPO_DOCUMENTO,"
	strSql = strSql & "	ISNULL(CUOTA.NRO_CLIENTE_DOC,'NO INGRESADO') AS NRO_CLIENTE_DOC,"
	strSql = strSql & "	CONVERT(VARCHAR(10),FECHA_CREACION,103) AS FECHA_ASIGNACION,"
	strSql = strSql & "	CONVERT(VARCHAR(10),FECHA_ESTADO,103) AS FECHA_ESTADO,"
	strSql = strSql & "	CONVERT(VARCHAR(10),FECHA_VENC,103) AS FECHA_VENC,"
	strSql = strSql & "	ISNULL(CONVERT(VARCHAR(10),FECHA_FACTURACION,103),'NO FACTURADO') AS FECHA_FACTURACION,"
	strSql = strSql & "	ISNULL(CUOTA.ADIC_92,'NO INGRESADO') AS CP,"
	strSql = strSql & "	ISNULL(SEDE.RUT,'NO DEFINIDO') AS RUT_CLIENTE,"
	strSql = strSql & "	ISNULL(SEDE.RAZON_SOCIAL,'NO DEFINIDO') AS RAZON_SOCIAL"

	strSql = strSql & " FROM CUOTA  INNER JOIN CLIENTE ON CUOTA.COD_CLIENTE = CLIENTE.COD_CLIENTE"
	strSql = strSql & " 			LEFT JOIN USUARIO ON CUOTA.USUARIO_ESTADO_FACT = USUARIO.ID_USUARIO"
	strSql = strSql & " 			LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
	strSql = strSql & " 			LEFT JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR  AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE"
	strSql = strSql & " 			LEFT JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql & " 			LEFT JOIN SEDE ON CUOTA.COD_CLIENTE = SEDE.COD_CLIENTE AND CUOTA.SUCURSAL = SEDE.SEDE"

	strSql = strSql & " WHERE  		(CUOTA.COD_CLIENTE =  '" & strCliente & "' AND ESTADO_FACTURA = '1' AND '" & Trim(strEstado) & "' IN (2) AND CAST( '" & dtmFechaProc & "' AS DATETIME) = FECHA_ENVIO_VISAR)"
	strSql = strSql & " 		OR  (CUOTA.COD_CLIENTE = '" & strCliente & "' AND ESTADO_FACTURA = '2' AND '" & strEstado & "' IN (3) AND CAST( '" & dtmFechaProc & "' AS DATETIME) = FECHA_ENVIO_FACTURAR)"
	strSql = strSql & " 		OR  (CUOTA.COD_CLIENTE = '" & strCliente & "' AND '" & strEstado & "' IN (1) AND (ESTADO_FACTURA IS NULL OR ESTADO_FACTURA IN ('4','5','7')) AND CUOTA.ESTADO_DEUDA = 3 AND (CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = ''))"

	''Response.write "strSql = " & strSql
	''Response.write "strSql = " & strEstado

		''Response.write "strSql = " & strSql

		set rsDet=Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsDet.Eof

		strTextoTercero = rsDet("ID_CUOTA")& ";" &rsDet("NRO_CLIENTE_DOC")& ";" & rsDet("INTERLOCUTOR")& ";" &rsDet("NRO_DOC")& ";" &rsDet("DESCRIPCION")& ";" &rsDet("NOM_TIPO_DOCUMENTO")& ";" & rsDet("FECHA_ASIGNACION")& ";" &rsDet("FECHA_VENC")& ";" & rsDet("RUT_DEUDOR")& ";" &rsDet("NOMDEUDOR")& ";" & rsDet("SUCURSAL")& ";" & rsDet("VALOR_CUOTA")& ";" & rsDet("DESCRIPT")& ";" & rsDet("FECHA_ESTADO")& ";" & rsDet("CP")& ";" & rsDet("HONORARIOS_FACT")& ";" & rsDet("RUT_CLIENTE")& ";" & rsDet("RAZON_SOCIAL")
		cantSiniestroC = cantSiniestroC + 1

		fichCA.writeline(strTextoTercero)

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

