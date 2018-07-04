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

strTipoGestion = request("cmb_tipogestion")

	if strTipoGestion = "" then strTipoGestion = "0"

	termino = request("termino")
	inicio = request("inicio")
	intCOD_CLIENTE = session("ses_codcli")
	usuario = request("CB_EJECUTIVO")
	strCobranza = Request("CB_COBRANZA")

AbrirScg()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivo = "export_Busqueda_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivo
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivo

	''Response.write "terceroCSV=" & terceroCSV


	''terceroCSV = "F:\" & strNomArchivo

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero=""

	strTextoTercero = "PRIORIDAD;GESTION;UBICABILIDAD;ESTADO_UBICABILIDAD;RUT_CLIENTE;NOMBR_CLIENTE;RUT_DEUDOR;NOMBRE_DEUDOR;SALDO;DOC_ACTIVOS;DIAS_MORA;EJECUTIVO_ASIG;SEDE;FONOS_DEUDOR;EMAIL DEUDOR;DIRECCIONES_DEUDOR"

	fichCA.writeline(strTextoTercero)


		strSql = "select max([dbo].[fun_PrioridadCuotaDocActivo] (CUOTA.COD_CLIENTE,CUOTA.RUT_DEUDOR,0)) AS PRIORIDAD,"

		strSql = strSql & "	max(CASE WHEN GESTIONES_TIPO_GESTION.GESTION_MODULOS = 9 AND '" & strTipoGestion & "' IN (0,1) "
		strSql = strSql & "	THEN 'SOLICITUD BUSQUEDA'"
		strSql = strSql & " WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 5 AND '" & strTipoGestion & "' IN (0,2) )"
		strSql = strSql & "	THEN 'SOLICITUD NO RESP.'"
		strSql = strSql & " WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4 AND '" & strTipoGestion & "' IN (0,4) )"
		strSql = strSql & "	THEN 'INUBICABLE 2'"
		strSql = strSql & "	WHEN [dbo].[fun_ubicabilidad_telefono_email] (cuota.rut_deudor) ='INUBICABLE' AND '" & strTipoGestion & "' IN (0,3) "
		strSql = strSql & "	THEN 'INUBICABLE 1' "
		strSql = strSql & "	ELSE 'OTRO' "
		strSql = strSql & "	END) as GESTION,"

		strSql = strSql & "	UBICABILIDAD =MAX([dbo].[fun_ubicabilidad_tipo]  ('TODOS',CUOTA.RUT_DEUDOR)),"
		strSql = strSql & "	UBICABILIDAD2 =MAX([dbo].[fun_ubicabilidad_tipo]  ('TODOS2',CUOTA.RUT_DEUDOR)),"
		strSql = strSql & " RUT_CLIENTE = MAX(CUOTA.RUT_SUBCLIENTE),"
		strSql = strSql & " NOMCLIENTE = MAX(CUOTA.NOMBRE_SUBCLIENTE),"
		strSql = strSql & " RUT_DEUDOR = DEUDOR.RUT_DEUDOR,"
		strSql = strSql & " NOMBRE_DEUDOR = MAX(DEUDOR.NOMBRE_DEUDOR),"
		strSql = strSql & " SALDO = CAST(SUM (CUOTA.SALDO) AS BIGINT),"
		strSql = strSql & " DOC = CAST(COUNT(CUOTA.ID_CUOTA) AS INT),"
		strSql = strSql & " DM = MAX(CAST(GETDATE()-CUOTA.FECHA_VENC AS INT)),"
		strSql = strSql & " SEDE = MAX(CUOTA.SUCURSAL),"
		strSql = strSql & " MAX(USUARIO.LOGIN) as USUARIO,"
		strSql = strSql & " [dbo].[concatena_reg_ubic] (DEUDOR.RUT_DEUDOR,'TELEFONO') AS TELEFONOS,"
		strSql = strSql & " [dbo].[concatena_reg_ubic] (DEUDOR.RUT_DEUDOR,'MAIL') AS MAIL,"
		strSql = strSql & " [dbo].[concatena_reg_ubic] (DEUDOR.RUT_DEUDOR,'DIRECCION') AS DIRECCIONES"


		strSql = strSql & "	FROM CUOTA		 INNER JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE"
		strSql = strSql & "			   		 LEFT JOIN GESTIONES ON CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION"
		strSql = strSql & "					 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
		strSql = strSql & "			  		 LEFT JOIN USUARIO ON DEUDOR.USUARIO_ASIG = USUARIO.ID_USUARIO"
		strSql = strSql & "					 LEFT JOIN GESTIONES_TIPO_GESTION ON SUBSTRING(COD_ULT_GEST,1,1)  = GESTIONES_TIPO_GESTION.COD_CATEGORIA"
		strSql = strSql & "						  								 AND SUBSTRING(COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
		strSql = strSql & "						  								 AND SUBSTRING(COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
		strSql = strSql & "						 								 AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"

		strSql = strSql & " WHERE	(  		GESTIONES_TIPO_GESTION.GESTION_MODULOS = 9"
		strSql = strSql & " 			OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4"
		strSql = strSql & " 			OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 5 AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"
		strSql = strSql & " 			OR ([dbo].[fun_ubicabilidad_telefono_email] (cuota.rut_deudor) ='INUBICABLE' AND ISNULL(GESTIONES_TIPO_GESTION.GESTION_MODULOS,0) NOT IN (5,10)))"
		strSql = strSql & " 			AND ESTADO_DEUDA.ACTIVO = 1"
		strSql = strSql & "	           	AND CUOTA.COD_CLIENTE = '" & intCOD_CLIENTE & "'"

		If Trim(strCobranza) = "INTERNA" Then
			strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
			strParametro = "1"
		End if

		If Trim(strCobranza) = "EXTERNA" Then
			strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
			strParametro = "1"
		End if

		if Trim(strTipoGestion) = "1" Then

		strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 9 "

		End If

		if Trim(strTipoGestion) = "2" Then

		strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 5 "

		End If

		if Trim(strTipoGestion) = "4" Then

		strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4 "

		End If

		if Trim(strTipoGestion) = "3" Then

		strSql = strSql & " AND [dbo].[fun_ubicabilidad_telefono_email] (cuota.rut_deudor) ='INUBICABLE' "

		End If


		If Trim(usuario) <> "" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & usuario & "'"
		End if

		If inicio <> "" then

		strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) > = '" & inicio & " 00:00:00'"

		End If

		If termino <> "" then

		strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) < = '" & termino & " 23:59:59'"

		End If

		strSql = strSql & " GROUP BY "
		strSql = strSql & " deudor.RUT_DEUDOR"


		strSql = strSql & " ORDER BY GESTION DESC, PRIORIDAD ASC, MAX(CAST(GETDATE()-CUOTA.FECHA_VENC AS INT)) DESC, SALDO DESC, UBICABILIDAD"


		'Response.write "strSql = " & strSql

		set rsDet=Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0

	Do While Not rsDet.Eof


		strTextoTercero = rsDet("PRIORIDAD")& ";" &rsDet("GESTION")& ";" & rsDet("UBICABILIDAD")& ";" & rsDet("UBICABILIDAD2")& ";" & rsDet("RUT_CLIENTE")& ";" & rsDet("NOMCLIENTE")& ";" & rsDet("RUT_DEUDOR")& ";" & rsDet("NOMBRE_DEUDOR")& ";" & rsDet("SALDO")& ";" & rsDet("DOC")& ";" & rsDet("DM")& ";" & rsDet("USUARIO")& ";" & rsDet("SEDE")& ";" & rsDet("TELEFONOS")& ";" & rsDet("MAIL")& ";" & rsDet("DIRECCIONES")

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

