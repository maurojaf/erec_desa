<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/lib.asp"-->
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
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
	strEstadoCC = request("CMB_ESTADO_CC")
	strCobranza = Request("CB_COBRANZA")

AbrirScg()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivo = "export_Normalizacion_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivo
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivo

	''Response.write "terceroCSV=" & terceroCSV


	''terceroCSV = "F:\" & strNomArchivo

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero=""

	strTextoTercero = "ESTADO REAL;Nº CP;FECHA PAGO;OBSERVACION;FECHA_INGRESO;FECHA_CONSULTA;ID_CUOTA;TIPO_GESTION;RUT_CLIENTE;NOMBR_CLIENTE;RUT_DEUDOR;NOMBRE_DEUDOR;NRO_DOC;NRO_CUOTA;TIPODO_DOC;FECHA_VENC;MONTO_CAPITAL;SALDO_ACTIVO;INTERLOCUTOR;SEDE;MONTO_PAGADO;FORMA_PAGO;FECHA_PAGO;LUGAR_PAGO;NRO_CP;OBSERVACION;EJECUTIVO;"

	fichCA.writeline(strTextoTercero)


	strSql = "SELECT CUOTA.ID_CUOTA AS ID_CUOTA,"
	strSql = strSql & " DEUDOR.NOMBRE_DEUDOR,"
	strSql = strSql & "	USUARIO.LOGIN AS EJEC_ASIG,"
	strSql = strSql & "	CAST(GETDATE()-CUOTA.FECHA_VENC AS INT) AS DM,"
	strSql = strSql & " CUOTA.RUT_SUBCLIENTE,"
	strSql = strSql & " CUOTA.NOMBRE_SUBCLIENTE,"
	strSql = strSql & " CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103) as 'FECHA_INGRESO',"
	strSql = strSql & " (CASE WHEN GESTIONES.NRO_DOC_PAGO = '' THEN 'NO ESPEC'ELSE GESTIONES.NRO_DOC_PAGO END) AS NRO_DOC_PAGO,"
	strSql = strSql & " 'FECHA CONSULTA: ' + ISNULL(CONVERT(VARCHAR(10),CUOTA.FECHA_CONSULTA_NORM,103),'NO CONSULTADO') AS FECHA_CONSULTA,"
	strSql = strSql & " ISNULL(CONVERT(VARCHAR(10),CUOTA.FECHA_CONSULTA_NORM,103),'NO CONSULTADO') AS FECHA_CONSULTA2,"
	strSql = strSql & " GESTIONSOLA = CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL)"
	strSql = strSql & " 				   THEN 'INDICA QUE PAGO'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP')"
	strSql = strSql & "                    THEN 'COMPROMISO D & T'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
	strSql = strSql & "                    THEN 'INDICA PAGO EN CONSULTA'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
	strSql = strSql & "                    THEN 'INDICA PAGO NO RESP.'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME))"
	strSql = strSql & "                    THEN 'REITERA INDICA PAGO'"
	strSql = strSql & " 				   ELSE 'PAGO NO APLICADO'"
	strSql = strSql & "					   END,"
	strSql = strSql & " CUOTA.RUT_DEUDOR AS RUT_DEUDOR,"
	strSql = strSql & " CUOTA.SUCURSAL AS SUCURSAL,"
	strSql = strSql & " CUOTA.NRO_DOC,"
	strSql = strSql & " CUOTA.NRO_CUOTA,"
	strSql = strSql & " TIPO_DOCUMENTO.NOM_TIPO_DOCUMENTO,"
	strSql = strSql & " CONVERT(VARCHAR(10),CUOTA.FECHA_VENC,103) as 'FECHA_VENC',"
	strSql = strSql & " CUOTA.VALOR_CUOTA,"
	strSql = strSql & " CUOTA.SALDO,"
	strSql = strSql & "	ISNULL(CUOTA.NRO_CLIENTE_DEUDOR,'NO INGRESADO') AS INTERLOCUTOR,"
	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES,CHAR(13),' '),CHAR(10),' ') as OBSERVACIONES,"

	strSql = strSql & " CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (1,11) )"
	strSql = strSql & " 	 THEN ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'')"
	strSql = strSql & " 	 WHEN ( GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (2) AND GESTIONES.FECHA_PAGO IS NOT NULL)"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103)"
	strSql = strSql & " 	 WHEN ( GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (6) AND GESTIONES.FECHA_PAGO IS NOT NULL)"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & "		 ELSE 'NO ESPEC'"
	strSql = strSql & " 	 END AS FECHA_NORMALIZACION,"

	strSql = strSql & " LUGAR_PAGO = UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))),"

	strSql = strSql & " ISNULL(CAJA_FORMA_PAGO.DESC_FORMA_PAGO,'NO ESPEC.') AS 'FORMA_PAGO',"

	strSql = strSql & " (CASE WHEN REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')<>'' THEN "
	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')"
	strSql = strSql & " ELSE 'SIN OBSERVACIÓN'"
	strSql = strSql & " END) as OBSERVACIONES_CAMPO,"

	strSql = strSql & " CASE WHEN ISNULL(GESTIONES.MONTO_CANCELADO,'')= 0"
	strSql = strSql & " THEN ''"
	strSql = strSql & " ELSE ISNULL(GESTIONES.MONTO_CANCELADO,'NO ESPEC')"
	strSql = strSql & " END AS 'MONTO_REGULARIZADO',"
	strSql = strSql & " '' AS VACIO"


	strSql = strSql & " FROM GESTIONES			INNER JOIN GESTIONES_CUOTA ON GESTIONES.ID_GESTION = GESTIONES_CUOTA.ID_GESTION"

	strSql = strSql & " LEFT JOIN FORMA_RECAUDACION RE ON RE.ID_FORMA_RECAUDACION= GESTIONES.ID_FORMA_RECAUDACION "
	strSql = strSql & " LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION= GESTIONES.ID_DIRECCION_COBRO_DEUDOR "


	strSql = strSql & "                         INNER JOIN CUOTA ON GESTIONES_CUOTA.ID_CUOTA = CUOTA.ID_CUOTA"
	strSql = strSql & "                         LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
	strSql = strSql & "                         INNER JOIN DEUDOR ON GESTIONES.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND GESTIONES.COD_CLIENTE = DEUDOR.COD_CLIENTE"
	strSql = strSql & "                         INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql & "                         LEFT JOIN CAJA_FORMA_PAGO ON GESTIONES.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO"
	strSql = strSql & "                         INNER JOIN GESTIONES_TIPO_GESTION ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_GESTION = GESTIONES_TIPO_GESTION.COD_GESTION AND"
	strSql = strSql & "                                                                             GESTIONES.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_CATEGORIA.COD_CATEGORIA"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_CATEGORIA AND"
	strSql = strSql & " 																	GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_SUB_CATEGORIA"
	strSql = strSql & " 						LEFT JOIN USUARIO ON CUOTA.USUARIO_ASIG = USUARIO.ID_USUARIO"


	strSql = strSql & " WHERE   (	 ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') AND '" & strTipoGestion & "' IN (0,2) AND GESTIONES.FECHA_COMPROMISO <= (GETDATE()))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 6 AND '" & strTipoGestion & "' IN (0,3) AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE () AND '" & strTipoGestion & "' IN (0,5))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND FECHA_CONSULTA_NORM IS NOT NULL AND '" & strTipoGestion & "' IN (0,6))"

	If Trim(strEstadoNorm) = "0" or Trim(strEstadoNorm) = "2" Then
		strSql = strSql & " OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2  AND FECHA_CONSULTA_NORM IS NULL AND '" & strTipoGestion & "' IN (0,1)))"
	Else
		strSql = strSql & " OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2  AND '" & strTipoGestion & "' IN (0,1)))"
	End If

	strSql = strSql & " 		  AND ESTADO_DEUDA.ACTIVO = 1 "

	If Trim(strCobranza) = "INTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
		strParametro = "1"
	End if

	If Trim(strCobranza) = "EXTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
		strParametro = "1"
	End if



	If Trim(strEstadoNorm) = "1" Then
		strSql = strSql & "		AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
	End If

	If inicio <> "" then

	strSql = strSql & " AND (CASE WHEN ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11))"
	strSql = strSql & " 	 THEN ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'')"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (2)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (6)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & "		 ELSE ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103),'')"
	strSql = strSql & " 	 END) > = CAST('" & inicio & " 00:00:00'AS DATETIME)"

	End If

	If termino <> "" then

	strSql = strSql & " AND (CASE WHEN ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11))"
	strSql = strSql & " 	 THEN ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'')"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (2)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (6)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & "		 ELSE ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103),'')"
	strSql = strSql & " 	 END) < = CAST('" & termino & " 23:59:59' AS DATETIME)"

	End If

	strSql = strSql & " AND CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION "
	strSql = strSql & "	AND CUOTA.COD_CLIENTE = '" & intCOD_CLIENTE & "'"


	if Trim(strTipoGestion) = "1" Then
		strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL)"
	End If

	if Trim(strTipoGestion) = "2" Then
		strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) "
		strSql = strSql & " 	AND  (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') AND GESTIONES.FECHA_COMPROMISO <= (GETDATE())"
	End If

	if Trim(strTipoGestion) = "3" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 6) "
	End If

	if Trim(strTipoGestion) = "4" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
	End If

	if Trim(strTipoGestion) = "5" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
	End If

	if Trim(strTipoGestion) = "6" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND FECHA_CONSULTA_NORM IS NOT NULL )"
	End If


	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
		strSql = strSql & " 	AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
	Else
		if Trim(usuario) <> "" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & usuario & "'"
		End if
	End if

	strSql = strSql & " ORDER BY  GESTIONSOLA, CUOTA.RUT_DEUDOR, FECHA_NORMALIZACION"

	'Response.write "strSql = " & strSql
	'ID_DIRECCION_COBRO_DEUDORresponse.end()
		set rsDet=Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsDet.Eof

		strTextoTercero = rsDet("VACIO")& ";" &rsDet("VACIO")& ";" &rsDet("VACIO")& ";" &rsDet("VACIO")& ";" &rsDet("FECHA_INGRESO")& ";" &rsDet("FECHA_CONSULTA2")& ";" & rsDet("ID_CUOTA")& ";" & rsDet("GESTIONSOLA")& ";" & rsDet("RUT_SUBCLIENTE") & ";" & rsDet("NOMBRE_SUBCLIENTE")& ";" & rsDet("RUT_DEUDOR") & ";" & rsDet("NOMBRE_DEUDOR")& ";" & rsDet("NRO_DOC")& ";" & rsDet("NRO_CUOTA") & ";" & rsDet("NOM_TIPO_DOCUMENTO")& ";" & rsDet("FECHA_VENC") & ";" & rsDet("VALOR_CUOTA")& ";" &rsDet("SALDO")& ";" &rsDet("INTERLOCUTOR")& ";" &rsDet("SUCURSAL")& ";" &rsDet("MONTO_REGULARIZADO")& ";" &rsDet("FORMA_PAGO")& ";" &rsDet("FECHA_NORMALIZACION")& ";" & rsDet("LUGAR_PAGO")& ";" &rsDet("NRO_DOC_PAGO")& ";" &rsDet("OBSERVACIONES_CAMPO")& ";" &rsDet("EJEC_ASIG")

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

