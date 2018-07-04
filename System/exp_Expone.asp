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

AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivo = "export_Expone_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivo
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivo

	''Response.write "terceroCSV=" & terceroCSV


	''terceroCSV = "F:\" & strNomArchivo

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero=""

	strTextoTercero = "FECHA_INGRESO;FECHA_CONSULTA;ID_CUOTA;TIPO_GESTION;RUT_CLIENTE;NOMBR_CLIENTE;RUT_DEUDOR;NOMBRE_DEUDOR;NRO_DOC;NRO_CUOTA;TIPODO_DOC;FECHA_VENC;SALDO_ACTIVO;INTERLOCUTOR;SEDE;MONTO_REGULARIZADO;FORMA_PAGO;FECHA_NORMALIZACION;LUGAR_PAGO;NRO_CP;OBSERVACION;EJECUTIVO;"

	fichCA.writeline(strTextoTercero)

		strSql = "SELECT DEUDOR.NOMBRE_DEUDOR, USUARIO.LOGIN AS EJEC_ASIG, "
		strSql = strSql & " CUOTA.ID_CUOTA AS ID_CUOTA,"
		strSql = strSql & " CUOTA.RUT_SUBCLIENTE,"
		strSql = strSql & " CUOTA.NOMBRE_SUBCLIENTE,"
		strSql = strSql & " CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103) as 'FECHA_INGRESO', (CASE WHEN GESTIONES.NRO_DOC_PAGO = '' THEN 'NO ESPEC'ELSE GESTIONES.NRO_DOC_PAGO END) AS NRO_DOC_PAGO,ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103),'NO ESPEC') AS FECHA_NORMALIZACION,"
		strSql = strSql & " 'FECHA CONSULTA: ' + ISNULL(CONVERT(VARCHAR(10),CUOTA.FECHA_CONSULTA_NORM,103),'NO CONSULTADO') AS FECHA_CONSULTA,"
		strSql = strSql & " ISNULL(CONVERT(VARCHAR(10),CUOTA.FECHA_CONSULTA_NORM,103),'NO CONSULTADO') AS FECHA_CONSULTA2,"
		strSql = strSql & " GESTIONSOLA = CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND FECHA_CONSULTA_NORM IS NULL)"
		strSql = strSql & " 				   THEN 'EXPONE REQUERIMIENTO'"
		strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
		strSql = strSql & "                    THEN 'EXPONE EN CONSULTA'"
		strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
		strSql = strSql & "                    THEN 'EXPONE NO RESPONDIDO'"
		strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME))"
		strSql = strSql & "                    THEN 'REITERA EXPONE REQUERIMIENTO'"
		strSql = strSql & " 				   ELSE 'OTRO'"
		strSql = strSql & "					   END,"

		strSql = strSql & " CUOTA.RUT_DEUDOR AS RUT_DEUDOR,"
		strSql = strSql & " CUOTA.NRO_DOC,"
		strSql = strSql & " CUOTA.NRO_CUOTA,"
		strSql = strSql & " TIPO_DOCUMENTO.NOM_TIPO_DOCUMENTO,"
		strSql = strSql & " CONVERT(VARCHAR(10),CUOTA.FECHA_VENC,103) as 'FECHA_VENC',"
		strSql = strSql & " CUOTA.SALDO,"
		strSql = strSql & "	ISNULL(CUOTA.NRO_CLIENTE_DEUDOR,'NO INGRESADO') AS INTERLOCUTOR,"
		strSql = strSql & " CUOTA.SUCURSAL AS SUCURSAL,"

		strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES,CHAR(13),' '),CHAR(10),' ') as OBSERVACIONES,"

		strSql = strSql & " (CASE WHEN REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')<>'' THEN "
		strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')"
		strSql = strSql & " ELSE 'SIN OBSERVACIÓN'"
		strSql = strSql & " END) as OBSERVACIONES_CAMPO, "
		strSql = strSql & " ISNULL(UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))),'NO ESPEC') AS LUGAR_PAGO,"

		strSql = strSql & " ISNULL(CAJA_FORMA_PAGO.DESC_FORMA_PAGO,'NO ESPEC.') AS 'FORMA_PAGO',"

		strSql = strSql & " CASE WHEN ISNULL(GESTIONES.MONTO_CANCELADO,'')= 0"
		strSql = strSql & " THEN ''"
		strSql = strSql & " ELSE ISNULL(GESTIONES.MONTO_CANCELADO,'NO ESPEC.')"
		strSql = strSql & " END AS 'MONTO_REGULARIZADO',"
		strSql = strSql & " GESTIONES_TIPO_GESTION.DESCRIPCION AS DESCRIP"


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


		strSql = strSql & " WHERE   (  (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND FECHA_CONSULTA_NORM IS NULL AND '" & strTipoGestion & "' IN (0,1))"
		strSql = strSql & " 	    OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE () AND '" & strTipoGestion & "' IN (2))"
		strSql = strSql & " 	    OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ()) AND '" & strTipoGestion & "' IN (0,3))"

		strSql = strSql & " 		AND ESTADO_DEUDA.ACTIVO = 1 "

		If inicio <> "" then

		strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) > = '" & inicio & " 00:00:00'"

		End If

		If termino <> "" then

		strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) < = '" & termino & " 23:59:59'"

		End If

		strSql = strSql & " 		AND CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION "
		strSql = strSql & "			AND CUOTA.COD_CLIENTE = '" & intCOD_CLIENTE & "'"


		if Trim(strTipoGestion) = "1" Then
			strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND FECHA_CONSULTA_NORM IS NULL)"
		End If

		if Trim(strTipoGestion) = "2" Then
			strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ()) "
		End If

		if Trim(strTipoGestion) = "3" Then
				strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
		End If


		If Trim(usuario) <> "" Then
				strSql = strSql & " 	AND  DEUDOR.USUARIO_ASIG = '" & usuario & "'"
		End if

		strSql = strSql & " ORDER BY  USUARIO.login,GESTIONES.FECHA_INGRESO"

	'Response.write "strSql = " & strSql
	'response.end()

	set rsTemp= Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsTemp.Eof

		strTextoTercero = rsTemp("FECHA_INGRESO")& ";" & rsTemp("FECHA_CONSULTA2")& ";" & rsTemp("ID_CUOTA")& ";" & rsTemp("DESCRIP")& ";" & rsTemp("RUT_SUBCLIENTE") & ";" & rsTemp("NOMBRE_SUBCLIENTE")& ";" & rsTemp("RUT_DEUDOR") & ";" & rsTemp("NOMBRE_DEUDOR")& ";" & rsTemp("NRO_DOC")& ";" & rsTemp("NRO_CUOTA") & ";" & rsTemp("NOM_TIPO_DOCUMENTO")& ";" & rsTemp("FECHA_VENC") & ";" & rsTemp("SALDO")& ";" &rsTemp("INTERLOCUTOR")& ";" & rsTemp("SUCURSAL")& ";" &rsTemp("MONTO_REGULARIZADO")& ";" &rsTemp("FORMA_PAGO")& ";" &rsTemp("FECHA_NORMALIZACION")& ";" & rsTemp("LUGAR_PAGO")& ";" &rsTemp("NRO_DOC_PAGO")& ";" &rsTemp("OBSERVACIONES_CAMPO")& ";" &rsTemp("EJEC_ASIG")


		cantSiniestroC = cantSiniestroC + 1

		fichCA.writeline(strTextoTercero)

		rsTemp.movenext

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

