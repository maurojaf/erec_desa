<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!DOCTYPE html>
<html lang="es">
<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/lib.asp"-->
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->

<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
<title>CRM RSA</title>
<style type="text/css">
<!--body {	background-color: #cccccc;}-->
</style>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border=0 bgcolor= "#FFFFFF" width="100%">

<tr>
<td>
<%
	Response.CodePage=65001
	Response.charset ="utf-8"


'******************************
'*	INICIO CODIGO PARTICULAR  *
''*****************************

strReprocesar = Request("strReprocesar")
strProcesar = Request("strProcesar")

intTotalRutCarga = Request("intTotalRutCarga")
intTotalDoc = Request("intTotalDoc")
intCambioCustodio_Llacruz = Request("intCambioCustodio_Llacruz")

strCodCliente=Request("CB_CLIENTE")

if Request("archivo") <> "" then
	strArchivo=Request("archivo")
End if

if Request("strTipoProceso") <> "" then
	strTipoProceso=Request("strTipoProceso")
End if

'response.write "<BR>strCodCliente=" & Request("CB_CLIENTE")

AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

If strReprocesar = "SI" Then

		'CUENTA LOS CAMBIOS DE CUSTODIO DE LLACRUZ A UMA (NO SE PERMITEN)'

		strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.COD_SAP = TMP_CARGA_UMA.COD_SAP" 
		strSql = strSql &" WHERE CARGA_UMA.CUSTODIO = 'CARGA_EXTERNA' AND '"& strTipoProceso &"' = 'CARGA_INTERNA'"
		set rsTemp= Conn.execute(strSql)
		
		if not rsTemp.eof then
			intCambioCustodio_UMA = rsTemp("CANTIDAD")
		Else
			intCambioCustodio_UMA = 0
		End if

		'CAMBIA EL CUSTODIO DE LLACRUZ A UMA'

		strSql = "UPDATE CARGA_UMA SET CUSTODIO = '"& strTipoProceso &"',USUARIO_CUSTODIO = '"& session("session_idusuario") &"',FECHA_CUSTODIO = GETDATE()" 
		strSql = strSql &" FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.COD_SAP = TMP_CARGA_UMA.COD_SAP"
		strSql = strSql &" WHERE CARGA_UMA.CUSTODIO = 'CARGA_EXTERNA' AND '"& strTipoProceso &"' = 'CARGA_INTERNA'"	
		Conn.Execute strSql,64

		strSql = "EXEC [dbo].[proc_CambiaCustodio] '1'"
		Conn.Execute strSql,64
		
		strSql1 = "EXEC Proc_Des_Asignacion_cobradores '" & strCodCliente & "'," & session("session_idusuario")
		set rsDesAsig = Conn.execute(strSql1)

		strSql1 = "EXEC Proc_Cambia_Custodio_Deudor '" & strCodCliente & "'," & session("session_idusuario")
		'response.write "<BR>strSql1 = " & strSql1
		
		set rsCambiaCustodio = Conn.execute(strSql1)

	%>

		<table border=1 bordercolor="#000000" width="500">

			<tr><td colspan=2 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13"><b>Estatus Reproceso (Esto solo cambio Custodio, si desea cargar procese nuevamente)</b></td></tr>
			<tr><td width="270">Cambio Custodio de LLACRUZ a UMA</td><td width="30" align="right"><%=intCambioCustodio_UMA%></td></tr>
			<tr><td colspan=2 align="center" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<br>
			<input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga_Cliente.asp');return false;">

			</td></tr>
		</table>

		<SCRIPT>
			alert('Cambio de Custodio de Llacruz a UMA realizado Correctamente');
		</SCRIPT>


<%ElseIf strProcesar = "SI" Then

		'CAMBIA LOS NRO_CLIENTE_DOC DISTINTOS CON IGUAL LLAVE AL ARCHIVO DE CARGA'

		strSql = "			UPDATE CUOTA"
		strSql = strSql &"  SET NRO_CLIENTE_DOC = TMP_CARGA_UMA.COD_SAP"
		strSql = strSql &"  FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.RUT_ALUMNO = TMP_CARGA_UMA.RUT_ALUMNO"
		strSql = strSql &"  								  		   AND CARGA_UMA.FOLIO = TMP_CARGA_UMA.FOLIO"
		strSql = strSql &" 								  			   AND CARGA_UMA.VENCIMI = TMP_CARGA_UMA.VENCIMI"
		strSql = strSql &" 			   	   INNER JOIN CUOTA ON REVERSE(SUBSTRING(REVERSE(CARGA_UMA.RUT_ALUMNO),2,10))+'-'+SUBSTRING(REVERSE(CARGA_UMA.RUT_ALUMNO),1,1) = RUT_DEUDOR"
		strSql = strSql &" 								  			   AND CARGA_UMA.FOLIO = CAST(NRO_DOC AS INT) AND FECHA_VENC = CARGA_UMA.VENCIMI"
		strSql = strSql &" WHERE CARGA_UMA.COD_SAP <> TMP_CARGA_UMA.COD_SAP"

		Conn.execute(strSql)
		
		'CAMBIA LOS COD_SAP DISTINTOS CON IGUAL LLAVE AL ARCHIVO DE CARGA'

		strSql = "			UPDATE CARGA_UMA"
		strSql = strSql &"  SET CARGA_UMA.COD_SAP = TMP_CARGA_UMA.COD_SAP"
		strSql = strSql &"  FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.RUT_ALUMNO = TMP_CARGA_UMA.RUT_ALUMNO"
		strSql = strSql &"  								  		   AND CARGA_UMA.FOLIO = TMP_CARGA_UMA.FOLIO"
		strSql = strSql &" 								  			   AND CARGA_UMA.VENCIMI = TMP_CARGA_UMA.VENCIMI"
		strSql = strSql &" 			   	   INNER JOIN CUOTA ON REVERSE(SUBSTRING(REVERSE(CARGA_UMA.RUT_ALUMNO),2,10))+'-'+SUBSTRING(REVERSE(CARGA_UMA.RUT_ALUMNO),1,1) = RUT_DEUDOR"
		strSql = strSql &" 								  			   AND CARGA_UMA.FOLIO = CAST(NRO_DOC AS INT) AND FECHA_VENC = CARGA_UMA.VENCIMI"
		strSql = strSql &" WHERE CARGA_UMA.COD_SAP <> TMP_CARGA_UMA.COD_SAP"

		Conn.execute(strSql)

        'INSERTA LOS REGISTROS NUEVOS A LA TABLA CARGA_UMA'

		strSqlFile = "INSERT INTO CARGA_UMA SELECT " & strCodCliente & ",*, '"& strTipoProceso &"',GETDATE(),'"& session("session_idusuario") &"','"& session("session_idusuario") &"',GETDATE(),null,null,null,null FROM TMP_CARGA_UMA WHERE COD_SAP NOT IN (SELECT COD_SAP FROM CARGA_UMA)"
		
		'response.write "<BR>strSqlFile = " & strSqlFile
		Conn.Execute strSqlFile,64
		
		'CAMBIA EL CUSTODIO DE LA BASE_ESTADO'

		strSql = "UPDATE CARGA_UMA SET CUSTODIO = '"& strTipoProceso &"',USUARIO_CUSTODIO = '"& session("session_idusuario") &"',FECHA_CUSTODIO = GETDATE()"
		strSql = strSql &" FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.COD_SAP = TMP_CARGA_UMA.COD_SAP" 
		strSql = strSql &" WHERE CARGA_UMA.CUSTODIO = 'CARGA_INTERNA' AND '"& strTipoProceso &"' = 'CARGA_EXTERNA'"
		Conn.Execute strSql,64

		strSql = "EXEC [dbo].[proc_CambiaCustodio] '1' "
		Conn.Execute strSql,64

		'CARGA DEUDORES Y DOCUMENTOS'

		strSql = "EXEC [dbo].[proc_Carga_Actualizacion_UMA] '1'"
		''Response.write strSql
		Conn.execute(strSql)
		
		strSql1 = "EXEC Proc_Des_Asignacion_cobradores '" & strCodCliente & "'," & session("session_idusuario")
		set rsDesAsig = Conn.execute(strSql1)

		strSql1 = "EXEC Proc_Cambia_Custodio_Deudor '" & strCodCliente & "'," & session("session_idusuario")
		set rsCambiaCustodio = Conn.execute(strSql1)
				
		strSql1 = "EXEC Proc_Asigna_Cobrador_Carga '" & strCodCliente & "'," & session("session_idusuario")
		set rsAsignaCarga = Conn.execute(strSql1)
		
		
		%>

		<table border=1 bordercolor="#000000" width="500">

			<tr><td colspan=2 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13"><b>Estatus Proceso</b></td></tr>

			<tr>
				<td width="400" height = "30">TOTAL DEUDORES CARGA</td>
				<td width="40" align="right"><%=intTotalRutCarga%></td>
			</tr>
				<td width="400" height = "30">TOTAL DOCUMENTOS CARGA</td>
				<td width="40" align="right"><%=intTotalDoc%></td>
			</tr>
			</tr>
				<td width="400" height = "30">TOTAL DOCUMENTOS CON CUSTODIO CAMBIADO A LLACRUZ</td>
				<td width="40" align="right"><%=intCambioCustodio_Llacruz%></td>
			</tr>
			<tr>
				<td colspan=2 align="center" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					<input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga_Cliente.asp');return false;">
				</td>
			</tr>

		</table>

		<SCRIPT>
			alert('Proceso de Carga y Cambio de custodio realizado Correctamente');
		</SCRIPT>

<%End If

' La actualización del importe debe ser siempre (tanto para interno como externo)
		
strSql = "			UPDATE CUOTA"
strSql = strSql &"  SET SALDO = TMP_CARGA_UMA.IMPORTE, VALOR_CUOTA = TMP_CARGA_UMA.IMPORTE"
strSql = strSql &"  FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.RUT_ALUMNO = TMP_CARGA_UMA.RUT_ALUMNO"
strSql = strSql &"  								  		   AND CARGA_UMA.FOLIO = TMP_CARGA_UMA.FOLIO"
strSql = strSql &" 								  			   AND CARGA_UMA.VENCIMI = TMP_CARGA_UMA.VENCIMI"
strSql = strSql &" 			   	   INNER JOIN CUOTA ON REVERSE(SUBSTRING(REVERSE(CARGA_UMA.RUT_ALUMNO),2,10))+'-'+SUBSTRING(REVERSE(CARGA_UMA.RUT_ALUMNO),1,1) = RUT_DEUDOR"
strSql = strSql &" 								  			   AND CARGA_UMA.FOLIO = CAST(NRO_DOC AS BIGINT) AND FECHA_VENC = CARGA_UMA.VENCIMI"

Conn.execute(strSql)

' La actualización del importe debe ser siempre (tanto para interno como externo)

strSql = "			UPDATE CARGA_UMA"
strSql = strSql &"  SET CARGA_UMA.IMPORTE = TMP_CARGA_UMA.IMPORTE"
strSql = strSql &"  FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.RUT_ALUMNO = TMP_CARGA_UMA.RUT_ALUMNO"
strSql = strSql &"  								  		   AND CARGA_UMA.FOLIO = TMP_CARGA_UMA.FOLIO"
strSql = strSql &" 								  			   AND CARGA_UMA.VENCIMI = TMP_CARGA_UMA.VENCIMI"
strSql = strSql &" 			   	   INNER JOIN CUOTA ON REVERSE(SUBSTRING(REVERSE(CARGA_UMA.RUT_ALUMNO),2,10))+'-'+SUBSTRING(REVERSE(CARGA_UMA.RUT_ALUMNO),1,1) = RUT_DEUDOR"
strSql = strSql &" 								  			   AND CARGA_UMA.FOLIO = CAST(NRO_DOC AS BIGINT) AND FECHA_VENC = CARGA_UMA.VENCIMI"

Conn.execute(strSql)

If strArchivo <> "" Then

	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "Terceros_cargados_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros

	strTextoArchivoCC = ""
	strTextoArchivoCNC = ""
	strTextoArchivoCA = ""
	
	strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"&strCodCliente &"/" & strArchivo

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CARGA_UMA]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_CARGA_UMA]"
	Conn.Execute strSql,64

	strSql = " CREATE TABLE TMP_CARGA_UMA ( COD_SAP BIGINT NOT NULL,"
	strSql = strSql &" 	FPO_4 BIGINT NOT NULL,"
	strSql = strSql &"  IMPORTE BIGINT NOT NULL,"
	strSql = strSql &" 	CORE VARCHAR(20) NULL,"
	strSql = strSql &" 	DIVI INT NOT NULL,"
	strSql = strSql &"  TI VARCHAR(20) NULL,"
	strSql = strSql &"  RUT_ALUMNO VARCHAR(20) NOT NULL,"
	strSql = strSql &"  ALUMNO BIGINT NOT NULL,"
	strSql = strSql &"  NOMBRE_ALUMNO VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL,"
	strSql = strSql &" 	RUT_GIRADOR VARCHAR(20)  NOT NULL,"
	strSql = strSql &" 	GIRADOR BIGINT NOT NULL,"
	strSql = strSql &" 	NOMBRE_GIRADOR VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL,"
	strSql = strSql &"  RUT_AVAL VARCHAR(20) NULL,"
	strSql = strSql &"  AVAL BIGINT NULL,"
	strSql = strSql &"  NOMBRE_AVAL VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL,"
	strSql = strSql &"  TIPO_DOCUM VARCHAR(20) NULL,"
	strSql = strSql &" 	DOCUMENTO_SA BIGINT NOT NULL,"
	strSql = strSql &" 	POSI INT NOT NULL,"
	strSql = strSql &" 	A VARCHAR(20) NULL,"
	strSql = strSql &"  FOLIO VARCHAR(20) NULL,"
	strSql = strSql &"  BANCO VARCHAR(20) NULL,"
	strSql = strSql &"  NOMBRE_DEL_BANCO VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL,"
	strSql = strSql &"  PLAZA VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL,"
	strSql = strSql &" 	CUENTA VARCHAR(100) NULL,"
	strSql = strSql &" 	MONEDA VARCHAR(20) NULL,"
	strSql = strSql &"  MONTO VARCHAR(20) NOT NULL,"
	strSql = strSql &"  VENCIMI DATETIME NOT NULL,"
	strSql = strSql &"  VENCIMI_2 DATETIME NOT NULL,"
	strSql = strSql &"  LUGAR VARCHAR(100) NULL,"
	strSql = strSql &"  SITUACION VARCHAR(100) NULL,"
	strSql = strSql &"  ESTADO VARCHAR(100) NULL,"
	strSql = strSql &"  GARANTIA VARCHAR(100) NULL,"
	strSql = strSql &"  REMESA VARCHAR(100) NULL,"
	strSql = strSql &"  LLAVE_OPER VARCHAR(100) NULL,"
	strSql = strSql &"  USUARIO VARCHAR(100) NULL,"
	strSql = strSql &"  FECHA VARCHAR(20) NULL,"
	strSql = strSql &"  NUMERO_DE_PR VARCHAR(100),"
	strSql = strSql &"  POSI_2 INT NULL,"
	strSql = strSql &"  TXT_EXPLCATIVO_P_POS VARCHAR(100),"
	strSql = strSql &"  TIPO INT NULL,"
	strSql = strSql &"  PROTESTO VARCHAR(50) NULL,"
	strSql = strSql &"  GASTO_DE_PROTESTO INT NULL,"
	strSql = strSql &"  TEL_1_A VARCHAR(100),"
	strSql = strSql &"  TEL_2_A VARCHAR(100),"
	strSql = strSql &"  DIRECCION_A VARCHAR(100),"
	strSql = strSql &"  COMUNA VARCHAR(100),"
	strSql = strSql &"  CIUDAD VARCHAR(100),"
	strSql = strSql &"  REGION VARCHAR(100),"
	strSql = strSql &"  EMAIL_A VARCHAR(100),"
	strSql = strSql &"  TEL_1_G VARCHAR(100),"
	strSql = strSql &"  TEL_2_G VARCHAR(100),"
	strSql = strSql &"  DIRECCION_G VARCHAR(100),"
	strSql = strSql &"  EMAIL_G VARCHAR(100))"

	Conn.Execute strSql,64

	'response.write "Conn = " & Conn
	'response.write "strSql " & strSql

	'**********CARGA ARCHIVO************'

	strSqlFile = "BULK INSERT TMP_CARGA_UMA FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2, CODEPAGE = 'ACP')"
	Conn.Execute strSqlFile,64
	
	'CUENTA LOS DOCUMENTOS A CARGAR'

	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_UMA WHERE COD_SAP NOT IN (SELECT COD_SAP FROM CARGA_UMA)"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intTotalDoc = rsTemp("CANTIDAD")
	Else
		intTotalDoc = 0
	End if

	'CUENTA LOS DEUDORES A CARGAR'

	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_UMA WHERE RUT_ALUMNO NOT IN (SELECT RUT_ALUMNO FROM CARGA_UMA) GROUP BY RUT_ALUMNO "
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intTotalRutCarga = rsTemp("CANTIDAD")
	Else
		intTotalRutCarga = 0
	End if

	'CUENTA LOS CAMBIOS DE CUSTODIO DE UMA A LLACRUZ'

	strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.COD_SAP = TMP_CARGA_UMA.COD_SAP WHERE CARGA_UMA.CUSTODIO = 'CARGA_INTERNA' AND '"& strTipoProceso &"' = 'CARGA_EXTERNA'"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCambioCustodio_Llacruz = rsTemp("CANTIDAD")
	Else
		intCambioCustodio_Llacruz = 0
	End if

	'CUENTA LOS CAMBIOS DE CUSTODIO DE LLACRUZ A UMA (NO SE PERMITEN)'

	strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.COD_SAP = TMP_CARGA_UMA.COD_SAP WHERE CARGA_UMA.CUSTODIO = 'CARGA_EXTERNA' AND '"& strTipoProceso &"' = 'CARGA_INTERNA'"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCambioCustodio_UMA = rsTemp("CANTIDAD")
	Else
		intCambioCustodio_UMA = 0
	End if

	'CUENTA LOS CAMBIOS DE CUSTODIO DE DOCUMENTOS NO ACTIVOS'

	strSql = " SELECT COUNT(*) AS CANTIDAD"
	strSql = strSql &" FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.COD_SAP = TMP_CARGA_UMA.COD_SAP"
	strSql = strSql &" 			   	  INNER JOIN CUOTA ON CUOTA.COD_CLIENTE = '1070' AND CARGA_UMA.COD_SAP = CUOTA.NRO_CLIENTE_DOC" 
	strSql = strSql &" 			      INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"

	strSql = strSql &" WHERE CARGA_UMA.CUSTODIO <> '"& strTipoProceso &"' AND ESTADO_DEUDA.ACTIVO = 0"

	set rsTemp= Conn.execute(strSql)

	if not rsTemp.eof then
		intCambioCustodio_NoActivo = rsTemp("CANTIDAD")
	Else
		intCambioCustodio_NoActivo = 0
	End if

	'CUENTA LOS REGISTROS DUPLICADOS EN SISTEMA'

	strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.COD_SAP = TMP_CARGA_UMA.COD_SAP WHERE CARGA_UMA.CUSTODIO = '"& strTipoProceso &"'"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaDuplicados = rsTemp("CANTIDAD")
	Else
		intCargaDuplicados = 0
	End if


	'CUENTA LOS REGISTROS DUPLICADOS EN BASE DE CARGA'

	strSql = "SELECT COUNT(REPETIDOS) AS REPETIDOS FROM"
	strSql = strSql &" (SELECT ROW_NUMBER() OVER (PARTITION BY COD_SAP ORDER BY COD_SAP ASC) AS REPETIDOS FROM TMP_CARGA_UMA) AS REP"
	strSql = strSql &" WHERE REPETIDOS > 1"

	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaDuplicadosBase = rsTemp("REPETIDOS")
	Else
		intCargaDuplicadosBase = 0
	End if

	'CUENTA LOS REGISTROS CON COD_SAP DISTINTO A DOCUMENTO_SA + POSI'

	strSql = "SELECT COUNT(*) AS ERROR"
	strSql = strSql &" FROM TMP_CARGA_UMA"
	strSql = strSql &" WHERE COD_SAP <> CAST((CONVERT(VARCHAR(50),DOCUMENTO_SA,103) + CONVERT(VARCHAR(50),POSI,103)) AS BIGINT)"

	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaErrorCodsap1= rsTemp("ERROR")
	Else
		intCargaErrorCodsap1 = 0
	End if

	'CUENTA LOS REGISTROS CON GASTO PROTESTO MAYOR A 25.000'

	strSql = "SELECT COUNT(*) AS ERROR2"
	strSql = strSql &" FROM TMP_CARGA_UMA"
	strSql = strSql &" WHERE ISNULL(GASTO_DE_PROTESTO,0) > 25000"

	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaErrorGastoProt= rsTemp("ERROR2")
	Else
		intCargaErrorGastoProt = 0
	End if


	'CUENTA LOS REGISTROS CON COD_SAP DISTINTO Y LA MISMA LLAVE (SEGÚN PROCEDIMIENTO, ESTO NO SE DEBE CARGAR SINO QUE SE DEBE ACTUALIZAR EL COD_SAP DE SISTEMA'

	strSql = "SELECT COUNT(*) AS CAMBIOCODSAP"
	strSql = strSql &" FROM CARGA_UMA INNER JOIN TMP_CARGA_UMA ON CARGA_UMA.RUT_ALUMNO = TMP_CARGA_UMA.RUT_ALUMNO"
	strSql = strSql &" 											  AND CARGA_UMA.FOLIO = TMP_CARGA_UMA.FOLIO"
	strSql = strSql &" 											  AND CARGA_UMA.VENCIMI = TMP_CARGA_UMA.VENCIMI"
	strSql = strSql &" WHERE CARGA_UMA.COD_SAP <> TMP_CARGA_UMA.COD_SAP"


	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaCambioCodSasp= rsTemp("CAMBIOCODSAP")
	Else
		intCargaCambioCodSasp = 0
	End if

	
	'-----------------------------SOLO SE PERMITE CARGAR Y REPROCESAR SI EN EL PROCESO NO HAY ERRORES-----------------------------'

		If Trim(intTotalRutCarga) = "" or IsNull(intTotalRutCarga) Then intTotalRutCarga = 0
		If Trim(intTotalDoc) = "" or IsNull(intTotalDoc) Then intTotalDoc = 0

		%>

		<table border=1 bgcolor="#<%=session("COLTABBG2")%>" class="Estilo28" width="700" align="center">

		<%If intCambioCustodio_UMA > 0 or intCargaDuplicados > 0 or intCargaDuplicadosBase >0 or intCambioCustodio_NoActivo >0 or intCargaErrorCodsap1 > 0 then %>

			<tr><td colspan=2 width="600" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13"><b>Estatus Carga (Existen errores, NO SE PUEDE CARGAR NADA NI CAMBIAR CUSTODIOS)</b></td></tr>
			<tr><td width="400">Tipo Carga: </td><td width="30" align="right"><%= strTipoProceso %></td></tr>
			<tr><td width="400">Cantidad Deudores a Cargar: </td><td width="30" align="right"><%=intTotalRutCarga%></td></tr>
			<tr><td width="400">Cantidad Documentos a Cargar: </td><td width="30" align="right"><%= intTotalDoc %></td></tr>
			<tr><td width="400">Cantidad Documentos a Actualizar COD_SAP: </td><td width="30" align="right"><%= intCargaCambioCodSasp %></td></tr>
			<tr><td width="400">Cambio Custodio a Llacruz: </td><td width="30" align="right"><%=intCambioCustodio_Llacruz%></td></tr>


			<tr><td colspan=2 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13"><b>Errores Carga</b></td></tr>

		<%If intCargaDuplicados > 0 then%>
			<tr><td width="400">Documentos ya Cargados en sistema: </td><td width="30" align="right"><%=intCargaDuplicados%></td></tr>
		<%end if%>

		<%If intCargaDuplicadosBase > 0 then%>
			<tr><td width="400">Documentos Duplicados en Base de Carga: </td><td width="30" align="right"><%=intCargaDuplicadosBase%></td></tr>
		<%end if%>

		<%If intCargaErrorCodsap1 > 0 then%>
			<tr><td width="400">COD_SAP distinto a DOCUMENTO_SA + POSI: </td><td width="30" align="right"><%=intCargaErrorCodsap1%></td></tr>
		<%end if%>

		<%If intCargaErrorGastoProt > 0 then%>
			<tr><td width="400">Gasto de Protesto mayor a 25.000: </td><td width="30" align="right"><%=intCargaErrorGastoProt%></td></tr>
		<%end if%>

		<%If intCambioCustodio_UMA > 0 then%>
			<tr><td width="400">Cambio Custodio a UMA Incorrecto (Llacruz a UMA): </td><td width="30" align="right"><%=intCambioCustodio_UMA%></td></tr>
		<%end if%>

		<%If intCambioCustodio_NoActivo > 0 then%>
			<tr><td width="400">Cambio Custodio de documentos no activos: </td><td width="30" align="right"><%=intCambioCustodio_NoActivo%></td></tr>
		<%end if%>

			<tr><td colspan=2 align="center">

		<% Else %>

			<tr><td colspan=2 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13"><b>Estatus Carga (Aun no se han realizado Cambios, debe procesar para Cargar y Actualizar Custodios)</b></td></tr>
			<tr><td width="400">Tipo Carga: </td><td width="30" align="right"><%= strTipoProceso %></td></tr>
			<tr><td width="400">Cantidad Deudores a Cargar: </td><td width="30" align="right"><%=intTotalRutCarga%></td></tr>
			<tr><td width="400">Cantidad Documentos a Cargar: </td><td width="30" align="right"><%= intTotalDoc %></td></tr>
			<tr><td width="400">Cantidad Documentos a Actualizar COD_SAP: </td><td width="30" align="right"><%= intCargaCambioCodSasp %></td></tr>
			<tr><td width="400">Cambio Custodio a Llacruz: </td><td width="30" align="right"><%=intCambioCustodio_Llacruz%></td></tr>

		<% end if%>

		<tr><td colspan=2 align="center" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
		<br>
		<input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga_Cliente.asp');return false;">

		<%If intCambioCustodio_UMA > 0 then %>
		<input type="BUTTON" value="Procesar Igualmente" name="terminar" onClick="Reprocesar('<%=strTipoProceso%>',<%=strCodCliente%>);return false;"><br><br>
		<% end if%>

		<%If intCambioCustodio_UMA = 0 and intCargaDuplicados = 0 and intCargaDuplicadosBase = 0 and intCambioCustodio_NoActivo = 0 and intCargaErrorCodsap1 = 0 then %>
		<input type="BUTTON" value="Procesar" name="terminar" onClick="Procesar(<%=intTotalRutCarga%>,<%=intTotalDoc%>,<%=strCodCliente%>,'<%=strTipoProceso%>',<%=intCambioCustodio_Llacruz%>);return false;"><br><br>
		<% end if%>

		</td></tr>
		</table>
	 <%

End if

%>

</td>
</tr>
</table>

</body>
</html>
<script language="JavaScript" type="text/JavaScript">

	function Terminar( sintPaginaTerminar ) {
		self.location.href = sintPaginaTerminar
	}
	function Reprocesar(strTipoProceso,strCodCliente)
	{
		if (confirm("¿ Está seguro de procesar nuevamente ? Este proceso cambia el custodio de Llacruz a UMA, esto no es algo que se debe hacer dado a que afecta al proceso normal de cobranza."))
		{
			if (confirm("¿ Está REALMENTE seguro de procesar ?"))
			{
				self.location.href = "Man_UploadDoc_UMA.asp?strReprocesar=SI&strTipoProceso=" + strTipoProceso + "&CB_CLIENTE=" + strCodCliente
			}
		}
	}

	function Procesar(intTotalRutCarga,intTotalDoc,strCodCliente,strTipoProceso,intCambioCustodio_Llacruz)


		{
			if (confirm("¿ Está REALMENTE seguro de cargar y actualizar custodios de los documentos ?"))
			{
				self.location.href = "Man_UploadDoc_UMA.asp?strProcesar=SI&intTotalRutCarga=" + intTotalRutCarga + "&intTotalDoc=" + intTotalDoc + "&CB_CLIENTE=" + strCodCliente + "&strTipoProceso=" + strTipoProceso + "&intCambioCustodio_Llacruz=" + intCambioCustodio_Llacruz
			}
		}
</script>
