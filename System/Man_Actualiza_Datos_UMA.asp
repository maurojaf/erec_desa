<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->    
	<!--#include file="arch_utils.asp"-->

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
	function Terminar( sintPaginaTerminar ) {
		self.location.href = sintPaginaTerminar
	}

	function Reprocesar(strInput)
	{
		if (confirm("¿ Está seguro de procesar nuevamente ? Este proceso cambia el custodio de Llacruz a UMA, esto no es algo que se debe hacer comunmente"))
		{
			if (confirm("¿ Está REALMENTE seguro de reprocesar nuevamente ?"))
			{
				self.location.href = "Man_Actualiza_UMA.asp?strReprocesar=SI&strTipoProceso=" + strInput
			}
		}
	}

	function Procesar()


		{
			if (confirm("¿ Está REALMENTE seguro de cargar los documentos ?"))
			{
				datos.action='Man_Actualiza_UMA.asp?strProcesar=SI';
				datos.submit();
			}
		}
</script>


<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
<title>CRM Cobros</title>
<style type="text/css">
<!--body {	background-color: #cccccc;}-->
</style>
</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<%

'******************************
'*	INICIO CODIGO PARTICULAR  *
''******************************
%>
<%

strReprocesar = Request("strReprocesar")
strProcesar = Request("strProcesar")

intTotalRutCarga = Request("intTotalRutCarga")
intTotalDoc = Request("intTotalDoc")

if Request("CB_CLIENTE") <> "" then
	strCliente=Request("CB_CLIENTE")
End if

if Request("archivo") <> "" then
	strArchivo=Request("archivo")
End if

if Request("strTipoProceso") <> "" then
	strTipoProceso=Request("strTipoProceso")
End if

if Request("opAc")= "0" then
	sIopAc=0
else
	sIopAc=1
End if

AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

'Response.write "<br>strProcesar=" & strProcesar
'Response.write "<br>strArchivo=" & strArchivo


If strArchivo <> "" Then


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "Terceros_cargados_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros

	strTextoArchivoCC = ""
	strTextoArchivoCNC = ""
	strTextoArchivoCA = ""


		'strFileDir = server.mappath("../Archivo/CargaActualizaciones/"&strCliente &"/" & strArchivo)
		strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"&strCliente &"/" & strArchivo

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_ACTUALIZA_DATOS_UMA]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_ACTUALIZA_DATOS_UMA]"
	Conn.Execute strSql,64


	strSql = " CREATE TABLE TMP_ACTUALIZA_DATOS_UMA ("
	strSql = strSql &" 	ID_CUOTA INT NOT NULL,"
	strSql = strSql &" 	COD_SAP BIGINT NOT NULL,"
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
	strSql = strSql &" 	DOCUMENTO_SA BIGINT NULL,"
	strSql = strSql &" 	POSI INT NULL,"
	strSql = strSql &" 	A VARCHAR(20) NULL,"
	strSql = strSql &"  FOLIO VARCHAR(20) NULL,"
	strSql = strSql &"  BANCO VARCHAR(20) NULL,"
	strSql = strSql &"  NOMBRE_DEL_BANCO VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL,"
	strSql = strSql &"  PLAZA VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL,"
	strSql = strSql &" 	CUENTA VARCHAR(100) NULL,"
	strSql = strSql &" 	MONEDA VARCHAR(20) NULL,"
	strSql = strSql &"  MONTO VARCHAR(20) NOT NULL,"
	strSql = strSql &"  VENCIMI SMALLDATETIME NOT NULL,"
	strSql = strSql &"  VENCIMI_2 SMALLDATETIME NOT NULL,"
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

	strSqlFile = "BULK INSERT TMP_ACTUALIZA_DATOS_UMA FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2, CODEPAGE = 'ACP')"
	Conn.Execute strSqlFile,64%>

	<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="datos" method="post">
	<table width="990" border="1" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
		<tr>
			<TD height="20" ALIGN=LEFT class="pasos2_i">
				<B>INFORME ACTUALIZACIÓN DE DEUDA MASIVA</B>
			</TD>
		</tr>
	<tr>
		<td align="center">


	<%'VERIFICA QUE TODOS LOS ID_CUOTA ASOCIADOS Y LOS DOCUMENTOS SEAN CORRESPONDIENTES A LA LLAVE DEL CLIENTE'

	strSql = "	SELECT COUNT(*) AS CANTIDAD FROM TMP_ACTUALIZA_DATOS_UMA LEFT JOIN CUOTA ON TMP_ACTUALIZA_DATOS_UMA.ID_CUOTA = CUOTA.ID_CUOTA"
	strSql = strSql &" 																	  AND TMP_ACTUALIZA_DATOS_UMA.COD_SAP = CUOTA.NRO_CLIENTE_DOC"
	strSql = strSql &" 																	  AND CUOTA.RUT_DEUDOR = REVERSE(SUBSTRING(REVERSE(TMP_ACTUALIZA_DATOS_UMA.RUT_ALUMNO),2,10))+'-'+SUBSTRING(REVERSE(TMP_ACTUALIZA_DATOS_UMA.RUT_ALUMNO),1,1)"
	strSql = strSql &" 																	  AND CUOTA.NRO_DOC = TMP_ACTUALIZA_DATOS_UMA.FOLIO"
	strSql = strSql &"																	  AND CUOTA.FECHA_VENC = CAST(TMP_ACTUALIZA_DATOS_UMA.VENCIMI AS DATETIME)"
	strSql = strSql &"	WHERE CUOTA.ID_CUOTA IS NULL"


	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCuentaErroresID = rsTemp("CANTIDAD")
	Else
		intCuentaErroresID = 0
	End if

	'CUENTA LOS REGISTROS DUPLICADOS EN BASE DE ACTUALIZACION'

	strSql = "SELECT COUNT(REPETIDOS) AS REPETIDOS FROM"
	strSql = strSql &" (SELECT ROW_NUMBER() OVER (PARTITION BY COD_SAP ORDER BY COD_SAP ASC) AS REPETIDOS FROM TMP_ACTUALIZA_DATOS_UMA) AS REP"
	strSql = strSql &" WHERE REPETIDOS > 1"

	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaDuplicadosBase = rsTemp("REPETIDOS")
	Else
		intCargaDuplicadosBase = 0
	End if


	'CUENTA LOS REGISTROS DUPLICADOS EN BASE DE ACTUALIZACION'

	strSql = "SELECT COUNT(REPETIDOS) AS REPETIDOS FROM"
	strSql = strSql &" (SELECT ROW_NUMBER() OVER (PARTITION BY COD_SAP ORDER BY COD_SAP ASC) AS REPETIDOS FROM TMP_ACTUALIZA_DATOS_UMA) AS REP"
	strSql = strSql &" WHERE REPETIDOS > 1"

	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaDuplicadosBase = rsTemp("REPETIDOS")
	Else
		intCargaDuplicadosBase = 0
	End if%>


<table border=0 width="700" height = "50" bgcolor="#<%=session("COLTABBG")%>" class="Estilo28">

		<%If intCuentaErroresID = 0 and intCargaDuplicadosBase = 0 and intEstadosNoReconocidos = 0 and intEstadosconError = 0 then%>
		<td colspan=1 align="center" width="700" >
		<input type="BUTTON" value="Actualizar Deuda" name="terminar" onClick="Procesar();return false;">
		</td>
		<%End If%>

		<td colspan=1 align="right">
		<input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga_Cliente.asp');return false;">
		</td>

		</td></tr>
</table>
	 <%

End if



%>


				</td>
			  </tr>
			</table>
</form>
</body>
</html>

