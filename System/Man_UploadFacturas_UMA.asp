<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
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

	function Terminar( sintPaginaTerminar ) {
		self.location.href = sintPaginaTerminar
	}

	function Reprocesar(strInput)
	{
		if (confirm("¿ Está seguro de procesar nuevamente ? Este proceso cambia el custodio de Llacruz a UMA, esto no es algo que se debe hacer dado a que afecta al proceso normal de cobranza"))
		{
			if (confirm("¿ Está REALMENTE seguro de procesar ?"))
			{
				self.location.href = "Man_UploadFacturas_UMA.asp?strReprocesar=SI&strTipoProceso=" + strInput
			}
		}
	}

	function Procesar(intTotalRutCarga,intTotalDoc,strCliente,strTipoProceso,intCambioCustodio_Llacruz)
		{
			if (confirm("¿ Está REALMENTE seguro de cargar y actualizar custodios de los documentos ?"))
			{
				self.location.href = "Man_UploadFacturas_UMA.asp?strProcesar=SI&intTotalRutCarga=" + intTotalRutCarga + "&intTotalDoc=" + intTotalDoc + "&strCliente=" + strCliente + "&strTipoProceso=" + strTipoProceso + "&intCambioCustodio_Llacruz=" + intCambioCustodio_Llacruz
			}
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

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border=0 bgcolor= "#FFFFFF" width="100%">

<tr>
<td>
<%

'******************************
'*	INICIO CODIGO PARTICULAR  *
''******************************

strReprocesar = Request("strReprocesar")
strProcesar = Request("strProcesar")

intTotalRutCarga = Request("intTotalRutCarga")
intTotalDoc = Request("intTotalDoc")
intCambioCustodio_Llacruz = Request("intCambioCustodio_Llacruz")

'strCodCliente=session("ses_codcli")
strCodCliente = Request("CB_CLIENTE")
strCliente= Request("CB_CLIENTE")

if Request("archivo") <> "" then
	strArchivo=Request("archivo")
End if

if Request("strTipoProceso") <> "" then
	strTipoProceso=Request("strTipoProceso")
End if

if strProcesar = "SI"  then
	strCliente=1070
End if

strCodCliente=Request("CB_CLIENTE")

AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

If strReprocesar = "SI" Then

		'CUENTA LOS CAMBIOS DE CUSTODIO DE LLACRUZ A UMA (NO SE PERMITEN)'

		strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_UMA_FACT INNER JOIN TMP_CARGA_UMA_FACT ON CARGA_UMA_FACT.COD_SAP = TMP_CARGA_UMA_FACT.COD_SAP WHERE CARGA_UMA_FACT.CUSTODIO = 'CARGA_EXTERNA' AND '"& strTipoProceso &"' = 'CARGA_INTERNA'"
		set rsTemp= Conn.execute(strSql)
		if not rsTemp.eof then
			intCambioCustodio_UMA = rsTemp("CANTIDAD")
		Else
			intCambioCustodio_UMA = 0
		End if


		'CAMBIA EL CUSTODIO DE UMA A LLACRUZ'

		strSqlFile = "UPDATE CARGA_UMA_FACT SET CUSTODIO = '"& strTipoProceso &"',USUARIO_CUSTODIO = '"& session("session_idusuario") &"',FECHA_CUSTODIO = GETDATE() FROM CARGA_UMA_FACT INNER JOIN TMP_CARGA_UMA_FACT ON CARGA_UMA_FACT.COD_SAP = TMP_CARGA_UMA_FACT.COD_SAP WHERE CARGA_UMA_FACT.CUSTODIO = 'CARGA_EXTERNA' AND '"& strTipoProceso &"' = 'CARGA_INTERNA'"
		Conn.Execute strSqlFile,64

		strSql = "EXEC [dbo].[proc_CambiaCustodio] '2'"
		Conn.Execute strSql,64
		
		strSql1 = "EXEC Proc_Des_Asignacion_cobradores '" & strCodCliente & "'," & session("session_idusuario")
		set rsDesAsig = Conn.execute(strSql1)

		strSql1 = "EXEC Proc_Cambia_Custodio_Deudor '" & strCodCliente & "'," & session("session_idusuario")
		set rsCambiaCustodio = Conn.execute(strSql1)
				
		strSql1 = "EXEC Proc_Asigna_Cobrador_Carga '" & strCodCliente & "'," & session("session_idusuario")
		set rsAsignaCarga = Conn.execute(strSql1)
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

		'INSERTA LOS EGISTROS NUEVOS A LA TABLA CARGA_UMA'

		strSqlFile = "INSERT INTO CARGA_UMA_FACT SELECT " & strCliente & ",*, '"& strTipoProceso &"',GETDATE(),'"& session("session_idusuario") &"','"& session("session_idusuario") &"',GETDATE(),null,null,null,null,null FROM TMP_CARGA_UMA_FACT WHERE COD_SAP NOT IN (SELECT COD_SAP FROM CARGA_UMA_FACT)"
		''Response.write strSqlFile
		Conn.Execute strSqlFile,64

		'CAMBIA EL CUSTODIO DE LA BASE_ESTADO'

		strSqlFile = "UPDATE CARGA_UMA_FACT SET CUSTODIO = '"& strTipoProceso &"',USUARIO_CUSTODIO = '"& session("session_idusuario") &"',FECHA_CUSTODIO = GETDATE() FROM CARGA_UMA_FACT INNER JOIN TMP_CARGA_UMA_FACT ON CARGA_UMA_FACT.COD_SAP = TMP_CARGA_UMA_FACT.COD_SAP WHERE CARGA_UMA_FACT.CUSTODIO = 'CARGA_INTERNA' AND '"& strTipoProceso &"' = 'CARGA_EXTERNA'"
		Conn.Execute strSqlFile,64

		strSql = "EXEC [dbo].[proc_CambiaCustodio] '2' "
		Conn.Execute strSql,64

		'CARGA DEUDORES Y DOCUMENTOS'

		strSql = "EXEC [dbo].[proc_Carga_Actualizacion_UMA] '2'"
		'Response.write strSql
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

If strArchivo <> "" Then


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "Terceros_cargados_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros

	strTextoArchivoCC = ""
	strTextoArchivoCNC = ""
	strTextoArchivoCA = ""


	strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"&strCodCliente&"/" & strArchivo

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CARGA_UMA_FACT]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_CARGA_UMA_FACT]"
	Conn.Execute strSql,64


	 strSql = " CREATE TABLE TMP_CARGA_UMA_FACT ( COD_SAP BIGINT NOT NULL,"
	 strSql = strSql &"  CRUCE BIGINT NULL,"
	 strSql = strSql &"  NRO_DOC BIGINT NOT NULL,"
	 strSql = strSql &"  POS_1 INT NOT NULL,"
	 strSql = strSql &"  IMPORTE_ML BIGINT NOT NULL,"
	 strSql = strSql &"  MON VARCHAR(20) NULL,"
	 strSql = strSql &"  SOC VARCHAR(20) NULL,"
	 strSql = strSql &"  DIV INT NULL,"
	 strSql = strSql &"  INT_CIAL BIGINT NULL,"
	 strSql = strSql &"  CONTRATO BIGINT NULL,"
	 strSql = strSql &"  CTA_MAYOR BIGINT NULL,"
	 strSql = strSql &"  FECHA_DOC SMALLDATETIME NOT NULL,"
	 strSql = strSql &"  FECHA_CONTABLE SMALLDATETIME NULL,"
	 strSql = strSql &"  TEXTO VARCHAR(50) NULL,"
	 strSql = strSql &"  VENC_NETO SMALLDATETIME NULL,"
	 strSql = strSql &"  COMPENS VARCHAR(50) NULL,"
	 strSql = strSql &"  FE_CON_COMPT SMALLDATETIME NULL,"
	 strSql = strSql &"  CD VARCHAR(50) NULL,"
	 strSql = strSql &"  REFERENCIA BIGINT NULL,"
	 strSql = strSql &"  TEXTO_P_OP_PRINCIPAL VARCHAR(100) NULL,"
	 strSql = strSql &"  TEXTO_P_OP_PARCIAL VARCHAR(100) NULL,"
	 strSql = strSql &"  OP_PRAL BIGINT NULL,"
	 strSql = strSql &"  OP_PARC BIGINT NULL,"
	 strSql = strSql &"  MBC VARCHAR(20) NULL,"
	 strSql = strSql &"  IMPORTE BIGINT NULL,"
	 strSql = strSql &"  MON_2 VARCHAR(20) NULL,"
	 strSql = strSql &"  CLOB VARCHAR(20) NULL,"
	 strSql = strSql &"  APLAZAM VARCHAR(50) NULL,"
	 strSql = strSql &"  MB VARCHAR(50) NULL,"
	 strSql = strSql &"  SEGMENTO VARCHAR(20) NULL,"
	 strSql = strSql &"  CBP VARCHAR(20) NULL,"
	 strSql = strSql &"  CTA_CONTR BIGINT NULL,"
	 strSql = strSql &"  TCC VARCHAR(20) NULL,"
	 strSql = strSql &"  FECHA_DE_ASIGNACION SMALLDATETIME NULL,"
	 strSql = strSql &"  COBRANZA VARCHAR(20) NULL,"
	 strSql = strSql &"  EMPRESA VARCHAR(70) NULL,"
	 strSql = strSql &"  RUT VARCHAR(20) NULL,"
	 strSql = strSql &"  TELEFONO_1 VARCHAR(100) NULL,"
	 strSql = strSql &"  TELEFONO_2 VARCHAR(100) NULL,"
	 strSql = strSql &"  EMAIL VARCHAR(100) NULL,"
	 strSql = strSql &"  DIRECCION VARCHAR(100) NULL,"
	 strSql = strSql &"  DETALLE VARCHAR(100) NULL,"
	 strSql = strSql &"  E_COBRANZA VARCHAR(50) NULL,"
	 strSql = strSql &"  SITUACION VARCHAR(20) NULL)"

	'response.write "strSql " & strSql

	Conn.Execute strSql,64

	'response.write "Conn = " & Conn


	'**********CARGA ARCHIVO************'

	strSqlFile = "BULK INSERT TMP_CARGA_UMA_FACT FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2, CODEPAGE = 'ACP')"
	Conn.Execute strSqlFile,64


	'CUENTA LOS DOCUMENTOS A CARGAR'

	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_UMA_FACT WHERE COD_SAP NOT IN (SELECT COD_SAP FROM CARGA_UMA_FACT)"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intTotalDoc = rsTemp("CANTIDAD")
	Else
		intTotalDoc = 0
	End if

	'CUENTA LOS DEUDORES A CARGAR'

	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_UMA_FACT WHERE RUT NOT IN (SELECT RUT FROM CARGA_UMA_FACT) GROUP BY RUT "
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intTotalRutCarga = rsTemp("CANTIDAD")
	Else
		intTotalRutCarga = 0
	End if

	'CUENTA LOS CAMBIOS DE CUSTODIO DE UMA A LLACRUZ'

	strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_UMA_FACT INNER JOIN TMP_CARGA_UMA_FACT ON CARGA_UMA_FACT.COD_SAP = TMP_CARGA_UMA_FACT.COD_SAP WHERE CARGA_UMA_FACT.CUSTODIO = 'CARGA_INTERNA' AND '"& strTipoProceso &"' = 'CARGA_EXTERNA'"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCambioCustodio_Llacruz = rsTemp("CANTIDAD")
	Else
		intCambioCustodio_Llacruz = 0
	End if

	'CUENTA LOS CAMBIOS DE CUSTODIO DE LLACRUZ A UMA (NO SE PERMITEN)'

	strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_UMA_FACT INNER JOIN TMP_CARGA_UMA_FACT ON CARGA_UMA_FACT.COD_SAP = TMP_CARGA_UMA_FACT.COD_SAP WHERE CARGA_UMA_FACT.CUSTODIO = 'CARGA_EXTERNA' AND '"& strTipoProceso &"' = 'CARGA_INTERNA'"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCambioCustodio_UMA = rsTemp("CANTIDAD")
	Else
		intCambioCustodio_UMA = 0
	End if

	'CUENTA LOS CAMBIOS DE CUSTODIO DE DOCUMENTOS NO ACTIVOS'

	strSql = " SELECT COUNT(*) AS CANTIDAD"
	strSql = strSql &" FROM CARGA_UMA_FACT INNER JOIN TMP_CARGA_UMA_FACT ON CARGA_UMA_FACT.COD_SAP = TMP_CARGA_UMA_FACT.COD_SAP"
	strSql = strSql &" 			   	  INNER JOIN CUOTA ON CARGA_UMA_FACT.COD_SAP = CUOTA.NRO_CLIENTE_DOC"
	strSql = strSql &" 			      INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"

	strSql = strSql &" WHERE CARGA_UMA_FACT.CUSTODIO <> '"& strTipoProceso &"' AND ESTADO_DEUDA.ACTIVO = 0"

	set rsTemp= Conn.execute(strSql)

	if not rsTemp.eof then
		intCambioCustodio_NoActivo = rsTemp("CANTIDAD")
	Else
		intCambioCustodio_NoActivo = 0
	End if

	'CUENTA LOS REGISTROS DUPLICADOS EN SISTEMA'

	strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_UMA_FACT INNER JOIN TMP_CARGA_UMA_FACT ON CARGA_UMA_FACT.COD_SAP = TMP_CARGA_UMA_FACT.COD_SAP WHERE CARGA_UMA_FACT.CUSTODIO = '"& strTipoProceso &"'"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaDuplicados = rsTemp("CANTIDAD")
	Else
		intCargaDuplicados = 0
	End if


	'CUENTA LOS REGISTROS DUPLICADOS EN BASE DE CARGA'

	strSql = "SELECT COUNT(REPETIDOS) AS REPETIDOS FROM"
	strSql = strSql &" (SELECT ROW_NUMBER() OVER (PARTITION BY COD_SAP ORDER BY COD_SAP ASC) AS REPETIDOS FROM TMP_CARGA_UMA_FACT) AS REP"
	strSql = strSql &" WHERE REPETIDOS > 1"

	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaDuplicadosBase = rsTemp("REPETIDOS")
	Else
		intCargaDuplicadosBase = 0
	End if


	'-----------------------------SOLO SE PERMITE CARGAR Y REPROCESAR SI EN EL PROCESO NO HAY ERRORES-----------------------------'


		If Trim(intTotalRutCarga) = "" or IsNull(intTotalRutCarga) Then intTotalRutCarga = 0
		If Trim(intTotalDoc) = "" or IsNull(intTotalDoc) Then intTotalDoc = 0

		%>

		<table border=1 bgcolor="#<%=session("COLTABBG2")%>" class="Estilo28" width="700" align="center">

		<%If intCambioCustodio_UMA > 0 or intCargaDuplicados > 0 or intCargaDuplicadosBase >0 or intCambioCustodio_NoActivo >0 then %>

			<tr><td colspan=2 width="600" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13"><b>Estatus Carga (Existen errores, NO SE PUEDE CARGAR NADA NI CAMBIAR CUSTODIOS)</b></td></tr>
			<tr><td width="400">Tipo Carga: </td><td width="30" align="right"><%= strTipoProceso %></td></tr>
			<tr><td width="400">Cantidad Deudores a Cargar: </td><td width="30" align="right"><%=intTotalRutCarga%></td></tr>
			<tr><td width="400">Cantidad Documentos a Cargar: </td><td width="30" align="right"><%= intTotalDoc %></td></tr>
			<tr><td width="400">Cambio Custodio a Llacruz: </td><td width="30" align="right"><%=intCambioCustodio_Llacruz%></td></tr>


			<tr><td colspan=2 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13"><b>Errores Carga</b></td></tr>

		<%If intCargaDuplicados > 0 then%>
			<tr><td width="400">Documentos ya Cargados en sistema: </td><td width="30" align="right"><%=intCargaDuplicados%></td></tr>
		<%end if%>

		<%If intCargaDuplicadosBase > 0 then%>
			<tr><td width="400">Documentos Duplicados en Base de Carga: </td><td width="30" align="right"><%=intCargaDuplicadosBase%></td></tr>
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
			<tr><td width="400">Cambio Custodio a Llacruz: </td><td width="30" align="right"><%=intCambioCustodio_Llacruz%></td></tr>

		<% end if%>

		<tr><td colspan=2 align="center" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
		<br>
		<input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga_Cliente.asp');return false;">

		<%If intCambioCustodio_UMA > 0 then %>
		<input type="BUTTON" value="Procesar Igualmente" name="terminar" onClick="Reprocesar('<%=strTipoProceso%>');return false;"><br><br>
		<% end if%>

		<%If intCambioCustodio_UMA = 0 and intCargaDuplicados = 0 and intCargaDuplicadosBase = 0 and intCambioCustodio_NoActivo = 0 then %>
		<input type="BUTTON" value="Procesar" name="terminar" onClick="Procesar(<%=intTotalRutCarga%>,<%=intTotalDoc%>,<%=strCliente%>,'<%=strTipoProceso%>',<%=intCambioCustodio_Llacruz%>);return false;"><br><br>
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

