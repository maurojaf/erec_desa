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
function AbreArchivo(nombre){
window.open(nombre,"INFORMACION","width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes");
}
function Terminar( sintPaginaTerminar ) {
        self.location.href = sintPaginaTerminar
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

'******************************
'*	INICIO CODIGO PARTICULAR  *
''******************************
%>
<%

 if Request("CB_CLIENTE") <> "" then
	strCliente=Request("CB_CLIENTE")
End if


if Request("Fecha") <> "" then
	Fecha=Request("Fecha")
End if


if Request("Asignacion") <> "Seleccionar" then
	strAsignacion=Request("Asignacion")
else
	strAsignacion = 0
End if

if Request("archivo") <> "" then
	strArchivo=Request("archivo")
End if

if Request("opAc")= "0" then
	sIopAc=0
else
	sIopAc=1
End if



AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

If strArchivo <> "" Then


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "Terceros_cargados_"&Fecha&".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros

	strTextoTercero = strTextoTercero & "ID_TERCERO;PATENTE;RUT;NOMBRE;MARCA;MODELO;TELEFONO_1;TELEFONO_2;TELEFONO_3;DIRECCION;COMUNA;CIUDAD" & chr(13) & chr(10)

	strTextoArchivoCC = ""
	strTextoArchivoCNC = ""
	strTextoArchivoCA = ""


	strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"&strCliente &"/" & strArchivo

'response.write strFileDir
	strSqlFile = "DELETE FROM CARGA_DEUDOR WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "'"
	Conn.Execute strSqlFile,64

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CARGA_DEUDOR]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_CARGA_DEUDOR]"
	Conn.Execute strSql,64


	strSql = " CREATE TABLE TMP_CARGA_DEUDOR ( RUT VARCHAR(10) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NOT NULL, NOMBRE VARCHAR(150) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL,"
	strSql = strSql &" RUT_REP_LEGAL VARCHAR(12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, NOM_REP_LEGAL VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL, ETAPA_COBRANZA INT NULL, ADIC_1 VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL, ADIC_2 VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL, ADIC_3 VARCHAR(100) COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS NULL)"
	Conn.Execute strSql,64

	'response.write "Conn = " & Conn
	'response.write "strSql " & strSql

	'**********CARGA ARCHIVO************'

	strSqlFile = "BULK INSERT TMP_CARGA_DEUDOR FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2, CODEPAGE = 'ACP')"
	Conn.Execute strSqlFile,64

	strSqlFile = "INSERT INTO CARGA_DEUDOR SELECT " & Request("CB_CLIENTE") & ",0, * FROM TMP_CARGA_DEUDOR"
	Conn.Execute strSqlFile,64


	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_DEUDOR"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intDeudoresCarga = rsTemp("CANTIDAD")
	Else
		intDeudoresCarga = 0
	End if


	strSql = "SELECT COUNT(*) AS CANTIDAD FROM CARGA_DEUDOR WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "' GROUP BY RUT HAVING COUNT(*) > 1"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intDuplicados = rsTemp("CANTIDAD")
	Else
		intDuplicados = 0
	End if


	if intDuplicados > 0 then
		Response.write "Existen RUT duplicados, favor validar, no puede existir rut duplicados en el archivo a cargar."
		%>
		<br><br><input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga.asp');return false;">
		<%

		Response.End
	End If

	''Response.End

	strObsCarga = now

	strSql = "INSERT INTO DEUDOR (RUT_DEUDOR, NOMBRE_DEUDOR,FECHA_INGRESO, USUARIO_INGRESO, COD_CLIENTE, REPLEG_RUT, REPLEG_NOMBRE, OBS_CARGA, ETAPA_COBRANZA, ADIC_1, ADIC_2, ADIC_3)"
	strSql = strSql & " SELECT DISTINCT RUT,NOMBRE,GETDATE(),1, COD_CLIENTE, RUT_REP_LEGAL, NOM_REP_LEGAL, '" & strObsCarga & "', ETAPA_COBRANZA, ADIC_1, ADIC_2, ADIC_3"
	strSql = strSql & " FROM CARGA_DEUDOR WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "' AND RUT NOT IN ( "
	strSql = strSql & " SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "') "
	'Response.write "strSql = " & strSql
	'Response.eND
	Conn.Execute strSql,64

	strSql = "UPDATE REMESA SET OBS_CARGA = '" & strObsCarga & "' WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "'"
	Conn.Execute strSql,64

	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM DEUDOR WHERE OBS_CARGA = '" & strObsCarga & "'"
		set rsTemp= Conn.execute(strSql)
		if not rsTemp.eof then
			intDeudoresNuevos = rsTemp("CANTIDAD")
		Else
			intDeudoresNuevos = 0
	End if

	'response.write "now = " & now
	'response.End
	

	%>
	<table border=1 bordercolor="#000000" width="300">

	<tr><td colspan=2><b>Estatus Carga</b></td></tr>
	<tr><td width="270">Cantidad Registros Totales: </td><td width="30" align="right"><%= intDeudoresCarga %></td></tr>
	<tr><td>Cantidad Registros Nuevos (cargados): </td><td align="right"><%= intDeudoresNuevos %></td></tr>

	<tr><td colspan=2 align="center">
	<br><input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga.asp');return false;"><br><br>
	</td></tr>



	<!--tr><td>Terceros Cargados : <%= cantTercerosC %>&nbsp;<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTerceros%>')">Ver</a></td></tr>
	<tr><td>Terceros Actualizados : <%= cantTercerosA %>&nbsp;<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTercerosA%>')">Ver</a></td></tr-->

	<% if sIopAc = 1 then %>

	<tr><td>Terceros Actualizados : <%= cantSiniestroC %>&nbsp;<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoSiniestrosA%>')">Ver</a></td></tr>

	<% end if%>

	</table>
 <%


	'conectamos con el FSO
	set confile = createObject("scripting.filesystemobject")
	'creamos el objeto TextStream

	''response.write "terceroCSV=" & terceroCSV
	''response.End


	''ELIMINACION TEMPORAL ALG
	''set fichCA = confile.CreateTextFile(terceroCSV)
	''fichCA.write(strTextoTercero)
	''fichCA.close()
End if


%>







				</td>
			  </tr>
			</table>


		</td>

	</tr>

</table>

</body>
</html>

