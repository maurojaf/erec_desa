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
''*****************************
%>
<%

 if Request("CB_CLIENTE") <> "" then
	strCliente=Request("CB_CLIENTE")
End if


if Request("Fecha") <> "" then
	Fecha=Request("Fecha")
End if

if Request("TX_FECHACREACION") <> "" then
	dtmFecCreacion = "'" & Request("TX_FECHACREACION") & "'"
Else
	dtmFecCreacion="getdate()"
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


Server.ScriptTimeout = 90000
Conn.ConnectionTimeout = 90000


''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

If strArchivo <> "" Then


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "Terceros_cargados_"&Fecha&".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros

	strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"&strCliente &"/" & strArchivo

	AbrirScg()

	strSqlFile = "DELETE FROM CARGA_DEUDA WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "' AND COD_REMESA = 0"
	Conn.Execute strSqlFile,64

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CARGA_DEUDA]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_CARGA_DEUDA]"
	Conn.Execute strSql,64


	strSql = " CREATE TABLE TMP_CARGA_DEUDA ( RUT varchar(12) NULL, NRO_DOC varchar(25) NULL,"
	strSql = strSql & " CUOTA tinyint NULL, TIPO_DOC varchar(30) NULL, PROTESTO varchar(60) NULL, FEC_VENC datetime NULL, FEC_PROTESTO datetime NULL,"
	strSql = strSql & " FEC_EMISION datetime NULL, MONTO numeric(12, 2) NULL, MONTO_PROTESTO numeric(12, 2) NULL, BANCO varchar(30) NULL,"
	strSql = strSql & " RUT_DEUDOR varchar(12) NULL, OBSERVACIONES varchar(80) NULL, ADICIONAL1 varchar(50) NULL,"
	strSql = strSql & " ADICIONAL2 varchar(50) NULL, ADICIONAL3 varchar(50) NULL, ADICIONAL4 varchar(50) NULL, ADICIONAL5 varchar(50) NULL) "


	Conn.Execute strSql,64

	CerrarSCG()

	'response.write "Conn = " & Conn
	'response.write "strSql " & strSql

	'**********CARGA ARCHIVO************'

	AbrirScg()

	strSqlFile = "BULK INSERT TMP_CARGA_DEUDA FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2)"
	Conn.Execute strSqlFile,64

	strSqlFile = "INSERT INTO CARGA_DEUDA SELECT '" & Request("CB_CLIENTE") & "',0, *, '" & Request("CB_CLIENTE") & "' + '-' + CAST(RUT AS VARCHAR(12)) + '-' + NRO_DOC + '-' + CAST(CUOTA AS VARCHAR(10)) FROM TMP_CARGA_DEUDA"

	'rESPONSE.WRITE "strSql=" & strSqlFile
	'rESPONSE.eND

	Conn.Execute strSqlFile,64

	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_DEUDA"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intDeudaCarga = rsTemp("CANTIDAD")
	Else
		intDeudaCarga = 0
	End if

	CerrarSCG()

	strObsCarga = now


	AbrirScg()

	strSql = "INSERT INTO CUOTA (RUT_DEUDOR, COD_CLIENTE, NRO_DOC, NRO_CUOTA, TIPO_DOCUMENTO, VALOR_CUOTA, SALDO, "
	strSql = strSql & " FECHA_VENC, FECHA_EMISION, ADIC_1, ADIC_2, ADIC_3, ADIC_4, ADIC_5, COD_REMESA,"
	strSql = strSql & " USUARIO_CREACION, FECHA_CREACION,ESTADO_DEUDA,FECHA_ESTADO, OBSERVACION,GASTOS_PROTESTOS) "
	strSql = strSql & " SELECT RUT, COD_CLIENTE, NRO_DOC, CUOTA, TIPO_DOC,  SUM(MONTO), SUM(MONTO), "
	strSql = strSql & " FEC_VENC, FEC_EMISION,ADICIONAL1, ADICIONAL2,ADICIONAL3,ADICIONAL4,ADICIONAL5, " & strAsignacion & ", "
	strSql = strSql & " 1," & dtmFecCreacion & ", 1, " & dtmFecCreacion & ",OBSERVACIONES,MONTO_PROTESTO " "
	strSql = strSql & " FROM CARGA_DEUDA WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "' AND COD_REMESA = 0 "


	strSql = strSql & " AND RUT IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "')"
	strSql = strSql & " GROUP BY RUT, COD_CLIENTE, NRO_DOC, CUOTA, TIPO_DOC, FEC_VENC,FEC_EMISION, FEC_PROTESTO,ADICIONAL1, ADICIONAL2,ADICIONAL3,ADICIONAL4,ADICIONAL5, COD_REMESA, OBSERVACIONES, MONTO_PROTESTO "

	'rESPONSE.WRITE "strSql=" & strSql
	'rESPONSE.eND
	Conn.Execute strSql,64


	'conectamos con el FSO
	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero = "COD_CLIENTE;RUT_DEUDOR;NRO_DOC;NRO_CUOTA;TIPO_DOCUMENTO;SALDO"
	fichCA.writeline(strTextoTercero)

	strSql = "SELECT COD_CLIENTE, RUT_DEUDOR, NRO_DOC, NRO_CUOTA, TIPO_DOCUMENTO, SALDO FROM CUOTA WHERE COD_REMESA = 0 AND COD_CLIENTE = '" & Request("CB_CLIENTE") & "'"
	set rsTemp= Conn.execute(strSql)
	Do While Not rsTemp.eof



		rsTemp.movenext
	Loop


	strSql = "UPDATE A SET A.COD_REMESA = B.COD_REMESA"
	strSql = strSql & " FROM CUOTA A, CUOTA B"
	strSql = strSql & " WHERE A.RUT_DEUDOR = B.RUT_DEUDOR"
	strSql = strSql & " AND A.COD_REMESA = 0 AND B.COD_REMESA <> 0 AND A.COD_CLIENTE = B.COD_CLIENTE AND A.COD_CLIENTE = '" & Request("CB_CLIENTE") & "'"

	'rESPONSE.WRITE "strSql=" & strSql
	'rESPONSE.eND


	Conn.Execute strSql,64

	CerrarSCG()

	AbrirScg()

	strSql = "UPDATE C2 SET USUARIO_ASIG = C1.USUARIO_ASIG"
	strSql = strSql & " FROM CUOTA C1, CUOTA C2 "
	strSql = strSql & " WHERE C1.COD_CLIENTE = C2.COD_CLIENTE"
	strSql = strSql & " AND C1.RUT_DEUDOR = C2.RUT_DEUDOR "
	strSql = strSql & " AND C2.USUARIO_ASIG IS NULL"
	strSql = strSql & " AND C1.USUARIO_ASIG IS NOT NULL"


	Conn.Execute strSql,64

	CerrarSCG()

	'response.write "now = " & now
	'response.End

	%>
	<table>
	<tr><td>Cantidad Registros Totales: <%= intDeudaCarga %>&nbsp;<td>
	<!--tr><td>Terceros Cargados : <%= cantTercerosC %>&nbsp;<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTerceros%>')">Ver</a></td></tr>
	<tr><td>Terceros Actualizados : <%= cantTercerosA %>&nbsp;<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTercerosA%>')">Ver</a></td></tr-->

	<% if sIopAc = 1 then %>

	<tr><td>Terceros Actualizados : <%= cantSiniestroC %>&nbsp;<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoSiniestrosA%>')">Ver</a></td></tr>

	<% end if%>

	</table>
 <%



	'creamos el objeto TextStream

	'response.write "terceroCSV=" & terceroCSV
	'response.End


	fichCA.close()

End if


function lipiatexto(texto)

	if isnull(texto) then
		texto = ""
	end if


	texto = replace(texto,"'","")
	texto = replace(texto,"  "," ")
	texto = replace(texto,".","")
	texto = replace(texto,chr(44)," ")
	texto = replace(texto,"_","")
	texto = replace(texto,"--","-")

lipiatexto =texto

End function

function lipiatelefono(texto)

	if isnull(texto) then
		texto = ""
	end if


	texto = replace(texto,"(","")
	texto = replace(texto,")","")
	texto = replace(texto,".","")
	texto = replace(texto,"_","")
	texto = replace(texto,"--","-")
	texto = replace(texto,"/","")
	texto = replace(texto,"\","\")

lipiatelefono =texto

End function


function fechaYYMMDD(fechaI)

FechaInv= Year(fechaI) & "-" & right("00"&Day(fechaI), 2) & "-" &  right("00"&(Month(fechaI)), 2)

fechaYYMMDD = FechaInv

End function

function SioNo(valor)

	min = LCase(valor)

	if min = "si" OR min = "s" then
		ValorI = 1
	else
		ValorI = 0
	End if

SioNo = ValorI

End function

Function codigo_veri(ruts)
	rut= lipiatelefono(ruts)

	tur=strreverse(rut)
	mult = 2

	for i = 1 to len(tur)
	if mult > 7 then mult = 2 end if

	suma = mult * mid(tur,i,1) + suma
	mult = mult +1
	next

	valor = 11 - (suma mod 11)

	if valor = 11 then
	codigo_veri = "0"
	elseif valor = 10 then
	codigo_veri = "k"
	else
	codigo_veri = valor
	end if

End function








%>







				</td>
			  </tr>
			</table>


		</td>

	</tr>

</table>

</body>
</html>

