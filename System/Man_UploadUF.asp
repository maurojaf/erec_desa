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

	strTextoTercero = strTextoTercero & "ID_TERCERO;PATENTE;RUT;NOMBRE;MARCA;MODELO;TELEFONO_1;TELEFONO_2;TELEFONO3;DIRECCION;COMUNA;CIUDAD" & chr(13) & chr(10)

	strTextoArchivoCC = ""
	strTextoArchivoCNC = ""
	strTextoArchivoCA = ""


	strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"&strCliente &"/" & strArchivo


	strSqlFile = "DELETE FROM CARGA_UF"
	Conn.Execute strSqlFile,64

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CARGA_UF]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_CARGA_UF]"
	Conn.Execute strSql,64


	strSql = " CREATE TABLE TMP_CARGA_UF ( FECHA DATETIME NOT NULL, VALOR FLOAT NULL)"
	Conn.Execute strSql,64

	'response.write "Conn = " & Conn
	'response.write "strSql " & strSql

	'**********CARGA ARCHIVO************'

	strSqlFile = "BULK INSERT TMP_CARGA_UF FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2)"
	'Response.write strSqlFile

	Conn.Execute strSqlFile,64

	strSqlFile = "INSERT INTO CARGA_UF SELECT * FROM TMP_CARGA_UF"
	Conn.Execute strSqlFile,64


	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_UF"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intDeudoresCarga = rsTemp("CANTIDAD")
	Else
		intDeudoresCarga = 0
	End if


	strSql = "SELECT COUNT(*),FECHA AS CANTIDAD FROM CARGA_UF GROUP BY FECHA HAVING COUNT(*) > 1"
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

	strSql = "INSERT INTO UNIDAD_FOMENTO (FECHA, VALOR)"
	strSql = strSql & " SELECT DISTINCT FECHA,VALOR"
	strSql = strSql & " FROM CARGA_UF WHERE FECHA NOT IN ( "
	strSql = strSql & " SELECT FECHA FROM UNIDAD_FOMENTO ) "
	'Response.write "strSql = " & strSql
	'Response.eND
	Conn.Execute strSql,64


	'response.write "now = " & now
	'response.End

	%>
	<table>
	<tr><td>Cantidad Registros Totales: <%= intDeudoresCarga %>&nbsp;<td>
	<tr><td>Cantidad Registros Nuevos: <%= intDeudoresNuevos %>&nbsp;<td>
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

	'response.write "terceroCSV=" & terceroCSV
	'response.End

	'set fichCA = confile.CreateTextFile(terceroCSV)
	'fichCA.write(strTextoTercero)
	'fichCA.close()

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

