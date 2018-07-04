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
''******************************
%>
<%

 if Request("CB_CLIENTE") <> "" then
	strCliente=Request("CB_CLIENTE")
End if


if Request("Fecha") <> "" then
	Fecha=Request("Fecha")
End if

	strAsignacion = "100"

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

	strSqlFile = "DELETE FROM CARGA_UBICABILIDAD WHERE COD_CLIENTE = '" & Request("CB_CLIENTE") & "' AND COD_REMESA = " & strAsignacion
	Conn.Execute strSqlFile,64

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CARGA_UBICABILIDAD]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_CARGA_UBICABILIDAD]"
	Conn.Execute strSql,64


	strSql = " CREATE TABLE TMP_CARGA_UBICABILIDAD ( RUT varchar(12) NULL, CALLE varchar(200) NULL, NUMERO varchar(10) NULL, RESTO varchar(200) NULL, COMUNA varchar(50) NULL,"
	strSql = strSql & " CIUDAD varchar(50) NULL, COD_AREA_TP int NULL, TELEFONO_PARTICULAR varchar(10) NULL, COD_AREA_TC int NULL,"
	strSql = strSql & " TELEFONO_COMERCIAL varchar(15) NULL, TELEFONO_CELULAR varchar(15) NULL, COD_AREA_TA_1 int NULL, TELEFONO_ADIC_1 varchar(15) NULL,"
	strSql = strSql & " COD_AREA_TA_2 int NULL, TELEFONO_ADIC_2 varchar(15) NULL, EMAIL varchar(100) NULL, PROVEEDOR varchar(100) NULL )"

	Conn.Execute strSql,64

	'response.write "Conn = " & Conn
	'response.write "strSql " & strSql

	'**********CARGA ARCHIVO************'

	strSqlFile = "BULK INSERT TMP_CARGA_UBICABILIDAD FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2 ,CODEPAGE = 'ACP')"
	Conn.Execute strSqlFile,64

	'Response.write "strSqlFile=" & strSqlFile

	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_UBICABILIDAD WHERE LEN(TELEFONO_PARTICULAR) > 10 OR LEN(TELEFONO_COMERCIAL) > 10 OR LEN(TELEFONO_CELULAR) > 10 OR LEN(TELEFONO_ADIC_1) > 10 OR LEN(TELEFONO_ADIC_2) > 10"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intValida = rsTemp("CANTIDAD")
	Else
		intValida = 0
	End if

	If intValida > 0 Then
	%>
		<script>
			alert('Largo de los telefonos no puede exceder de 10 caracteres');
			history.back();
		</script>

	<%
		Response.End
	End if


	strSqlFile = "INSERT INTO CARGA_UBICABILIDAD SELECT " & Request("CB_CLIENTE") & "," & strAsignacion & ", * FROM TMP_CARGA_UBICABILIDAD"
	Conn.Execute strSqlFile,64


	'Response.write "strSqlFile=" & strSqlFile

'Response.End


	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_UBICABILIDAD"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intDeudaCarga = rsTemp("CANTIDAD")
	Else
		intDeudaCarga = 0
	End if

	strSql="SELECT (REPLACE(REPLACE(REPLACE(REPLACE(CONVERT([VARCHAR](30),GETDATE(),(121)),'-',''),' ',''),':',''),'.','')) AS FM"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		dtmFM = rsTemp("FM")
	Else
		dtmFM = "0"
	End if

	strObsCarga = now


	strSql = "exec proc_IngresaDirecciones " & strAsignacion & ",'" & Request("CB_CLIENTE") & "','" & Trim(session("session_login")) &"'"
	Conn.Execute strSql,64

	strSql = "exec proc_IngresaFonos_cel " & strAsignacion & ",'" & Request("CB_CLIENTE") & "','" & Trim(session("session_login")) &"'"
	Conn.Execute strSql,64

	strSql = "exec proc_IngresaFonos_tc " & strAsignacion & ",'" & Request("CB_CLIENTE") & "','" & Trim(session("session_login")) &"'"
	Conn.Execute strSql,64

	strSql = "exec proc_IngresaFonos_tp " & strAsignacion & ",'" & Request("CB_CLIENTE") & "','" & Trim(session("session_login")) &"'"
	Conn.Execute strSql,64

	strSql = "exec proc_IngresaFonos_ta " & strAsignacion & ",'" & Request("CB_CLIENTE") & "','" & Trim(session("session_login")) &"'"
	'Response.write strSql
	Conn.Execute strSql,64

	strSql = "exec proc_IngresaFonos_ta2 " & strAsignacion & ",'" & Request("CB_CLIENTE") & "','" & Trim(session("session_login")) &"'"
	Conn.Execute strSql,64

	strSql = "exec proc_IngresaEmail " & strAsignacion & ",'" & Request("CB_CLIENTE") & "','" & Trim(session("session_login")) &"'"
	Conn.Execute strSql,64

	strSql = "exec proc_IngresaEmail_adic " & strAsignacion & ",'" & Request("CB_CLIENTE") & "','" & Trim(session("session_login")) &"'"
	Conn.Execute strSql,64



	strFuente = "CLIENTE : " & Request("CB_CLIENTE") & " REMESA : " & strAsignacion


	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM DEUDOR_DIRECCION WHERE FUENTE = '" & strFuente & "'"

	'response.write "<br>strSql = " & strSql

	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intDirNueva = rsTemp("CANTIDAD")
	Else
		intDirNueva = 0
	End if

	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM DEUDOR_TELEFONO WHERE FUENTE = '" & strFuente & "'"
	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intTelNuevo = rsTemp("CANTIDAD")
	Else
		intTelNuevo = 0
	End if
	
	'response.write "<br>strSql = " & strSql
	'response.End

	%>
	<table>
	<!--tr><td>Direcciones Nuevas: <%= intDirNueva %>&nbsp;<td>
	<tr><td>Telefonos Nuevos: <%= intTelNuevo %>&nbsp;<td-->
	<!--tr><td>Terceros Cargados : <%= cantTercerosC %>&nbsp;<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTerceros%>')">Ver</a></td></tr-->
	Proceso realizado correctamente
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

