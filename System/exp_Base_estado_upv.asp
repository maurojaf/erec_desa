<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<!--#include file="arch_utils.asp"-->
<!--#include file="arch_utils_upv.asp"-->
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
'******************************
%>
<%

 if Request("CB_CLIENTE") <> "" then
	strCliente=Trim(Request("CB_CLIENTE"))
End if

 if Request("CB_USUARIO") <> "" then
	strUsuario=Request("CB_USUARIO")
End if

dtmInicio = Request("dtmInicio")
dtmTermino = Request("dtmTermino")

	'Response.write "inicio=" & Request("dtmInicio")
	'Response.write "dtmTermino=" & Request("dtmTermino")
CH_EFECTIVA=UCASE(Request("CH_EFECTIVA"))

'Response.write "<b>CH_DIASGTE=" & Request("CH_DIASGTE")


	'Response.End


if Request("Fecha") <> "" then
	Fecha=Request("Fecha")
End if

if Request("archivo") <> "" then
	strArchivo=Request("archivo")
End if
strArchivo=1

if Request("CH_ACTIVO")= "true" then
	sIopAc=1
else
	sIopAc=0
End if

abrirscg()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc



If strArchivo <> "" Then

	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "export_Base_Estado_Upv_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivoTerceros


	''terceroCSV = "F:\" & strNomArchivoTerceros
	'response.write terceroCSV

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)
	strTextoTercero=""

	strTextoTercero = "[ID_CUOTA];CUSTODIO;FECHA_CUSTODIO;FECHA_ESTADO;ID_DEUDOR;TIPO_DOC;NUMERO_DOCUMENTO;CODAPOD;NOMBRE_AVAL;CODCLI;NOMBRE_ALUMNO;MONTO_DEUDO;FECHA_VENCIMIENTO;GASTO_PROTESTO;SEDE;CAMPUS;CARRERA;DIRECCION;COMUNA;FONO1;FONO2;FONO3;EMAIL1;EMAIL2;ESTADO;NOMBRE_ESTADO;UBICACION;ANO;FECHA_CANCELACION"
	fichCA.writeline(strTextoTercero)

	strSql="SELECT [id_deudor],"
	strSql = strSql & " [tipo_doc],"
	strSql = strSql & " [numero_documento],"
	strSql = strSql & " [codapod],"
	strSql = strSql & " [nombre_aval],"
	strSql = strSql & " [codcli],"
	strSql = strSql & " [nombre_alumno],"
	strSql = strSql & " [monto_deudo],"
	strSql = strSql & " [fecha_vencimiento],"
	strSql = strSql & " [gasto_protesto],"
	strSql = strSql & " [Sede],"
	strSql = strSql & " [Campus],"
	strSql = strSql & " [carrera],"
	strSql = strSql & " [direccion],"
	strSql = strSql & " [comuna],"
	strSql = strSql & " [fono1],"
	strSql = strSql & " [fono2],"
	strSql = strSql & " [fono3],"
	strSql = strSql & " [email1],"
	strSql = strSql & " [email2],"
	strSql = strSql & " [estado],"
	strSql = strSql & " [nombre_estado],"
	strSql = strSql & " [ubicacion],"
	strSql = strSql & " [ano],"
	strSql = strSql & " [fecha_cancelacion] "
	strSql = strSql & " FROM mt_llacruz "


	strSql="SELECT ISNULL(cast(B.id_cuota as varchar(12)),'NO_CARGADO') AS ID_CUOTA,ISNULL(B.CUSTODIO,'LLACRUZ') AS CUSTODIO,"
	strSql = strSql & " ISNULL(CONVERT(VARCHAR(10),ISNULL(B.FECHA_ESTADO_CUSTODIO,B.FECHA_CREACION),103),'NO CARGADO') AS FECHACUSTODIO, "
	strSql = strSql & " ISNULL(CONVERT(VARCHAR(10),B.FECHA_ESTADO,103),'NO CARGADO') AS FECHAESTADO, "
	strSql = strSql & " A.[id_deudor],A.TIPO_DOC, "
	strSql = strSql & " (CASE WHEN A.TIPO_DOC = 8 THEN 4  WHEN A.TIPO_DOC = 9 THEN 2  WHEN A.TIPO_DOC = 2 THEN 18  WHEN A.TIPO_DOC = 3 THEN 19 ELSE A.TIPO_DOC END) AS TIPO_DOC_LLACRUZ,"
	strSql = strSql & " A.[numero_documento], A.[codapod], A.[nombre_aval], A.[codcli], A.[nombre_alumno], A.[monto_deudo], A.[fecha_vencimiento], A.[gasto_protesto], A.[Sede], A.[Campus], A.[carrera], A.[direccion], A.[comuna], A.[fono1], A.[fono2], A.[fono3], A.[email1], A.[email2], A.[estado], A.[nombre_estado], A.[ubicacion], A.[ano], A.[fecha_cancelacion] "
	strSql = strSql & " FROM [200.54.110.124].[matricula].[matricula].[mt_llacruz] A"
	strSql = strSql & " left join Cuota B on  B.COD_CLIENTE = '1200'"
	strSql = strSql & " AND SUBSTRING(RUT_DEUDOR,1,len(RUT_DEUDOR)-2) = cast(A.codcli as varchar(12)) collate SQL_Latin1_General_CP1_CI_AS "
	strSql = strSql & " and B.NRO_DOC = cast(a.numero_documento as varchar(12)) collate SQL_Latin1_General_CP1_CI_AS "
	strSql = strSql & " and B.VALOR_CUOTA = A.monto_deudo"
	strSql = strSql & " AND (CASE WHEN A.TIPO_DOC = 8 THEN 4  WHEN A.TIPO_DOC = 9 THEN 2  WHEN A.TIPO_DOC = 2 THEN 18  WHEN A.TIPO_DOC = 3 THEN 19 ELSE A.TIPO_DOC END) = B.TIPO_DOCUMENTO"

	'Response.write "<br>strSql=" & strSql

	Conn.Execute strSql,64
	set rsTemp= Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsTemp.Eof

		strTextoTercero = 	rsTemp("ID_CUOTA") & ";" & rsTemp("CUSTODIO") & ";" & rsTemp("FECHACUSTODIO") & ";" & rsTemp("FECHAESTADO") & ";" & rsTemp("id_deudor")& ";" &rsTemp("tipo_doc")& ";" &rsTemp("numero_documento")& ";" &rsTemp("codapod")& ";" &rsTemp("nombre_aval")& ";" &rsTemp("codcli")& ";" &rsTemp("nombre_alumno")& ";" &rsTemp("monto_deudo")& ";" &rsTemp("fecha_vencimiento")& ";" &rsTemp("gasto_protesto")& ";" &rsTemp("Sede")& ";" &rsTemp("Campus")& ";" &rsTemp("carrera")& ";" &rsTemp("direccion")& ";" &rsTemp("comuna")& ";" &rsTemp("fono1")& ";" &rsTemp("fono2")& ";" &rsTemp("fono3")& ";" &rsTemp("email1")& ";" &rsTemp("email2")& ";" &rsTemp("estado")& ";" &rsTemp("nombre_estado")& ";" &rsTemp("ubicacion")& ";" &rsTemp("ano")& ";" &rsTemp("fecha_cancelacion")
		cantSiniestroC = cantSiniestroC + 1
		fichCA.writeline(strTextoTercero)
		rsTemp.movenext

	Loop

	%>
	<table>
	<tr><td>Cantidad de registros generados : <%= cantSiniestroC %></td></tr>
	<tr><td>
	<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTerceros%>')">Descargar</a>
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

