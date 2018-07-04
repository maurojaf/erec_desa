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

dtmFecIniEstado=Request("dtmInicio")

dtmFecTerEstado=Request("dtmTermino")


'Response.write "<b>dtmFecIniEstado=" & dtmFecIniEstado

'Response.write "<b>dtmFecTerEstado=" & dtmFecTerEstado

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

AbrirScg()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc



If strArchivo <> "" Then

	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "export_Base_Estado_UMA_" & session("session_idusuario") & "_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivoTerceros

'response.write terceroCSV
	'terceroCSV = "F:\" & strNomArchivoTerceros


	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero=""

	strTextoTercero = "USUARIO_CARGA;FECHA_ASIGNACION;ESTADO_DEUDA;USUARIO_ESTADO;FECHA_ESTADO;CUSTODIO;FECHA_CUSTODIO;USUARIO_CUSTODIO;ID_CUOTA;COD_CLIENTE;COD_SAP;FPO_4;IMPORTE;CORE;DIVI;TI;RUT_ALUMNO;ALUMNO;NOMBRE_ALUMNO;RUT_GIRADOR;GIRADOR;NOMBRE_GIRADOR;RUT_AVAL;AVAL;NOMBRE_AVAL;TIPO_DOCUM;DOCUMENTO_SA;POSI;A;FOLIO;BANCO;NOMBRE_DEL_BANCO;PLAZA;CUENTA;MONEDA;MONTO;VENCIMI;VENCIMI_2;LUGAR;SITUACION;ESTADO;GARANTIA;REMESA;LLAVE_OPER;USUARIO;FECHA;NUMERO_DE_PR;POSI_2;TXT_EXPLCATIVO_P_POS;TIPO;PROTESTO;GASTO_DE_PROTESTO;TEL_1_A;TEL_2_A;DIRECCION_A;COMUNA;CIUDAD;REGION;EMAIL_A;TEL_1_G;TEL_2_G;DIRECCION_G;EMAIL_G"

	fichCA.writeline(strTextoTercero)

	strSql="SELECT CUOTA.[COD_CLIENTE],"
	strSql = strSql & " [COD_SAP],"
	strSql = strSql & " [FPO_4],"
	strSql = strSql & " [IMPORTE],"
	strSql = strSql & " [CORE],"
	strSql = strSql & " [DIVI],"
	strSql = strSql & " [TI],"
	strSql = strSql & " [RUT_ALUMNO],"
	strSql = strSql & " [ALUMNO],"
	strSql = strSql & " [NOMBRE_ALUMNO],"
	strSql = strSql & " [RUT_GIRADOR],"
	strSql = strSql & " [GIRADOR],"
	strSql = strSql & " [NOMBRE_GIRADOR],"
	strSql = strSql & " [RUT_AVAL],"
	strSql = strSql & " [AVAL],"
	strSql = strSql & " [NOMBRE_AVAL],"
	strSql = strSql & " [TIPO_DOCUM],"
	strSql = strSql & " [DOCUMENTO_SA],"
	strSql = strSql & " [POSI],"
	strSql = strSql & " [A],"
	strSql = strSql & " [FOLIO],"
	strSql = strSql & " [BANCO],"
	strSql = strSql & " [NOMBRE_DEL_BANCO],"
	strSql = strSql & " [PLAZA],"
	strSql = strSql & " CARGA_UMA.CUENTA AS CUENTA,"
	strSql = strSql & " [MONEDA],"
	strSql = strSql & " [MONTO],"
	strSql = strSql & " [VENCIMI],"
	strSql = strSql & " [VENCIMI_2],"
	strSql = strSql & " [LUGAR],"
	strSql = strSql & " [SITUACION],"
	strSql = strSql & " [ESTADO],"
	strSql = strSql & " [GARANTIA],"
	strSql = strSql & " [REMESA],"
	strSql = strSql & " [LLAVE_OPER],"
	strSql = strSql & " [USUARIO],"
	strSql = strSql & " [FECHA],"
	strSql = strSql & " [NUMERO_DE_PR],"
	strSql = strSql & " [POSI_2],"
	strSql = strSql & " [TXT_EXPLCATIVO_P_POS],"
	strSql = strSql & " [TIPO],"
	strSql = strSql & " [PROTESTO],"
	strSql = strSql & " [GASTO_DE_PROTESTO],"
	strSql = strSql & " [TEL_1_A],"
	strSql = strSql & " [TEL_2_A],"
	strSql = strSql & " [DIRECCION_A],"
	strSql = strSql & " [COMUNA],"
	strSql = strSql & " [CIUDAD],"
	strSql = strSql & " [REGION],"
	strSql = strSql & " [EMAIL_A],"
	strSql = strSql & " [TEL_1_G],"
	strSql = strSql & " [TEL_2_G],"
	strSql = strSql & " [DIRECCION_G],"
	strSql = strSql & " [EMAIL_G],"
	strSql = strSql & " (CASE WHEN CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = '' THEN 'LLACRUZ' ELSE CUOTA.CUSTODIO END) AS CUSTODIO,"
	strSql = strSql & " [FECHA_CARGA],"
	strSql = strSql & " U2.LOGIN AS USUARIO_CARGA,"
	strSql = strSql & " U1.LOGIN AS USUARIO_CUSTODIO,"
	strSql = strSql & " CUOTA.OBSERVACION AS USUARIO_ESTADO,"
	strSql = strSql & " CONVERT(VARCHAR(19),FECHA_CUSTODIO,103) AS FECHA_CUSTODIO,"
	strSql = strSql & " CONVERT(VARCHAR(19),CUOTA.FECHA_CREACION,103) AS FECHA_ASIGNACION,"
	strSql = strSql & " (CASE WHEN (ESTADO_DEUDA.DESCRIPCION = 'PAGADA EN EMPRESA' OR ESTADO_DEUDA.DESCRIPCION = 'CONVENIO') THEN 'PAGADA EN LLACRUZ' WHEN ESTADO_DEUDA.DESCRIPCION = 'PAGADA EN CLIENTE' THEN 'PAGADA EN UMA' ELSE ESTADO_DEUDA.DESCRIPCION END) AS ESTADO_DEUDA,"
	strSql = strSql & " CONVERT(VARCHAR(19),CUOTA.FECHA_ESTADO,103) AS FECHA_ESTADO,"
	strSql = strSql & " CUOTA.ID_CUOTA AS ID_CUOTA"


	strSql = strSql & " FROM CUOTA INNER JOIN CARGA_UMA ON CUOTA.NRO_CLIENTE_DOC = CARGA_UMA.COD_SAP"
	strSql = strSql & "		   	   INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql & "		   	   LEFT JOIN USUARIO AS U1 ON CARGA_UMA.USUARIO_CUSTODIO = U1.ID_USUARIO"
	strSql = strSql & "		   	   LEFT JOIN USUARIO AS U2 ON CARGA_UMA.USUARIO_CARGA = U2.ID_USUARIO"


	strSql = strSql & " WHERE CUOTA.COD_CLIENTE = '" & strCliente & "'"

	If Trim(dtmFecIniEstado) <> "" Then
		strSql = strSql & " AND (CUOTA.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) OR CONVERT(VARCHAR(10),CUOTA.FECHA_ESTADO,103) >= CAST('" & dtmFecIniEstado & "' AS DATETIME)) "
	End If

	If Trim(dtmFecIniEstado) <> "" AND Trim(dtmFecTerEstado) <> "" Then
		strSql = strSql & " AND CONVERT(VARCHAR(10),CUOTA.FECHA_ESTADO,103) <= CAST('" & dtmFecTerEstado & "' AS DATETIME) "
	End If

	If Trim(dtmFecIniEstado) = "" Then
		strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
	End If

	''Response.write "<b>strSql=" & strSql

	Conn.Execute strSql,64

	set rsTemp= Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsTemp.Eof

		strTextoTercero = rsTemp("USUARIO_CARGA")& ";" &rsTemp("FECHA_ASIGNACION")& ";" &rsTemp("ESTADO_DEUDA")& ";" &rsTemp("USUARIO_ESTADO")& ";" &rsTemp("FECHA_ESTADO")& ";" &rsTemp("CUSTODIO")& ";" &rsTemp("FECHA_CUSTODIO")& ";" &rsTemp("USUARIO_CUSTODIO")& ";" &rsTemp("ID_CUOTA")& ";" &rsTemp("COD_CLIENTE")& ";" &rsTemp("COD_SAP")& ";" &rsTemp("FPO_4")& ";" &rsTemp("IMPORTE")& ";" &rsTemp("CORE")& ";" &rsTemp("DIVI")& ";" &rsTemp("TI")& ";" &rsTemp("RUT_ALUMNO")& ";" &rsTemp("ALUMNO")& ";" &rsTemp("NOMBRE_ALUMNO")& ";" &rsTemp("RUT_GIRADOR")& ";" &rsTemp("GIRADOR")& ";" &rsTemp("NOMBRE_GIRADOR")& ";" &rsTemp("RUT_AVAL")& ";" &rsTemp("AVAL")& ";" &rsTemp("NOMBRE_AVAL")& ";" &rsTemp("TIPO_DOCUM")& ";" &rsTemp("DOCUMENTO_SA")& ";" &rsTemp("POSI")& ";" &rsTemp("A")& ";" &rsTemp("FOLIO")& ";" &rsTemp("NOMBRE_DEL_BANCO")& ";" &rsTemp("NOMBRE_DEL_BANCO")& ";" &rsTemp("PLAZA")& ";" &rsTemp("CUENTA")& ";" &rsTemp("MONEDA")& ";" &rsTemp("MONTO")& ";" &rsTemp("VENCIMI")& ";" &rsTemp("VENCIMI_2")& ";" &rsTemp("LUGAR")& ";" &rsTemp("SITUACION")& ";" &rsTemp("ESTADO")& ";" &rsTemp("GARANTIA")& ";" &rsTemp("REMESA")& ";" &rsTemp("LLAVE_OPER")& ";" &rsTemp("USUARIO")& ";" &rsTemp("FECHA")& ";" &rsTemp("NUMERO_DE_PR")& ";" &rsTemp("POSI_2")& ";" &rsTemp("TXT_EXPLCATIVO_P_POS")& ";" &rsTemp("TIPO")& ";" &rsTemp("PROTESTO")& ";" &rsTemp("GASTO_DE_PROTESTO")& ";" &rsTemp("TEL_1_A")& ";" &rsTemp("TEL_2_A")& ";" &rsTemp("DIRECCION_A")& ";" &rsTemp("COMUNA")& ";" &rsTemp("CIUDAD")& ";" &rsTemp("REGION")& ";" &rsTemp("EMAIL_A")& ";" &rsTemp("TEL_1_G")& ";" &rsTemp("TEL_2_G")& ";" &rsTemp("DIRECCION_G")& ";" &rsTemp("EMAIL_G")

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

