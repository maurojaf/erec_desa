<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
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
<html xmlns="http:www.w3.org/1999/xhtml">
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

 if Request("CB_CUSTODIO") <> "" then
	strCobranza=Request("CB_CUSTODIO")
End if

 if Request("CB_USUARIO") <> "" then
	strUsuario=Request("CB_USUARIO")
End if

dtmFecIniEstado=Request("dtmInicio")

dtmFecTerEstado=Request("dtmTermino")

	'Response.Write "dtmFecIniEstado=" & dtmFecIniEstado
	'Response.Write "<br>dtmFecTerEstado=" & dtmFecTerEstado

strArchivo=1

AbrirScg()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

'Response.write "<br>strCobranza=" & strCobranza

If strArchivo <> "" Then

	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())
	'Response.write "<br>Fecha=" & Fecha
	strNomArchivoTerceros = "export_Deuda_" & strCliente & "_" & Fecha & ".csv"
	'terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros

	terceroCSV = session("ses_PathFisicoSistema")  & "\Logs\" & strNomArchivoTerceros

	'terceroCSV = request.serverVariables("PATH_INFO") & "Logs\" & strNomArchivoTerceros
	'Response.write "<br>dtmFecFinCiclo=" & dtmFecFinCiclo
	'Response.write "<br>strCobranza=" & strCobranza

	'Response.write "<br>terceroCSV=" & terceroCSV

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strSql="SELECT IsNull(ADIC_1,'ADIC_1') as ADIC_1, IsNull(ADIC_2,'ADIC_2') as ADIC_2, IsNull(ADIC_3,'ADIC_3') as ADIC_3, IsNull(ADIC_4,'ADIC_4') as ADIC_4, IsNull(ADIC_5,'ADIC_5') as ADIC_5, IsNull(ADIC_91,'ADIC_91') as ADIC_91, IsNull(ADIC_92,'ADIC_92') as ADIC_92, IsNull(ADIC_93,'ADIC_93') as ADIC_93, IsNull(ADIC_94,'ADIC_94') as ADIC_94, IsNull(ADIC_95,'ADIC_95') as ADIC_95 , IsNull(ADIC_96,'ADIC_96') as ADIC_96 , IsNull(ADIC_97,'ADIC_97') as ADIC_97 , IsNull(ADIC_98,'ADIC_98') as ADIC_98, IsNull(ADIC_99,'ADIC_99') as ADIC_99 , IsNull(ADIC_100,'ADIC_100') as ADIC_100 FROM CLIENTE WHERE COD_CLIENTE = '" & strCliente & "'"

	set rsDET=Conn.execute(strSql)
	if Not rsDET.eof Then
		strNombreAdic1 = UCASE(rsDET("ADIC_1"))
		strNombreAdic2 = UCASE(rsDET("ADIC_2"))
		strNombreAdic3 = UCASE(rsDET("ADIC_3"))
		strNombreAdic4 = UCASE(rsDET("ADIC_4"))
		strNombreAdic5 = UCASE(rsDET("ADIC_5"))

		strNombreAdic91 = UCASE(rsDET("ADIC_91"))
		strNombreAdic92 = UCASE(rsDET("ADIC_92"))
		strNombreAdic93 = UCASE(rsDET("ADIC_93"))
		strNombreAdic94 = UCASE(rsDET("ADIC_94"))
		strNombreAdic95 = UCASE(rsDET("ADIC_95"))

		strNombreAdic96 = UCASE(rsDET("ADIC_96"))
		strNombreAdic97 = UCASE(rsDET("ADIC_97"))
		strNombreAdic98 = UCASE(rsDET("ADIC_98"))
		strNombreAdic99 = UCASE(rsDET("ADIC_99"))
		strNombreAdic100 = UCASE(rsDET("ADIC_100"))
	End If

	If trim(strNombreAdic1) = "" Then strNombreAdic1 = "ADIC_1"
	If trim(strNombreAdic2) = "" Then strNombreAdic2 = "ADIC_2"
	If trim(strNombreAdic3) = "" Then strNombreAdic3 = "ADIC_3"
	If trim(strNombreAdic4) = "" Then strNombreAdic4 = "ADIC_4"
	If trim(strNombreAdic5) = "" Then strNombreAdic5 = "ADIC_5"

	If trim(strNombreAdic91) = "" Then strNombreAdic91 = "ADIC_91"
	If trim(strNombreAdic92) = "" Then strNombreAdic92 = "ADIC_92"
	If trim(strNombreAdic93) = "" Then strNombreAdic93 = "ADIC_93"
	If trim(strNombreAdic94) = "" Then strNombreAdic94 = "ADIC_94"
	If trim(strNombreAdic95) = "" Then strNombreAdic95 = "ADIC_95"

	If trim(strNombreAdic96) = "" Then strNombreAdic96 = "ADIC_96"
	If trim(strNombreAdic97) = "" Then strNombreAdic97 = "ADIC_97"
	If trim(strNombreAdic98) = "" Then strNombreAdic98 = "ADIC_98"
	If trim(strNombreAdic99) = "" Then strNombreAdic99 = "ADIC_99"
	If trim(strNombreAdic100) = "" Then strNombreAdic100 = "ADIC_100"

	strTextoTercero=""
	strTextoTercero = "[ID_DEUDA];COD_ASIGNACION;ACREEDOR;NOM_ACREEDOR;TIPO_DOCUMENTO;FECHA_CARGA;FECHA_LLEGADA;FECHA_CREACION;NRO_DOC;NRO_CUOTA;FECHA_VENC;RUT_DEUDOR;NOMBRE_DEUDOR;REPLEG_NOMBRE;SUCURSAL;" & strNombreAdic1 & ";" & strNombreAdic2 & ";" & strNombreAdic3 & ";" & strNombreAdic4 & ";" & strNombreAdic5 & ";" & strNombreAdic91 & ";" & strNombreAdic92 & ";" & strNombreAdic93 & ";" & strNombreAdic94 & ";" & strNombreAdic95 & ";" & strNombreAdic96 & ";" & strNombreAdic97 & ";" & strNombreAdic98 & ";" & strNombreAdic99 & ";" & strNombreAdic100 & ";OBSERVACION;DEUDACAPITAL;SALDO;GASTOS_PROTESTOS;ESTADO_DEUDA;DESCRIPCION;FECHA_ESTADO;USUARIO_ASIG;LOGIN;CAMPAÑA;CUSTODIO" '& chr(13) & chr(10)

	'response.write "strTextoTercero=" & strTextoTercero
	'Response.End

	fichCA.writeline(strTextoTercero)

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'TMP_EXPORT_DEUDA_" & session("session_idusuario") & "') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE TMP_EXPORT_DEUDA_" & session("session_idusuario")
	Conn.Execute strSql,64

	'**********CREO TABLA Y LA LLENO************'

	strSql ="SELECT C.ID_CUOTA,C.COD_REMESA, C.COD_CLIENTE, CL.DESCRIPCION AS NOM_CLIENTE, TIPO_DOCUMENTO,CONVERT(VARCHAR(10),R.FECHA_CARGA,103) AS FECHA_CARGA, "
	strSql = strSql & " CONVERT(VARCHAR(10),R.FECHA_LLEGADA,103) AS FECHA_LLEGADA,CONVERT(VARCHAR(10),C.FECHA_CREACION,103) AS FECHA_CREACION,NRO_DOC, NRO_CUOTA, CONVERT(VARCHAR(10),C.FECHA_VENC,103) AS FECHA_VENC,"
	strSql = strSql & " C.RUT_DEUDOR, NOMBRE_DEUDOR, D.ID_CAMPANA, REPLEG_NOMBRE, C.SUCURSAL, C.ADIC_1, C.ADIC_2, C.ADIC_3, C.ADIC_4, C.ADIC_5, C.ADIC_91, C.ADIC_92, C.ADIC_93, C.ADIC_94, C.ADIC_95, C.ADIC_96, REPLACE(REPLACE(C.ADIC_97,char(13),' '),char(10),' ') AS ADIC_97, REPLACE(REPLACE(C.ADIC_98,char(13),' '),char(10),' ') AS ADIC_98, C.ADIC_99, C.ADIC_100, OBSERVACION, CAST(VALOR_CUOTA AS BIGINT) AS DEUDACAPITAL,"
	strSql = strSql & " CAST(SALDO AS BIGINT) AS SALDO, CAST(GASTOS_PROTESTOS AS INT) AS GASTOS_PROTESTOS,C.ESTADO_DEUDA, CONVERT(VARCHAR(10),C.FECHA_ESTADO,103) AS FECHA_ESTADO,E.DESCRIPCION, TD.NOM_TIPO_DOCUMENTO,"
	strSql = strSql & " C.USUARIO_ASIG, U.LOGIN ,ISNULL(C.CUSTODIO,'LLACRUZ') AS CUSTODIO"


	strSql = strSql & " INTO TMP_EXPORT_DEUDA_" & session("session_idusuario")
	strSql = strSql & " FROM CUOTA C, REMESA R, DEUDOR D, ESTADO_DEUDA E, USUARIO U, CLIENTE CL, TIPO_DOCUMENTO TD"

	strSql = strSql & " WHERE C.COD_REMESA = R.COD_REMESA AND C.COD_CLIENTE = R.COD_CLIENTE AND C.RUT_DEUDOR = D.RUT_DEUDOR AND C.COD_CLIENTE = D.COD_CLIENTE AND CL.COD_CLIENTE = D.COD_CLIENTE"
	strSql = strSql & " AND C.ESTADO_DEUDA = E.CODIGO AND C.USUARIO_ASIG *= U.ID_USUARIO AND CL.COD_CLIENTE = C.COD_CLIENTE AND C.TIPO_DOCUMENTO = TD.COD_TIPO_DOCUMENTO"
	strSql = strSql & " AND CL.ACTIVO = 1"


	If Trim(strCliente) <> "" Then
		strSql = strSql & " AND C.COD_CLIENTE = '" & strCliente & "'"
	End If

	If Trim(strUsuario) <> "" Then
		strSql = strSql & " AND C.USUARIO_ASIG = '" & strUsuario & "'"
	End If

	If Trim(strCobranza) = "INTERNA" Then
		strSql = strSql & " AND C.CUSTODIO IS NOT NULL "
	End If

	If Trim(strCobranza) = "EXTERNA"  Then
		strSql = strSql & " AND C.CUSTODIO IS NULL "
	End If

	If Trim(dtmFecIniEstado) <> "" Then
		strSql = strSql & " AND (C.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) OR CONVERT(VARCHAR(10),C.FECHA_ESTADO,103) >= CAST('" & dtmFecIniEstado & "' AS DATETIME)) "
	End If

	If Trim(dtmFecIniEstado) <> "" AND Trim(dtmFecTerEstado) <> "" Then
		strSql = strSql & " AND CONVERT(VARCHAR(10),C.FECHA_ESTADO,103) <= CAST('" & dtmFecTerEstado & "' AS DATETIME) "
	End If

	If Trim(dtmFecIniEstado) = "" Then
		strSql = strSql & " AND C.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
	End If

	'Response.Write "strSql=" & strSql
	'Response.End

	Conn.Execute strSql,64

	strSql = "SELECT * FROM TMP_EXPORT_DEUDA_" & session("session_idusuario")

	set rsTemp= Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsTemp.Eof
		strTextoTercero = rsTemp("ID_CUOTA") & ";" & rsTemp("COD_REMESA") & ";" & rsTemp("COD_CLIENTE") & ";" & rsTemp("NOM_CLIENTE") & ";" & rsTemp("NOM_TIPO_DOCUMENTO") & ";" & rsTemp("FECHA_CARGA") & ";" & rsTemp("FECHA_LLEGADA")  & ";" & rsTemp("FECHA_CREACION")  & ";" & rsTemp("NRO_DOC") & ";" & rsTemp("NRO_CUOTA") & ";"
		strTextoTercero = strTextoTercero & rsTemp("FECHA_VENC") & ";" & rsTemp("RUT_DEUDOR") & ";" & rsTemp("NOMBRE_DEUDOR") & ";" & rsTemp("REPLEG_NOMBRE")  & ";" & rsTemp("SUCURSAL") & ";" & rsTemp("ADIC_1") & ";"
		strTextoTercero = strTextoTercero & rsTemp("ADIC_2") & ";" & rsTemp("ADIC_3") & ";" & rsTemp("ADIC_4") & ";" & rsTemp("ADIC_5")  & ";" & rsTemp("ADIC_91") & ";" & rsTemp("ADIC_92") & ";" & rsTemp("ADIC_93") & ";" & rsTemp("ADIC_94") & ";" & rsTemp("ADIC_95")  & ";" & rsTemp("ADIC_96")  & ";" & rsTemp("ADIC_97")  & ";" & rsTemp("ADIC_98")  & ";" & rsTemp("ADIC_99")  & ";" & rsTemp("ADIC_100")  & ";" & rsTemp("OBSERVACION") & ";" & rsTemp("DEUDACAPITAL") & ";"
		strTextoTercero = strTextoTercero & rsTemp("SALDO") & ";" & rsTemp("GASTOS_PROTESTOS") & ";" & rsTemp("ESTADO_DEUDA") & ";" & rsTemp("DESCRIPCION")  & ";" & rsTemp("FECHA_ESTADO")  & ";" & rsTemp("USUARIO_ASIG") & ";" & rsTemp("LOGIN") & ";" & rsTemp("ID_CAMPANA") & ";" & rsTemp("CUSTODIO")

		cantSiniestroC = cantSiniestroC + 1

		fichCA.writeline(strTextoTercero)

		rsTemp.movenext

	Loop


	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'TMP_EXPORT_DEUDA_" & session("session_idusuario") & "') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE TMP_EXPORT_DEUDA_" & session("session_idusuario")
	Conn.Execute strSql,64

	%>
	<table>
	<tr><td>Cantidad de registros generados : <%= cantSiniestroC %></td></tr>
	<tr><td><a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTerceros%>')">Descargar</a>

	&nbsp;
	<a href="#" onClick="history.back()">Volver</a>

	</td></tr>

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


%>







				</td>
			  </tr>
			</table>


		</td>

	</tr>

</table>

</body>
</html>

