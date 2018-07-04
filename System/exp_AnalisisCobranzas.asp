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
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>

	<script language="JavaScript" type="text/JavaScript">
		function AbreArchivo(nombre){
		window.open(nombre,"INFORMACION","width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes");
		}
	</script>

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


if Request("CB_CLIENTE") <> "" then
	strCliente=Request("CB_CLIENTE")
End if

 if Request("CB_ASIGNACION") <> "" then
	strAsignacion=Request("CB_ASIGNACION")
End if


if Request("Fecha") <> "" then
	Fecha=Request("Fecha")
End if

if Request("dtmFecIniCiclo") <> "" then
	dtmFecIniCiclo=Request("dtmFecIniCiclo")
End if

if Request("dtmFecFinCiclo") <> "" then
	dtmFecFinCiclo=Request("dtmFecFinCiclo")
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

Server.ScriptTimeout = 90000
Conn.ConnectionTimeout = 90000


''ACA DEBERIA TRAER LOS REGISTROS

If strArchivo <> "" Then

	intValorUF = session("valor_moneda")

	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "export_AnalisisCob_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivoTerceros

	''Response.write "terceroCSV=" & terceroCSV

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero = "ID_DEUDOR;RUT_DEUDOR;FECHA_INGRESO;ACREEDOR;NOM_ACREEDOR;ESTADO;DEUDACAPITAL;SALDO;USUARIO_ASIG;LOGIN;CAMPAÑA;FEC_SUBIDA_ARCH;FECHA_CREACION;ETAPA_COBRANZA;FECHA_ESTADO_ETAPA;UBIC_TELEFONICA;FECHA_PRORROGA;ADIC_1;ADIC_2;ADIC_3;DOCUMENTOS;VENC_INF_ACT;ASIG_INF_ACT;FECHA_ULT_NORMALIZACION;FECHA_ENVIO_CONSULTA;HON_ACTIVOS;COD_ULT_GEST_TEL;NOM_AGEN_ULT_GEST_TEL;FECHA_INGRESO_ULT_GEST_TEL;FECHA_AGEN_ULT_GEST_TEL;LET_ACTIVOS;CHE_ACTIVOS;CPAG_ACTIVOS;LETGAR_ACTIVOS"

	fichCA.writeline(strTextoTercero)

	AbrirScg()
	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario") & "') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario")
	Conn.Execute strSql,64

	'**********CREO TABLA Y LA LLENO************'

	strSqlSel="SELECT D.COD_CLIENTE, CL.DESCRIPCION AS NOM_CLIENTE, D.ID_DEUDOR, D.RUT_DEUDOR, D.FECHA_INGRESO, NOMBRE_DEUDOR, D.ID_CAMPANA,[dbo].[fun_trae_fecha_venc_inf_activa] (D.COD_CLIENTE,D.RUT_DEUDOR) AS VENC_INF_ACT, [dbo].[fun_trae_fecha_creacion_inf_activa] (D.COD_CLIENTE,D.RUT_DEUDOR) AS ASIG_INF_ACT, [dbo].[fun_trae_fecha_ult_normalizacion] (D.COD_CLIENTE,D.RUT_DEUDOR) AS FECHA_ULT_NORMALIZACION, [dbo].[fun_calc_honorarios] (SUM(SALDO) , " & intValorUF & ", MAX(FECHA_ESTADO)) AS HON_ACTIVOS, [DBO].[FUN_CUENTA_DOC_ACTIVA] (D.COD_CLIENTE, D.RUT_DEUDOR,4) AS LET_ACTIVOS, [DBO].[FUN_CUENTA_DOC_ACTIVA] (D.COD_CLIENTE, D.RUT_DEUDOR,2) AS CHE_ACTIVOS, [DBO].[FUN_CUENTA_DOC_ACTIVA] (D.COD_CLIENTE, D.RUT_DEUDOR,5) AS PAG_ACTIVOS, [DBO].[FUN_CUENTA_DOC_ACTIVA] (D.COD_CLIENTE, D.RUT_DEUDOR,3) AS LETGAR_ACTIVOS,"
	strSqlSel = strSqlSel & " CAST(SUM(VALOR_CUOTA) AS INT) AS DEUDACAPITAL, CAST(SUM(SALDO) AS INT) AS SALDO, [dbo].[fun_ubicabilidad_telefono_email] (D.RUT_DEUDOR) AS UBIC_TELEFONICA, [dbo].[fun_trae_Fecha_Envío_Consulta] (D.COD_CLIENTE,D.RUT_DEUDOR) AS FECHA_ENVIO_CONSULTA, [dbo].[fun_trae_Ultima_Gestion_Telefonica] (D.COD_CLIENTE,D.RUT_DEUDOR,'CODIGO') AS COD_ULT_GEST_TEL, [dbo].[fun_trae_Ultima_Gestion_Telefonica] (D.COD_CLIENTE,D.RUT_DEUDOR,'NOMBRE') AS NOM_ULT_GEST_TEL, [dbo].[fun_trae_Ultima_Gestion_Telefonica] (D.COD_CLIENTE,D.RUT_DEUDOR,'FECHA_AGENDAMIENTO') AS FECHA_AGEN_ULT_GEST_TEL, [dbo].[fun_trae_Ultima_Gestion_Telefonica] (D.COD_CLIENTE,D.RUT_DEUDOR,'FECHA') AS FECHA_INGRESO_UG,"
	strSqlSel = strSqlSel & " [DBO].[FUN_CUENTA_DOC_ACTIVA] (D.COD_CLIENTE,D.RUT_DEUDOR,9999) AS DOCUMENTOS, CONVERT(VARCHAR(10),FEC_SUBIDA_ULT_ARCHIVO,103) AS FEC_SUBIDA_ARCH, CONVERT(VARCHAR(10),MAX(C.FECHA_CREACION),103) AS FECHA_CREACION,NOM_ESTADO_COBRANZA, CONVERT(VARCHAR(10),FECHA_ESTADO_ETAPA,103) AS FECHA_ESTADO_ETAPA, FECHA_PRORROGA, D.ADIC_1, D.ADIC_2, D.ADIC_3"

	strSqlInto = strSqlInto & " INTO TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario")

	strSqlFrom = strSqlFrom & " FROM CUOTA C, DEUDOR D, CLIENTE CL, ESTADO_COBRANZA EC"

	strSqlWhere = strSqlWhere & " WHERE C.RUT_DEUDOR = D.RUT_DEUDOR AND C.COD_CLIENTE = D.COD_CLIENTE AND "
	strSqlWhere = strSqlWhere & " CL.COD_CLIENTE = D.COD_CLIENTE AND CL.COD_CLIENTE = C.COD_CLIENTE AND D.ETAPA_COBRANZA = EC.COD_ESTADO_COBRANZA"
	strSqlWhere = strSqlWhere & " AND CL.ACTIVO = 1"



	If sIopAc=1 Then
		strSql = strSql & " AND G.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM CUOTA WHERE G.COD_CLIENTE = CUOTA.COD_CLIENTE AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1))"
	End If




	If Trim(strCliente) <> "" Then
		strSqlWhere = strSqlWhere & " AND C.COD_CLIENTE = '" & strCliente & "'"
	End If

	If Trim(strAsignacion) <> "" Then
		strSqlWhere = strSqlWhere & " AND C.COD_REMESA = '" & strAsignacion & "'"
	End If

	If Trim(dtmFecIniCiclo) <> "" Then
		strSqlCond1 = strSqlCond1 & " AND C.FECHA_ESTADO >= '" & dtmFecIniCiclo & "'"
	End If

	If Trim(dtmFecFinCiclo) <> "" Then
		strSqlCond1 = strSqlCond1 & " AND C.FECHA_ESTADO <= '" & dtmFecFinCiclo & "'"
	End If

	If Trim(dtmFecFinCiclo) <> "" or Trim(dtmFecFinCiclo) <> "" Then
		strSqlCond1 = strSqlCond1 & "  AND C.ESTADO_DEUDA NOT IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
	End If

	strSqlCond2 = strSqlCond2 & "  AND C.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"

	strSqlGroup = strSqlGroup & " GROUP BY D.COD_CLIENTE, D.FECHA_INGRESO, CL.DESCRIPCION,  D.ID_DEUDOR, D.RUT_DEUDOR, NOMBRE_DEUDOR, D.ID_CAMPANA, FEC_SUBIDA_ULT_ARCHIVO,NOM_ESTADO_COBRANZA, FECHA_PRORROGA, D.ADIC_1, D.ADIC_2, D.ADIC_3, FECHA_ESTADO_ETAPA  "

	'Response.Write "strSql=" & strSql
	'Response.End

	strSql = strSqlSel & strSqlInto & strSqlFrom & strSqlWhere & strSqlCond1 & strSqlGroup

	'Response.Write "strSql=" & strSql
	Conn.Execute strSql,64


	If Trim(dtmFecFinCiclo) <> "" or Trim(dtmFecFinCiclo) <> "" Then
		strSql = "INSERT INTO TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario") & " " & strSqlSel & strSqlFrom & strSqlWhere & strSqlCond2 & strSqlGroup
		'Response.Write "strSql=" & strSql
		Conn.Execute strSql,64
	End If


	strSql = "UPDATE TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario") & " SET COD_ULT_GEST_TEL = '', NOM_ULT_GEST_TEL = '' , FECHA_AGEN_ULT_GEST_TEL = '' WHERE ASIG_INF_ACT > FECHA_INGRESO_UG"
	'Response.Write "<BR>strSql=" & strSql
	Conn.Execute strSql,64


	strSql = "UPDATE TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario") & " SET COD_ULT_GEST_TEL = '', NOM_ULT_GEST_TEL = '' , FECHA_AGEN_ULT_GEST_TEL = '' WHERE ASIG_INF_ACT IS NULL"
	'Response.Write "<BR>strSql=" & strSql
	Conn.Execute strSql,64



	'Response.End
	CerrarScg()

	AbrirSCG()
	strSql = "SELECT * FROM TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario")
	set rsTemp= Conn.execute(strSql)


	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsTemp.Eof

		AbrirSCG1()
		strSql = "SELECT SALDO FROM CUOTA WHERE COD_CLIENTE = '" & rsTemp("COD_CLIENTE") & "' AND RUT_DEUDOR = '" & rsTemp("RUT_DEUDOR")  & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
		set rsEstado =  Conn1.execute(strSql)
		If Not rsEstado.Eof Then
			strEstado = "ACTIVO"
		Else
			strEstado = "NO ACTIVO"
		End If
		CerrarSCG1()

		AbrirSCG1()
		strSql = "SELECT USUARIO_ASIG, LOGIN FROM CUOTA C, USUARIO U WHERE C.COD_CLIENTE = '" & rsTemp("COD_CLIENTE") & "' AND C.RUT_DEUDOR = '" & rsTemp("RUT_DEUDOR")  & "' AND C.USUARIO_ASIG = U.ID_USUARIO AND C.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) ORDER BY ESTADO_DEUDA"
		set rsEstado =  Conn1.execute(strSql)
		If Not rsEstado.Eof Then
			intCodAsig = rsEstado("USUARIO_ASIG")
			strLogin = rsEstado("LOGIN")
		Else
			intCodAsig = ""
			strLogin = "SIN ASIGNACION"
		End If
		CerrarSCG1()


		AbrirSCG1()

			strSql="SELECT COD_TIPODOCUMENTO_HON, MESES_TD_HON FROM CLIENTE WHERE COD_CLIENTE = '" & rsTemp("COD_CLIENTE") & "'"
			'response.write "strSql=" & strSql
			'Response.End
			set rsDET=Conn1.execute(strSql)
			if Not rsDET.eof Then
				intTipoDocHono = ValNulo(rsDET("COD_TIPODOCUMENTO_HON"),"C")
				intMesHon = ValNulo(rsDET("MESES_TD_HON"),"C")
			end if
		CerrarSCG1()

		AbrirSCG2()
			strSql = "SELECT TIPO_DOCUMENTO, RUT_DEUDOR, CAST(VALOR_CUOTA AS INT) AS VALOR_CUOTA, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES FROM CUOTA WHERE COD_CLIENTE = '" & rsTemp("COD_CLIENTE") & "' AND RUT_DEUDOR = '" & rsTemp("RUT_DEUDOR")  & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
			set rsHon =  Conn2.execute(strSql)
			intToTHonorarios=0
			Do While Not rsHon.Eof

				If Trim(intTipoDocHono) = Trim(rsHon("TIPO_DOCUMENTO")) Then
					intCMeses = rsHon("ANT_MESES")
					intCDias = rsHon("ANT_DIAS")
					intCMeses = Fix((intCDias/30))

					If Cint(intCMeses) < Cint(intMesHon) Then
						intHonorarios = GASTOS_COBRANZAS(HonorariosEspeciales1(intSaldo,intCMeses,intMesHon)) * intCMeses
					Else
						intHonorarios = GASTOS_COBRANZAS(HonorariosEspeciales1(intSaldo,intCMeses,intMesHon))
					End If
				Else
					intHonorarios = GASTOS_COBRANZAS(ValNulo(rsHon("VALOR_CUOTA"),"N"))
				End If

				If intHonorarios < 900 Then intHonorarios = 900
				intHonorarios = Round(intHonorarios,0)
				intToTHonorarios = intToTHonorarios + intHonorarios
				rsHon.movenext
			Loop
		CerrarSCG2()


		'intHonorarios = Trim(ValNulo(rsTemp("HON_ACTIVOS"),"N"))
		intHonorarios = intToTHonorarios

		strTextoTercero = rsTemp("ID_DEUDOR") & ";" & rsTemp("RUT_DEUDOR") & ";" & rsTemp("FECHA_INGRESO") & ";" & rsTemp("COD_CLIENTE") & ";" & rsTemp("NOM_CLIENTE") & ";" & strEstado & ";" & rsTemp("DEUDACAPITAL") & ";" & rsTemp("SALDO") & ";"
		strTextoTercero = strTextoTercero & intCodAsig & ";" & strLogin & ";" & rsTemp("ID_CAMPANA") & ";" & rsTemp("FEC_SUBIDA_ARCH") & ";" & rsTemp("FECHA_CREACION")  & ";" & rsTemp("NOM_ESTADO_COBRANZA") & ";" & rsTemp("FECHA_ESTADO_ETAPA") & ";" & rsTemp("UBIC_TELEFONICA") & ";" & rsTemp("FECHA_PRORROGA") & ";" & rsTemp("ADIC_1") & ";" & rsTemp("ADIC_2") & ";" & rsTemp("ADIC_3") & ";" & rsTemp("DOCUMENTOS") & ";"
		strTextoTercero = strTextoTercero & Trim(rsTemp("VENC_INF_ACT")) & ";" & Trim(rsTemp("ASIG_INF_ACT")) & ";" & Trim(rsTemp("FECHA_ULT_NORMALIZACION")) & ";" & Trim(rsTemp("FECHA_ENVIO_CONSULTA")) & ";" & intHonorarios & ";" & ValNulo(Trim(rsTemp("COD_ULT_GEST_TEL")),"C") & ";" & ValNulo(Trim(rsTemp("NOM_ULT_GEST_TEL")),"C") & ";" & ValNulo(Trim(rsTemp("FECHA_INGRESO_UG")),"C") & ";" & ValNulo(Trim(rsTemp("FECHA_AGEN_ULT_GEST_TEL")),"C") & ";"
		strTextoTercero = strTextoTercero & rsTemp("LET_ACTIVOS") & ";" & rsTemp("CHE_ACTIVOS")  & ";" & rsTemp("PAG_ACTIVOS") & ";" & rsTemp("LETGAR_ACTIVOS")

		cantSiniestroC = cantSiniestroC + 1

		fichCA.writeline(strTextoTercero)


		rsTemp.movenext

	Loop

	CerrarScg()

	AbrirScg()
	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario") & "') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario")
	Conn.Execute strSql,64
	CerrarScg()

	%>
	<table>
	<tr><td>Cantidad de registros generados : <%= cantSiniestroC %></td></tr>
	<tr><td><a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTerceros%>')">Descargar</a></td></tr>
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




</body>
</html>

