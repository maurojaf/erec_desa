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

	function AbreArchivo(nombre)
	{
		window.open(nombre,"INFORMACION","width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes");
	}
	function Terminar( sintPaginaTerminar )
	{
		self.location.href = sintPaginaTerminar
	}
	function Procesar()
	{
		if (confirm("¿ Está REALMENTE seguro de cargar y actualizar la información ?"))
		{
			self.location.href = "Man_UploadDoc_UPV.asp?strProcesar=SI"
		}
	}
	function ModificarTipoDoc()
	{
		if (confirm("¿ Está REALMENTE seguro de modificar los tipos de documentos ?"))
		{
			self.location.href = "Man_UploadDoc_UPV.asp?strProcesar=MODTD"
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
%>

<%
abrirscg()
If Trim(intCodUsuario) = "" Then intCodUsuario = session("session_idusuario")
%>

<%
strReprocesar = Request("strReprocesar")
strProcesar = Request("strProcesar")
intTotalRutCarga = Request("intTotalRutCarga")
intTotalDoc = Request("intTotalDoc")
intCambioCustodio_Llacruz = Request("intCambioCustodio_Llacruz")

if Request("CB_CLIENTE") <> "" then
	strCliente=Request("CB_CLIENTE")
End if

if Request("strTipoProceso") <> "" then
	strTipoProceso=Request("strTipoProceso")
End if

if strProcesar = "SI"  then
	strCliente=Request("strCliente")
End if

AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
Dim ConnectDBQ,rsPlanilla,dbc

'Response.write "<br>cli=" & strCliente
'Response.write "<br>TipProc=" & strTipoProceso

%>


<%

If 1 = 1 Then

	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CARGA_UPV]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE TMP_CARGA_UPV"
	'***********************************************'
	Conn.Execute strSql,64
	'***********************************************'
	strSql = " CREATE TABLE TMP_CARGA_UPV ( ID_DEUDOR VARCHAR(53) NULL,"
	strSql = strSql &" 	TIPO_DOC DECIMAL(2,0) NOT NULL,"
	strSql = strSql &"  NUMERO_DOCUMENTO DECIMAL(15,0) NOT NULL,"
	strSql = strSql &" 	CODAPOD VARCHAR(30) NOT NULL,"
	strSql = strSql &" 	NOMBRE_AVAL VARCHAR(104) NOT NULL,"
	strSql = strSql &"  CODCLI VARCHAR(30) NOT NULL,"
	strSql = strSql &"  NOMBRE_ALUMNO VARCHAR(107) NULL,"
	strSql = strSql &"  MONTO_DEUDO DECIMAL(10,0) NOT NULL,"
	strSql = strSql &"  FECHA_VENCIMIENTO DATETIME  NULL,"
	strSql = strSql &" 	GASTO_PROTESTO INT NOT NULL,"
	strSql = strSql &" 	SEDE VARCHAR(60) NULL,"
	strSql = strSql &" 	CAMPUS VARCHAR(60) NULL,"
	strSql = strSql &"  CARRERA VARCHAR(300) NULL,"
	strSql = strSql &"  DIRECCION VARCHAR(300) NOT NULL,"
	strSql = strSql &"  COMUNA VARCHAR(30) NOT NULL,"
	strSql = strSql &"  FONO1 VARCHAR(32) NULL,"
	strSql = strSql &" 	FONO2 VARCHAR(32) NULL,"
	strSql = strSql &" 	FONO3 VARCHAR(32) NULL,"
	strSql = strSql &" 	EMAIL1 VARCHAR(100) NULL,"
	strSql = strSql &"  EMAIL2 VARCHAR(100) NULL,"
	strSql = strSql &"  ESTADO DECIMAL(2,0) NOT NULL,"
	strSql = strSql &"  NOMBRE_ESTADO VARCHAR(23) NOT NULL,"
	strSql = strSql &"  UBICACION DECIMAL(2,0) NULL,"
	strSql = strSql &" 	ANO DECIMAL(4,0) NOT NULL,"
	strSql = strSql &" 	FECHA_CANCELACION DATETIME NULL,"
	strSql = strSql &"  CLAVE_PRINCIPAL VARCHAR(50) NULL,"
	strSql = strSql &"  CLAVE_SECUNDARIA VARCHAR(50) NULL)"
	'***********************************************'
	Conn.Execute strSql,64
	'***********************************************'
	strSql = " INSERT INTO TMP_CARGA_UPV ( ID_DEUDOR, TIPO_DOC,"
	strSql = strSql &"  NUMERO_DOCUMENTO, CODAPOD,"
	strSql = strSql &" 	NOMBRE_AVAL, CODCLI,"
	strSql = strSql &"  NOMBRE_ALUMNO, MONTO_DEUDO,"
	strSql = strSql &"  FECHA_VENCIMIENTO, GASTO_PROTESTO,"
	strSql = strSql &" 	SEDE, CAMPUS, CARRERA,"
	strSql = strSql &"  DIRECCION, COMUNA, FONO1,"
	strSql = strSql &" 	FONO2, FONO3, EMAIL1,"
	strSql = strSql &"  EMAIL2, ESTADO, NOMBRE_ESTADO,"
	strSql = strSql &"  UBICACION, ANO, FECHA_CANCELACION,"
	strSql = strSql &"  CLAVE_PRINCIPAL,"
	strSql = strSql &"  CLAVE_SECUNDARIA)"

	strSql = strSql &"  SELECT ID_DEUDOR, TIPO_DOC,"
	strSql = strSql &"  NUMERO_DOCUMENTO, CODAPOD,"
	strSql = strSql &" 	NOMBRE_AVAL, CODCLI,"
	strSql = strSql &"  NOMBRE_ALUMNO, MONTO_DEUDO,"
	strSql = strSql &"  FECHA_VENCIMIENTO, GASTO_PROTESTO,"
	strSql = strSql &" 	SEDE, CAMPUS, CARRERA,"
	strSql = strSql &"  DIRECCION, COMUNA, FONO1,"
	strSql = strSql &" 	FONO2, FONO3, EMAIL1,"
	strSql = strSql &"  EMAIL2, ESTADO, NOMBRE_ESTADO,"
	strSql = strSql &"  UBICACION, ANO, FECHA_CANCELACION,"
	strSql = strSql &"	(CONVERT(VARCHAR(20),TIPO_DOC) + '-' +  CODCLI + '-' +  CONVERT(VARCHAR(20),NUMERO_DOCUMENTO) +'-'+ CONVERT(VARCHAR(20),MONTO_DEUDO)) AS CLAVE_PRINCIPAL,"
	strSql = strSql &"	(CODCLI +'-'+ CONVERT(VARCHAR(20),NUMERO_DOCUMENTO) +'-'+ CONVERT(VARCHAR(20),MONTO_DEUDO)) AS CLAVE_SECUNDARIA"
	strSql = strSql &" 	FROM [200.73.64.148].[matricula].[matricula].[mt_llacruz] "



	'***********************************************'
	Conn.Execute strSql,64

	''Response.write "<<<<<<<<<<<<<<<<strSql=" & strSql


	'***********************************************'
	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CUOTA_LLACRUZ]') AND type in (N'U'))"
	strSql = strSql &" DROP TABLE TMP_CUOTA_LLACRUZ"
	'***********************************************'
	Conn.Execute strSql,64

	strSql = " CREATE TABLE TMP_CUOTA_LLACRUZ"
	strSql = strSql &" (LLAVE1 VARCHAR(100) NULL,"
	strSql = strSql &" SUB_LLAVE VARCHAR(100) NULL,"
	strSql = strSql &" RUT_DEUDOR_SD VARCHAR(25) NOT NULL,"
	strSql = strSql &" COD_TIPO_DOC VARCHAR(60) NULL,"
	strSql = strSql &" ID_CUOTA INT NOT NULL,"
	strSql = strSql &" COD_REMESA INT NULL,"
	strSql = strSql &" ACREEDOR VARCHAR(20) NOT NULL,"
	strSql = strSql &" NOM_ACREEDOR VARCHAR(60) NULL,"
	strSql = strSql &" NOM_TIPO_DOCUMENTO VARCHAR(50) NULL,"
	strSql = strSql &" FECHA_CARGA DATETIME NULL,"
	strSql = strSql &" FECHA_LLEGADA DATETIME NULL,"
	strSql = strSql &" FECHA_CREACION DATETIME NULL,"
	strSql = strSql &" NRO_DOC VARCHAR(30) NOT NULL,"
	strSql = strSql &" NRO_CUOTA INT NOT NULL,"
	strSql = strSql &" FECHA_VENCIMIENTO DATETIME NULL,"
	strSql = strSql &" RUT_DEUDOR VARCHAR(15)NOT NULL,"
	strSql = strSql &" NOMBRE_DEUDOR VARCHAR(120) NULL,"
	strSql = strSql &" REPLEG_NOMBRE VARCHAR(120) NULL,"
	strSql = strSql &" SUCURSAL VARCHAR(50) NULL,"
	strSql = strSql &" ADIC1 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC2 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC3 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC4 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC5 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC91 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC92 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC93 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC94 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC95 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC96 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC97 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC98 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC99 VARCHAR(60) NULL,"
	strSql = strSql &" ADIC100 VARCHAR(60) NULL,"
	strSql = strSql &" OBSERVACION VARCHAR(70) NULL,"
	strSql = strSql &" VALOR_CUOTA MONEY NULL,"
	strSql = strSql &" SALDO MONEY NULL,"
	strSql = strSql &" GASTOSPROTESTOS MONEY NULL,"
	strSql = strSql &" ESTADO_DEUDA VARCHAR(10) NULL,"
	strSql = strSql &" DESCRIPCION VARCHAR(50) NULL,"
	strSql = strSql &" FECHA_ESTADO DATETIME NULL,"
	strSql = strSql &" USUARIO_ASIG INT NULL,"
	strSql = strSql &" USER_LOGIN VARCHAR(12) NULL,"
	strSql = strSql &" CAMPAÑA VARCHAR(20) NULL,"
	strSql = strSql &" CUSTODIO VARCHAR(60) NULL,"
	strSql = strSql &" COMP1 VARCHAR(10) NULL,"
	strSql = strSql &" COMP2 VARCHAR(10) NULL)"
	'***********************************************'
	Conn.Execute strSql,64
	'***********************************************'
	strSql = " INSERT INTO TMP_CUOTA_LLACRUZ"
	strSql = strSql &"	(LLAVE1, SUB_LLAVE,"
	strSql = strSql &"	RUT_DEUDOR_SD, COD_TIPO_DOC,"
	strSql = strSql &"	ID_CUOTA, COD_REMESA,"
	strSql = strSql &"	ACREEDOR, NOM_ACREEDOR,"
	strSql = strSql &"	NOM_TIPO_DOCUMENTO, FECHA_CARGA,"
	strSql = strSql &"	FECHA_LLEGADA, FECHA_CREACION,"
	strSql = strSql &"	NRO_DOC, NRO_CUOTA,"
	strSql = strSql &"	FECHA_VENCIMIENTO, RUT_DEUDOR,"
	strSql = strSql &"	NOMBRE_DEUDOR, REPLEG_NOMBRE,"
	strSql = strSql &"	SUCURSAL, ADIC1, ADIC2,"
	strSql = strSql &"	ADIC3, ADIC4, ADIC5,"
	strSql = strSql &"	ADIC91, ADIC92, ADIC93,"
	strSql = strSql &"	ADIC94, ADIC95, ADIC96,"
	strSql = strSql &"	ADIC97, ADIC98, ADIC99,"
	strSql = strSql &"	ADIC100, OBSERVACION,"
	strSql = strSql &"	VALOR_CUOTA, SALDO,"
	strSql = strSql &"	GASTOSPROTESTOS, ESTADO_DEUDA,"
	strSql = strSql &"	DESCRIPCION, FECHA_ESTADO,"
	strSql = strSql &"	USUARIO_ASIG, USER_LOGIN,"
	strSql = strSql &"	CAMPAÑA, CUSTODIO,"
	strSql = strSql &"	COMP1, COMP2)"

	strSql = strSql &" SELECT"
	strSql = strSql &"	CAST(COD_TIPO_DOC AS VARCHAR(10))+'-'+CAST(RUT_DEUDOR_SD AS VARCHAR(10))+'-'+CAST(NRO_DOC AS VARCHAR(10))+'-'+CAST(CAST(VALOR_CUOTA AS INT) AS VARCHAR(10))AS LLAVE1,"
	strSql = strSql &"	CAST(RUT_DEUDOR_SD AS VARCHAR(10))+'-'+CAST(NRO_DOC AS VARCHAR(10))+'-'+CAST(CAST(VALOR_CUOTA AS INT) AS VARCHAR(10))AS SUB_LLAVE,"
	strSql = strSql &"	* FROM"
	strSql = strSql &"	(SELECT"
	strSql = strSql &"		CASE WHEN LEN(CUOTA.RUT_DEUDOR)= 10 OR LEN(CUOTA.RUT_DEUDOR)= 9"
	strSql = strSql &"			 THEN SUBSTRING(CUOTA.RUT_DEUDOR,1,LEN(CUOTA.RUT_DEUDOR)-2)"
	strSql = strSql &"			 ELSE CUOTA.RUT_DEUDOR"
	strSql = strSql &"			 END AS RUT_DEUDOR_SD,"
	strSql = strSql &"		CASE WHEN CUOTA.TIPO_DOCUMENTO = '2'"
	strSql = strSql &"			 THEN '9'"
	strSql = strSql &"			 WHEN CUOTA.TIPO_DOCUMENTO = '4'"
	strSql = strSql &"			 THEN '8'"
	strSql = strSql &"			 WHEN CUOTA.TIPO_DOCUMENTO = '5'"
	strSql = strSql &"			 THEN '5'"
	strSql = strSql &"			 WHEN CUOTA.TIPO_DOCUMENTO = '18'"
	strSql = strSql &"			 THEN '2'"
	strSql = strSql &"			 WHEN CUOTA.TIPO_DOCUMENTO = '19'"
	strSql = strSql &"			 THEN '3'"
	strSql = strSql &"			 ELSE NULL"
	strSql = strSql &"			 END AS COD_TIPO_DOC,"
	strSql = strSql &"	CUOTA.ID_CUOTA,"
	strSql = strSql &"	CUOTA.COD_REMESA,"
	strSql = strSql &"	CUOTA.COD_CLIENTE AS ACREEDOR,"
	strSql = strSql &"	CLIENTE.DESCRIPCION AS NOM_ACREEDOR,"
	strSql = strSql &"	TIPO_DOCUMENTO.NOM_TIPO_DOCUMENTO,"
	strSql = strSql &"	CONVERT(VARCHAR(10),REMESA.FECHA_CARGA,103) AS FECHA_CARGA,"
	strSql = strSql &"	CONVERT(VARCHAR(10),REMESA.FECHA_LLEGADA,103) AS FECHA_LLEGADA,"
	strSql = strSql &"	CONVERT(VARCHAR(10),CUOTA.FECHA_CREACION,103) AS FECHA_CREACION,"
	strSql = strSql &"	CUOTA.NRO_DOC,"
	strSql = strSql &"	CUOTA.NRO_CUOTA,"
	strSql = strSql &"	CONVERT(VARCHAR(10),CUOTA.FECHA_VENC,103) AS FECHA_VENCIMIENTO,"
	strSql = strSql &"	CUOTA.RUT_DEUDOR,"
	strSql = strSql &"	DEUDOR.NOMBRE_DEUDOR,"
	strSql = strSql &"	DEUDOR.REPLEG_NOMBRE,"
	strSql = strSql &"	ISNULL(CUOTA.SUCURSAL,'NO INGRESADA') AS SUCURSAL,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_1,'') AS ADIC1,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_2,'')AS ADIC2,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_3,'')AS ADIC3,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_4,'')AS ADIC4,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_5,'')AS ADIC5,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_91,'')AS ADIC91,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_92,'')AS ADIC92,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_93,'')AS ADIC93,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_94,'')AS ADIC94,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_95,'')AS ADIC95,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_96,'')AS ADIC96,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_97,'')AS ADIC97,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_98,'')AS ADIC98,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_99,'')AS ADIC99,"
	strSql = strSql &"	ISNULL(CUOTA.ADIC_100,'')AS ADIC100,"
	strSql = strSql &"	ISNULL(CUOTA.OBSERVACION,'') AS OBSERVACION,"
	strSql = strSql &"	CUOTA.VALOR_CUOTA,"
	strSql = strSql &"	CUOTA.SALDO,"
	strSql = strSql &"	ISNULL(CUOTA.GASTOS_PROTESTOS,'') AS GASTOSPROTESTOS,"
	strSql = strSql &"	CUOTA.ESTADO_DEUDA,"
	strSql = strSql &"	ESTADO_DEUDA.DESCRIPCION,"
	strSql = strSql &"	CONVERT(VARCHAR(10),CUOTA.FECHA_ESTADO,103) AS FECHA_ESTADO,"
	strSql = strSql &"	CUOTA.USUARIO_ASIG,"
	strSql = strSql &"	ISNULL(USUARIO.LOGIN,'SINASIG') AS LOGIN,"
	strSql = strSql &"	'' AS CAMPAÑA,"
	strSql = strSql &"	ISNULL(CUOTA.CUSTODIO,'') AS CUSTODIO,"
	strSql = strSql &"	'' AS COMP1, '' AS COMP2"
	strSql = strSql &" FROM CUOTA	INNER JOIN CLIENTE			ON CUOTA.COD_CLIENTE = CLIENTE.COD_CLIENTE"
	strSql = strSql &"			INNER JOIN TIPO_DOCUMENTO	ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
	strSql = strSql &"			INNER JOIN REMESA			ON CUOTA.COD_REMESA = REMESA.COD_REMESA AND REMESA.COD_CLIENTE = CUOTA.COD_CLIENTE"
	strSql = strSql &"			INNER JOIN DEUDOR			ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE"
	strSql = strSql &"			INNER JOIN ESTADO_DEUDA		ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql &"			LEFT JOIN  USUARIO			ON CUOTA.USUARIO_ASIG = USUARIO.ID_USUARIO"
	strSql = strSql &" WHERE CUOTA.COD_CLIENTE IN (1200))AS UPV"
	Conn.Execute strSql,64

	'Response.write "<br><br><br>strSql=" & strSql

	''Response.End

	'***********************************************'
	strSql = " EXEC proc_Carga_Actualizacion_UPV 1"
	'***********************************************'
	Conn.Execute strSql,64

	set rsTemp = Conn.execute(strSql)
	if not rsTemp.eof then
		intActUpvCacelado = rsTemp("ActUpvCacelado")
		intActUpvMoroso = rsTemp("ActUpvMoroso")
		intActUpvCancelado = rsTemp("ActUpvCancelado")
		intActUpvProtestado = rsTemp("ActUpvProtestado")
		intActUpvProrrogado = rsTemp("ActUpvProrrogado")
		intActUpvPendBanco = rsTemp("ActUpvPendBanco")
		intPagUpvProrrogado = rsTemp("PagUpvProrrogado")
		intPagUpvCacelado = rsTemp("PagUpvCacelado")
		intNoAsigUpvPendBanco = rsTemp("NoAsigUpvPendBanco")
		intNoAsigUpvProrrogado = rsTemp("NoAsigUpvProrrogado")
		intNoRegUpvCacelado = rsTemp("NoRegUpvCacelado")
		intNoRegUpvCaceladoPorResol = rsTemp("NoRegUpvCaceladoPorResol")
		intNoRegUpvMoroso = rsTemp("NoRegUpvMoroso")
		intNoRegUpvProtestado = rsTemp("NoRegUpvProtestado")
		intPagCliUpvCacelado = rsTemp("PagCliUpvCacelado")
		intPendUpvProtestado = rsTemp("PendUpvProtestado")
		intNoRegUpvCancelado = rsTemp("NoRegUpvCancelado")
		intPagCliUpvMoroso = rsTemp("PagCliUpvMoroso")
		intNoAsigUpvProtestado = rsTemp("NoAsigUpvProtestado")
		intNoAsigUpvMoroso = rsTemp("NoAsigUpvMoroso")
		strMostrarCuadroSuperior = "SI"
		strMostrarCuadroInferior = "NO"
	Else
		intActUpvCacelado = 0
		intActUpvMoroso = 0
		intActUpvCancelado = 0
		intActUpvProtestado = 0
		intActUpvProrrogado = 0
		intActUpvPendBanco = 0
		intPagUpvProrrogado = 0
		intPagUpvCacelado = 0
		intNoAsigUpvPendBanco = 0
		intNoAsigUpvProrrogado = 0
		intNoRegUpvCacelado = 0
		intNoRegUpvCaceladoPorResol = 0
		intNoRegUpvMoroso = 0
		intNoRegUpvProtestado = 0
		intPagCliUpvCacelado = 0
		intPendUpvProtestado = 0
		intNoRegUpvCancelado = 0
		intPagCliUpvMoroso = 0
		intNoAsigUpvProtestado = 0
		intNoAsigUpvMoroso = 0
		strMostrarCuadroSuperior = "SI"
		strMostrarCuadroInferior = "NO"
	End if

	'***********************************************'
		strSql = "SELECT ISNULL(COUNT(*),0) AS ContActTipoDoc, "
		strSql = strSql &" TIPO_DOCUMENTO = "
		strSql = strSql &" (CASE WHEN CUOTA.TIPO_DOCUMENTO = '2' "
		strSql = strSql &" THEN '18' "
		strSql = strSql &" WHEN CUOTA.TIPO_DOCUMENTO = '3' "
		strSql = strSql &" THEN '19' "
		strSql = strSql &" WHEN CUOTA.TIPO_DOCUMENTO = '8' "
		strSql = strSql &" THEN '4' "
		strSql = strSql &" WHEN CUOTA.TIPO_DOCUMENTO = '9' "
		strSql = strSql &" THEN '2' "
		strSql = strSql &" END) "
		strSql = strSql &" FROM CUOTA INNER JOIN TMP_CUOTA_LLACRUZ tempCuota ON CUOTA.ID_CUOTA = tempCuota.ID_CUOTA "
		strSql = strSql &"	INNER JOIN TMP_CARGA_UPV tempUPV ON tempCuota.LLAVE1 = tempUPV.CLAVE_PRINCIPAL "
		strSql = strSql &"		WHERE	tempCuota.COMP1 = '' AND tempCuota.COMP2 = 1 "
		strSql = strSql &"		AND tempUPV.nombre_estado IN ('PROTESTOS','MOROSO','PENDIENTE BANCO','PRORROGADO') "
		strSql = strSql &"		AND CUOTA.TIPO_DOCUMENTO = tempUPV.TIPO_DOC "
		strSql = strSql &"		GROUP BY CUOTA.TIPO_DOCUMENTO "

		Set rsTemp2= Conn.execute(strSql)
			If Not rsTemp2.eof Then
				intContActTipoDoc = rsTemp2("ContActTipoDoc")
			Else
				intContActTipoDoc = 0
			End If

	'***********************************************'

	intTotalGeneral = intActUpvCacelado + intActUpvMoroso + intActUpvCancelado + intActUpvProtestado + intActUpvProrrogado + intActUpvPendBanco + intPagUpvProrrogado + intPagUpvCacelado + intNoAsigUpvPendBanco + intNoAsigUpvProrrogado + intNoRegUpvCacelado + intNoRegUpvCaceladoPorResol + intNoRegUpvCancelado + intNoRegUpvMoroso + intNoRegUpvProtestado + intPagCliUpvCacelado + intPendUpvProtestado + intPagCliUpvMoroso + intNoAsigUpvProtestado + intNoAsigUpvMoroso

	If strProcesar = "MODTD" Then

		strSql = " UPDATE CUOTA SET TIPO_DOCUMENTO = "
		strSql = strSql & " CASE WHEN CUOTA.TIPO_DOCUMENTO = '2' "
		strSql = strSql & " THEN '18' "
		strSql = strSql & " WHEN CUOTA.TIPO_DOCUMENTO = '3' "
		strSql = strSql & " THEN '19' "
		strSql = strSql & " WHEN CUOTA.TIPO_DOCUMENTO = '8' "
		strSql = strSql & " THEN '4' "
		strSql = strSql & " WHEN CUOTA.TIPO_DOCUMENTO = '9' "
		strSql = strSql & " THEN '2' "
		strSql = strSql & " END "
		strSql = strSql & " FROM CUOTA INNER JOIN TMP_CUOTA_LLACRUZ tempCuota		ON CUOTA.ID_CUOTA = tempCuota.ID_CUOTA "
		strSql = strSql & "		INNER JOIN TMP_CARGA_UPV tempUPV ON tempCuota.LLAVE1 = tempUPV.CLAVE_PRINCIPAL "
		strSql = strSql & "		WHERE	tempCuota.COMP1 = '' AND tempCuota.COMP2 = 1 "
		strSql = strSql & "			AND tempUPV.nombre_estado IN ('PROTESTOS','MOROSO','PENDIENTE BANCO','PRORROGADO') "
		strSql = strSql & "			AND CUOTA.TIPO_DOCUMENTO = tempUPV.TIPO_DOC "

		Conn.Execute strSql,64
	End If

	If strProcesar = "SI" Then
		'***********************************************'
		' PROCESAR INFORMACION
		'***********************************************'
		'CRM = ACTIVA - UPV = CACELADO'
		strSql = " SELECT COUNT (*) AS Count_1"
		strSql = strSql &" UPDATE CUOTA SET SALDO = 0, FECHA_ESTADO = getdate(), ESTADO_DEUDA = 3, OBSERVACION = 'PAGO EN CLIENTE TOTAL POR " & session("session_idusuario") & "' WHERE ID_CUOTA IN "
		strSql = strSql &" (SELECT ID_CUOTA FROM TMP_CUOTA_LLACRUZ tmpLlacruz INNER JOIN TMP_CARGA_UPV tmpUpv ON tmpLlacruz.LLAVE1 = tmpUpv.CLAVE_PRINCIPAL "
		strSql = strSql &" WHERE tmpLlacruz.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO= 1) AND tmpUpv.NOMBRE_ESTADO = 'CACELADO') "

		Set rsTemp= Conn.execute(strSql)
		If Not rsTemp.eof Then
			intCount_1 = rsTemp("Count_1")
		Else
			intCount_1 = 0
		End If

		'CRM = ACTIVA - UPV = CANCELADO'
		strSql = " SELECT COUNT (*) AS Count_2"
		strSql = strSql & " UPDATE CUOTA SET SALDO = 0, FECHA_ESTADO = getdate(), ESTADO_DEUDA = 3, OBSERVACION = 'PAGO EN CLIENTE TOTAL POR " & session("session_idusuario") & "' WHERE ID_CUOTA IN "
		strSql = strSql &" (SELECT ID_CUOTA FROM TMP_CUOTA_LLACRUZ tmpLlacruz INNER JOIN TMP_CARGA_UPV tmpUpv ON tmpLlacruz.LLAVE1 = tmpUpv.CLAVE_PRINCIPAL "
		strSql = strSql &" WHERE tmpLlacruz.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO= 1) AND tmpUpv.NOMBRE_ESTADO = 'CANCELADO') "

		Set rsTemp= Conn.execute(strSql)
		If Not rsTemp.eof Then
			intCount_2 = rsTemp("Count_2")
		Else
			intCount_2 = 0
		End If

		'CRM = ACTIVA - UPV = PRORROGADO'
		strSql = " SELECT COUNT (*) AS Count_3"
		strSql = strSql & " UPDATE CUOTA SET ESTADO_DEUDA = 13 , FECHA_ESTADO = getdate(), SALDO = 0, OBSERVACION = 'NO ASIGNABLE POR " & session("session_idusuario") & "' WHERE ID_CUOTA IN "
		strSql = strSql &" (SELECT ID_CUOTA FROM TMP_CUOTA_LLACRUZ tmpLlacruz INNER JOIN TMP_CARGA_UPV tmpUpv ON tmpLlacruz.LLAVE1 = tmpUpv.CLAVE_PRINCIPAL "
		strSql = strSql &" WHERE tmpLlacruz.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO= 1) AND tmpUpv.NOMBRE_ESTADO = 'PRORROGADO') "

		Set rsTemp= Conn.execute(strSql)
		If Not rsTemp.eof Then
			intCount_2 = rsTemp("Count_3")
		Else
			intCount_2 = 0
		End If

		'CRM = ACTIVA - UPV = PENDIENTE BANCO'
		strSql = " SELECT COUNT (*) AS Count_4"
		strSql = strSql & " UPDATE CUOTA SET ESTADO_DEUDA = 13 , FECHA_ESTADO = getdate(), SALDO = 0, OBSERVACION = 'NO ASIGNABLE POR " & session("session_idusuario") & "' WHERE ID_CUOTA IN "
		strSql = strSql &" (SELECT ID_CUOTA FROM TMP_CUOTA_LLACRUZ tmpLlacruz INNER JOIN TMP_CARGA_UPV tmpUpv ON tmpLlacruz.LLAVE1 = tmpUpv.CLAVE_PRINCIPAL "
		strSql = strSql &" WHERE tmpLlacruz.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO= 1) AND tmpUpv.NOMBRE_ESTADO = 'PENDIENTE BANCO') "

		Set rsTemp= Conn.execute(strSql)
		If Not rsTemp.eof Then
			intCount_4 = rsTemp("Count_4")
		Else
			intCount_4 = 0
		End If

		'CRM = PAGADA EN CLIENTE - UPV = PRORROGADO'
		strSql = " SELECT COUNT (*) AS Count_5"
		strSql = strSql & " UPDATE CUOTA SET ESTADO_DEUDA = 13 , FECHA_ESTADO = getdate(), SALDO = 0, OBSERVACION = 'NO ASIGNABLE POR " & session("session_idusuario") & "' WHERE ID_CUOTA IN "
		strSql = strSql &" (SELECT ID_CUOTA FROM TMP_CUOTA_LLACRUZ tmpLlacruz INNER JOIN TMP_CARGA_UPV tmpUpv ON tmpLlacruz.LLAVE1 = tmpUpv.CLAVE_PRINCIPAL "
		strSql = strSql &" WHERE tmpLlacruz.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE DESCRIPCION = 'PAGADA EN CLIENTE') AND tmpUpv.NOMBRE_ESTADO = 'PRORROGADO') "

		''Response.write "strSql="&strSql
		Set rsTemp= Conn.execute(strSql)
		If Not rsTemp.eof Then
			intCount_4 = rsTemp("Count_5")
		Else
			intCount_4 = 0
		End If

		'CRM = PAGADA EN CLIENTE - UPV = MOROSO'
		strSql = " SELECT COUNT (*) AS Count_8"
		strSql = strSql & " UPDATE CUOTA"
		strSql = strSql & " SET ESTADO_DEUDA = 1 , FECHA_ESTADO = getdate(), SALDO = VALOR_CUOTA, OBSERVACION = 'VUELTO A ACTIVAR POR " & session("session_idusuario") & "', FECHA_AGEND_ULT_GES = NULL, HORA_AGEND_ULT_GES = NULL, CUOTA.USUARIO_ASIG = DEUDOR.USUARIO_ASIG,CUOTA.FECHA_ASIGNACION = getdate() "
		strSql = strSql & " FROM CUOTA INNER JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE "
		strSql = strSql & " WHERE ID_CUOTA IN (SELECT ID_CUOTA FROM TMP_CUOTA_LLACRUZ tmpLlacruz INNER JOIN TMP_CARGA_UPV tmpUpv ON tmpLlacruz.LLAVE1 = tmpUpv.CLAVE_PRINCIPAL "
		strSql = strSql & " WHERE tmpLlacruz.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE DESCRIPCION = 'PAGADA EN CLIENTE') "
		strSql = strSql & " AND tmpUpv.NOMBRE_ESTADO = 'MOROSO') "
		''Response.write "strSql="&strSql
		Set rsTemp= Conn.execute(strSql)
		If Not rsTemp.eof Then
			intCount_8 = rsTemp("Count_8")
		Else
			intCount_8 = 0
		End If

		'CRM = NO ASIGNABLE - UPV = PROTESTADO'
		strSql = " SELECT COUNT (*) AS Count_9"
		strSql = strSql & " UPDATE CUOTA"
		strSql = strSql & " SET ESTADO_DEUDA = 1 , FECHA_ESTADO = getdate(), SALDO = VALOR_CUOTA, OBSERVACION = 'VUELTO A ACTIVAR POR " & session("session_idusuario") & "', FECHA_AGEND_ULT_GES = NULL, HORA_AGEND_ULT_GES = NULL, CUOTA.USUARIO_ASIG = DEUDOR.USUARIO_ASIG,CUOTA.FECHA_ASIGNACION = getdate() "
		strSql = strSql & " FROM CUOTA INNER JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE"
		strSql = strSql & " WHERE ID_CUOTA IN (SELECT ID_CUOTA FROM TMP_CUOTA_LLACRUZ tmpLlacruz INNER JOIN TMP_CARGA_UPV tmpUpv ON tmpLlacruz.LLAVE1 = tmpUpv.CLAVE_PRINCIPAL "
		strSql = strSql & " WHERE tmpLlacruz.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE DESCRIPCION = 'NO ASIGNABLE') "
		strSql = strSql & " AND tmpUpv.NOMBRE_ESTADO = 'PROTESTADO')"

		Set rsTemp= Conn.execute(strSql)
		If Not rsTemp.eof Then
			intCount_9 = rsTemp("Count_9")
		Else
			intCount_9 = 0
		End If

		'CRM = NO ASIGNABLE - UPV = MOROSO'
		strSql = " SELECT COUNT (*) AS Count_10"
		strSql = strSql & " UPDATE CUOTA"
		strSql = strSql & " SET ESTADO_DEUDA = 1 , FECHA_ESTADO = getdate(), SALDO = VALOR_CUOTA, OBSERVACION = 'VUELTO A ACTIVAR POR " & session("session_idusuario") & "', FECHA_AGEND_ULT_GES = NULL, HORA_AGEND_ULT_GES = NULL, CUOTA.USUARIO_ASIG = DEUDOR.USUARIO_ASIG,CUOTA.FECHA_ASIGNACION = getdate() "
		strSql = strSql & " FROM CUOTA INNER JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE "
		strSql = strSql & " WHERE ID_CUOTA IN (SELECT ID_CUOTA FROM TMP_CUOTA_LLACRUZ tmpLlacruz INNER JOIN TMP_CARGA_UPV tmpUpv ON tmpLlacruz.LLAVE1 = tmpUpv.CLAVE_PRINCIPAL "
		strSql = strSql & " WHERE tmpLlacruz.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE DESCRIPCION = 'NO ASIGNABLE') "
		strSql = strSql & " AND tmpUpv.NOMBRE_ESTADO = 'MOROSO') "

		Set rsTemp= Conn.execute(strSql)
		If Not rsTemp.eof Then
			intCount_10 = rsTemp("Count_10")
		Else
			intCount_10 = 0
		End If

		Response.write "ssss3333333333333333333"

		strMostrarCuadroSuperior = "NO"
		strMostrarCuadroInferior = "SI"
	End If

	If strMostrarCuadroSuperior = "SI" and strMostrarCuadroInferior = "NO" Then
		%>
			<table border=1 bordercolor="#000000" width="500" ALIGN="CENTER">
				<tr>	<td colspan=2><b>COMPARACION DE ESTADOS</b>
						</td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = ACTIVA - UPV = CACELADO</td>
					<td width="40" align="right"><%=intActUpvCacelado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = ACTIVA - UPV = MOROSO</td>
					<td width="40" align="right"><%=intActUpvMoroso%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = ACTIVA - UPV = CANCELADO</td>
					<td width="40" align="right"><%=intActUpvCancelado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = ACTIVA - UPV = PROTESTADO</td>
					<td width="40" align="right"><%=intActUpvProtestado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = ACTIVA - UPV = PRORROGADO</td>
					<td width="40" align="right"><%=intActUpvProrrogado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = ACTIVA - UPV = PENDIENTE BANCO</td>
					<td width="40" align="right"><%=intActUpvPendBanco%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = PAGADA EN CLIENTE - UPV = PRORROGADO</td>
					<td width="40" align="right"><%=intPagUpvProrrogado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO ASIGNABLE - UPV = CACELADO</td>
					<td width="40" align="right"><%=intPagUpvCacelado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO ASIGNABLE - UPV = PENDIENTE BANCO</td>
					<td width="40" align="right"><%=intNoAsigUpvPendBanco%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO ASIGNABLE - UPV = PRORROGADO</td>
					<td width="40" align="right"><%=intNoAsigUpvProrrogado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO REGISTRA EN CRM - UPV = CACELADO</td>
					<td width="40" align="right"><%=intNoRegUpvCacelado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO REGISTRA EN CRM - UPV = CACELADO POR RESOLUCION</td>
					<td width="40" align="right"><%=intNoRegUpvCaceladoPorResol%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO REGISTRA EN CRM - UPV = CANCELADO</td>
					<td width="40" align="right"><%=intNoRegUpvCancelado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO REGISTRA EN CRM - UPV = MOROSO</td>
					<td width="40" align="right"><%=intNoRegUpvMoroso%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO REGISTRA EN CRM - UPV = PROTESTADO</td>
					<td width="40" align="right"><%=intNoRegUpvProtestado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = PAGADA EN CLIENTE - UPV = CACELADO</td>
					<td width="40" align="right"><%=intPagCliUpvCacelado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = PENDIENTE - UPV = PROTESTADO</td>
					<td width="40" align="right"><%=intPendUpvProtestado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = PAGADA EN CLIENTE - UPV = MOROSO</td>
					<td width="40" align="right"><%=intPagCliUpvMoroso%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO ASIGNABLE - UPV = PROTESTADO</td>
					<td width="40" align="right"><%=intNoAsigUpvProtestado%></td>
				</tr>
				<tr>
					<td width="400" height = "30">CRM = NO ASIGNABLE - UPV = MOROSO</td>
					<td width="40" align="right"><%=intNoAsigUpvMoroso%></td>
				</tr>
				<tr>
					<td width="400" height = "30">TOTAL GENERAL</td>
					<td width="40" align="right"><%=intTotalGeneral%></td>
				</tr>
				<tr>
					<td width="400" height = "30"></td>
					<td width="40" align="right"></td>
				</tr>
				<tr>
					<td width="400" height = "30">TOTAL DOC. A MODIFICAR</td>
					<td width="40" align="right"><%=intContActTipoDoc%></td>
				</tr>
				<tr>
					<td colspan="2" ALIGN="CENTER">
						<input TYPE="BUTTON" value="Modificar Tipo Doc." name="ModificarTipoDoc" onClick="ModificarTipoDoc();return false;">
						<input TYPE="BUTTON" value="Procesar" name="Procesar" onClick="Procesar();return false;">
						<input type="BUTTON" value="Volver" name="TerminarCuadroSuperior" onClick="Terminar('man_carga_Cliente.asp');return false;">
					</td>
				</tr>
			</table>
	 <%
	 ElseIf strMostrarCuadroSuperior = "NO" and strMostrarCuadroInferior = "SI" Then
	 %>
		 <table border=1 bordercolor="#000000" width="400">
			<tr>
				<td colspan=2><b>RESULTADO DE PROCESO</b>
				</td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = ACTIVA - UPV = CACELADO</td>
				<td width="40" align="right"><%=intCount_1%></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = ACTIVA - UPV = CANCELADO</td>
				<td width="40" align="right"><%=intCount_2%></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = ACTIVA - UPV = PRORROGADO</td>
				<td width="40" align="right"><%=intCount_3%></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = ACTIVA - UPV = PENDIENTE BANCO</td>
				<td width="40" align="right"><%=intCount_4%></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = PAGADA EN CLIENTE - UPV = PRORROGADO</td>
				<td width="40" align="right"><%=intCount_5%></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = NO REGISTRA EN CRM - UPV = MOROSO</td>
				<td width="40" align="right"></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = NO REGISTRA EN CRM - UPV = PROTESTADO</td>
				<td width="40" align="right"></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = PAGADA EN CLIENTE - UPV = MOROSO</td>
				<td width="40" align="right"><%=intCount_8%></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = NO ASIGNABLE - UPV = PROTESTADO</td>
				<td width="40" align="right"><%=intCount_9%></td>
			</tr>
			<tr>
				<td width="400" height = "30">CRM = NO ASIGNABLE - UPV = MOROSO</td>
				<td width="40" align="right"><%=intCount_10%></td>
			</tr>
			<tr>
				<td colspan="2" ALIGN="CENTER">
					<input type="BUTTON" value="Volver" name="TerminarCuadroInferior" onClick="Terminar('man_carga_Cliente.asp');return false;">
				</td>
			</tr>
		</table>
	<%
	 End If
End if
%>
</td>
</tr>
</table>
</body>
</html>

