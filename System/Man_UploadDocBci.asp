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
function Terminar( sintPaginaTerminar ) {
	self.location.href = sintPaginaTerminar
}

function Procesar(intTotalRutCarga,intTotalDoc)

{
	if (confirm("¿ Está Cargar y Actualizar? Este proceso carga los documentos luego actualiza el estado de deuda"))
	{
		if (confirm("¿ Está REALMENTE seguro de cargar los documentos ?"))
		{
			self.location.href = "Man_UploadDocBci.asp?strProcesar=SI&intTotalRutCarga=" + intTotalRutCarga + "&intTotalDoc=" + intTotalDoc
		}
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

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<%

'******************************
'*	INICIO CODIGO PARTICULAR  *
''******************************

strProcesar = Request("strProcesar")

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

if Request("CB_CLIENTE") <> "" then
	strCodCliente=Request("CB_CLIENTE")
End if

AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

If strArchivo <> "" Then


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoDocBCI = "archDocBCI_"&Fecha&".csv"
	docBCICSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoDocBCI

	strTextoTercero = strTextoTercero & "ID_TERCERO;PATENTE;RUT;NOMBRE;MARCA;MODELO;TELEFONO_1;TELEFONO_2;TELEFONO3;DIRECCION;COMUNA;CIUDAD" & chr(13) & chr(10)
	 

	strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"&strCodCliente &"/" & strArchivo
	
	strSqlFile = "TRUNCATE TABLE CARGA_BCI_TXT"
	Conn.Execute strSqlFile,64

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_CARGA_BCI_TXT]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_CARGA_BCI_TXT]"
	Conn.Execute strSql,64

	strSql = " CREATE TABLE TMP_CARGA_BCI_TXT ( TEXTO VARCHAR(1000) NULL)"
	Conn.Execute strSql,64

	'response.write "Conn = " & Conn
	'response.write "strSql " & strSql

	'**********CARGA ARCHIVO************'

	strSqlFile = "BULK INSERT TMP_CARGA_BCI_TXT FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2, CODEPAGE = 'ACP')"
	'Response.write strSqlFile

	Conn.Execute strSqlFile,64


	strSqlFile = "INSERT INTO CARGA_BCI_TXT "
	strSqlFile = strSqlFile & " SELECT cast(cast(SUBSTRING(texto,1,12) as int) as varchar(10))+'-'+ SUBSTRING(texto,13,1) , ltrim(rtrim(substring(texto,14,80))),"
	strSqlFile = strSqlFile & " cast(cast(SUBSTRING(texto,94,12) as int) as varchar(10))+'-'+ SUBSTRING(texto,106,1), ltrim(rtrim(substring(texto,107,80))),"
	strSqlFile = strSqlFile & " SUBSTRING(texto,187,10),SUBSTRING(texto,197,10),SUBSTRING(texto,207,5),SUBSTRING(texto,212,2),SUBSTRING(texto,214,12)"
	strSqlFile = strSqlFile & " ,SUBSTRING(texto,226,10),ltrim(rtrim(SUBSTRING(texto,236,50))),ltrim(rtrim(SUBSTRING(texto,286,30))),ltrim(rtrim(SUBSTRING(texto,316,10))),ltrim(rtrim(SUBSTRING(texto,326,250)))"
	strSqlFile = strSqlFile & " ,SUBSTRING(texto,576,12),SUBSTRING(texto,588,12),SUBSTRING(texto,600,4),ltrim(rtrim(SUBSTRING(texto,604,16))),ltrim(rtrim(SUBSTRING(texto,620,150)))"
	strSqlFile = strSqlFile & " ,SUBSTRING(texto,770,10),SUBSTRING(texto,780,1)"
 	strSqlFile = strSqlFile & " FROM TMP_CARGA_BCI_TXT"

 	''Response.write "strSqlFile=" & strSqlFile

	Conn.Execute strSqlFile,64


	strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD FROM TMP_CARGA_BCI_TXT"
	set rsTemp= Conn.execute(strSql)

	if not rsTemp.eof then
		intTotalBase = rsTemp("CANTIDAD")
	Else
		intTotalBase = 0
	End if

	strObsCarga = now

	strSqlFile = "				SELECT COUNT(DISTINCT CARGA_BCI_TXT.RUT_DEUDOR ) AS TOTAL_RUT"
	strSqlFile = strSqlFile & "	FROM CARGA_BCI_TXT				LEFT JOIN DEUDOR ON  CARGA_BCI_TXT.RUT_DEUDOR = DEUDOR.RUT_DEUDOR"
	strSqlFile = strSqlFile & "	WHERE CARGA_BCI_TXT.RUT_DEUDOR NOT IN ( SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '1100')"

	set rsTotalRut= Conn.execute(strSqlFile)

	intTotalRutCarga = rsTotalRut("TOTAL_RUT")

	strSqlFile = "				SELECT COUNT(*) AS TOTAL_DOC"
	strSqlFile = strSqlFile & "	FROM CARGA_BCI_TXT AS B			LEFT JOIN TIPO_DOCUMENTO AS TD ON B.TIPO_DOCTO = TD.NOM_TIPO_DOCUMENTO"
	strSqlFile = strSqlFile & "							        LEFT JOIN CUOTA AS C ON (RUT_SUBCLIENTE+C.RUT_DEUDOR+NRO_DOC+CAST(NRO_CUOTA AS VARCHAR(2))+ CAST(C.TIPO_DOCUMENTO AS VARCHAR(10))) = (RUT_CLIENTE+B.RUT_DEUDOR+CAST(NRO_DOCTO AS VARCHAR(30))+CAST(CUOTA_DOCUMENTO AS VARCHAR(2))+ CAST (TD.COD_TIPO_DOCUMENTO AS VARCHAR(10)))"
	strSqlFile = strSqlFile & "	WHERE C.ID_CUOTA IS NULL AND B.RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = 1100)"

	set rsTotalDoc= Conn.execute(strSqlFile)

	intTotalDoc = rsTotalDoc("TOTAL_DOC")


	strSqlFile = "				SELECT ESTADO, COUNT(ESTADO) AS TOTAL FROM"
	strSqlFile = strSqlFile & " (SELECT DATEDIFF(DAY,FECHA_VENCTO_REAL,GETDATE()) as DIAS_MORA,ID_CUOTA,SALDO_DEUDOR AS ABONO,CODIGO_COBRANZA,estado_deuda,COBRANZA_ANTICIPADA,"
	strSqlFile = strSqlFile & " CASE"
	strSqlFile = strSqlFile & " WHEN (ESTADO_DEUDA IN (1,7) AND CODIGO_COBRANZA IS NULL)"
	strSqlFile = strSqlFile & " THEN '1-ERROR'"

	strSqlFile = strSqlFile & " WHEN (ESTADO_DEUDA IN (1,7) AND CODIGO_COBRANZA IN ('0400','0410','0440','0460','0510','0530','0540','0570','0590','0690','0700','0710','0725','0790','0800','0311'))"
	strSqlFile = strSqlFile & " THEN '5-RETIRAR DE ASIGNACION'"

	strSqlFile = strSqlFile & " WHEN (ESTADO_DEUDA IN (1,7) AND DATEDIFF(DAY,FECHA_VENCTO_REAL,GETDATE())<=-22 AND CODIGO_COBRANZA IN ('0000','0310','0320','0420','0421','0455','0415','0430')"
	strSqlFile = strSqlFile & " 	  AND COBRANZA_ANTICIPADA = 'N' AND ISNULL(SUBSTRING(CUOTA.OBSERVACION,1,6),'Ñ')<>'VUELTO')"
	strSqlFile = strSqlFile & " THEN '6-RETIRAR NO ASIGNABLE'"

	strSqlFile = strSqlFile & " WHEN (ESTADO_DEUDA IN (3) AND CODIGO_COBRANZA IN ('0000','0310','0320','0420','0421','0455','0415','0430') AND DATEDIFF(DAY,FECHA_VENCTO_REAL,GETDATE())<=45)"
	strSqlFile = strSqlFile & " THEN '7-ACTIVAR DOC CANCELADO'"

	strSqlFile = strSqlFile & " WHEN (ESTADO_DEUDA IN (13) AND CODIGO_COBRANZA IN ('0000','0310','0320','0420','0421','0455','0415','0430') AND (DATEDIFF(DAY,FECHA_VENCTO_REAL,GETDATE())>=-21 OR COBRANZA_ANTICIPADA = 'S') AND DATEDIFF(DAY,FECHA_VENCTO_REAL,GETDATE())<=45)"
	strSqlFile = strSqlFile & " THEN '8-ACTIVAR DOC NO ASIGNABLE'"

	strSqlFile = strSqlFile & " WHEN (	   ESTADO_DEUDA IN (1,7) AND CODIGO_COBRANZA IN ('0330','0110','0120','0130','0140','0550')) OR (ESTADO_DEUDA IN (2) AND CODIGO_COBRANZA IN ('0330','0110','0120','0130','0140','0550') AND CAST(CONVERT(VARCHAR(10),FECHA_ESTADO,103) AS DATETIME) = CAST(CONVERT(VARCHAR(10),GETDATE(),103) AS DATETIME))"
	strSqlFile = strSqlFile & " THEN '3-CANCELAR'"
	strSqlFile = strSqlFile & " WHEN ESTADO_DEUDA IN (1,7) AND (CUOTA.SALDO-CARGA_BCI_TXT.SALDO_DEUDOR)> 0 AND CODIGO_COBRANZA IN ('0000','0310','0320','0420','0421','0455','0415','0430')"
	strSqlFile = strSqlFile & " THEN '4-ABONAR'"

	strSqlFile = strSqlFile & " WHEN ((ESTADO_DEUDA IN (1,7) AND CODIGO_COBRANZA IN ('0000','0310','0320','0420','0421','0455','0415','0430') AND DATEDIFF(DAY,FECHA_VENCTO_REAL,GETDATE())>=-21) OR (ESTADO_DEUDA IN (1,7) AND CODIGO_COBRANZA IN ('0000','0310','0320','0420','0421','0455','0415','0430') AND COBRANZA_ANTICIPADA = 'S') OR (SUBSTRING(CUOTA.OBSERVACION,1,6)='VUELTO'AND DATEDIFF(DAY,FECHA_VENCTO_REAL,GETDATE())<=-22 AND ESTADO_DEUDA IN (1,7,8) AND CODIGO_COBRANZA IN ('0000','0310','0320','0420','0421','0455','0415','0430')))"
	strSqlFile = strSqlFile & " THEN 'ACTIVO'"

	strSqlFile = strSqlFile & " WHEN (ESTADO_DEUDA IN (13) AND DATEDIFF(DAY,FECHA_VENCTO_REAL,GETDATE())<-21 AND COBRANZA_ANTICIPADA = 'N')"
	strSqlFile = strSqlFile & " 	  OR (ESTADO_DEUDA IN (2,3) AND CODIGO_COBRANZA IN ('0400','0410','0440','0460','0510','0530','0540','0570','0590','0690','0700','0710','0725','0790','0800','0311','0425'))"
	strSqlFile = strSqlFile & " 	  OR (ESTADO_DEUDA IN (2,3) AND CODIGO_COBRANZA IN ('0330','0110','0120','0130','0140','0550'))"
	strSqlFile = strSqlFile & " THEN 'OK'"

	strSqlFile = strSqlFile & " ELSE '2-REVISAR'"
	strSqlFile = strSqlFile & " END AS ESTADO"

	strSqlFile = strSqlFile & " FROM CUOTA				LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
	strSqlFile = strSqlFile & " 						LEFT JOIN CARGA_BCI_TXT ON (RUT_SUBCLIENTE+CUOTA.RUT_DEUDOR+NRO_DOC+CAST(NRO_CUOTA AS VARCHAR(2))+TIPO_DOCUMENTO.NOM_TIPO_DOCUMENTO) = (RUT_CLIENTE+CARGA_BCI_TXT.RUT_DEUDOR+CAST(NRO_DOCTO AS VARCHAR(30))+CAST(CUOTA_DOCUMENTO AS VARCHAR(2))+TIPO_DOCTO)"
	strSqlFile = strSqlFile & " 						WHERE (ESTADO_DEUDA IN (1,7) OR (ESTADO_DEUDA IN (2,3,13) AND CODIGO_COBRANZA IS NOT NULL)) AND COD_CLIENTE = 1100) as dd"
	strSqlFile = strSqlFile & " 						WHERE ESTADO <> 'OK' AND ESTADO <> 'ACTIVO'"
	strSqlFile = strSqlFile & " GROUP BY ESTADO"
	strSqlFile = strSqlFile & " ORDER BY ESTADO ASC"


 	'Response.write "strSqlFile=" & strSqlFile

	set rsInf= Conn.execute(strSqlFile)%>

	<table border=1 bordercolor="#000000">

	<tr>
		<td width="400" height = "30">TOTAL REGISTROS BASE</td>
		<td width="40" align="right"><%=intTotalBase%></td>
	</tr>

	<tr>
		<td width="400" height = "30">TOTAL DEUDORES CARGA</td>
		<td width="40" align="right"><%=intTotalRutCarga%></td>
	</tr>

		<td width="400" height = "30">TOTAL DOCUMENTOS CARGA</td>
		<td width="40" align="right"><%=intTotalDoc%></td>
	</tr>

	</tr>

		<td width="400" height = "40">&nbsp;</td>
		<td width="40" align="right">&nbsp;</td>
	</tr>

	<%if not rsInf.eof then
		do while not rsInf.eof%>

	<tr>
		<td width="400"><%=rsInf("ESTADO")%></td>
		<td width="40" align="right"><%=rsInf("TOTAL")%></td>
	</tr>

	<%
			rsInf.movenext
		loop
	End if%>


	<tr><td colspan="2" align="center">
		<input type="BUTTON" value="Cargar y Actualizar" name="terminar" onClick="Procesar(<%=intTotalRutCarga%>,<%=intTotalDoc%>);return false;"><br>
	</tr>

	</table>

<% End if %>


	<%	If strProcesar = "SI" then

		intTotalRutCarga = Request("intTotalRutCarga")
		intTotalDoc = Request("intTotalDoc")

		'CARGA DEUDORES Y DOCUMENTOS'

		strSql = "EXEC [dbo].[proc_Carga_Actualizacion_BCI] "
		''Response.write strSql

		strSql1 = "EXEC Proc_Asigna_Cobrador_Carga '" & strCodCliente & "'," & session("session_idusuario")
		set rsAsignaCarga = Conn.execute(strSql1)	
		
		Conn.execute(strSql)

		%>

		<table border=1 bordercolor="#000000" width="300">

			<tr><td colspan=2><b>Estatus Proceso</b></td></tr>

			<tr>
				<td width="400" height = "30">TOTAL DEUDORES CARGA</td>
				<td width="40" align="right"><%=intTotalRutCarga%></td>
			</tr>

				<td width="400" height = "30">TOTAL DOCUMENTOS CARGA</td>
				<td width="40" align="right"><%=intTotalDoc%></td>
			</tr>

	</tr>

		</table>


	<%End if%>


<%

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

