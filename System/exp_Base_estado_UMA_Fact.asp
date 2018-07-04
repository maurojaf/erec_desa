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

AbrirScg()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc



If strArchivo <> "" Then

	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "export_Base_Estado_UMA_Fact_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivoTerceros


	''terceroCSV = "F:\" & strNomArchivoTerceros


	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero=""

	strTextoTercero = "USUARIO_CARGA;FECHA_ASIGNACION;ESTADO_DEUDA;FECHA_ESTADO;CUSTODIO;FECHA_CUSTODIO;USUARIO_CUSTODIO;ID_CUOTA;COD_CLIENTE;COD_SAP;CRUCE;NRO_DOC;POS_1;IMPORTE_ML;MON;SOC;DIV;INT_CIAL;CONTRATO;CTA_MAYOR;FECHA_DOC;FECHA_CONTABLE;TEXTO;VENC_NETO;COMPENS;FE_CON_COMPT;CD;REFERENCIA;TEXTO_P_OP_PRINCIPAL;TEXTO_P_OP_PARCIAL;OP_PRAL;OP_PARC;MBC;IMPORTE;MON_2;CLOB;APLAZAM;MB;SEGMENTO;CBP;CTA_CONTR;TCC;FECHA_DE_ASIGNACION;COBRANZA;EMPRESA;RUT;TELEFONO_1;TELEFONO_2;EMAIL;DIRECCION;DETALLE;E_COBRANZA;SITUACION"

	fichCA.writeline(strTextoTercero)

	strSql="SELECT  CUOTA.[COD_CLIENTE],"
	strSql = strSql & "  [CRUCE],"
	strSql = strSql & "  [COD_SAP],"
	strSql = strSql & "  CUF.NRO_DOC AS NRO_DOC,"
	strSql = strSql & "  [POS_1],"
	strSql = strSql & "  [IMPORTE_ML],"
	strSql = strSql & "  [MON],"
	strSql = strSql & "  [SOC],"
	strSql = strSql & "  [DIV],"
	strSql = strSql & "  [INT_CIAL],"
	strSql = strSql & "  [CONTRATO],"
	strSql = strSql & "  [CTA_MAYOR],"
	strSql = strSql & "  [FECHA_DOC],"
	strSql = strSql & "  [FECHA_CONTABLE],"
	strSql = strSql & "  [TEXTO],"
	strSql = strSql & "  [VENC_NETO],"
	strSql = strSql & "  [COMPENS],"
	strSql = strSql & "  [FE_CON_COMPT],"
	strSql = strSql & "  [CD],"
	strSql = strSql & "  [REFERENCIA],"
	strSql = strSql & "  [TEXTO_P_OP_PRINCIPAL],"
	strSql = strSql & "  [TEXTO_P_OP_PARCIAL],"
	strSql = strSql & "  [OP_PRAL],"
	strSql = strSql & "  [OP_PARC],"
	strSql = strSql & "  [MBC],"
	strSql = strSql & "  [IMPORTE],"
	strSql = strSql & "  [MON_2],"
	strSql = strSql & "  [CLOB],"
	strSql = strSql & "  [APLAZAM],"
	strSql = strSql & "  [MB],"
	strSql = strSql & "  [SEGMENTO],"
	strSql = strSql & "  [CBP],"
	strSql = strSql & "  [CTA_CONTR],"
	strSql = strSql & "  [TCC],"
	strSql = strSql & "  [FECHA_DE_ASIGNACION],"
	strSql = strSql & "  [COBRANZA],"
	strSql = strSql & "  [EMPRESA],"
	strSql = strSql & "  [RUT],"
	strSql = strSql & "  [TELEFONO_1],"
	strSql = strSql & "  [TELEFONO_2],"
	strSql = strSql & "  [EMAIL],"
	strSql = strSql & "  [DIRECCION],"
	strSql = strSql & "  [DETALLE],"
	strSql = strSql & "  [E_COBRANZA],"
	strSql = strSql & "  [SITUACION],"
	strSql = strSql & " (CASE WHEN CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = '' THEN 'LLACRUZ' ELSE CUOTA.CUSTODIO END) AS CUSTODIO,"
	strSql = strSql & " [FECHA_CARGA],"
	strSql = strSql & " U2.LOGIN AS USUARIO_CARGA,"
	strSql = strSql & " U1.LOGIN AS USUARIO_CUSTODIO,"
	strSql = strSql & " CONVERT(VARCHAR(19),FECHA_CUSTODIO,103) AS FECHA_CUSTODIO,"
	strSql = strSql & " CONVERT(VARCHAR(19),CUOTA.FECHA_CREACION,103) AS FECHA_ASIGNACION,"
	strSql = strSql & " (CASE WHEN ESTADO_DEUDA.DESCRIPCION = 'PAGADA EN EMPRESA' THEN 'PAGADA EN LLACRUZ' WHEN ESTADO_DEUDA.DESCRIPCION = 'PAGADA EN CLIENTE' THEN 'PAGADA EN UMA' ELSE ESTADO_DEUDA.DESCRIPCION END) AS ESTADO_DEUDA,"
	strSql = strSql & " CONVERT(VARCHAR(19),CUOTA.FECHA_ESTADO,103) AS FECHA_ESTADO,"
	strSql = strSql & " CUOTA.ID_CUOTA AS ID_CUOTA"


	strSql = strSql & " FROM CUOTA INNER JOIN CARGA_UMA_FACT CUF ON CUOTA.NRO_CLIENTE_DOC = CUF.COD_SAP"
	strSql = strSql & "		   	   INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql & "		   	   LEFT JOIN USUARIO AS U1 ON CUF.USUARIO_CUSTODIO = U1.ID_USUARIO"
	strSql = strSql & "		   	   LEFT JOIN USUARIO AS U2 ON CUF.USUARIO_CARGA = U2.ID_USUARIO"


	strSql = strSql & " WHERE CUOTA.COD_CLIENTE = '" & strCliente & "'"

	Conn.Execute strSql,64

	set rsTemp= Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsTemp.Eof

		strTextoTercero = rsTemp("USUARIO_CARGA")& ";" &rsTemp("FECHA_ASIGNACION")& ";" &rsTemp("ESTADO_DEUDA")& ";" &rsTemp("FECHA_ESTADO")& ";" &rsTemp("CUSTODIO")& ";" &rsTemp("FECHA_CUSTODIO")& ";" &rsTemp("USUARIO_CUSTODIO")& ";" &rsTemp("ID_CUOTA")& ";" &rsTemp("COD_CLIENTE")& ";" &rsTemp("COD_SAP")& ";" &rsTemp("CRUCE")& ";" &rsTemp("NRO_DOC")& ";" &rsTemp("POS_1")& ";" &rsTemp("IMPORTE_ML")& ";" &rsTemp("MON")& ";" &rsTemp("SOC")& ";" &rsTemp("DIV")& ";" &rsTemp("INT_CIAL")& ";" &rsTemp("CONTRATO")& ";" &rsTemp("CTA_MAYOR")& ";" &rsTemp("FECHA_DOC")& ";" &rsTemp("FECHA_CONTABLE")& ";" &rsTemp("TEXTO")& ";" &rsTemp("VENC_NETO")& ";" &rsTemp("COMPENS")& ";" &rsTemp("FE_CON_COMPT")& ";" &rsTemp("CD")& ";" &rsTemp("REFERENCIA")& ";" &rsTemp("TEXTO_P_OP_PRINCIPAL")& ";" &rsTemp("TEXTO_P_OP_PARCIAL")& ";" &rsTemp("OP_PRAL")& ";" &rsTemp("OP_PARC")& ";" &rsTemp("MBC")& ";" &rsTemp("IMPORTE")& ";" &rsTemp("MON_2")& ";" &rsTemp("CLOB")& ";" &rsTemp("APLAZAM")& ";" &rsTemp("MB")& ";" &rsTemp("SEGMENTO")& ";" &rsTemp("CBP")& ";" &rsTemp("CTA_CONTR")& ";" &rsTemp("TCC")& ";" &rsTemp("FECHA_DE_ASIGNACION")& ";" &rsTemp("COBRANZA")& ";" &rsTemp("EMPRESA")& ";" &rsTemp("RUT")& ";" &rsTemp("TELEFONO_1")& ";" &rsTemp("TELEFONO_2")& ";" &rsTemp("EMAIL")& ";" &rsTemp("DIRECCION")& ";" &rsTemp("DETALLE")& ";" &rsTemp("E_COBRANZA")& ";" &rsTemp("SITUACION")

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

