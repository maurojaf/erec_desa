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
''******************************
%>
<%

 if Request("CB_CLIENTE") <> "" then
	strCliente=Request("CB_CLIENTE")
End if

 if Request("CB_ASIGNACION") <> "" then
	strAsignacion=Request("CB_ASIGNACION")
End if

 if Request("CB_CUSTODIO") <> "" then
	strCobranza=Request("CB_CUSTODIO")
End if

 if Request("CB_USUARIO") <> "" then
	strUsuario=Request("CB_USUARIO")
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
if Request("opAc")= "0" then
	sIopAc=0
else
	sIopAc=1
End if

rut_especifico 			=Trim(Request("rut_especifico"))
opcion_rut 				=Trim(Request("opcion_rut"))
rut_especifico_masivo 	=Trim(Request("rut_especifico_masivo"))

'response.write "<BR>strCobranza=" & strCobranza
'response.write "<BR>strUsuario=" & strUsuario

if trim(rut_especifico)="1" then
	if trim(rut_especifico_masivo)<>"" then

		ValorNuevo_rut 	=split(rut_especifico_masivo,CHR(13))	
		total_rut 		=ubound(ValorNuevo_rut)	
		valor_rut 		=""

		For indice = 0 to total_rut
			VALOR 		= Replace(ValorNuevo_rut(indice), chr(13),"")
			VALOR 		= Replace(ValorNuevo_rut(indice), chr(10),"")
			
			if trim(VALOR)<>"" then

				if trim(opcion_rut)="1" then 'SIN GUION
					valor_rut 	= valor_rut &"*''"&REPLACE(VALOR,"-","")&"''"

				elseif trim(opcion_rut)="2" then ' SIN DV'
					valor_rut 	= valor_rut &"*''"&mid( VALOR, 1, (len(VALOR)-2) )&"''"

				else					
					valor_rut 	= valor_rut &"*''"&TRIM(VALOR)&"''"

				end if	
			end if

		next	
		'response.write valor_rut&"<BR>"

	end if
	
	rut_masivo =mid(valor_rut,2,len(valor_rut))

end if


AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc



If strArchivo <> "" Then

	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "export_DeudaAgrup_" & session("session_idusuario") & "_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivoTerceros


	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero = "[ID_DEUDOR];RUT_DEUDOR;PRIORIDAD;CUSTODIO;FECHA_INGRESO;ACREEDOR;NOM_ACREEDOR;ESTADO;SALDO;COD_USUARIO_ASIG;NOM_USUARIO_ASIG;CAMPAÑA;FEC_SUBIDA_ARCH;ETAPA_COBRANZA;FECHA_ESTADO_ETAPA;UBIC_TELEFONICA;FECHA_PRORROGA;ADIC_1;ADIC_2;ADIC_3;FECHA_VENC_INF;DM_VENC_INF;FECHA_ASIGNACION_INF;DM_ASIG_INF;DOCUMENTOS;FECHA_INGRESO_UG;NOMBRE_UG;AGENDAMIENTO_UG;FECHA_COMP_UG;FECHA_INGRESO_UGS;NOMBRE_UGS;AGENDAMIENTO_UGS;FECHA_COMP_UGS;FECHA_INGRESO_UGT;NOMBRE_UGT;AGENDAMIENTO_UTG;FECHA_COMP_UGT;FONOS;MAIL"

	fichCA.writeline(strTextoTercero)

	strSql = "EXEC [proc_inf_Deuda_Agrupada_Str] '" & strCliente & "', '" & strCobranza &"','" & strUsuario & "', '"&trim(rut_especifico)&"','"&trim(opcion_rut)&"','"&trim(rut_masivo)&"'"

	'response.write "<BR>strSql=" & strSql

	set rsTemp= Conn.execute(strSql)

	strTextoTercero=""
	cantSiniestroC = 0
	Do While Not rsTemp.Eof

		AbrirSCG1()
		strSql = "SELECT SALDO FROM CUOTA WHERE COD_CLIENTE = '" & rsTemp("COD_CLIENTE") & "' AND RUT_DEUDOR = '" & rsTemp("RUT_DEUDOR")  & "' AND (SALDO > 0 OR ESTADO_DEUDA IN (1,7,8)) "
		set rsEstado =  Conn1.execute(strSql)
		If Not rsEstado.Eof Then
			strEstado = "ACTIVO"
		Else
			strEstado = "NO ACTIVO"
		End If
		CerrarSCG1()

		AbrirSCG1()
		strSql = "SELECT ISNULL(USUARIO_ASIG,'SIN ASIGNAR') AS USUARIO_ASIG, ISNULL(LOGIN,'SIN ASIGNAR') AS LOGIN FROM CUOTA C, USUARIO U WHERE C.COD_CLIENTE = '" & rsTemp("COD_CLIENTE") & "' AND C.RUT_DEUDOR = '" & rsTemp("RUT_DEUDOR")  & "' AND C.USUARIO_ASIG = U.ID_USUARIO AND C.ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) ORDER BY ESTADO_DEUDA"
		set rsEstado =  Conn1.execute(strSql)
		If Not rsEstado.Eof Then
			intCodAsig = rsEstado("USUARIO_ASIG")
			strLogin = rsEstado("LOGIN")
		Else
			intCodAsig = ""
			strLogin = "SIN ASIGNACION"
		End If
		CerrarSCG1()

		strTextoTercero = rsTemp("ID_DEUDOR")& ";" & rsTemp("RUT_DEUDOR")& ";" & rsTemp("PRIORIDAD")& ";" & rsTemp("CUSTODIO") & ";" & rsTemp("FECHA_INGRESO") & ";" & rsTemp("COD_CLIENTE") & ";" & rsTemp("NOM_CLIENTE") & ";" & strEstado & ";" & rsTemp("SALDO") & ";"
		strTextoTercero = strTextoTercero & intCodAsig & ";" & strLogin & ";" & rsTemp("ID_CAMPANA") & ";" & rsTemp("FEC_SUBIDA_ARCH") & ";" & rsTemp("NOM_ESTADO_COBRANZA") & ";" & rsTemp("FECHA_ESTADO_ETAPA")  & ";" & rsTemp("UBIC_TELEFONICA") & ";" & rsTemp("FECHA_PRORROGA") & ";" & rsTemp("ADIC_1") & ";" & rsTemp("ADIC_2") & ";" & rsTemp("ADIC_3") & ";" & rsTemp("FECHA_VENC_INF_ACTIVO") & ";" & rsTemp("DM_VENC") & ";" & rsTemp("FECHA_CREACION_INF") & ";" & rsTemp("DM_ASIG")  & ";" & rsTemp("DOCUMENTOS") & ";"
		strTextoTercero = strTextoTercero & rsTemp("FECHA_INGR_UG") & ";" & rsTemp("NOMBRE_INGR_UG") & ";" & rsTemp("FECHA_AGEND_UG") & ";" & rsTemp("FECHA_COMP_UG") & ";" & rsTemp("FECHA_INGR_UGS") & ";" & rsTemp("NOMBRE_INGR_UGS") & ";" & rsTemp("FECHA_AGEND_UGS") & ";" & rsTemp("FECHA_COMP_UGS") & ";" & rsTemp("FECHA_INGR_UGT") & ";" & rsTemp("NOMBRE_INGR_UGT") & ";" & rsTemp("FECHA_AGEND_UGT") & ";" & rsTemp("FECHA_COMP_UGT") & ";" & rsTemp("FONOS")& ";" & rsTemp("MAIL")

		cantSiniestroC = cantSiniestroC + 1

		fichCA.writeline(strTextoTercero)


		rsTemp.movenext

	Loop

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario") & "') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE TMP_EXPORT_DEUDA_AGRUP_" & session("session_idusuario")
	Conn.Execute strSql,64

	%>
	<table>
	<tr><td>Cantidad de registros generados : <%= cantSiniestroC %></td></tr>
	<tr><td><a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTerceros%>')">Descargar</a>

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

