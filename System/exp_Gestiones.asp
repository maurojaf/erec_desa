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

If Request("CB_CLIENTE") <> "" then
	strCliente=Trim(Request("CB_CLIENTE"))
End If

If Request("CB_USUARIO") <> "" then
	strUsuario=Request("CB_USUARIO")
End If

If Request("dtmInicio") <> "" then
	dtmInicio = Request("dtmInicio")
End If

If Request("dtmTermino") <> "" then
	dtmTermino = Request("dtmTermino")
End If

If Request("CB_CUSTODIO") <> "" then
	strCobranza=Request("CB_CUSTODIO")
End If

CH_EFECTIVA 			=UCASE(Request("CH_EFECTIVA"))

rut_especifico 			=Trim(Request("rut_especifico"))
opcion_rut 				=Trim(Request("opcion_rut"))
rut_especifico_masivo 	=Trim(Request("rut_especifico_masivo"))

'r'esponse.write  rut_especifico&"<br>"&sin_guion&"<br>"&sin_dv&"<br>"&sin_dv&"<br>"&rut_especifico_masivo


'Response.write "<b>CH_DIASGTE=" & Request("CH_DIASGTE")


	'Response.End


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
					valor_rut 	= valor_rut &",'"&REPLACE(VALOR,"-","")&"'"

				elseif trim(opcion_rut)="2" then ' SIN DV'
					valor_rut 	= valor_rut &",'"&mid( VALOR, 1, (len(VALOR)-2) )&"'"

				else					
					valor_rut 	= valor_rut &",'"&TRIM(VALOR)&"'"

				end if	
			end if

		next	
		'response.write valor_rut&"<BR>"
		
		rut_masivo =mid(valor_rut,2,len(valor_rut))
		
	end if
	
	

end if


'Response.WRITE rut_masivo  & "<BR>"

if Request("Fecha") <> "" then
	Fecha=Request("Fecha")
End if

if Request("archivo") <> "" then
	strArchivo=Request("archivo")
End if
strArchivo=1

AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc



If strArchivo <> "" Then

	strFechaPrev=replace(replace(replace(now(),":","")," ","_"),"/","_")
	Fecha= mid(strFechaPrev,1,len(strFechaPrev)-2)

	strNomArchivoTerceros = "export_Gestiones_" & session("session_idusuario") & "_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros
	terceroCSV = session("ses_ruta_sitio")  & "\Logs\" & strNomArchivoTerceros

	''terceroCSV = "F:\" & strNomArchivoTerceros
	
	'Response.write "terceroCSV=" & terceroCSV

	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)

	strTextoTercero=""

	strTextoTercero = "[ID_DEUDOR];[ID_GESTION];RUT_DEUDOR;INTERLOCUTOR;CLIENTE;NOM_CLIENTE;CORRELATIVO;CATEGORIA;SUBCATEGORIA;GESTION;COD_GESTION;NOM_CATEGORIA;NOM_SUBCATEGORIA;NOM_GESTION;GESTION_CONCATENADA;FECHA_INGRESO;HORA_INGRESO;FECHA_COMPROMISO;FECHA_AGENDAMIENTO;HORA_AGENDAMIENTO;OBSERVACIONES;MEDIO_ASOCIADO; MEDIO_AGENDAMIENTO;ID_USUARIO;NOM_USUARIO;CAMPAÑA"

	fichCA.writeline(strTextoTercero)


	strSql="SELECT CL.DESCRIPCION AS NOM_CLIENTE,D.ID_DEUDOR,"
	strSql = strSql & " (SELECT max(nro_cliente_deudor) from cuota where cod_cliente = 1070 and cuota.RUT_DEUDOR=G.rut_deudor GROUP BY RUT_DEUDOR) AS NRO_DEUDOR_CLIENTE, "
	strSql = strSql & " G.ID_GESTION,"

	if trim(opcion_rut)="1" then
		strSql = strSql & " substring(G.RUT_DEUDOR,1,LEN(G.RUT_DEUDOR)-2)+SUBSTRING(G.RUT_DEUDOR,LEN(G.RUT_DEUDOR),len(G.RUT_DEUDOR)) RUT_DEUDOR , "

	elseif trim(opcion_rut)="2" then
		strSql = strSql & " substring(G.RUT_DEUDOR,1,LEN(G.RUT_DEUDOR)-2) RUT_DEUDOR,	"

	else
		strSql = strSql & " G.RUT_DEUDOR, "

	end if	

	strSql = strSql & " G.COD_CLIENTE, G.ID_CAMPANA, G.CORRELATIVO, G.COD_CATEGORIA, G.COD_SUB_CATEGORIA,"
	strSql = strSql & " G.COD_GESTION, G.FECHA_INGRESO,SUBSTRING(G.HORA_INGRESO,1,5) AS HORA_INGRESO,G.FECHA_COMPROMISO, G.FECHA_AGENDAMIENTO, ISNULL(OBSERVACIONES,'') AS OBSERVACIONES,G.ID_USUARIO AS ID_USUARIO,ISNULL(G.HORA_AGENDAMIENTO,'') AS HORA_AGENDAMIENTO,"
	strSql = strSql & " GTG.DESCRIPCION AS NOMCATEGORIA, GTG.DESCRIPCION AS NOMSUBCATEGORIA , GTG.DESCRIPCION AS NOMGESTION,"
	strSql = strSql & " U.LOGIN,"
	strSql = strSql & " 	CASE "	
	strSql = strSql & " 	WHEN TIPO_MEDIO_GESTION =1 THEN ( ISNULL((CONVERT(VARCHAR(5),DT.COD_AREA) + '-' + DT.TELEFONO),'')) "
	strSql = strSql & " 	WHEN TIPO_MEDIO_GESTION =2 THEN ( ISNULL(UPPER(DE.EMAIL),'') ) "
	strSql = strSql & " 	WHEN TIPO_MEDIO_GESTION =3 THEN ( ISNULL(UPPER(DD.CALLE + ' ' + DD.NUMERO + ' ' + DD.RESTO + ', ' + DD.COMUNA + ', ' + CIUDAD),'') ) "
	strSql = strSql & " 	END MEDIO_GESTION_ASOCIADO, "
	strSql = strSql & " 	CASE "	
	strSql = strSql & " 	WHEN TIPO_MEDIO_AGENDAMIENTO =1 THEN ( ISNULL((CONVERT(VARCHAR(5),DT.COD_AREA) + '-' + DT.TELEFONO),'')) "
	strSql = strSql & " 	WHEN TIPO_MEDIO_AGENDAMIENTO =2 THEN ( ISNULL(UPPER(DE.EMAIL),'') ) "
	strSql = strSql & " 	WHEN TIPO_MEDIO_AGENDAMIENTO =3 THEN ( ISNULL(UPPER(DD.CALLE + ' ' + DD.NUMERO + ' ' + DD.RESTO + ', ' + DD.COMUNA + ', ' + CIUDAD),'') ) "
	strSql = strSql & " 	END MEDIO_AGENDAMIENTO_ASOCIADO "	
	strSql = strSql & " FROM GESTIONES G  INNER JOIN CLIENTE CL ON G.COD_CLIENTE=CL.COD_CLIENTE AND CL.ACTIVO = 1"
	strSql = strSql & " 				  INNER JOIN DEUDOR D ON G.RUT_DEUDOR=D.RUT_DEUDOR AND G.COD_CLIENTE=D.COD_CLIENTE"
	strSql = strSql & " 				  INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA AND"
	strSql = strSql & " 						   G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA AND"
	strSql = strSql & " 						   G.COD_GESTION = GTG.COD_GESTION AND"
	strSql = strSql & " 					       G.COD_CLIENTE = GTG.COD_CLIENTE"
	strSql = strSql & " 				 LEFT JOIN USUARIO U ON G.ID_USUARIO=U.ID_USUARIO"
	strSql = strSql & " 				 LEFT JOIN DEUDOR_TELEFONO DT ON G.ID_MEDIO_GESTION=DT.ID_TELEFONO"
	strSql = strSql & " 				 LEFT JOIN DEUDOR_EMAIL DE ON G.ID_MEDIO_GESTION=DE.ID_EMAIL"
	strSql = strSql & " 				 LEFT JOIN DEUDOR_DIRECCION DD ON G.ID_MEDIO_GESTION=DD.ID_DIRECCION"

	If Trim(strCliente) <> "" and Not IsNull(strCliente) Then
		strSql = strSql & " 	WHERE G.COD_CLIENTE = '" & strCliente & "'"
	End If

	If Trim(dtmInicio) <> "" Then
		strSql = strSql & " AND CONVERT(VARCHAR(10),G.FECHA_INGRESO,103) >= CAST('" & dtmInicio & "' AS DATETIME) "
	End If

	If Trim(dtmTermino) <> "" Then
		strSql = strSql & " AND CONVERT(VARCHAR(10),G.FECHA_INGRESO,103) <= CAST('" & dtmTermino & "' AS DATETIME) "
	End If

	If Trim(strCobranza) = "INTERNA" Then
		strSql = strSql & " AND G.RUT_DEUDOR IN( SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCliente & "' AND CUSTODIO IS NOT NULL ) "
	End If

	If Trim(strCobranza) = "EXTERNA"  Then
		strSql = strSql & " AND G.RUT_DEUDOR IN( SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCliente & "' AND CUSTODIO IS NULL) "
	End If

	If Trim(strUsuario) <> "" Then
		strSql = strSql & " AND G.ID_USUARIO = " & strUsuario
	End If


	if trim(rut_especifico)="1" then
		if trim(rut_especifico_masivo)<>"" then
			if trim(opcion_rut)="1" then
				strSql = strSql & " AND (substring(g.RUT_DEUDOR,1,LEN(g.RUT_DEUDOR)-2)+SUBSTRING(g.RUT_DEUDOR,LEN(g.RUT_DEUDOR),len(g.RUT_DEUDOR))) IN ("&CSTR(rut_masivo)&") "

			elseif trim(opcion_rut)="2" then
				strSql = strSql & " AND substring(g.RUT_DEUDOR,1,LEN(g.RUT_DEUDOR)-2)  IN ("&CSTR(rut_masivo)&") 	"

			else
				strSql = strSql & " AND g.RUT_DEUDOR  IN ("&CSTR(rut_masivo)&")  "

			end if	
		end if	
	end if	

	'Response.write "strSql=" & strSql
	
	set rsTemp= Conn.execute(strSql)
	strTextoTercero=""
	cantSiniestroC = 0

	Do While Not rsTemp.Eof

		intCampana 	= rsTemp("ID_CAMPANA")
		intIdDeudor = rsTemp("ID_DEUDOR")
			
		strGestConcatenada = rsTemp("NOMCATEGORIA") & "-" & rsTemp("NOMSUBCATEGORIA") & "-" & rsTemp("NOMGESTION")
		intGestConcatenada = rsTemp("COD_CATEGORIA") & "*" & rsTemp("COD_SUB_CATEGORIA") & "*" & rsTemp("COD_GESTION")

		strTextoTercero = intIdDeudor & ";" & rsTemp("ID_GESTION")& ";" & rsTemp("RUT_DEUDOR") & ";" & rsTemp("NRO_DEUDOR_CLIENTE")  & ";" & rsTemp("NOM_CLIENTE") & ";" & rsTemp("COD_CLIENTE") & ";" & rsTemp("CORRELATIVO") & ";" & rsTemp("COD_CATEGORIA")  & ";" & rsTemp("COD_SUB_CATEGORIA") & ";" & rsTemp("COD_GESTION") & ";" & intGestConcatenada  & ";"
		strTextoTercero = strTextoTercero & rsTemp("NOMCATEGORIA") & ";" & rsTemp("NOMSUBCATEGORIA") & ";" & rsTemp("NOMGESTION") & ";" & strGestConcatenada & ";" & rsTemp("FECHA_INGRESO")  & ";" & rsTemp("HORA_INGRESO") & ";"
		strTextoTercero = strTextoTercero & rsTemp("FECHA_COMPROMISO") & ";" & rsTemp("FECHA_AGENDAMIENTO") & ";" & rsTemp("HORA_AGENDAMIENTO") & ";" & Replace(Replace(rsTemp("OBSERVACIONES"),chr(13),""),chr(10),"") & ";" & rsTemp("MEDIO_GESTION_ASOCIADO") & ";" & rsTemp("MEDIO_AGENDAMIENTO_ASOCIADO") & ";" & rsTemp("ID_USUARIO") & ";" & rsTemp("LOGIN") & ";"
		strTextoTercero = strTextoTercero & intCampana

		cantSiniestroC = cantSiniestroC + 1

		fichCA.writeline(strTextoTercero)

	rsTemp.movenext
	Loop

	rut_especifico 			=""
	opcion_rut 				=""
	rut_especifico_masivo 	=""

	%>
	<table>
	<tr><td>Cantidad de registros generados : <%= cantSiniestroC %></td></tr>
	<tr><td>
	<a href="#" onClick="AbreArchivo('../logs/<%=strNomArchivoTerceros%>')">Descargar</a>
	&nbsp;
	<a href="#" onClick="location.href='man_export.asp'">Volver</a>


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

