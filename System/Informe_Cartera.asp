<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%

Response.CodePage=65001
Response.charset ="utf-8"
	
inicio= request("inicio")
termino= request("termino")

If Request("CB_TIPO_INF") <> "" then strFiltroInforme=Request("CB_TIPO_INF") Else strFiltroInforme=0 End If

strCodCliente=session("ses_codcli")
strEjeAsig = request("CB_EJECUTIVO")
intEtapaCobranza=Request("CB_ETAPACOB")

If Request("intTipoInforme") <> "" then intTipoInforme=Request("intTipoInforme") else intTipoInforme = "1" End If

If intTipoInforme = 1 then
strColor1 = "boton_rojo"
else
strColor1 = "boton_azul"
End if

If intTipoInforme = 2 then
strColor2 = "boton_rojo"
else
strColor2 = "boton_azul"
End if

If intTipoInforme = 3 then
strColor3 = "boton_rojo"
else
strColor3 = "boton_azul"
End if

'--Calcula Objetos relacionados al tipo de cobranza (Interna, Externa), CB_COBRANZA Y CB_EJECUTIVO--'

	strCobranza = Request("CB_COBRANZA")

	abrirscg()

			strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
			strSql = strSql & " FROM CLIENTE CL"
			strSql = strSql & " WHERE CL.COD_CLIENTE = '" & strCodCliente & "'"

			set RsCli=Conn.execute(strSql)
			If not RsCli.eof then
				intUsaCobInterna = RsCli("USA_COB_INTERNA")
			End if
			RsCli.close
			set RsCli=nothing

	cerrarscg()

	intVerCobExt = "1"
	intVerEjecutivos = "1"

	If TraeSiNo(session("perfil_emp")) = "Si" and strCobranza = "" and intUsaCobInterna = "1" Then

		strCobranza="INTERNA"

	ElseIf TraeSiNo(session("perfil_emp")) = "No" and strCobranza = "" then

		strCobranza="EXTERNA"

	End If

	If TraeSiNo(session("perfil_emp")) = "Si" Then

		intVerEjecutivos="0"

	End If

	If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

		sinCbUsario="0"

	End If

	'Response.write "<br>intTipoInforme=" & intTipoInforme
'---Fin codigo tipo de cobranza---'

%>
<title>INFORME CARTERA</title>
</head>
<body>
<form name="datos" method="post">
<div class="titulo_informe">INFORME CARTERA VIGENTE</div>
<br>
<table width="90%" align="CENTER" border="0">

	<table width="90%" border="0" class="estilo_columnas" align="center">
		<thead>
		<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td height="20">COBRANZA</td>
			<td>TIPO INFORME</td>
			<td>FILTRO INFORME</td>
			<td>ETAPA COBRANZA</td>
			<td>CAMPAÑA</td>

		<% If sinCbUsario = "0" Then %>
			<td>EJECUTIVO</td>
		<% End If %>

			<td width="50">&nbsp;</td>
		</tr>
		</thead>
		<tr>
			<td>
				<select name="CB_COBRANZA" style="width:100px;" <%If sinCbUsario = "0" then%> onChange="CargaUsuarios(CB_COBRANZA.value);" <%End If%> >

					<%If Trim(intVerCobExt) = "1" and Trim(intUsaCobInterna) = "1" Then%>
						<option value="0" <%If Trim(strCobranza) ="" Then Response.write "SELECTED"%>>TODOS</option>
					<%End If%>

					<%If Trim(intUsaCobInterna) = "1" Then%>
						<option value="INTERNA" <%If Trim(strCobranza) ="INTERNA" Then Response.write "SELECTED"%>>INTERNA</option>
					<%End If%>

					<%If Trim(intVerCobExt) = "1" Then%>
						<option value="EXTERNA" <%If Trim(strCobranza) ="EXTERNA" Then Response.write "SELECTED"%>>EXTERNA</option>
					<%End If%>

				</select>
			</td>

			<td>
					  <input name="fi_" class="fondo_boton_100 <%=strColor1%>" style="width:70px;"type="button" onClick="window.navigate('informe_Cartera.asp?intTipoInforme=1');" value="General">

					  <input name="fi_"  class="fondo_boton_100 <%=strColor2%>" style="width:70px;" type="button" onClick="window.navigate('informe_Cartera.asp?intTipoInforme=2');" value="Tramos caso">
					  
					  <input name="fi_"  class="fondo_boton_100 <%=strColor3%>" style="width:70px;" type="button" onClick="window.navigate('informe_Cartera.asp?intTipoInforme=3');" value="Tramos doc">

			</td>

			<%If intTipoInforme = 1 then%>
			<td>
				<select name="CB_TIPO_INF" >
						<option value="0" <%If Trim(strFiltroInforme) ="0" Then Response.write "SELECTED"%>>TODOS</option>
						<option value="GENERAL" <%If Trim(strFiltroInforme) ="GENERAL" Then Response.write "SELECTED"%>>GENERAL</option>
						<option value="TIPO_DOC" <%If Trim(strFiltroInforme) ="TIPO_DOC" Then Response.write "SELECTED"%>>TIPO DOCUMENTO</option>
						<option value="TIPO_SUCURSAL" <%If Trim(strFiltroInforme) ="TIPO_SUCURSAL" Then Response.write "SELECTED"%>>SEDE</option>
						<option value="ESTADO_COB" <%If Trim(strFiltroInforme) ="ESTADO_COB" Then Response.write "SELECTED"%>>ETAPA COBRANZA</option>
				</select>
			</td>
			<%End If%>

			<%If intTipoInforme = 2 then%>
			<td>
				<select name="CB_TIPO_INF" >
						<option value="0" <%If Trim(strFiltroInforme) ="0" Then Response.write "SELECTED"%>>TODOS</option>
						<option value="TRAMO_DD_CASO" <%If Trim(strFiltroInforme) ="TRAMO_DD_CASO" Then Response.write "SELECTED"%>>TRAMOS DEUDA MORA</option>
						<option value="TRAMO_DM_CASO" <%If Trim(strFiltroInforme) ="TRAMO_DM_CASO" Then Response.write "SELECTED"%>>TRAMOS DIA MORA</option>
						<option value="TRAMO_DA_CASO" <%If Trim(strFiltroInforme) ="TRAMO_DA_CASO" Then Response.write "SELECTED"%>>TRAMOS DIA ASIGIGNACION</option>
						<option value="TRAMO_DOC_CASO" <%If Trim(strFiltroInforme) ="TRAMO_DOC_CASO" Then Response.write "SELECTED"%>>TRAMOS DOC POR CASOS</option>
				</select>
			</td>
			<%End If%>

			<%If intTipoInforme = 3 then%>
			<td>
				<select name="CB_TIPO_INF" >
						<option value="0" <%If Trim(strFiltroInforme) ="0" Then Response.write "SELECTED"%>>TODOS</option>
						<option value="TRAMO_DD_DOC" <%If Trim(strFiltroInforme) ="TRAMO_DD_DOC" Then Response.write "SELECTED"%>>TRAMOS DEUDA MORA</option>
						<option value="TRAMO_DM_DOC" <%If Trim(strFiltroInforme) ="TRAMO_DM_DOC" Then Response.write "SELECTED"%>>TRAMOS DIA MORA</option>
						<option value="TRAMO_DA_DOC" <%If Trim(strFiltroInforme) ="TRAMO_DA_DOC" Then Response.write "SELECTED"%>>TRAMOS DIA ASIGIGNACION</option>
				</select>
			</td>
			<%End If%>

			<td>
				<select name="CB_ETAPACOB" >
				<option value="">TODOS</option>
					<%
					abrirscg()

						ssql="SELECT COD_ESTADO_COBRANZA, NOM_ESTADO_COBRANZA FROM ESTADO_COBRANZA"
					set rsTemp= Conn.execute(ssql)
					if not rsTemp.eof then
						do until rsTemp.eof%>
						<option value="<%=rsTemp("COD_ESTADO_COBRANZA")%>"<%if Trim(intEtapaCobranza)=Trim(rsTemp("COD_ESTADO_COBRANZA")) then response.Write("Selected") End If%>><%=rsTemp("NOM_ESTADO_COBRANZA")%></option>
						<%
						rsTemp.movenext
						loop
					end if
					rsTemp.close
					set rsTemp=nothing
					cerrarscg()
					%>
				</select>
			</td>

			<td>
				<select name="CB_CAMPANA" >
					<option value="">TODAS</option>
					<%
					AbrirSCG()
						strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
						set rsCampana=Conn.execute(strSql)
						Do While not rsCampana.eof
							If Trim(intCodCampana)=Trim(rsCampana("ID_CAMPANA")) Then strSelCam = "SELECTED" Else strSelCam = ""
							%>
							<option value="<%=rsCampana("ID_CAMPANA")%>" <%=strSelCam%>> <%=rsCampana("ID_CAMPANA") & " - " & rsCampana("NOMBRE")%></option>
							<%
							rsCampana.movenext
						Loop
						rsCampana.close
						set rsCampana=nothing
					CerrarSCG()
					''Response.End
					%>
				</select>
			</td>

		<% If sinCbUsario="0" Then %>
			<td>
				<select name="CB_EJECUTIVO"  id="CB_EJECUTIVO"  >
				</select>
			</td>
		<% End If %>

			<td ALIGN="CENTER">
				<input type="button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();">
			</td>
		</tr>
    </table>
</form>
<br>
<%

If (strFiltroInforme = "0" or strFiltroInforme = "GENERAL") and intTipoInforme = 1 then

	AbrirSCG()

				'--Obtiene la información relacionada con los documentos--'

				strSql = "SELECT COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS, SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL,"
				strSql= strSql & " MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO, MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"
				strSql= strSql & " FROM (SELECT D.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"
				strSql= strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"
				strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"

				strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

				If Trim(strCobranza) = "INTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
				End if

				If Trim(strCobranza) = "EXTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NULL"
				End if

				strSql= strSql & " GROUP BY D.RUT_DEUDOR) AS PP"

				'Response.write "<br>strSql=" & strSql

				set RsInf=Conn.execute(strSql)
				
				intTotalCasos= "0"
				intMinCaso= "0"
				intMaxCaso= "0"
				intTotalDoc= "0"
				intTotalMonto= "0"
				intMinDoc= "0"
				intMaxDoc= "0"

				If not RsInf.eof then

				intTotalCasos= RsInf("TOTAL_CASOS")
				intMinCaso= RsInf("MIN_CAPITAL_CASO")
				intMaxCaso= RsInf("MAX_CAPITAL_CASO")
				intTotalDoc= RsInf("TOTAL_DOC")
				intTotalMonto= RsInf("MONTO_TOTAL")
				intMinDoc= RsInf("MIN_CAPITAL_DOC")
				intMaxDoc= RsInf("MAX_CAPITAL_DOC")

				intPromMontoCaso= intTotalMonto/intTotalCasos

				intPromDocCaso= intTotalDoc/intTotalCasos

				intPromMontoDcumento= intTotalMonto/intTotalDoc

				End If

	CerrarSCG()
%>

	<table width="90%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER">
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="4" ALIGN="Left" class="subtitulo_informe">
							> Informe General
						</TD>
					</tr>
					<tr class="totales">
						<TD WIDTH="20%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1"  WIDTH="20%" ALIGN="CENTER">
							&nbsp;
						</TD>
					</tr>
					<tr bordercolor="#999999">
						<TD height="30" ALIGN="CENTER">
							<%=FN(intTotalCasos,0)%>
						</TD>
						<TD height="30" ALIGN="CENTER">
							<%=FN(intTotalDoc,0)%>
						</TD>
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intTotalMonto,0)%>
						</TD>
					</tr>
					<tr class="totales">
						<TD WIDTH="20%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD ALIGN="CENTER">
							Promedio Monto / Documento
						</TD>
						<TD ALIGN="CENTER">
							&nbsp;
						</TD>
					</tr>
					<tr bordercolor="#999999">
						<TD height="30"ALIGN="CENTER">
							$&nbsp;<%=FN(intMinDoc,0)%>
						</TD>
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intMaxDoc,0)%>
						</TD>
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intPromMontoDcumento,0)%>
						</TD>
					</tr>
					<tr class="totales">
						<TD WIDTH="20%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Caso máximo
						</TD>
						<TD ALIGN="CENTER">
							Promedio Monto / Caso
						</TD>
						<TD ALIGN="CENTER">
							Promedio Documento / Caso
						</TD>
					</tr>
					<tr bordercolor="#999999">
						<TD height="30" ALIGN="CENTER">
							$&nbsp;<%=FN(intMinCaso,0)%>
						</TD>
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intMaxCaso,0)%>
						</TD>
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intPromMontoCaso,0)%>
						</TD>
						<TD ALIGN="CENTER">
							<%=FN(intPromDocCaso,2)%>
						</TD>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<br>
<%End if

If (strFiltroInforme = "0" or strFiltroInforme = "TIPO_DOC") and intTipoInforme = 1 then%>


	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
				<thead>	
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="8" class="subtitulo_informe">
							> Informe Tipo Documento
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Tipo Documento
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
				</thead>
				<tbody>

	<%AbrirSCG()
					strSql = "SELECT NOM_TIPO_DOCUMENTO,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS, SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL,"
					strSql= strSql & " MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO, MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT TD.NOM_TIPO_DOCUMENTO,D.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"
					strSql= strSql & " FROM CUOTA C  INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"
					strSql= strSql & "	        	 INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & "			     LEFT JOIN TIPO_DOCUMENTO TD ON C.TIPO_DOCUMENTO = TD.COD_TIPO_DOCUMENTO"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY D.RUT_DEUDOR,TD.NOM_TIPO_DOCUMENTO) AS PP"

					strSql= strSql & " GROUP BY PP.NOM_TIPO_DOCUMENTO"
					strSql= strSql & " ORDER BY PP.NOM_TIPO_DOCUMENTO ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf3=Conn.execute(strSql)

					intTTCasosTD=0
					intTTDocTD=0
					intTTMontoTD=0

					if not RsInf3.eof then
						do until RsInf3.eof

						strNomTipoDoc = RsInf3("NOM_TIPO_DOCUMENTO")
						intTotalCasosTD= RsInf3("TOTAL_CASOS")
						intMinCasoTD= RsInf3("MIN_CAPITAL_CASO")
						intMaxCasoTD= RsInf3("MAX_CAPITAL_CASO")
						intTotalDocTD= RsInf3("TOTAL_DOC")
						intTotalMontoTD= RsInf3("MONTO_TOTAL")
						intMinDocTD= RsInf3("MIN_CAPITAL_DOC")
						intMaxDocTD= RsInf3("MAX_CAPITAL_DOC")

						intTTCasosTD= intTTCasosTD + intTotalCasosTD
						intTTDocTD= intTTDocTD + intTotalDocTD
						intTTMontoTD= intTTMontoTD + intTotalMontoTD

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strNomTipoDoc%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosTD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocTD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoTD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocTD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocTD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoTD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoTD,0)%>
							</TD>
						</tr>

						<%RsInf3.movenext
						loop
					end if
					RsInf3.close
					set RsInf3=nothing

	CerrarSCG()%>
			</tbody>
			<thead>
					<tr bgcolor="#380ACD" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosTD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocTD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoTD,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>

<%End if

If strFiltroInforme = "0" and intTipoInforme = 1 then %>

<br>

<%End if

If (strFiltroInforme = "0" or strFiltroInforme = "TIPO_SUCURSAL") and intTipoInforme = 1 then%>

	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
				<thead>	
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="8" class="subtitulo_informe">
							> Informe por Sucursal
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Sucursal
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
				</thead>	
				<tbody>

	<%AbrirSCG()
					strSql = "SELECT SUCURSAL,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS, SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, "
					strSql= strSql & " SUM(CAPITAL_CASO) AS MONTO_TOTAL, MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO, MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,"
					strSql= strSql & " MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC, MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC FROM (SELECT C.SUCURSAL,D.RUT_DEUDOR,"
					strSql= strSql & " SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"
					strSql= strSql & " 				INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY D.RUT_DEUDOR,C.SUCURSAL) AS PP"

					strSql= strSql & " GROUP BY PP.SUCURSAL"
					strSql= strSql & " ORDER BY PP.SUCURSAL ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf5=Conn.execute(strSql)

					intTTCasosSU=0
					intTTDocSU=0
					intTTMontoSU=0

					if not RsInf5.eof then
						do until RsInf5.eof

						strNomSucursal = RsInf5("SUCURSAL")
						intTotalCasosSU= RsInf5("TOTAL_CASOS")
						intMinCasoSU= RsInf5("MIN_CAPITAL_CASO")
						intMaxCasoSU= RsInf5("MAX_CAPITAL_CASO")
						intTotalDocSU= RsInf5("TOTAL_DOC")
						intTotalMontoSU= RsInf5("MONTO_TOTAL")
						intMinDocSU= RsInf5("MIN_CAPITAL_DOC")
						intMaxDocSU= RsInf5("MAX_CAPITAL_DOC")

						intTTCasosSU= intTTCasosSU + intTotalCasosSU
						intTTDocSU= intTTDocSU + intTotalDocSU
						intTTMontoSU= intTTMontoSU + intTotalMontoSU

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strNomSucursal%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosSU,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocSU,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoSU,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocSU,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocSU,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoSU,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoSU,0)%>
							</TD>
						</tr>

						<%RsInf5.movenext
						loop
					end if
					RsInf5.close
					set RsInf5=nothing

	CerrarSCG()%>
				</tbody>
				<thead>

					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosSU,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocSU,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoSU,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>

<%End if

If strFiltroInforme = "0" and intTipoInforme = 1 then %>


<br>

<%End if

If (strFiltroInforme = "0" or strFiltroInforme = "ESTADO_COB") and intTipoInforme = 1 then%>


	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
				<thead>
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="8" class="subtitulo_informe">
							> Informe por Etapa Cobranza
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Estado Cobranza
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
				</thead>
				<tbody>
<%

		AbrirSCG()
					strSql = "SELECT ETAPA_COBRANZA,NOM_ESTADO_COBRANZA,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS, SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL,"
					strSql= strSql & " MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO, MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT D.ETAPA_COBRANZA,EC.NOM_ESTADO_COBRANZA,D.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"
					strSql= strSql & " FROM CUOTA C  INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"
					strSql= strSql & "	        	 INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & "			     INNER JOIN ESTADO_COBRANZA EC ON D.ETAPA_COBRANZA=EC.COD_ESTADO_COBRANZA"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY D.RUT_DEUDOR,D.ETAPA_COBRANZA,EC.NOM_ESTADO_COBRANZA) AS PP"

					strSql= strSql & " GROUP BY PP.ETAPA_COBRANZA,PP.NOM_ESTADO_COBRANZA"
					strSql= strSql & " ORDER BY PP.ETAPA_COBRANZA ASC"

					set RsInf3=Conn.execute(strSql)

					intTTCasosEC=0
					intTTDocEC=0
					intTTMontoEC=0

					if not RsInf3.eof then
						do until RsInf3.eof

						strEstadoCobranza= RsInf3("NOM_ESTADO_COBRANZA")
						intTotalCasosEC= RsInf3("TOTAL_CASOS")
						intMinCasoEC= RsInf3("MIN_CAPITAL_CASO")
						intMaxCasoEC= RsInf3("MAX_CAPITAL_CASO")
						intTotalDocEC= RsInf3("TOTAL_DOC")
						intTotalMontoEC= RsInf3("MONTO_TOTAL")
						intMinDocEC= RsInf3("MIN_CAPITAL_DOC")
						intMaxDocEC= RsInf3("MAX_CAPITAL_DOC")

						intTTCasosEC= intTTCasosEC + intTotalCasosEC
						intTTDocEC= intTTDocEC + intTotalDocEC
						intTTMontoEC= intTTMontoEC + intTotalMontoEC

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strEstadoCobranza%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosEC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocEC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoEC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocEC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocEC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoEC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoEC,0)%>
							</TD>
						</tr>

						<%RsInf3.movenext
						loop
					end if
					RsInf3.close
					set RsInf3=nothing

	CerrarSCG()%>
				</tbody>
				<thead>

					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosEC,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocEC,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoEC,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>
<%End if

If strFiltroInforme = "0" and intTipoInforme = 1 then %>

<br>

<%End If

If (strFiltroInforme = "0" or strFiltroInforme = "TRAMO_DD_CASO") and intTipoInforme = 2 then%>

	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
					<thead>
				   	<tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="7" class="subtitulo_informe">
							> Tramo Deuda Caso
						</TD>
						<TD align="left">
						  </acronym>&nbsp;&nbsp;<acronym title="FICHA DEUDOR">
						  <input name="fi_"  class="fondo_boton_100" type="button" onClick="window.navigate('man_carga_tramos_informes.asp?intTipoInforme=2&intCliente=<%=strCodCliente%>&intTipoTramo=1');" value="Auditar Tramos">
						  </acronym>
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Tramo Deuda
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
				</thead>
				<tbody>

<%		AbrirSCG()
					strSql = "SELECT * FROM TRAMOS_DEUDA"
					strSql = strSql & " WHERE COD_CLIENTE= '" & strCodCliente & "' AND TIPO_TRAMOS = 1"
					strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

					'Response.write "strSql = " & strSql
						set rsDet=Conn.execute(strSql)

						if not rsDet.eof then
							intReg = 0
							do while not rsDet.eof

							intOrdenTramo=rsDet("ORDEN_TRAMO")
							intTramo=rsDet("TRAMO")

								If intOrdenTramo = 1 then
								intTramo_1= intTramo
								End If

								If intOrdenTramo = 2 then
								intTramo_2= intTramo
								End If

								If intOrdenTramo = 3 then
								intTramo_3= intTramo
								End If

								If intOrdenTramo = 4 then
								intTramo_4= intTramo
								End If

							rsDet.movenext
							loop
						end if
						rsDet.close
						set rsDet=nothing

					strSql = "SELECT SUBSTRING(TRAMO_DEUDA,3,20) AS TRAMO_DEUDA_CASO,SUBSTRING(TRAMO_DEUDA,1,1) AS ORDEN_TRAMO_DEUDA,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS,"
					strSql= strSql & " SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO,"
					strSql= strSql & " MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT (CASE WHEN SUM(C.VALOR_CUOTA) <= TR1.TRAMO"
					strSql= strSql & " THEN '1-' +  CAST(TR1.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN SUM(C.VALOR_CUOTA) <= TR2.TRAMO"
					strSql= strSql & " THEN '2-' + CAST(TR2.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN SUM(C.VALOR_CUOTA) <= TR3.TRAMO"
					strSql= strSql & " THEN '3-' +  CAST(TR3.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN SUM(C.VALOR_CUOTA) <= TR4.TRAMO"
					strSql= strSql & " THEN '4-' +  CAST(TR4.TRAMO AS VARCHAR(10))"
					strSql= strSql & " ELSE '5-' +  CAST(TR4.TRAMO AS VARCHAR(10)) END) AS TRAMO_DEUDA,"
					strSql= strSql & " C.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"
					strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & " INNER JOIN CLIENTE CL ON C.COD_CLIENTE = CL.COD_CLIENTE"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR1 ON CL.COD_CLIENTE=TR1.COD_CLIENTE AND TR1.ORDEN_TRAMO=1 AND TR1.TIPO_TRAMOS=1"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR2 ON CL.COD_CLIENTE=TR2.COD_CLIENTE AND TR2.ORDEN_TRAMO=2 AND TR2.TIPO_TRAMOS=1"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR3 ON CL.COD_CLIENTE=TR3.COD_CLIENTE AND TR3.ORDEN_TRAMO=3 AND TR3.TIPO_TRAMOS=1"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR4 ON CL.COD_CLIENTE=TR4.COD_CLIENTE AND TR4.ORDEN_TRAMO=4 AND TR4.TIPO_TRAMOS=1"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY C.RUT_DEUDOR,TR1.TRAMO,TR2.TRAMO,TR3.TRAMO,TR4.TRAMO) AS PP GROUP BY TRAMO_DEUDA ORDER BY PP.TRAMO_DEUDA ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf4=Conn.execute(strSql)

					intTTCasosTRD=0
					intTTDocTRD=0
					intTTMontoTRD=0


					if not RsInf4.eof then
						do until RsInf4.eof

						intOrdenInforme = RsInf4("ORDEN_TRAMO_DEUDA")

						Orden=Orden+1

						'Response.write "<br>intOrdenInforme=" & intOrdenInforme

						If intOrdenInforme = "1" then

						intTramoFinal = intTramo_1
						strTramoDeudaCaso= "0" + " - " + CStr(FN(intTramo_1,0))

						ElseIf intOrdenInforme = "2" then

						intTramoFinal = intTramo_2
						strTramoDeudaCaso= CStr(FN(Clng(intTramo_1) + 1,0)) + " - " + CStr(FN(intTramo_2,0))

						ElseIf intOrdenInforme = "3" then

						intTramoFinal = intTramo_3
						strTramoDeudaCaso= CStr(FN(Clng(intTramo_2) + 1,0)) + " - " + CStr(FN(intTramo_3,0))

						ElseIf intOrdenInforme = "4" then

						intTramoFinal = intTramo_4
						strTramoDeudaCaso= CStr(FN(Clng(intTramo_3) + 1,0)) + " - " + CStr(FN(intTramo_4,0))

						Else

						strTramoDeudaCaso= "MAYOR A " + CStr(FN(Clng(intTramoFinal) + 1,0))

						End If

						strTramoDeudaAnt= FN(RsInf4("TRAMO_DEUDA_CASO") + 1,0)
						intTotalCasosTRD= RsInf4("TOTAL_CASOS")
						intMinCasoTRD= RsInf4("MIN_CAPITAL_CASO")
						intMaxCasoTRD= RsInf4("MAX_CAPITAL_CASO")
						intTotalDocTRD= RsInf4("TOTAL_DOC")
						intTotalMontoTRD= RsInf4("MONTO_TOTAL")
						intMinDocTRD= RsInf4("MIN_CAPITAL_DOC")
						intMaxDocTRD= RsInf4("MAX_CAPITAL_DOC")

						intTTCasosTRD= intTTCasosTRD + intTotalCasosTRD
						intTTDocTRD= intTTDocTRD + intTotalDocTRD
						intTTMontoTRD= intTTMontoTRD + intTotalMontoTRD

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strTramoDeudaCaso%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosTRD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocTRD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoTRD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocTRD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocTRD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoTRD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoTRD,0)%>
							</TD>
						</tr>

						<%RsInf4.movenext
						loop
					Else%>
						<TD Height="20" bgcolor="#ff6666" Colspan="8" align="CENTER">
							<B>SELECCIONE TRAMOS PARA VISUALIZAR INFORME<B>
						</TD>
					<%end if
					RsInf4.close
					set RsInf4=nothing

	CerrarSCG()%>
					</tbody>
					<thead>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosTRD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocTRD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoTRD,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>

<%End if

If strFiltroInforme = "0" and intTipoInforme = 2 then %>

<br>

<%End If

If (strFiltroInforme = "0" or strFiltroInforme = "TRAMO_DM_CASO") and intTipoInforme = 2 then%>

	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
					<thead>
				   	<tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="7" class="subtitulo_informe">
							> Tramo Días Mora Caso (Segmenta según el documento con mayor vencimiento)
						</TD>
						<TD align="left">
						  </acronym>&nbsp;&nbsp;<acronym title="FICHA DEUDOR">
						  <input name="fi_" class="fondo_boton_100"  type="button" onClick="window.navigate('man_carga_tramos_informes.asp?intTipoInforme=2&intCliente=<%=strCodCliente%>&intTipoTramo=2');" value="Auditar Tramos">
						  </acronym>
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Tramo Dias Mora
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
					</thead>
					<tbody>
<%		AbrirSCG()

					strSql = "SELECT * FROM TRAMOS_DEUDA"
					strSql = strSql & " WHERE COD_CLIENTE= '" & strCodCliente & "' AND TIPO_TRAMOS = 2"
					strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

					'Response.write "strSql = " & strSql
						set rsDet=Conn.execute(strSql)

						if not rsDet.eof then
							intReg = 0
							do while not rsDet.eof

							intOrdenTramo=rsDet("ORDEN_TRAMO")
							intTramo=rsDet("TRAMO")

								If intOrdenTramo = 1 then
								intTramo_1= intTramo
								End If

								If intOrdenTramo = 2 then
								intTramo_2= intTramo
								End If

								If intOrdenTramo = 3 then
								intTramo_3= intTramo
								End If

								If intOrdenTramo = 4 then
								intTramo_4= intTramo
								End If

							rsDet.movenext
							loop
						end if
						rsDet.close
						set rsDet=nothing

					'Response.write "<br>intTramo_1=" & intTramo_1
					'Response.write "<br>intTramo_2=" & intTramo_2
					'Response.write "<br>intTramo_3=" & intTramo_3
					'Response.write "<br>intTramo_4=" & intTramo_4

					strSql = "SELECT SUBSTRING(TRAMO_DEUDA,3,20) AS TRAMO_DM_CASO,SUBSTRING(TRAMO_DEUDA,1,1) AS ORDEN_TRAMO_DEUDA,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS,"
					strSql= strSql & " SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO,"
					strSql= strSql & " MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT (CASE WHEN MAX(DATEDIFF(DAY,C.FECHA_VENC,GETDATE())) <= TR1.TRAMO"
					strSql= strSql & " THEN '1-' + CAST(TR1.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,C.FECHA_VENC,GETDATE())) <= TR2.TRAMO "
					strSql= strSql & " THEN '2-' + CAST(TR2.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,C.FECHA_VENC,GETDATE()))  <= TR3.TRAMO"
					strSql= strSql & " THEN '3-' + CAST(TR3.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,C.FECHA_VENC,GETDATE())) <= TR4.TRAMO"
					strSql= strSql & " THEN '4-' + CAST(TR4.TRAMO AS VARCHAR(10)) ELSE '5-' + CAST(TR4.TRAMO AS VARCHAR(10)) END) AS TRAMO_DEUDA,"
					strSql= strSql & " C.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"

					strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & " INNER JOIN CLIENTE CL ON C.COD_CLIENTE = CL.COD_CLIENTE"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR1 ON CL.COD_CLIENTE=TR1.COD_CLIENTE AND TR1.ORDEN_TRAMO=1 AND TR1.TIPO_TRAMOS=2"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR2 ON CL.COD_CLIENTE=TR2.COD_CLIENTE AND TR2.ORDEN_TRAMO=2 AND TR2.TIPO_TRAMOS=2"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR3 ON CL.COD_CLIENTE=TR3.COD_CLIENTE AND TR3.ORDEN_TRAMO=3 AND TR3.TIPO_TRAMOS=2"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR4 ON CL.COD_CLIENTE=TR4.COD_CLIENTE AND TR4.ORDEN_TRAMO=4 AND TR4.TIPO_TRAMOS=2"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY C.RUT_DEUDOR,TR1.TRAMO,TR2.TRAMO,TR3.TRAMO,TR4.TRAMO) AS PP GROUP BY TRAMO_DEUDA ORDER BY PP.TRAMO_DEUDA ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf4=Conn.execute(strSql)

					intTTCasosTRD=0
					intTTDocTRD=0
					intTTMontoTRD=0

					if not RsInf4.eof then
						do until RsInf4.eof

						intOrdenInforme = RsInf4("ORDEN_TRAMO_DEUDA")

						'Response.write "<br>intOrdenInforme=" & intOrdenInforme

						If intOrdenInforme = "1" then

						strTramoDMCaso= "0" + " - " + CStr(FN(intTramo_1,0))
						intTramoFinal2 = intTramo_1

						ElseIf intOrdenInforme = "2" then
						intTramoFinal2 = intTramo_2

						strTramoDMCaso= CStr(FN(Clng(intTramo_1) + 1,0)) + " - " + CStr(FN(intTramo_2,0))

						ElseIf intOrdenInforme = "3" then
						intTramoFinal2 = intTramo_3

						strTramoDMCaso= CStr(FN(Clng(intTramo_2) + 1,0)) + " - " + CStr(FN(intTramo_3,0))

						ElseIf intOrdenInforme = "4" then
						intTramoFinal2 = intTramo_4

						strTramoDMCaso= CStr(FN(Clng(intTramo_3) + 1,0)) + " - " + CStr(FN(intTramo_4,0))

						Else

						strTramoDMCaso= "MAYOR A " + CStr(FN(Clng(intTramoFinal2) + 1,0))

						End If

						strTramoDMAnt= FN(RsInf4("TRAMO_DM_CASO") + 1,0)
						intTotalCasosTDM= RsInf4("TOTAL_CASOS")
						intMinCasoTDM= RsInf4("MIN_CAPITAL_CASO")
						intMaxCasoTDM= RsInf4("MAX_CAPITAL_CASO")
						intTotalDocTDM= RsInf4("TOTAL_DOC")
						intTotalMontoTDM= RsInf4("MONTO_TOTAL")
						intMinDocTDM= RsInf4("MIN_CAPITAL_DOC")
						intMaxDocTDM= RsInf4("MAX_CAPITAL_DOC")

						intTTCasosTDM= intTTCasosTDM + intTotalCasosTDM
						intTTDocTDM= intTTDocTDM + intTotalDocTDM
						intTTMontoTDM= intTTMontoTDM + intTotalMontoTDM

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strTramoDMCaso%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosTDM,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocTDM,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoTDM,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocTDM,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocTDM,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoTDM,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoTDM,0)%>
							</TD>
						</tr>

						<%RsInf4.movenext
						loop
					Else%>
						<TD Height="20" bgcolor="#ff6666" Colspan="8" align="CENTER">
							<B>SELECCIONE TRAMOS PARA VISUALIZAR INFORME<B>
						</TD>
					<%end if
					RsInf4.close
					set RsInf4=nothing

	CerrarSCG()%>
				</tbody>
				<thead>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosTDM,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocTDM,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoTDM,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>

	<%End if

If strFiltroInforme = "0" and intTipoInforme = 2 then %>

<br>

<%End If

If (strFiltroInforme = "0" or strFiltroInforme = "TRAMO_DA_CASO") and intTipoInforme = 2 then%>

	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
					<thead>
				   	<tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="7" class="subtitulo_informe">
							> Tramo Día Asignación Caso (Segmenta según el documento con mayor días de asignación)
						</TD>
						<TD align="left">
						  </acronym>&nbsp;&nbsp;<acronym title="FICHA DEUDOR">
						  <input name="fi_"  class="fondo_boton_100" type="button" onClick="window.navigate('man_carga_tramos_informes.asp?intTipoInforme=2&intCliente=<%=strCodCliente%>&intTipoTramo=3');" value="Auditar Tramos">
						  </acronym>
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Tramo Dias Asignación
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
				</thead>
				<tbody>

<%		AbrirSCG()

					strSql = "SELECT * FROM TRAMOS_DEUDA"
					strSql = strSql & " WHERE COD_CLIENTE= '" & strCodCliente & "' AND TIPO_TRAMOS = 3"
					strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

					'Response.write "strSql = " & strSql
						set rsDet=Conn.execute(strSql)

						if not rsDet.eof then
							intReg = 0
							do while not rsDet.eof

							intOrdenTramo=rsDet("ORDEN_TRAMO")
							intTramo=rsDet("TRAMO")

								If intOrdenTramo = 1 then
								intTramo_1= intTramo
								End If

								If intOrdenTramo = 2 then
								intTramo_2= intTramo
								End If

								If intOrdenTramo = 3 then
								intTramo_3= intTramo
								End If

								If intOrdenTramo = 4 then
								intTramo_4= intTramo
								End If

							rsDet.movenext
							loop
						end if
						rsDet.close
						set rsDet=nothing

					'Response.write "<br>intTramo_1=" & intTramo_1
					'Response.write "<br>intTramo_2=" & intTramo_2
					'Response.write "<br>intTramo_3=" & intTramo_3
					'Response.write "<br>intTramo_4=" & intTramo_4

					strSql = "SELECT SUBSTRING(TRAMO_DEUDA,3,20) AS TRAMO_DA_CASO,SUBSTRING(TRAMO_DEUDA,1,1) AS ORDEN_TRAMO_DEUDA,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS,"
					strSql= strSql & " SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO,"
					strSql= strSql & " MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT (CASE WHEN MAX(DATEDIFF(DAY,ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION),GETDATE())) <= TR1.TRAMO"
					strSql= strSql & " THEN '1-' + CAST(TR1.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION),GETDATE())) <= TR2.TRAMO "
					strSql= strSql & " THEN '2-' + CAST(TR2.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION),GETDATE())) <= TR3.TRAMO"
					strSql= strSql & " THEN '3-' + CAST(TR3.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION),GETDATE())) <= TR4.TRAMO"
					strSql= strSql & " THEN '4-' + CAST(TR4.TRAMO AS VARCHAR(10)) ELSE '5-' + CAST(TR4.TRAMO AS VARCHAR(10)) END) AS TRAMO_DEUDA,"
					strSql= strSql & " C.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"

					strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & " INNER JOIN CLIENTE CL ON C.COD_CLIENTE = CL.COD_CLIENTE"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR1 ON CL.COD_CLIENTE=TR1.COD_CLIENTE AND TR1.ORDEN_TRAMO=1 AND TR1.TIPO_TRAMOS=3"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR2 ON CL.COD_CLIENTE=TR2.COD_CLIENTE AND TR2.ORDEN_TRAMO=2 AND TR2.TIPO_TRAMOS=3"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR3 ON CL.COD_CLIENTE=TR3.COD_CLIENTE AND TR3.ORDEN_TRAMO=3 AND TR3.TIPO_TRAMOS=3"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR4 ON CL.COD_CLIENTE=TR4.COD_CLIENTE AND TR4.ORDEN_TRAMO=4 AND TR4.TIPO_TRAMOS=3"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY C.RUT_DEUDOR,TR1.TRAMO,TR2.TRAMO,TR3.TRAMO,TR4.TRAMO) AS PP GROUP BY TRAMO_DEUDA ORDER BY PP.TRAMO_DEUDA ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf4=Conn.execute(strSql)

					intTTCasosTRA=0
					intTTDocTRA=0
					intTTMontoTRA=0

					if not RsInf4.eof then
						do until RsInf4.eof

						intOrdenInforme = RsInf4("ORDEN_TRAMO_DEUDA")

						'Response.write "<br>intOrdenInforme=" & intOrdenInforme

						If intOrdenInforme = "1" then

						strTramoDACaso= "0" + " - " + CStr(FN(intTramo_1,0))
						intTramoFinal3 = intTramo_1

						ElseIf intOrdenInforme = "2" then
						intTramoFinal3 = intTramo_2

						strTramoDACaso= CStr(FN(Clng(intTramo_1) + 1,0)) + " - " + CStr(FN(intTramo_2,0))

						ElseIf intOrdenInforme = "3" then
						intTramoFinal3 = intTramo_3

						strTramoDACaso= CStr(FN(Clng(intTramo_2) + 1,0)) + " - " + CStr(FN(intTramo_3,0))

						ElseIf intOrdenInforme = "4" then
						intTramoFinal3 = intTramo_4

						strTramoDACaso= CStr(FN(Clng(intTramo_3) + 1,0)) + " - " + CStr(FN(intTramo_4,0))

						Else

						strTramoDACaso= "MAYOR A " + CStr(FN(Clng(intTramoFinal3) + 1,0))

						End If

						strTramoDMAnt= FN(RsInf4("TRAMO_DA_CASO") + 1,0)
						intTotalCasosTDA= RsInf4("TOTAL_CASOS")
						intMinCasoTDA= RsInf4("MIN_CAPITAL_CASO")
						intMaxCasoTDA= RsInf4("MAX_CAPITAL_CASO")
						intTotalDocTDA= RsInf4("TOTAL_DOC")
						intTotalMontoTDA= RsInf4("MONTO_TOTAL")
						intMinDocTDA= RsInf4("MIN_CAPITAL_DOC")
						intMaxDocTDA= RsInf4("MAX_CAPITAL_DOC")

						intTTCasosTDA= intTTCasosTDA + intTotalCasosTDA
						intTTDocTDA= intTTDocTDA + intTotalDocTDA
						intTTMontoTDA= intTTMontoTDA + intTotalMontoTDA

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strTramoDACaso%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosTDA,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocTDA,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoTDA,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocTDA,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocTDA,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoTDA,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoTDA,0)%>
							</TD>
						</tr>

						<%RsInf4.movenext
						loop
					Else%>
						<TD Height="20" bgcolor="#ff6666" Colspan="8" align="CENTER">
							<B>SELECCIONE TRAMOS PARA VISUALIZAR INFORME<B>
						</TD>
					<%end if
					RsInf4.close
					set RsInf4=nothing

	CerrarSCG()%>
					</tbody>
					<thead>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosTDA,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocTDA,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoTDA,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>
	<%End if

If strFiltroInforme = "0" and intTipoInforme = 2 then %>

<br>

<%End If

If (strFiltroInforme = "0" or strFiltroInforme = "TRAMO_DOC_CASO") and intTipoInforme = 2 then%>

	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
				<thead>
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="7" class="subtitulo_informe">
							> Tramo Documentos Caso
						</TD>
						<TD align="left">
						  </acronym>&nbsp;&nbsp;<acronym title="FICHA DEUDOR">
						  <input name="fi_"  class="fondo_boton_100" type="button" onClick="window.navigate('man_carga_tramos_informes.asp?intTipoInforme=2&intCliente=<%=strCodCliente%>&intTipoTramo=4');" value="Auditar Tramos">
						  </acronym>
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Tramo Dcoumentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
				</thead>
				<tbody>
<%		AbrirSCG()

					strSql = "SELECT * FROM TRAMOS_DEUDA"
					strSql = strSql & " WHERE COD_CLIENTE= '" & strCodCliente & "' AND TIPO_TRAMOS = 4"
					strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

					'Response.write "strSql = " & strSql
						set rsDet=Conn.execute(strSql)

						if not rsDet.eof then
							intReg = 0
							do while not rsDet.eof

							intOrdenTramo=rsDet("ORDEN_TRAMO")
							intTramo=rsDet("TRAMO")

								If intOrdenTramo = 1 then
								intTramo_1= intTramo
								End If

								If intOrdenTramo = 2 then
								intTramo_2= intTramo
								End If

								If intOrdenTramo = 3 then
								intTramo_3= intTramo
								End If

								If intOrdenTramo = 4 then
								intTramo_4= intTramo
								End If

							rsDet.movenext
							loop
						end if
						rsDet.close
						set rsDet=nothing

					'Response.write "<br>intTramo_1=" & intTramo_1
					'Response.write "<br>intTramo_2=" & intTramo_2
					'Response.write "<br>intTramo_3=" & intTramo_3
					'Response.write "<br>intTramo_4=" & intTramo_4

					strSql = "SELECT SUBSTRING(TRAMO_DEUDA,3,20) AS TRAMO_DA_CASO,SUBSTRING(TRAMO_DEUDA,1,1) AS ORDEN_TRAMO_DEUDA,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS,"
					strSql= strSql & " SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO,"
					strSql= strSql & " MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT (CASE WHEN COUNT(C.RUT_DEUDOR) <= TR1.TRAMO"
					strSql= strSql & " THEN '1-' + CAST(TR1.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN COUNT(C.RUT_DEUDOR) <= TR2.TRAMO "
					strSql= strSql & " THEN '2-' + CAST(TR2.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN COUNT(C.RUT_DEUDOR) <= TR3.TRAMO"
					strSql= strSql & " THEN '3-' + CAST(TR3.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN COUNT(C.RUT_DEUDOR) <= TR4.TRAMO"
					strSql= strSql & " THEN '4-' + CAST(TR4.TRAMO AS VARCHAR(10)) ELSE '5-' + CAST(TR4.TRAMO AS VARCHAR(10)) END) AS TRAMO_DEUDA,"
					strSql= strSql & " C.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"

					strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & " INNER JOIN CLIENTE CL ON C.COD_CLIENTE = CL.COD_CLIENTE"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR1 ON CL.COD_CLIENTE=TR1.COD_CLIENTE AND TR1.ORDEN_TRAMO=1 AND TR1.TIPO_TRAMOS=4"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR2 ON CL.COD_CLIENTE=TR2.COD_CLIENTE AND TR2.ORDEN_TRAMO=2 AND TR2.TIPO_TRAMOS=4"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR3 ON CL.COD_CLIENTE=TR3.COD_CLIENTE AND TR3.ORDEN_TRAMO=3 AND TR3.TIPO_TRAMOS=4"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR4 ON CL.COD_CLIENTE=TR4.COD_CLIENTE AND TR4.ORDEN_TRAMO=4 AND TR4.TIPO_TRAMOS=4"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY C.RUT_DEUDOR,TR1.TRAMO,TR2.TRAMO,TR3.TRAMO,TR4.TRAMO) AS PP GROUP BY TRAMO_DEUDA ORDER BY PP.TRAMO_DEUDA ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf4=Conn.execute(strSql)

					intTTCasosTRA=0
					intTTDocTRA=0
					intTTMontoTRA=0

					if not RsInf4.eof then
						do until RsInf4.eof

						intOrdenInforme = RsInf4("ORDEN_TRAMO_DEUDA")

						'Response.write "<br>intOrdenInforme=" & intOrdenInforme

						If intOrdenInforme = "1" then

						strTramoDOCCaso= "0" + " - " + CStr(FN(intTramo_1,0))
						intTramoFinal4 = intTramo_1

						ElseIf intOrdenInforme = "2" then
						intTramoFinal4 = intTramo_2

						strTramoDOCCaso= CStr(FN(Clng(intTramo_1) + 1,0)) + " - " + CStr(FN(intTramo_2,0))

						ElseIf intOrdenInforme = "3" then
						intTramoFinal4 = intTramo_3

						strTramoDOCCaso= CStr(FN(Clng(intTramo_2) + 1,0)) + " - " + CStr(FN(intTramo_3,0))

						ElseIf intOrdenInforme = "4" then
						intTramoFinal4 = intTramo_4

						strTramoDOCCaso= CStr(FN(Clng(intTramo_3) + 1,0)) + " - " + CStr(FN(intTramo_4,0))

						Else

						strTramoDOCCaso= "MAYOR A " + CStr(FN(Clng(intTramoFinal4) + 1,0))

						End If

						strTramoTDOCAnt= FN(RsInf4("TRAMO_DA_CASO") + 1,0)
						intTotalCasosTDOC= RsInf4("TOTAL_CASOS")
						intMinCasoTDOC= RsInf4("MIN_CAPITAL_CASO")
						intMaxCasoTDOC= RsInf4("MAX_CAPITAL_CASO")
						intTotalDocTDOC= RsInf4("TOTAL_DOC")
						intTotalMontoTDOC= RsInf4("MONTO_TOTAL")
						intMinDocTDOC= RsInf4("MIN_CAPITAL_DOC")
						intMaxDocTDOC= RsInf4("MAX_CAPITAL_DOC")

						intTTCasosTDOC= intTTCasosTDOC + intTotalCasosTDOC
						intTTDocTDOC= intTTDocTDOC + intTotalDocTDOC
						intTTMontoTDOC= intTTMontoTDOC + intTotalMontoTDOC

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strTramoDOCCaso%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosTDOC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocTDOC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoTDOC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocTDOC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocTDOC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoTDOC,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoTDOC,0)%>
							</TD>
						</tr>

						<%RsInf4.movenext
						loop
					Else%>
						<TD Height="20" bgcolor="#ff6666" Colspan="8" align="CENTER">
							<B>SELECCIONE TRAMOS PARA VISUALIZAR INFORME<B>
						</TD>
					<%end if
					RsInf4.close
					set RsInf4=nothing

	CerrarSCG()%>
				</tbody>
				<thead>

					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosTDOC,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocTDOC,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoTDOC,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>

<%End If

If (strFiltroInforme = "0" or strFiltroInforme = "TRAMO_DD_DOC") and intTipoInforme = 3 then%>

	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
					<thead>
				   	<tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="7" class="subtitulo_informe">
							> Tramo Deuda Doc (Segmenta según el monto capital del documento)
						</TD>
						<TD align="left">
						  </acronym>&nbsp;&nbsp;<acronym title="FICHA DEUDOR">
						  <input name="fi_"  class="fondo_boton_100" type="button" onClick="window.navigate('man_carga_tramos_informes.asp?intTipoInforme=3&intCliente=<%=strCodCliente%>&intTipoTramo=5');" value="Auditar Tramos">
						  </acronym>
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Tramo Deuda
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
					</thead>
					<tbody>
<%		AbrirSCG()
					strSql = "SELECT * FROM TRAMOS_DEUDA"
					strSql = strSql & " WHERE COD_CLIENTE= '" & strCodCliente & "' AND TIPO_TRAMOS = 5"
					strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

					'Response.write "strSql = " & strSql
						set rsDet=Conn.execute(strSql)

						if not rsDet.eof then
							intReg = 0
							do while not rsDet.eof

							intOrdenTramo=rsDet("ORDEN_TRAMO")
							intTramo=rsDet("TRAMO")

								If intOrdenTramo = 1 then
								intTramo_1= intTramo
								End If

								If intOrdenTramo = 2 then
								intTramo_2= intTramo
								End If

								If intOrdenTramo = 3 then
								intTramo_3= intTramo
								End If

								If intOrdenTramo = 4 then
								intTramo_4= intTramo
								End If

							rsDet.movenext
							loop
						end if
						rsDet.close
						set rsDet=nothing

					strSql = "SELECT SUBSTRING(TRAMO_DEUDA,3,20) AS TRAMO_DEUDA_CASO,SUBSTRING(TRAMO_DEUDA,1,1) AS ORDEN_TRAMO_DEUDA,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS,"
					strSql= strSql & " SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO,"
					strSql= strSql & " MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT (CASE WHEN SUM(C.VALOR_CUOTA) <= TR1.TRAMO"
					strSql= strSql & " THEN '1-' +  CAST(TR1.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN SUM(C.VALOR_CUOTA) <= TR2.TRAMO"
					strSql= strSql & " THEN '2-' + CAST(TR2.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN SUM(C.VALOR_CUOTA) <= TR3.TRAMO"
					strSql= strSql & " THEN '3-' +  CAST(TR3.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN SUM(C.VALOR_CUOTA) <= TR4.TRAMO"
					strSql= strSql & " THEN '4-' +  CAST(TR4.TRAMO AS VARCHAR(10))"
					strSql= strSql & " ELSE '5-' +  CAST(TR4.TRAMO AS VARCHAR(10)) END) AS TRAMO_DEUDA,"
					strSql= strSql & " C.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"
					strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & " INNER JOIN CLIENTE CL ON C.COD_CLIENTE = CL.COD_CLIENTE"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR1 ON CL.COD_CLIENTE=TR1.COD_CLIENTE AND TR1.ORDEN_TRAMO=1 AND TR1.TIPO_TRAMOS=5"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR2 ON CL.COD_CLIENTE=TR2.COD_CLIENTE AND TR2.ORDEN_TRAMO=2 AND TR2.TIPO_TRAMOS=5"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR3 ON CL.COD_CLIENTE=TR3.COD_CLIENTE AND TR3.ORDEN_TRAMO=3 AND TR3.TIPO_TRAMOS=5"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR4 ON CL.COD_CLIENTE=TR4.COD_CLIENTE AND TR4.ORDEN_TRAMO=4 AND TR4.TIPO_TRAMOS=5"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY C.RUT_DEUDOR,C.ID_CUOTA,TR1.TRAMO,TR2.TRAMO,TR3.TRAMO,TR4.TRAMO) AS PP GROUP BY TRAMO_DEUDA ORDER BY PP.TRAMO_DEUDA ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf4=Conn.execute(strSql)

					intTTDOCTRDD=0
					intTTDocTRDD=0
					intTTMontoTRDD=0


					if not RsInf4.eof then
						do until RsInf4.eof

						intOrdenInforme = RsInf4("ORDEN_TRAMO_DEUDA")

						Orden=Orden+1

						'Response.write "<br>intOrdenInforme=" & intOrdenInforme

						If intOrdenInforme = "1" then

						intTramoFinal5 = intTramo_1
						strTramoDeudaCasoD= "0" + " - " + CStr(FN(intTramo_1,0))

						ElseIf intOrdenInforme = "2" then

						intTramoFinal5 = intTramo_2
						strTramoDeudaCasoD= CStr(FN(Clng(intTramo_1) + 1,0)) + " - " + CStr(FN(intTramo_2,0))

						ElseIf intOrdenInforme = "3" then

						intTramoFinal5 = intTramo_3
						strTramoDeudaCasoD= CStr(FN(Clng(intTramo_2) + 1,0)) + " - " + CStr(FN(intTramo_3,0))

						ElseIf intOrdenInforme = "4" then

						intTramoFinal5 = intTramo_4
						strTramoDeudaCasoD= CStr(FN(Clng(intTramo_3) + 1,0)) + " - " + CStr(FN(intTramo_4,0))

						Else

						strTramoDeudaCasoD= "MAYOR A " + CStr(FN(Clng(intTramoFinal5) + 1,0))

						End If

						strTramoDeudaAntD= FN(RsInf4("TRAMO_DEUDA_CASO") + 1,0)
						intTotalCasosTRDD= RsInf4("TOTAL_CASOS")
						intMinCasoTRDD= RsInf4("MIN_CAPITAL_CASO")
						intMaxCasoTRDD= RsInf4("MAX_CAPITAL_CASO")
						intTotalDocTRDD= RsInf4("TOTAL_DOC")
						intTotalMontoTRDD= RsInf4("MONTO_TOTAL")
						intMinDocTRDD= RsInf4("MIN_CAPITAL_DOC")
						intMaxDocTRDD= RsInf4("MAX_CAPITAL_DOC")

						intTTCasosTRDD= intTTCasosTRDD + intTotalCasosTRDD
						intTTDocTRDD= intTTDocTRDD + intTotalDocTRDD
						intTTMontoTRDD= intTTMontoTRDD + intTotalMontoTRDD

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strTramoDeudaCasoD%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosTRDD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocTRDD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoTRDD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocTRDD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocTRDD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoTRDD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoTRDD,0)%>
							</TD>
						</tr>

						<%RsInf4.movenext
						loop
					Else%>
						<TD Height="20" bgcolor="#ff6666" Colspan="8" align="CENTER">
							<B>SELECCIONE TRAMOS PARA VISUALIZAR INFORME<B>
						</TD>
					<%end if
					RsInf4.close
					set RsInf4=nothing

	CerrarSCG()%>
				</tbody>
				<thead>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosTRDD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocTRDD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoTRDD,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>

<%End if

If strFiltroInforme = "0" and intTipoInforme = 3 then %>

<br>

<%End If

If (strFiltroInforme = "0" or strFiltroInforme = "TRAMO_DM_DOC") and intTipoInforme = 3 then%>

	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
					<thead>
				   	<tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="7" class="subtitulo_informe">
							> Tramo Días Mora Doc (Segmenta según los días de vencimienton del documento)
						</TD>
						<TD align="left">
						  </acronym>&nbsp;&nbsp;<acronym title="FICHA DEUDOR">
						  <input name="fi_"  class="fondo_boton_100" type="button" onClick="window.navigate('man_carga_tramos_informes.asp?intTipoInforme=3&intCliente=<%=strCodCliente%>&intTipoTramo=6');" value="Auditar Tramos">
						  </acronym>
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Tramo Dias Mora
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
					</thead>
					<tbody>
<%		AbrirSCG()

					strSql = "SELECT * FROM TRAMOS_DEUDA"
					strSql = strSql & " WHERE COD_CLIENTE= '" & strCodCliente & "' AND TIPO_TRAMOS = 6"
					strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

					'Response.write "strSql = " & strSql
						set rsDet=Conn.execute(strSql)

						if not rsDet.eof then
							intReg = 0
							do while not rsDet.eof

							intOrdenTramo=rsDet("ORDEN_TRAMO")
							intTramo=rsDet("TRAMO")

								If intOrdenTramo = 1 then
								intTramo_1= intTramo
								End If

								If intOrdenTramo = 2 then
								intTramo_2= intTramo
								End If

								If intOrdenTramo = 3 then
								intTramo_3= intTramo
								End If

								If intOrdenTramo = 4 then
								intTramo_4= intTramo
								End If

							rsDet.movenext
							loop
						end if
						rsDet.close
						set rsDet=nothing

					'Response.write "<br>intTramo_1=" & intTramo_1
					'Response.write "<br>intTramo_2=" & intTramo_2
					'Response.write "<br>intTramo_3=" & intTramo_3
					'Response.write "<br>intTramo_4=" & intTramo_4

					strSql = "SELECT SUBSTRING(TRAMO_DEUDA,3,20) AS TRAMO_DM_CASO,SUBSTRING(TRAMO_DEUDA,1,1) AS ORDEN_TRAMO_DEUDA,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS,"
					strSql= strSql & " SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO,"
					strSql= strSql & " MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT (CASE WHEN MAX(DATEDIFF(DAY,C.FECHA_VENC,GETDATE())) <= TR1.TRAMO"
					strSql= strSql & " THEN '1-' + CAST(TR1.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,C.FECHA_VENC,GETDATE())) <= TR2.TRAMO "
					strSql= strSql & " THEN '2-' + CAST(TR2.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,C.FECHA_VENC,GETDATE()))  <= TR3.TRAMO"
					strSql= strSql & " THEN '3-' + CAST(TR3.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,C.FECHA_VENC,GETDATE()))  <= TR4.TRAMO"
					strSql= strSql & " THEN '4-' + CAST(TR4.TRAMO AS VARCHAR(10)) ELSE '5-' + CAST(TR4.TRAMO AS VARCHAR(10)) END) AS TRAMO_DEUDA,"
					strSql= strSql & " C.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"

					strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & " INNER JOIN CLIENTE CL ON C.COD_CLIENTE = CL.COD_CLIENTE"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR1 ON CL.COD_CLIENTE=TR1.COD_CLIENTE AND TR1.ORDEN_TRAMO=1 AND TR1.TIPO_TRAMOS=6"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR2 ON CL.COD_CLIENTE=TR2.COD_CLIENTE AND TR2.ORDEN_TRAMO=2 AND TR2.TIPO_TRAMOS=6"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR3 ON CL.COD_CLIENTE=TR3.COD_CLIENTE AND TR3.ORDEN_TRAMO=3 AND TR3.TIPO_TRAMOS=6"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR4 ON CL.COD_CLIENTE=TR4.COD_CLIENTE AND TR4.ORDEN_TRAMO=4 AND TR4.TIPO_TRAMOS=6"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY C.RUT_DEUDOR,C.ID_CUOTA,TR1.TRAMO,TR2.TRAMO,TR3.TRAMO,TR4.TRAMO) AS PP GROUP BY TRAMO_DEUDA ORDER BY PP.TRAMO_DEUDA ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf4=Conn.execute(strSql)

					intTTCasosTRMD=0
					intTTDocTRMD=0
					intTTMontoTRMD=0

					if not RsInf4.eof then
						do until RsInf4.eof

						intOrdenInforme = RsInf4("ORDEN_TRAMO_DEUDA")

						'Response.write "<br>intOrdenInforme=" & intOrdenInforme

						If intOrdenInforme = "1" then

						strTramoDMD= "0" + " - " + CStr(FN(intTramo_1,0))
						intTramoFinal6 = intTramo_1

						ElseIf intOrdenInforme = "2" then
						intTramoFinal6 = intTramo_2

						strTramoDMD= CStr(FN(Clng(intTramo_1) + 1,0)) + " - " + CStr(FN(intTramo_2,0))

						ElseIf intOrdenInforme = "3" then
						intTramoFinal6 = intTramo_3

						strTramoDMD= CStr(FN(Clng(intTramo_2) + 1,0)) + " - " + CStr(FN(intTramo_3,0))

						ElseIf intOrdenInforme = "4" then
						intTramoFinal6 = intTramo_4

						strTramoDMD= CStr(FN(Clng(intTramo_3) + 1,0)) + " - " + CStr(FN(intTramo_4,0))

						Else

						strTramoDMD= "MAYOR A " + CStr(FN(Clng(intTramoFinal6) + 1,0))

						End If

						strTramoDMDAnt= FN(RsInf4("TRAMO_DM_CASO") + 1,0)
						intTotalCasosTDMD= RsInf4("TOTAL_CASOS")
						intMinCasoTDMD= RsInf4("MIN_CAPITAL_CASO")
						intMaxCasoTDMD= RsInf4("MAX_CAPITAL_CASO")
						intTotalDocTDMD= RsInf4("TOTAL_DOC")
						intTotalMontoTDMD= RsInf4("MONTO_TOTAL")
						intMinDocTDMD= RsInf4("MIN_CAPITAL_DOC")
						intMaxDocTDMD= RsInf4("MAX_CAPITAL_DOC")

						intTTCasosTDMD= intTTCasosTDMD + intTotalCasosTDMD
						intTTDocTDMD= intTTDocTDMD + intTotalDocTDMD
						intTTMontoTDMD= intTTMontoTDMD + intTotalMontoTDMD

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strTramoDMD%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosTDMD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocTDMD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoTDMD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocTDMD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocTDMD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoTDMD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoTDMD,0)%>
							</TD>
						</tr>

						<%RsInf4.movenext
						loop
					Else%>
						<TD Height="20" bgcolor="#ff6666" Colspan="8" align="CENTER">
							<B>SELECCIONE TRAMOS PARA VISUALIZAR INFORME<B>
						</TD>
					<%end if
					RsInf4.close
					set RsInf4=nothing

	CerrarSCG()%>

				</tbody>
				<thead>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosTDMD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocTDMD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoTDMD,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>

<%End if

If strFiltroInforme = "0" and intTipoInforme = 3 then %>

<br>

<%End If

If (strFiltroInforme = "0" or strFiltroInforme = "TRAMO_DA_DOC") and intTipoInforme = 3 then%>

	<table width="100%" border="0" valign="top" align="center">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado">
					<thead>
				   	<tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="7" class="subtitulo_informe">
							> Tramo Día Asignación Doc (Segmenta según los días de asignación del documento)
						</TD>
						<TD align="left">
						  </acronym>&nbsp;&nbsp;<acronym title="FICHA DEUDOR">
						  <input name="fi_"  class="fondo_boton_100" type="button" onClick="window.navigate('man_carga_tramos_informes.asp?intTipoInforme=3&intCliente=<%=strCodCliente%>&intTipoTramo=7');" value="Auditar Tramos">
						  </acronym>
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="16%" ALIGN="CENTER">
							Tramo Dias Asignación Doc
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Casos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Documentos
						</TD>
						<TD width="12%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Documento máximo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD colspan="1" width="12%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
				</thead>
				<tbody>

<%		AbrirSCG()

					strSql = "SELECT * FROM TRAMOS_DEUDA"
					strSql = strSql & " WHERE COD_CLIENTE= '" & strCodCliente & "' AND TIPO_TRAMOS = 7"
					strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

					'Response.write "strSql = " & strSql
						set rsDet=Conn.execute(strSql)

						if not rsDet.eof then
							intReg = 0
							do while not rsDet.eof

							intOrdenTramo=rsDet("ORDEN_TRAMO")
							intTramo=rsDet("TRAMO")

								If intOrdenTramo = 1 then
								intTramo_1= intTramo
								End If

								If intOrdenTramo = 2 then
								intTramo_2= intTramo
								End If

								If intOrdenTramo = 3 then
								intTramo_3= intTramo
								End If

								If intOrdenTramo = 4 then
								intTramo_4= intTramo
								End If

							rsDet.movenext
							loop
						end if
						rsDet.close
						set rsDet=nothing

					'Response.write "<br>intTramo_1=" & intTramo_1
					'Response.write "<br>intTramo_2=" & intTramo_2
					'Response.write "<br>intTramo_3=" & intTramo_3
					'Response.write "<br>intTramo_4=" & intTramo_4

					strSql = "SELECT SUBSTRING(TRAMO_DEUDA,3,20) AS TRAMO_DA_CASO,SUBSTRING(TRAMO_DEUDA,1,1) AS ORDEN_TRAMO_DEUDA,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS,"
					strSql= strSql & " SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO,"
					strSql= strSql & " MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"

					strSql= strSql & " FROM (SELECT (CASE WHEN MAX(DATEDIFF(DAY,ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION),GETDATE())) <= TR1.TRAMO"
					strSql= strSql & " THEN '1-' + CAST(TR1.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION),GETDATE())) <= TR2.TRAMO "
					strSql= strSql & " THEN '2-' + CAST(TR2.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION),GETDATE())) <= TR3.TRAMO"
					strSql= strSql & " THEN '3-' + CAST(TR3.TRAMO AS VARCHAR(10))"
					strSql= strSql & " WHEN MAX(DATEDIFF(DAY,ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION),GETDATE())) <= TR4.TRAMO"
					strSql= strSql & " THEN '4-' + CAST(TR4.TRAMO AS VARCHAR(10)) ELSE '5-' + CAST(TR4.TRAMO AS VARCHAR(10)) END) AS TRAMO_DEUDA,"
					strSql= strSql & " C.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"

					strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"
					strSql= strSql & " INNER JOIN CLIENTE CL ON C.COD_CLIENTE = CL.COD_CLIENTE"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR1 ON CL.COD_CLIENTE=TR1.COD_CLIENTE AND TR1.ORDEN_TRAMO=1 AND TR1.TIPO_TRAMOS=7"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR2 ON CL.COD_CLIENTE=TR2.COD_CLIENTE AND TR2.ORDEN_TRAMO=2 AND TR2.TIPO_TRAMOS=7"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR3 ON CL.COD_CLIENTE=TR3.COD_CLIENTE AND TR3.ORDEN_TRAMO=3 AND TR3.TIPO_TRAMOS=7"
					strSql= strSql & " INNER JOIN TRAMOS_DEUDA TR4 ON CL.COD_CLIENTE=TR4.COD_CLIENTE AND TR4.ORDEN_TRAMO=4 AND TR4.TIPO_TRAMOS=7"

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if

					strSql= strSql & " GROUP BY C.RUT_DEUDOR,C.ID_CUOTA,TR1.TRAMO,TR2.TRAMO,TR3.TRAMO,TR4.TRAMO) AS PP GROUP BY TRAMO_DEUDA ORDER BY PP.TRAMO_DEUDA ASC"

					'Response.write "<br>strSql=" & strSql

					set RsInf4=Conn.execute(strSql)

					intTTCasosTRAD=0
					intTTDocTRAD=0
					intTTMontoTRAD=0

					if not RsInf4.eof then
						do until RsInf4.eof

						intOrdenInforme = RsInf4("ORDEN_TRAMO_DEUDA")

						'Response.write "<br>intOrdenInforme=" & intOrdenInforme

						If intOrdenInforme = "1" then

						strTramoDAD= "0" + " - " + CStr(FN(intTramo_1,0))
						intTramoFinal7 = intTramo_1

						ElseIf intOrdenInforme = "2" then
						intTramoFinal7 = intTramo_2

						strTramoDAD= CStr(FN(Clng(intTramo_1) + 1,0)) + " - " + CStr(FN(intTramo_2,0))

						ElseIf intOrdenInforme = "3" then
						intTramoFinal7 = intTramo_3

						strTramoDAD= CStr(FN(Clng(intTramo_2) + 1,0)) + " - " + CStr(FN(intTramo_3,0))

						ElseIf intOrdenInforme = "4" then
						intTramoFinal7 = intTramo_4

						strTramoDAD= CStr(FN(Clng(intTramo_3) + 1,0)) + " - " + CStr(FN(intTramo_4,0))

						Else

						strTramoDAD= "MAYOR A " + CStr(FN(Clng(intTramoFinal7) + 1,0))

						End If

						strTramoDMDAnt= FN(RsInf4("TRAMO_DA_CASO") + 1,0)
						intTotalCasosTDAD= RsInf4("TOTAL_CASOS")
						intMinCasoTDAD= RsInf4("MIN_CAPITAL_CASO")
						intMaxCasoTDAD= RsInf4("MAX_CAPITAL_CASO")
						intTotalDocTDAD= RsInf4("TOTAL_DOC")
						intTotalMontoTDAD= RsInf4("MONTO_TOTAL")
						intMinDocTDAD= RsInf4("MIN_CAPITAL_DOC")
						intMaxDocTDAD= RsInf4("MAX_CAPITAL_DOC")

						intTTCasosTDAD= intTTCasosTDAD + intTotalCasosTDAD
						intTTDocTDAD= intTTDocTDAD + intTotalDocTDAD
						intTTMontoTDAD= intTTMontoTDAD + intTotalMontoTDAD

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">
								<%=strTramoDAD%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasosTDAD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDocTDAD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMontoTDAD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDocTDAD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDocTDAD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCasoTDAD,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCasoTDAD,0)%>
							</TD>
						</tr>

						<%RsInf4.movenext
						loop
					Else%>
						<TD Height="20" bgcolor="#ff6666" Colspan="8" align="CENTER">
							<B>SELECCIONE TRAMOS PARA VISUALIZAR INFORME<B>
						</TD>
					<%end if
					RsInf4.close
					set RsInf4=nothing

	CerrarSCG()%>
				</tbody>
				<thead>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasosTDAD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTDocTDAD,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMontoTDAD,0)%>
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
						<TD >
							&nbsp;
						</TD>
						<TD>
							&nbsp;
						</TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>

<%End if%>
</body>
</html>

<script language="JavaScript1.2">

function envia(){
		//datos.action='cargando.asp';
		datos.action='informe_cartera.asp?intTipoInforme=<%=intTipoInforme%>';
		datos.submit();
}

function tipoinforme()
{
	datos.action='informe_cartera.asp';
	datos.submit();
}

function refrescar(){
		if (datos.CB_CLIENTE.value=='0'){
			alert('DEBE SELECCIONAR UN CLIENTE');
		}else
		{
		datos.action='informe_cartera.asp';
		datos.submit();
	}
}

function CargaUsuarios(subCat)
{
	//alert(subCat);

	var comboBox = document.getElementById('CB_EJECUTIVO');
	comboBox.options.length = 0;

		if (subCat=='INTERNA') {
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;
			<%

			AbrirSCG2()

			strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = '" & strCodCliente & "'"

			strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
			strSql = strSql & " AND U.PERFIL_EMP=1"

			'Response.write "<br>strSql=" & strSql

			set rsUsuario=Conn2.execute(strSql)
			If Not rsUsuario.Eof Then
				Do While Not rsUsuario.Eof
					%>
						var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					rsUsuario.movenext
				Loop
			Else
			%>
				var newOption = new Option('SIN USUARIO', '');
				comboBox.options[comboBox.options.length] = newOption;
			<%
			End if
			CerrarSCG2()
			%>
		}

		else if ((subCat=='EXTERNA') && (<%=intVerEjecutivos%>=='1')) {
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;
			<%

			AbrirSCG2()

			strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = '" & strCodCliente & "'"

			strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
			strSql = strSql & " AND U.PERFIL_EMP=0"

			'Response.write "<br>strSql=" & strSql

			set rsUsuario=Conn2.execute(strSql)
			If Not rsUsuario.Eof Then
				Do While Not rsUsuario.Eof
					%>
						var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					rsUsuario.movenext
				Loop
			Else
			%>
				var newOption = new Option('SIN USUARIO', '');
				comboBox.options[comboBox.options.length] = newOption;
			<%
			End if
			CerrarSCG2()
			%>
		}
		else if ((subCat=='EXTERNA') && (<%=intVerEjecutivos%>=='0')) {
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;

		}
		else {
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;
			<%

			AbrirSCG2()

			strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = '" & strCodCliente & "'"

			strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"

			If intVerEjecutivos = "0" then
			strSql = strSql & " AND U.PERFIL_EMP=1"
			end If

			set rsUsuario=Conn2.execute(strSql)
			If Not rsUsuario.Eof Then
				Do While Not rsUsuario.Eof
					%>
						var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					rsUsuario.movenext
				Loop
			Else
			%>
				var newOption = new Option('SIN USUARIO', '');
				comboBox.options[comboBox.options.length] = newOption;
			<%
			End if
			CerrarSCG2()
			%>
		}

}

function InicializaInforme()
{
		var comboBox = document.getElementById('CB_EJECUTIVO');
		comboBox.options.length = 0;
		var newOption = new Option('TODOS','');
		comboBox.options[comboBox.options.length] = newOption;
}

<%If sinCbUsario = "0" then%>
CargaUsuarios('<%=strCobranza%>');
<%End If%>

<%If strEjeAsig <> "" then%>
datos.CB_EJECUTIVO.value='<%=strEjeAsig%>';
<%End If%>

</script>