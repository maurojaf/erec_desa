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
	<!--#include file="../lib/comunes/js_css/top_tooltip.inc" -->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

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
<title>INFORME ASIGNACION COBRADORES</title>

</head>
<body>
<form name="datos" method="post">
<div class="titulo_informe">INFORME ASIGNACIÓN COBRADORES</div>
 <br>
<table width="90%" border="0" align="center" class="estilo_columnas">
	<thead>
	<tr >
		<td height="20">COBRANZA</td>
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
			<select name="CB_COBRANZA"  id="CB_COBRANZA"  <%If sinCbUsario = "0" then%> onChange="CargaUsuarios(CB_COBRANZA.value);" <%End If%> >

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
			<select name="CB_ETAPACOB">
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
			<select name="CB_CAMPANA">
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
			<select name="CB_EJECUTIVO" id="CB_EJECUTIVO">
			</select>
		</td>
	<% End If %>

		<td ALIGN="CENTER">
			<input type="button" name="Submit" class="fondo_boton_100" value="Ver" onClick="envia();">
		</td>
	</tr>

	<tr class="subtitulo_informe">
		<td Colspan="6"><br>> TIPO INFORME</td>
	</tr>
	
	<tr>
		<td Colspan="6" align="left">
				  <input name="fi_" style="width:120px" class="fondo_boton_100 <%=strColor1%>" type="button" onClick="window.navigate('Informe_asignacion_cobradores.asp?intTipoInforme=1');"  value="GENERAL">
					&nbsp; &nbsp; &nbsp;
				  <input name="fi_" style="width:120px" class="fondo_boton_100 <%=strColor2%>" type="button" onClick="window.navigate('Informe_asignacion_cobradores.asp?intTipoInforme=2');" value="SEDE">
				  
		</td>			
	</tr>
	
</table>
</form>

<%If (strFiltroInforme = "0" or strFiltroInforme = "TIPO_DOC") and intTipoInforme = 1 then%>

	<table width="90%" border="0"  ALIGN="CENTER">
		<tr>
			<td>
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado" style="width:100%;">
				<thead>	
				   <tr HEIGHT="20">
						<TD Colspan="11" class="subtitulo_informe">
							> Asiganción por ejecutivo
						</TD>
					</tr>
					<tr >
						<TD width="10%" ALIGN="CENTER">
							Ejecutivo
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Casos
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Documentos
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Monto
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Rut / Doc
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Prom. Doc.
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Prom. Caso
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Doc. mínimo
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Doc. míximo
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Caso mínimo
						</TD>
						<TD width="9%" ALIGN="CENTER">
							Caso máximo
						</TD>
					</tr>
				</thead>
				<tbody>
	<%AbrirSCG()
					strSql= "SELECT (CASE WHEN ISNULL(U.LOGIN,'SINASIG')='SINASIG' THEN 'SIN ASIGNAR' ELSE U.LOGIN END) AS USUARIO_ASIG,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS, SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, "
					strSql= strSql & " MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO, MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC, MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC,ISNULL((U.NOMBRES_USUARIO+' '+isnull(U.APELLIDO_PATERNO,'')+' '+isnull(U.APELLIDO_MATERNO,'')),'SIN ASIGNAR') AS NOMBRE_COMPLETO "

					strSql= strSql & " FROM (SELECT D.USUARIO_ASIG AS USUARIO_ASIG,D.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"
					strSql= strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO "
					strSql= strSql & " 				INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR "			

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if
					
					strSql= strSql & " GROUP BY D.RUT_DEUDOR,D.USUARIO_ASIG) PP LEFT JOIN USUARIO U ON PP.USUARIO_ASIG = U.ID_USUARIO " 
					strSql= strSql & " GROUP BY U.LOGIN,U.NOMBRES_USUARIO,U.APELLIDO_PATERNO, U.APELLIDO_MATERNO"
					strSql= strSql & " ORDER BY U.LOGIN ASC"
					
				

					set RsInf=Conn.execute(strSql)

					intTTCasosTD=0
					intTTDocTD=0
					intTTMontoTD=0

					if not RsInf.eof then
						do until RsInf.eof

						strNomUsuario = RsInf("USUARIO_ASIG")
						intTotalCasos= RsInf("TOTAL_CASOS")
						intMinCaso= RsInf("MIN_CAPITAL_CASO")
						intMaxCaso= RsInf("MAX_CAPITAL_CASO")
						intTotalDoc= RsInf("TOTAL_DOC")
						intTotalMonto= RsInf("MONTO_TOTAL")
						intMinDoc= RsInf("MIN_CAPITAL_DOC")
						intMaxDoc= RsInf("MAX_CAPITAL_DOC")

						intTTCasos= intTTCasos + intTotalCasos
						intTtDoc= intTtDoc + intTotalDoc
						intTTMonto= intTTMonto + intTotalMonto
						
						intDocRut= intTotalDoc/intTotalCasos
						
						intPromDoc= intTotalMonto/intTotalDoc
						
						intPromRut= intTotalMonto/intTotalCasos

						%>
						<tr >
							<TD ALIGN="left" title="<%=RsInf("NOMBRE_COMPLETO")%>">
								<%=strNomUsuario%>
								
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasos,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMonto,0)%>
							</TD>
							<TD ALIGN="CENTER">
								<%=FN(intDocRut,1)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intPromDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intPromRut,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCaso,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCaso,0)%>
							</TD>
						</tr>

						<%RsInf.movenext
						loop
					end if
					RsInf.close
					set RsInf=nothing
					
					intDocRutGeneral = intTtDoc / intTTCasos
					
					intPromDocGeneral = intTTMonto / intTtDoc
					
					intPromRutGeneral = intTTMonto / intTTCasos

	CerrarSCG()%>
				</tbody>
				<thead>
					<tr class="totales">
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasos,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTtDoc,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMonto,0)%>
						</TD>
						<TD ALIGN="CENTER">
							<%=FN(intDocRutGeneral,1)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intPromDocGeneral,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intPromRutGeneral,0)%>
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

If (strFiltroInforme = "0" or strFiltroInforme = "TRAMO_DD_CASO") and intTipoInforme = 2 then%>

	<table width="90%" border="0" valign="top" ALIGN="CENTER">
		<tr>
			<td>
				<table border="0" ALIGN="CENTER" class="intercalado" style="width:100%;">

	<%AbrirSCG()
					strSql= "SELECT ROW_NUMBER() OVER(PARTITION BY SUCURSAL ORDER BY SUCURSAL DESC) AS CORR,USUARIO_ASIG,SUCURSAL,NOMBRE_COMPLETO,COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS, SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL, "
					strSql= strSql & " MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO, MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC, MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC "

					strSql= strSql & " FROM (SELECT (CASE WHEN ISNULL(U.LOGIN,'SINASIG')='SINASIG' THEN 'SIN ASIGNAR' ELSE U.LOGIN END) AS USUARIO_ASIG,D.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,"
					strSql= strSql & " MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC,ISNULL((U.NOMBRES_USUARIO+' '+isnull(U.APELLIDO_PATERNO,'')+' '+isnull(U.APELLIDO_MATERNO,'')),'SIN ASIGNAR') AS NOMBRE_COMPLETO,MAX(UPPER(substring(SUCURSAL,1,1))+LOWER(substring(SUCURSAL,2,20))) AS SUCURSAL "
					strSql= strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO "
					strSql= strSql & " 				INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR "
					strSql= strSql & " 				LEFT JOIN USUARIO U ON D.USUARIO_ASIG = U.ID_USUARIO "

					strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND C.CUSTODIO IS NULL"
					End if
					
					strSql= strSql & " GROUP BY D.RUT_DEUDOR,U.LOGIN,U.NOMBRES_USUARIO,U.APELLIDO_PATERNO,U.APELLIDO_MATERNO) AS PP" 
					strSql= strSql & " GROUP BY USUARIO_ASIG,NOMBRE_COMPLETO,SUCURSAL" 
					strSql= strSql & " ORDER BY SUCURSAL ASC"
					
					'Response.write "<br>strSql=" & strSql

					set RsInf=Conn.execute(strSql)

					intTTCasosTD	=0
					intTTDocTD		=0
					intTTMontoTD	=0
					intCorrAnt		=0
					
					if not RsInf.eof then
						do until RsInf.eof

						intCorr= RsInf("CORR")
				
						If Cint(intCorr) <= Cint(intCorrAnt) Then
						
						intDocRutGeneral = intTtDoc / intTTCasos
						
						intPromDocGeneral = intTTMonto / intTtDoc
						
						intPromRutGeneral = intTTMonto / intTTCasos%>
						<thead>
						<tr >
							<TD ALIGN="CENTER">
								Totales
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTTCasos,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTtDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTTMonto,0)%>
							</TD>
							<TD ALIGN="CENTER">
								<%=FN(intDocRutGeneral,1)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intPromDocGeneral,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intPromRutGeneral,0)%>
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
						<%
						
						intTTCasos=0
						intTtDoc=0
						intTTMonto=0
						
						End If
						
						strNomUsuario = RsInf("USUARIO_ASIG")
						intTotalCasos= RsInf("TOTAL_CASOS")
						intMinCaso= RsInf("MIN_CAPITAL_CASO")
						intMaxCaso= RsInf("MAX_CAPITAL_CASO")
						intTotalDoc= RsInf("TOTAL_DOC")
						intTotalMonto= RsInf("MONTO_TOTAL")
						intMinDoc= RsInf("MIN_CAPITAL_DOC")
						intMaxDoc= RsInf("MAX_CAPITAL_DOC")
						strSucursal= RsInf("SUCURSAL")

						intTTCasos= intTTCasos + intTotalCasos
						intTtDoc= intTtDoc + intTotalDoc
						intTTMonto= intTTMonto + intTotalMonto
						
						intDocRut= intTotalDoc/intTotalCasos
						
						intPromDoc= intTotalMonto/intTotalDoc
						
						intPromRut= intTotalMonto/intTotalCasos
						
						'Response.write "<br>intCorr=" & intCorr
						'Response.write "<br>intCorrAnt=" & intCorrAnt
						
						If intCorr = "1" Then%>
							<thead>
							<tr HEIGHT="20" VALIGN="MIDDLE" >
								<TD Colspan="11" class="subtitulo_informe">
									> Asiganción sede <%=strSucursal%>
								</TD>
							</tr>
							<tr >
								<TD width="10%" ALIGN="CENTER">
									Ejecutivo
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Casos
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Documentos
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Monto
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Rut / Doc
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Prom. Doc.
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Prom. Caso
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Doc. mínimo
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Doc. míximo
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Caso mínimo
								</TD>
								<TD width="9%" ALIGN="CENTER">
									Caso máximo
								</TD>
							</tr>				
							</thead>
						<%End If%>
						<tbody>
						<tr >
							<TD ALIGN="left" title="<%=RsInf("NOMBRE_COMPLETO")%>">
								<%=strNomUsuario%>
								
							<TD ALIGN="RIGHT">
								<%=FN(intTotalCasos,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intTotalDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intTotalMonto,0)%>
							</TD>
							<TD ALIGN="CENTER">
								<%=FN(intDocRut,1)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intPromDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								<%=FN(intPromRut,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxDoc,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMinCaso,0)%>
							</TD>
							<TD ALIGN="RIGHT">
								$&nbsp;<%=FN(intMaxCaso,0)%>
							</TD>
						</tr>
						</tbody>
						<%
						intCorrAnt= RsInf("CORR")
						
						RsInf.movenext
						loop
					end if
					RsInf.close
					set RsInf=nothing

	CerrarSCG()%>
					<thead>
					<tr >
						<TD ALIGN="CENTER">
							Totales
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTCasos,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTtDoc,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTTMonto,0)%>
						</TD>
						<TD ALIGN="CENTER">
							<%=FN(intDocRutGeneral,1)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intPromDocGeneral,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intPromRutGeneral,0)%>
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

<%End If%>
</body>
</html>


<script type="text/javascript">
	$(document).ready(function(){
		$(document).tooltip();
	})
</script>

<script language="JavaScript1.2">

function envia(){
		//datos.action='cargando.asp';
		datos.action='Informe_asignacion_cobradores.asp?intTipoInforme=<%=intTipoInforme%>';
		datos.submit();
}

function tipoinforme()
{
	datos.action='Informe_asignacion_cobradores.asp';
	datos.submit();
}

function refrescar(){
		if (datos.CB_CLIENTE.value=='0'){
			alert('DEBE SELECCIONAR UN CLIENTE');
		}else
		{
		datos.action='Informe_asignacion_cobradores.asp';
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