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

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

	<script language="JavaScript">
	$(document).ready(function(){

		$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	 
	})

	function ventanaSecundaria (URL){
		window.open(URL,"DETALLE","width=200, height=200, scrollbars=no, menubar=no, location=no, resizable=no")
	}

	</script>

<%

	Response.CodePage=65001
	Response.charset ="utf-8"
	
	strFechaInicio= request("inicio")
	strFechaTermino= request("termino")
	intEtapaCobranza=Request("CB_ETAPACOB")
	strEjeAsig = request("CB_EJECUTIVO")
	resp=request("resp")

	If resp="si" then
	session("FtroIC_strFechaInicio") = strFechaInicio
	session("FtroIC_strFechaTermino") = strFechaTermino
	End If
	
	'Response.write "strSql = " & session("FtroIC_strFechaInicio")
	'Response.write "<br>resp = " & resp
	
	If strFechaInicio = "" Then strFechaInicio = session("FtroIC_strFechaInicio")
	If strFechaTermino = "" Then strFechaTermino = session("FtroIC_strFechaTermino")

	abrirscg()
		If Trim(strFechaInicio) = "" Then
			strFechaInicio = TraeFechaActual(Conn)
			strFechaInicio = "01/" & Mid(TraeFechaActual(Conn),4,10)
		End If

		If Trim(strFechaTermino) = "" Then
			strFechaTermino = TraeFechaActual(Conn)
		End If

	If Request("CB_TIPO_INF") <> "" then strFiltroInforme=Request("CB_TIPO_INF") Else strFiltroInforme=0 End If
	
	If Request("intTipoInforme") <> "" then intTipoInforme=Request("intTipoInforme") else intTipoInforme = "1" End If

	AbrirSCG()

	strSql = " SELECT COD_CLIENTE FROM CLIENTE WHERE COD_CLIENTE IN"
	strSql= strSql & " (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"

	set rsTemp= Conn.execute(strSql)
	'Response.write "strSql = " & strSql

	strTClienteUsu = "0"

	if not rsTemp.eof then
			do while not rsTemp.eof

			strClientesUsu = rsTemp("COD_CLIENTE")

			strTClienteUsu = strTClienteUsu + "," + strClientesUsu

			rsTemp.movenext
			Loop
		rsTemp.close
		set rsTemp=nothing
	End If

	CerrarSCG()
	
	If Request("CB_CLIENTE") = "" then
		strCodCliente = session("ses_codcli")
	Elseif Request("CB_CLIENTE") <> "0" then
		strCodCliente  = Request("CB_CLIENTE")
	Else
		strCodCliente = mid(strTClienteUsu,3,len(strTClienteUsu))
	End If

	'Response.write "<br>strCodCliente = " & strCodCliente

'--Calcula Objetos relacionados al tipo de cobranza (Interna, Externa), CB_COBRANZA Y CB_EJECUTIVO--'

strCobranza = Request("CB_COBRANZA")

abrirscg()

		strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
		strSql = strSql & " FROM CLIENTE CL"
		strSql = strSql & " WHERE CL.COD_CLIENTE IN ('" & strCodCliente & "')"
	
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
	intVerCobExt = "0"

End If

If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

	sinCbUsario="0"

End If

'---Fin codigo tipo de cobranza---'

%>

<style type="text/css">
.uno a {
	text-align:center;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #FFFFFF;
}
.uno a:hover {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	text-decoration: none;
	color: #FFFFFF;
}
</style>
</head>
<body>
<form name="datos" method="post">

<div class="titulo_informe">INFORME COMPROMISOS</div>
<BR>
<table width="90%" align="CENTER" border="0">
<tr>
	<td>

	<table width="100%" border="0" class="estilo_columnas">
	<thead>
		<tr height="20" >
			<td>CLIENTE</td>
			<td>COBRANZA</td>
			<td>ETAPA COBRANZA</td>
			<td>COMPROMISO DESDE</td>
			<td>COMPROMISO HASTA</td>
			<td>CAMPAÑA</td>

			<% If sinCbUsario = "0" Then %>
				<td>EJECUTIVO</td>
			<%Else%>
				<td>&nbsp;</td>
			<%End If%>
			<td>&nbsp;</td>
			
		</tr>
	</thead>
		<tr>
			<td>
				<SELECT NAME="CB_CLIENTE" id="CB_CLIENTE" onChange="envia();">
					<option value="0">TODOS</option>
					<%
					AbrirSCG()

						ssql="SELECT COD_CLIENTE,RAZON_SOCIAL, NOMBRE_FANTASIA FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
								<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCodCliente)=Trim(rsTemp("COD_CLIENTE")) then response.Write("Selected") End If%>><%=rsTemp("NOMBRE_FANTASIA")%></option>
									<%
								rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing

					CerrarSCG()
					%>
				</SELECT>
			</td>

			<td>
				<select name="CB_COBRANZA" <%If sinCbUsario = "0" then%> onChange="CargaUsuarios(this.value,CB_CLIENTE.value);" <%End If%> >
				
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
				<input name="inicio" readonly="true" type="text" id="inicio" value="<%=strFechaInicio%>" size="10" maxlength="10">
			</td>
			<td>
				<input name="termino" readonly="true" type="text" id="termino" value="<%=strFechaTermino%>" size="10" maxlength="10">
			</td>
			<td>
				<select name="CB_CAMPANA" >
					<option value="">TODAS</option>
					<%
					AbrirSCG()
						strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE = '" & strCodCliente & "'"
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
				<select name="CB_EJECUTIVO" id="CB_EJECUTIVO" >
				</select>
			</td>

		<% End If %>
		
			<td ALIGN="CENTER">
				<input type="button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();">
			</td>

		</tr>
    </table>

	<table width="100%" border="0">
		<tr>
		<TD>
		<table width="100%" border="0" ALIGN="CENTER" class="Estilo13">
		<tr >
			<TD class="subtitulo_informe">> Estado</TD>
			<td ALIGN="RIGHT" >

					  <input name="fi_" class="fondo_boton_100" id="fi_" style="font-size:11px;width:130px" type="button" onClick="cajas4();"  value="   Gestion   ">

					  <input name="fi_" class="fondo_boton_100" id="fi_" style="font-size:11px;width:130px" type="button" onClick="cajas5();" value="  Documentos">
					  
					  <input name="fi_" class="fondo_boton_100" id="fi_" style="font-size:11px;width:130px" type="button"  value="  Seguimiento">
					  
			</td>
		</tr>
		</table>
		</TD>
		</tr>
	</table>

	<div name="divGes" id="divGes" style="display:inline" >
	
	<table width="100%" border="0" valign="top" >
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado" style="width:100%;">
					<thead>
				   	<tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="8" class="subtitulo_informe">
							> Gestión
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="23%" ALIGN="CENTER">
							Estado Compromiso
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Gestiones
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Doc.
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Doc. Pagados
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Doc. Pednientes
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Monto Pagados
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Monto Pendiente
						</TD>
					</tr>
				</thead>
				<tbody>
	<%
		AbrirSCG()

				strSql = "SELECT ESTADO_GESTION_CP,SUM(TT_MONTO) AS TOTAL_MONTO,SUM(TT_MONTO_PAGADO) AS TOTAL_MONTO_PAGADO,SUM(TT_MONTO_ACTIVO) AS TOTAL_MONTO_ACTIVO,"
				strSql= strSql & " COUNT(*) AS TOTAL_GESTIONES,SUM(TT_DOC) AS TOTAL_DOCUMENTOS,SUM(DOC_PAGADOS) AS TOTAL_DOC_PAGADOS,SUM(TT_MONTO_ACTIVO) AS TOTAL_DOC_PEND"

				strSql= strSql & " 	FROM GESTIONES G	INNER JOIN DEUDOR D ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.COD_CLIENTE = D.COD_CLIENTE"
				strSql= strSql & " 						LEFT JOIN CAJA_FORMA_PAGO CFP ON G.FORMA_PAGO = CFP.ID_FORMA_PAGO"
				strSql= strSql & " 						LEFT JOIN USUARIO U ON G.ID_USUARIO = U.ID_USUARIO"
				strSql= strSql & " 						INNER JOIN"
				strSql= strSql & " (SELECT G.ID_GESTION,"
				strSql= strSql & " (CASE WHEN COUNT(C.ID_CUOTA) = SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END))"
				strSql= strSql & " 	  THEN 'COMPROMISOS CUMPLIDOS'"
				strSql= strSql & " 	  WHEN SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END)) > 0"
				strSql= strSql & " 	  THEN 'COMPROMISOS PARC. CUMPLIDOS'"
				strSql= strSql & " 	  ELSE 'COMPROMISOS NO CUMPLIDOS'"
 				strSql= strSql & " END) AS ESTADO_GESTION_CP,"
				strSql= strSql & " SUM(C.VALOR_CUOTA) AS TT_MONTO,"
				strSql= strSql & " COUNT(C.ID_CUOTA) AS TT_DOC,"
				strSql= strSql & " SUM((CASE WHEN ED.GRUPO='ACTIVOS' THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_ACTIVO,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_PAGADO,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='RETIROS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_RETIRADO,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='NO ASIGNABLES' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA  ELSE 0 END)) AS TT_MONTO_NO_ASIGNABLE,"

				strSql= strSql & " SUM((CASE WHEN ED.GRUPO='ACTIVOS' THEN 1 ELSE 0 END)) AS DOC_ACTIVOS,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END)) AS DOC_PAGADOS,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='RETIROS' AND G.ID_GESTION = C.ID_ULT_GEST_CP)  THEN 1 ELSE 0 END)) AS DOC_RETIRADOS,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='NO ASIGNABLES' AND G.ID_GESTION = C.ID_ULT_GEST_CP)  THEN 1 ELSE 0 END)) AS DOC_NO_ASIGNABLES,"

				strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST_TIT ELSE 0 END)) AS GEST_TIT,"
				strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST_GENERAL ELSE 0 END)) AS GEST_GENERAL,"
				strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST ELSE 0 END)) AS GEST_SCOR"

				strSql= strSql & " FROM GESTIONES G		INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION"
				strSql= strSql & " 						INNER JOIN CUOTA C ON GC.ID_CUOTA = C.ID_CUOTA"
				strSql= strSql & " 						INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA AND"
				strSql= strSql & " 																 G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA AND"
				strSql= strSql & " 																 G.COD_GESTION = GTG.COD_GESTION AND"
				strSql= strSql & " 																 G.COD_CLIENTE = GTG.COD_CLIENTE"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON G.COD_CATEGORIA = GTC.COD_CATEGORIA"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA AND"
				strSql= strSql & " 																	   G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA"

				strSql= strSql & " WHERE GTG.CATEGORIZACION IN (1,2) AND G.COD_CLIENTE IN (" & strCodCliente &")"

				If strEjeAsig <> "" then

				strSql = strSql & " AND  G.ID_USUARIO = " & strEjeAsig

				End If

				If strFechaInicio <> "" then

				strSql = strSql & " AND  G.FECHA_COMPROMISO > = '" & strFechaInicio & " 00:00:00'"

				End If

				If strFechaTermino <> "" then

				strSql = strSql & " AND G.FECHA_COMPROMISO < = '" & strFechaTermino & " 23:59:59'"

				End If

				If Trim(strCobranza) = "INTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
				End if

				If Trim(strCobranza) = "EXTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NULL"
				End if

				strSql= strSql & " GROUP BY G.ID_GESTION) AS PP ON G.ID_GESTION=PP.ID_GESTION"

				strSql= strSql & " GROUP BY ESTADO_GESTION_CP"

				'Response.write "<br>strSql=" & strSql

				set RsInf=Conn.execute(strSql)

				If not RsInf.eof then
					do until RsInf.eof

					strEstadoGestionCP=RsInf("ESTADO_GESTION_CP")

					If strEstadoGestionCP = "COMPROMISOS NO CUMPLIDOS" then

						intTTGestCNC = RsInf("TOTAL_GESTIONES")
						intTTDocCNC = RsInf("TOTAL_DOCUMENTOS")
						intTTMontoCNC = RsInf("TOTAL_MONTO")
						intTTDocPagadoCNC = RsInf("TOTAL_DOC_PAGADOS")
						intTTDocPendCNC = RsInf("TOTAL_DOC_PEND")
						intTTMontoPagadoCNC = RsInf("TOTAL_MONTO_PAGADO")
						intTTMontoPendCNC = RsInf("TOTAL_MONTO_ACTIVO")

					End If

					If strEstadoGestionCP = "COMPROMISOS PARC. CUMPLIDOS" then

						intTTGestCPC = RsInf("TOTAL_GESTIONES")
						intTTDocCPC  = RsInf("TOTAL_DOCUMENTOS")
						intTTMontoCPC  = RsInf("TOTAL_MONTO")
						intTTDocPagadoCPC  = RsInf("TOTAL_DOC_PAGADOS")
						intTTDocPendCPC  = RsInf("TOTAL_DOC_PEND")
						intTTMontoPagadoCPC  = RsInf("TOTAL_MONTO_PAGADO")
						intTTMontoPendCPC  = RsInf("TOTAL_MONTO_ACTIVO")

					End If

					If strEstadoGestionCP = "COMPROMISOS CUMPLIDOS" then

						intTTGestCC = RsInf("TOTAL_GESTIONES")
						intTTDocCC  = RsInf("TOTAL_DOCUMENTOS")
						intTTMontoCC  = RsInf("TOTAL_MONTO")
						intTTDocPagadoCC  = RsInf("TOTAL_DOC_PAGADOS")
						intTTDocPendCC  = RsInf("TOTAL_DOC_PEND")
						intTTMontoPagadoCC  = RsInf("TOTAL_MONTO_PAGADO")
						intTTMontoPendCC  = RsInf("TOTAL_MONTO_ACTIVO")

					End If

						RsInf.movenext
					loop
				end if
				RsInf.close
				set RsInf=nothing

				intTTGestC = intTTGestCNC + intTTGestCPC + intTTGestCC
				intTTDocC = intTTDocCNC + intTTDocCPC + intTTDocCC
				intTTMontoC = intTTMontoCNC + intTTMontoCPC + intTTMontoCC
				intTTDocPagadoC = intTTDocPagadoCNC + intTTDocPagadoCPC + intTTDocPagadoCC
				intTTDocPendC = intTTDocPendCNC + intTTDocPendCPC + intTTDocPendCC
				intTTMontoPagadoC = intTTMontoPagadoCNC + intTTMontoPagadoCPC + intTTMontoPagadoCC
				intTTMontoPendC = intTTMontoPendCNC + intTTMontoPendCPC + intTTMontoPendCC

	CerrarSCG()%>
						<TR bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">COMPROMISOS NO CUMPLIDOS</TD>
							
							<td align="right">
								<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=1&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>">
								<%=FN(intTTGestCNC,0)%>
								</A>
							</td>
							
							<TD ALIGN="RIGHT"><%=FN(intTTDocCNC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoCNC,0)%></TD>
							<TD ALIGN="RIGHT"><%=FN(intTTDocPagadoCNC,0)%></TD>
							<TD ALIGN="RIGHT"><%=FN(intTTDocPendCNC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPagadoCNC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPendCNC,0)%></TD>
						</TR>
						<TR bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">COMPROMISOS PARCIALMENTE CUMPLIDOS</TD>

							<td align="right">
								<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=2&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>">
								<%=FN(intTTGestCPC,0)%>
								</A>
							</td>
							
							<TD ALIGN="RIGHT"><%=FN(intTTDocCPC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoCPC,0)%></TD>
							<TD ALIGN="RIGHT"><%=FN(intTTDocPagadoCPC,0)%></TD>
							<TD ALIGN="RIGHT"><%=FN(intTTDocPendCPC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPagadoCPC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPendCPC,0)%></TD>
						</TR>
						<TR bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<TD ALIGN="left">COMPROMISOS CUMPLIDOS</TD>

							<td align="right">
								<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=3&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>">
								<%=FN(intTTGestCC,0)%>
								</A>
							</td>
							
							<TD ALIGN="RIGHT"><%=FN(intTTDocCC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoCC,0)%></TD>
							<TD ALIGN="RIGHT"><%=FN(intTTDocPagadoCC,0)%></TD>
							<TD ALIGN="RIGHT"><%=FN(intTTDocPendCC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPagadoCC,0)%></TD>
							<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPendCC,0)%></TD>
						</TR>
					</tbody>	
					<thead>
					<tr class="totales">
						<TD ALIGN="CENTER">Totales</TD>
						
						<td ALIGN="RIGHT" bordercolor="#999999">
							<div class="uno">
							<A HREF="Detalle_informe_compromisos.asp?strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>">
							<%=FN(intTTGestC,0)%>
							</A>
							<div>
						</td>
							
						<TD ALIGN="RIGHT"><%=FN(intTTDocC,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoC,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTDocPagadoC,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTDocPendC,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPagadoC,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPendC,0)%></TD>
					</tr>
				</thead>
				</table>
			</td>
		</tr>
	</table>
	
	</div>

	<div name="divDoc" id="divDoc" style="display:none" >
		
	<table width="100%" border="0" valign="top" >
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado" style="width:100%;" >
				<thead>
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="8" class="subtitulo_informe">
							> Documentos
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD ALIGN="CENTER" width="400">
							Estado Documentos
						</TD>
						<TD  ALIGN="CENTER">
							Total Gestiones
						</TD>
						<TD ALIGN="CENTER">
							Total Doc.
						</TD>
						<TD ALIGN="CENTER">
							Total Monto
						</TD>
					</tr>
				</thead>
				<tbody>
	<%
		AbrirSCG()
		
				strSql = "SELECT ESTADO_DOC,COUNT(DISTINCT ID_GESTION) AS TOTAL_GESTIONES,COUNT(ID_GESTION) AS TOTAL_DOCUMENTOS, SUM(PP.VALOR_CUOTA) AS TOTAL_MONTO"
				strSql= strSql & " FROM"
				strSql= strSql & " (SELECT" 
				strSql= strSql & " (CASE WHEN ED.GRUPO='PAGADOS' THEN '12-PAGADOS'"
				strSql= strSql & " 	 WHEN ED.GRUPO='RETIROS' THEN '11-RETIRO' "
				strSql= strSql & " 	 WHEN GTG2.CATEGORIZACION IN (5,6,17) THEN '03-EN NORMALIZACION'"
				strSql= strSql & " 	 WHEN GTG2.CATEGORIZACION IN (12) THEN '04-PAGO NO APLICADO'"
				strSql= strSql & " 	 WHEN GTG2.CATEGORIZACION IN (13) THEN '05-EN ESPERA DE APLICACION'"
				strSql= strSql & " 	 WHEN GTG2.CATEGORIZACION IN (11,15,14) THEN '06-INUBICABLE'"
				strSql= strSql & " 	 WHEN GTG2.CATEGORIZACION IN (3,4) THEN '07-REHUSA PAGO'"
				strSql= strSql & " 	 WHEN (G.ID_GESTION <> C.ID_ULT_GEST) AND GTG2.CATEGORIZACION IN (1,2) THEN '08-NUEVO COMPROMISO'"
				strSql= strSql & " ELSE '09-COMPROMISO ROTO' END) AS ESTADO_DOC,"
				strSql= strSql & " G.ID_GESTION,C.VALOR_CUOTA,"
				strSql= strSql & " (CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST_TIT ELSE 0 END)AS GEST_TIT, "
				strSql= strSql & " (CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST_GENERAL ELSE 0 END) AS GEST_GENERAL," 
				strSql= strSql & " (CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST ELSE 0 END) AS GEST_SCOR "
				strSql= strSql & " FROM GESTIONES G INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION" 
				strSql= strSql & " 				 INNER JOIN CUOTA C ON GC.ID_CUOTA = C.ID_CUOTA "
				strSql= strSql & " 				 INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO" 
				strSql= strSql & " 				 INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA" 
				strSql= strSql & " 							AND G.COD_GESTION = GTG.COD_GESTION AND G.COD_CLIENTE = GTG.COD_CLIENTE "
				strSql= strSql & " 				 INNER JOIN GESTIONES G2 ON C.ID_ULT_GEST = G2.ID_GESTION"
				strSql= strSql & " 				 INNER JOIN GESTIONES_TIPO_GESTION GTG2 ON G2.COD_CATEGORIA = GTG2.COD_CATEGORIA AND G2.COD_SUB_CATEGORIA = GTG2.COD_SUB_CATEGORIA "
				strSql= strSql & " 							AND G2.COD_GESTION = GTG2.COD_GESTION AND G2.COD_CLIENTE = GTG2.COD_CLIENTE "

				strSql= strSql & " WHERE GTG.CATEGORIZACION IN (1,2) AND G.COD_CLIENTE IN (" & strCodCliente &")"

				If strEjeAsig <> "" then

				strSql = strSql & " AND  G.ID_USUARIO = " & strEjeAsig

				End If

				If strFechaInicio <> "" then

				strSql = strSql & " AND  G.FECHA_COMPROMISO > = '" & strFechaInicio & " 00:00:00'"

				End If

				If strFechaTermino <> "" then

				strSql = strSql & " AND G.FECHA_COMPROMISO < = '" & strFechaTermino & " 23:59:59'"

				End If

				If Trim(strCobranza) = "INTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
				End if

				If Trim(strCobranza) = "EXTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NULL"
				End if

				strSql= strSql & " ) AS PP "

				strSql= strSql & " GROUP BY ESTADO_DOC"
				strSql= strSql & " ORDER BY ESTADO_DOC ASC"

				'Response.write "<br>strSql=" & strSql

				set RsInf=Conn.execute(strSql)


				If not RsInf.eof then
					do until RsInf.eof

							strEstadoGestionCD=RsInf("ESTADO_DOC")

							intTTGestCD = RsInf("TOTAL_GESTIONES")
							intTTDocCD = RsInf("TOTAL_DOCUMENTOS")
							intTTMontoCD = RsInf("TOTAL_MONTO")

							intTTGeneralGesCD = intTTGeneralGesCD + intTTGestCD
							intTTGeneralDocCD = intTTGeneralDocCD + intTTDocCD
							intTTGeneralMontoCD = intTTGeneralMontoCD + intTTMontoCD

							%>
							<TR bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
								<TD ALIGN="left"><%=Mid(strEstadoGestionCD,4,50)%></TD>
								<TD ALIGN="RIGHT"><%=FN(intTTGestCD,0)%></TD>
								<TD ALIGN="RIGHT"><%=FN(intTTDocCD,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoCD,0)%></TD>
							</TR>
							<%


						RsInf.movenext
					loop
				RsInf.close
				set RsInf=nothing
				
		CerrarSCG()%>
				</tbody>
				<thead>
					<tr class="totales">
						<TD ALIGN="CENTER">Totales</TD>

						<TD ALIGN="RIGHT"><%=FN(intTTGeneralGesCD,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocCD,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoCD,0)%></TD>

					</tr>

				<%Else%>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD HEIGHT="20" ALIGN="CENTER"><B>NO EXISTEN COMPROMISOS NO CUMPLIDOS<B></TD>
					</tr>
				<%End If%>
				</thead>
				</table>
			</td>
		</tr>
	</table>
	
	</div>
	
<br>

	<table width="100%" border="0">
		<tr>
		<TD>
		<table width="100%" border="0" ALIGN="CENTER" class="Estilo13">
		<tr >
			<TD class="subtitulo_informe">> Estado Gestion por Ejecutivo</TD>
			<td ALIGN="RIGHT" >

					  <input name="fi_" id="fi_" class="fondo_boton_100" style="font-size:11px;width:130px" type="button" onClick="cajas1();"  value="   No Cumplidos   ">
					  
					  <input name="fi_" id="fi_" class="fondo_boton_100" style="font-size:11px;width:130px" type="button" onClick="cajas2();" value="Parc. Cumplidos">

					  <input name="fi_" id="fi_" class="fondo_boton_100" style="font-size:11px;width:130px" type="button" onClick="cajas3();" value="Cumplidos">
			</td>
		</tr>
		</table>
		</TD>
		</tr>
	</table>

<%
				strSql = "SELECT U.LOGIN AS USUARIO_GES,G.ID_USUARIO,ESTADO_GESTION_CP,SUM(TT_MONTO) AS TOTAL_MONTO,SUM(TT_MONTO_PAGADO) AS TOTAL_MONTO_PAGADO,SUM(TT_MONTO_ACTIVO) AS TOTAL_MONTO_ACTIVO,"
				strSql= strSql & " COUNT(*) AS TOTAL_GESTIONES,SUM(TT_DOC) AS TOTAL_DOCUMENTOS,SUM(DOC_PAGADOS) AS TOTAL_DOC_PAGADOS,SUM(TT_MONTO_ACTIVO) AS TOTAL_DOC_PEND"

				strSql= strSql & " 	FROM GESTIONES G	INNER JOIN DEUDOR D ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.COD_CLIENTE = D.COD_CLIENTE"
				strSql= strSql & " 						LEFT JOIN CAJA_FORMA_PAGO CFP ON G.FORMA_PAGO = CFP.ID_FORMA_PAGO"
				strSql= strSql & " 						LEFT JOIN USUARIO U ON G.ID_USUARIO = U.ID_USUARIO"
				strSql= strSql & " 						INNER JOIN"
				strSql= strSql & " (SELECT G.ID_GESTION,"
				strSql= strSql & " (CASE WHEN COUNT(C.ID_CUOTA) = SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END))"
				strSql= strSql & " 	  THEN 'COMPROMISOS CUMPLIDOS'"
				strSql= strSql & " 	  WHEN SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END)) > 0"
				strSql= strSql & " 	  THEN 'COMPROMISOS PARC. CUMPLIDOS'"
				strSql= strSql & " 	  ELSE 'COMPROMISOS NO CUMPLIDOS'"
 				strSql= strSql & " END) AS ESTADO_GESTION_CP,"
				strSql= strSql & " SUM(C.VALOR_CUOTA) AS TT_MONTO,"
				strSql= strSql & " COUNT(C.ID_CUOTA) AS TT_DOC,"
				strSql= strSql & " SUM((CASE WHEN ED.GRUPO='ACTIVOS' THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_ACTIVO,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_PAGADO,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='RETIROS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_RETIRADO,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='NO ASIGNABLES' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA  ELSE 0 END)) AS TT_MONTO_NO_ASIGNABLE,"

				strSql= strSql & " SUM((CASE WHEN ED.GRUPO='ACTIVOS' THEN 1 ELSE 0 END)) AS DOC_ACTIVOS,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END)) AS DOC_PAGADOS,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='RETIROS' AND G.ID_GESTION = C.ID_ULT_GEST_CP)  THEN 1 ELSE 0 END)) AS DOC_RETIRADOS,"
				strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='NO ASIGNABLES' AND G.ID_GESTION = C.ID_ULT_GEST_CP)  THEN 1 ELSE 0 END)) AS DOC_NO_ASIGNABLES,"

				strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST_TIT ELSE 0 END)) AS GEST_TIT,"
				strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST_GENERAL ELSE 0 END)) AS GEST_GENERAL,"
				strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST ELSE 0 END)) AS GEST_SCOR"

				strSql= strSql & " FROM GESTIONES G		INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION"
				strSql= strSql & " 						INNER JOIN CUOTA C ON GC.ID_CUOTA = C.ID_CUOTA"
				strSql= strSql & " 						INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA AND"
				strSql= strSql & " 																 G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA AND"
				strSql= strSql & " 																 G.COD_GESTION = GTG.COD_GESTION AND"
				strSql= strSql & " 																 G.COD_CLIENTE = GTG.COD_CLIENTE"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON G.COD_CATEGORIA = GTC.COD_CATEGORIA"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA AND"
				strSql= strSql & " 																	   G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA"

				strSql= strSql & " WHERE GTG.CATEGORIZACION IN (1,2) AND G.COD_CLIENTE IN (" & strCodCliente &")"

				If strEjeAsig <> "" then

				strSql = strSql & " AND  G.ID_USUARIO = " & strEjeAsig

				End If

				If strFechaInicio <> "" then

				strSql = strSql & " AND  G.FECHA_COMPROMISO > = '" & strFechaInicio & " 00:00:00'"

				End If

				If strFechaTermino <> "" then

				strSql = strSql & " AND G.FECHA_COMPROMISO < = '" & strFechaTermino & " 23:59:59'"

				End If

				If Trim(strCobranza) = "INTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
				End if

				If Trim(strCobranza) = "EXTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NULL"
				End if

				strSql= strSql & " GROUP BY G.ID_GESTION) AS PP ON G.ID_GESTION=PP.ID_GESTION"

				strSql= strSql & " GROUP BY ESTADO_GESTION_CP,U.LOGIN,G.ID_USUARIO"

				strSql= strSql & " ORDER BY U.LOGIN"

				'Response.write "<br>strSql=" & strSql

		AbrirSCG()

				set RsCuenta=Conn.execute(strSql)

				intTTCasosCNC=0
				intTTCasosCPC=0
				intTTCasosCC=0

				If not RsCuenta.eof then
					do until RsCuenta.eof

						strEstadoGestionCP=RsCuenta("ESTADO_GESTION_CP")

						If strEstadoGestionCP = "COMPROMISOS NO CUMPLIDOS" then
							intTTCasosCNC = intTTCasosCNC + 1
						End If

						If strEstadoGestionCP = "COMPROMISOS PARC. CUMPLIDOS" then
						intTTCasosCPC = intTTCasosCPC + 1
						End If

						If strEstadoGestionCP = "COMPROMISOS CUMPLIDOS" then
						intTTCasosCC = intTTCasosCC + 1
						End If

						RsCuenta.movenext
					loop
				end if
				RsCuenta.close
				set RsCuenta=nothing

		CerrarSCG()%>

	<div name="divCNC" id="divCNC" style="display:inline" >

	<table width="100%" border="0" valign="top" >
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado" style="width:100%;">
				<thead>	
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="8" class="subtitulo_informe">
							> Compromiso NO Cumplidos
						</TD>
					</tr>

				<%If intTTCasosCNC > 0  then%>

					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="23%" ALIGN="CENTER">
							Estado Compromiso
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Gestiones
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Doc.
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Doc. Pagados
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Doc. Pednientes
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Monto Pagados
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Monto Pendiente
						</TD>
					</tr>
				</thead>
				<tbody>

<%			AbrirSCG()

				set RsInf=Conn.execute(strSql)

				intTTGeneralCNCE = 0
				intTTGeneralDocCNCE = 0
				intTTGeneralMontoCNCE = 0
				intTTGeneralDocPagadoCNCE = 0
				intTTGeneralDocPendCNCE = 0
				intTTGeneralMontoPagadoCNCE = 0
				intTTGeneralMontoPendCNCE = 0

				If not RsInf.eof then
					do until RsInf.eof

					strEstadoGestionCP=RsInf("ESTADO_GESTION_CP")

						If strEstadoGestionCP = "COMPROMISOS NO CUMPLIDOS" then

							strUsuarioGestCNC = RsInf("USUARIO_GES")
							intTTGestCNCE = RsInf("TOTAL_GESTIONES")
							intTTDocCNCE = RsInf("TOTAL_DOCUMENTOS")
							intTTMontoCNCE = RsInf("TOTAL_MONTO")
							intTTDocPagadoCNCE = RsInf("TOTAL_DOC_PAGADOS")
							intTTDocPendCNCE = RsInf("TOTAL_DOC_PEND")
							intTTMontoPagadoCNCE = RsInf("TOTAL_MONTO_PAGADO")
							intTTMontoPendCNCE = RsInf("TOTAL_MONTO_ACTIVO")

							intTTGeneralCNCE = intTTGeneralCNCE + intTTGestCNCE
							intTTGeneralDocCNCE = intTTGeneralDocCNCE + intTTDocCNCE
							intTTGeneralMontoCNCE = intTTGeneralMontoCNCE + intTTMontoCNCE
							intTTGeneralDocPagadoCNCE = intTTGeneralDocPagadoCNCE + intTTDocPagadoCNCE
							intTTGeneralDocPendCNCE = intTTGeneralDocPendCNCE + intTTDocPendCNCE
							intTTGeneralMontoPagadoCNCE = intTTGeneralMontoPagadoCNCE + intTTMontoPagadoCNCE
							intTTGeneralMontoPendCNCE = intTTGeneralMontoPendCNCE + intTTMontoPendCNCE


						%>
							<TR bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
								<TD ALIGN="left"><%=(strUsuarioGestCNC)%></TD>
								
								<td align="right">
									<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=1&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&strEjeAsig=<%=RsInf("ID_USUARIO")%>">
									<%=FN(intTTGestCNCE,0)%>
									</A>
								</td>
								
								<TD ALIGN="RIGHT"><%=FN(intTTDocCNCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoCNCE,0)%></TD>
								<TD ALIGN="RIGHT"><%=FN(intTTDocPagadoCNCE,0)%></TD>
								<TD ALIGN="RIGHT"><%=FN(intTTDocPendCNCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPagadoCNCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPendCNCE,0)%></TD>
							</TR>
							<%

						End If

						RsInf.movenext
					loop
				end if
				RsInf.close
				set RsInf=nothing

		CerrarSCG()%>
				</tbody>
				<thead>
					<tr class="totales">
						<TD ALIGN="CENTER">Totales</TD>

						<td ALIGN="RIGHT" bordercolor="#999999">
							<div class="uno">
							<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=1&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>">
							<%=FN(intTTGeneralCNCE,0)%>
							</A>
							<div>
						</td>

						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocCNCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoCNCE,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocPagadoCNCE,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocPendCNCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoPagadoCNCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoPendCNCE,0)%></TD>
					</tr>

				<%Else%>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD HEIGHT="20" ALIGN="CENTER"><B>NO EXISTEN COMPROMISOS NO CUMPLIDOS<B></TD>
					</tr>
				<%End If%>
			</thead>
				</table>
			</td>
		</tr>
	</table>

	</div>

	<div name="divCPC" id="divCPC" style="display:none" >

	<table width="100%" border="0" valign="top" >
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado" style="width:100%;">
				<thead>
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="8" class="subtitulo_informe">
							> Compromisos Parcialmente Cumplidos
						</TD>
					</tr>

					<%If intTTCasosCPC > 0  then%>

					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="23%" ALIGN="CENTER">
							Estado Compromiso
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Gestiones
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Doc.
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Doc. Pagados
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Doc. Pednientes
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Monto Pagados
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Monto Pendiente
						</TD>
					</tr>
				</thead>
				</tbody>
<%			AbrirSCG()

				set RsInf=Conn.execute(strSql)

				intTTGeneralCPCE = 0
				intTTGeneralDocCPCE = 0
				intTTGeneralMontoCPCE = 0
				intTTGeneralDocPagadoCPCE = 0
				intTTGeneralDocPendCPCE = 0
				intTTGeneralMontoPagadoCPCE = 0
				intTTGeneralMontoPendCPCE = 0

				If not RsInf.eof then
					do until RsInf.eof

					strEstadoGestionCP=RsInf("ESTADO_GESTION_CP")

						If strEstadoGestionCP = "COMPROMISOS PARC. CUMPLIDOS" then

							strUsuarioGestCPC = RsInf("USUARIO_GES")
							intTTGestCPCE = RsInf("TOTAL_GESTIONES")
							intTTDocCPCE = RsInf("TOTAL_DOCUMENTOS")
							intTTMontoCPCE = RsInf("TOTAL_MONTO")
							intTTDocPagadoCPCE = RsInf("TOTAL_DOC_PAGADOS")
							intTTDocPendCPCE = RsInf("TOTAL_DOC_PEND")
							intTTMontoPagadoCPCE = RsInf("TOTAL_MONTO_PAGADO")
							intTTMontoPendCPCE = RsInf("TOTAL_MONTO_ACTIVO")

							intTTGeneralCPCE = intTTGeneralCPCE + intTTGestCPCE
							intTTGeneralDocCPCE = intTTGeneralDocCPCE + intTTDocCPCE
							intTTGeneralMontoCPCE = intTTGeneralMontoCPCE + intTTMontoCPCE
							intTTGeneralDocPagadoCPCE = intTTGeneralDocPagadoCPCE + intTTDocPagadoCPCE
							intTTGeneralDocPendCPCE = intTTGeneralDocPendCPCE + intTTDocPendCPCE
							intTTGeneralMontoPagadoCPCE = intTTGeneralMontoPagadoCPCE + intTTMontoPagadoCPCE
							intTTGeneralMontoPendCPCE = intTTGeneralMontoPendCPCE + intTTMontoPendCPCE


						%>
							<TR bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
								<TD ALIGN="left"><%=(strUsuarioGestCPC)%></TD>

								<td align="right">
									<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=2&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&strEjeAsig=<%=RsInf("ID_USUARIO")%>">
									<%=FN(intTTGestCPCE,0)%>
									</A>
								</td>
								
								<TD ALIGN="RIGHT"><%=FN(intTTDocCPCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoCPCE,0)%></TD>
								<TD ALIGN="RIGHT"><%=FN(intTTDocPagadoCPCE,0)%></TD>
								<TD ALIGN="RIGHT"><%=FN(intTTDocPendCPCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPagadoCPCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPendCPCE,0)%></TD>
							</TR>
							<%

						End If

						RsInf.movenext
					loop
				end if
				RsInf.close
				set RsInf=nothing

		CerrarSCG()%>
				</tbody>
				<thead>
					<tr class="totales">
						<TD ALIGN="CENTER">Totales</TD>

						<td ALIGN="RIGHT" bordercolor="#999999">
							<div class="uno">
							<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=2&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>">
							<%=FN(intTTGeneralCPCE,0)%>
							</A>
							<div>
						</td>
						
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocCPCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoCPCE,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocPagadoCPCE,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocPendCPCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoPagadoCPCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoPendCPCE,0)%></TD>
					</tr>
				<%Else%>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD HEIGHT="20" ALIGN="CENTER"><B>NO EXISTEN COMPROMISOS PARCIALMENTE CUMPLIDOS<B></TD>
					</tr>
				<%End If%>
				</thead>
				</table>
			</td>
		</tr>
	</table>

	</div>

	<div name="divCC" id="divCC" style="display:none" >

	<table width="100%" border="0" valign="top" >
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado" style="width:100%;">
				<thead>
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="8" class="subtitulo_informe">
							> Compromiso Cumplidos
						</TD>
					</tr>

				<%If intTTCasosCC > 0  then%>

					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD width="23%" ALIGN="CENTER">
							Estado Compromiso
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Gestiones
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Doc.
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Total Monto
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Doc. Pagados
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Doc. Pednientes
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Monto Pagados
						</TD>
						<TD width="11%" ALIGN="CENTER">
							Monto Pendiente
						</TD>
					</tr>
					</thead>
					<tbody>
<%			AbrirSCG()

				set RsInf=Conn.execute(strSql)

				intTTGeneralCCE = 0
				intTTGeneralDocCCE = 0
				intTTGeneralMontoCCE = 0
				intTTGeneralDocPagadoCCE = 0
				intTTGeneralDocPendCCE = 0
				intTTGeneralMontoPagadoCCE = 0
				intTTGeneralMontoPendCCE = 0

				If not RsInf.eof then
					do until RsInf.eof

					strEstadoGestionCP=RsInf("ESTADO_GESTION_CP")

						If strEstadoGestionCP = "COMPROMISOS CUMPLIDOS" then

							strUsuarioGestCC = RsInf("USUARIO_GES")
							intTTGestCCE = RsInf("TOTAL_GESTIONES")
							intTTDocCCE = RsInf("TOTAL_DOCUMENTOS")
							intTTMontoCCE = RsInf("TOTAL_MONTO")
							intTTDocPagadoCCE = RsInf("TOTAL_DOC_PAGADOS")
							intTTDocPendCCE = RsInf("TOTAL_DOC_PEND")
							intTTMontoPagadoCCE = RsInf("TOTAL_MONTO_PAGADO")
							intTTMontoPendCCE = RsInf("TOTAL_MONTO_ACTIVO")

							intTTGeneralCCE = intTTGeneralCCE + intTTGestCCE
							intTTGeneralDocCCE = intTTGeneralDocCCE + intTTDocCCE
							intTTGeneralMontoCCE = intTTGeneralMontoCCE + intTTMontoCCE
							intTTGeneralDocPagadoCCE = intTTGeneralDocPagadoCCE + intTTDocPagadoCCE
							intTTGeneralDocPendCCE = intTTGeneralDocPendCCE + intTTDocPendCCE
							intTTGeneralMontoPagadoCCE = intTTGeneralMontoPagadoCCE + intTTMontoPagadoCCE
							intTTGeneralMontoPendCCE = intTTGeneralMontoPendCCE + intTTMontoPendCCE


						%>
							<TR bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
								<TD ALIGN="left"><%=(strUsuarioGestCC)%></TD>
								
								<td align="right">
									<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=3&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&strEjeAsig=<%=RsInf("ID_USUARIO")%>">
									<%=FN(intTTGestCCE,0)%>
									</A>
								</td>
								
								<TD ALIGN="RIGHT"><%=FN(intTTDocCCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoCCE,0)%></TD>
								<TD ALIGN="RIGHT"><%=FN(intTTDocPagadoCCE,0)%></TD>
								<TD ALIGN="RIGHT"><%=FN(intTTDocPendCCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPagadoCCE,0)%></TD>
								<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTMontoPendCCE,0)%></TD>
							</TR>
						<%

						End If

						RsInf.movenext
					loop
				end if
				RsInf.close
				set RsInf=nothing

		CerrarSCG()%>
				</tbody>
				<thead>
					<tr class="totales">
						<TD ALIGN="CENTER">Totales</TD>

						<td ALIGN="RIGHT" bordercolor="#999999">
							<div class="uno">
							<A HREF="Detalle_informe_compromisos.asp?intCodTipoGes=3&strCobranza=<%=strCobranza%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>">
							<%=FN(intTTGeneralCCE,0)%>
							</A>
							<div>
						</td>
						
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocCCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoCCE,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocPagadoCCE,0)%></TD>
						<TD ALIGN="RIGHT"><%=FN(intTTGeneralDocPendCCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoPagadoCCE,0)%></TD>
						<TD ALIGN="RIGHT">$&nbsp;<%=FN(intTTGeneralMontoPendCCE,0)%></TD>
					</tr>

				<%Else%>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<TD HEIGHT="20" ALIGN="CENTER"><B>NO EXISTEN COMPROMISOS CUMPLIDOS<B></TD>
					</tr>
				<%End If%>
				</thead>
				</table>
			</td>
		</tr>
	</table>

	</div>

	<br>

	<table width="100%" border="0" valign="top">

		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="intercalado" style="width:100%;">
				<thead>	
				   <tr HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
						<TD Colspan="4" class="subtitulo_informe">
							> Efectividad Compromisos
						</TD>
					</tr>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

						<TD WIDTH="20%" ALIGN="CENTER">Ejecutivo</TD>
						<TD WIDTH="20%" ALIGN="CENTER">Gestiones</TD>
						<TD WIDTH="20%" ALIGN="CENTER">Documentos</TD>
						<TD WIDTH="20%" ALIGN="CENTER">Monto</TD>
					</tr>
				</thead>
				<tbody>
<%		AbrirSCG()
						strSql = "SELECT U.LOGIN,SUM(GESTIONES_PAGADAS) AS GESTIONES_PAGADAS,SUM(TT_MONTO) AS TOTAL_MONTO,SUM(TT_MONTO_PAGADO) AS TOTAL_MONTO_PAGADO,SUM(TT_MONTO_ACTIVO) AS TOTAL_MONTO_ACTIVO,"
						strSql= strSql & " COUNT(*) AS TOTAL_GESTIONES,SUM(TT_DOC) AS TOTAL_DOCUMENTOS,SUM(DOC_PAGADOS) AS TOTAL_DOC_PAGADOS,SUM(TT_MONTO_ACTIVO) AS TOTAL_DOC_PEND"

						strSql= strSql & " 	FROM GESTIONES G	INNER JOIN DEUDOR D ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.COD_CLIENTE = D.COD_CLIENTE"
						strSql= strSql & " 						LEFT JOIN CAJA_FORMA_PAGO CFP ON G.FORMA_PAGO = CFP.ID_FORMA_PAGO"
						strSql= strSql & " 						LEFT JOIN USUARIO U ON G.ID_USUARIO = U.ID_USUARIO"
						strSql= strSql & " 						INNER JOIN"
						strSql= strSql & " (SELECT G.ID_GESTION,"
						strSql= strSql & " (CASE WHEN COUNT(C.ID_CUOTA) = SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP)"
						strSql= strSql & " THEN 1 ELSE 0 END)) THEN 1"
						strSql= strSql & " WHEN SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END)) > 0"
						strSql= strSql & " THEN 1"
						strSql= strSql & " ELSE 0 END) AS GESTIONES_PAGADAS,"
						strSql= strSql & " SUM(C.VALOR_CUOTA) AS TT_MONTO,"
						strSql= strSql & " COUNT(C.ID_CUOTA) AS TT_DOC,"
						strSql= strSql & " SUM((CASE WHEN ED.GRUPO='ACTIVOS' THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_ACTIVO,"
						strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_PAGADO,"
						strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='RETIROS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA ELSE 0 END)) AS TT_MONTO_RETIRADO,"
						strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='NO ASIGNABLES' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN C.VALOR_CUOTA  ELSE 0 END)) AS TT_MONTO_NO_ASIGNABLE,"

						strSql= strSql & " SUM((CASE WHEN ED.GRUPO='ACTIVOS' THEN 1 ELSE 0 END)) AS DOC_ACTIVOS,"
						strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END)) AS DOC_PAGADOS,"
						strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='RETIROS' AND G.ID_GESTION = C.ID_ULT_GEST_CP)  THEN 1 ELSE 0 END)) AS DOC_RETIRADOS,"
						strSql= strSql & " SUM((CASE WHEN (ED.GRUPO='NO ASIGNABLES' AND G.ID_GESTION = C.ID_ULT_GEST_CP)  THEN 1 ELSE 0 END)) AS DOC_NO_ASIGNABLES,"

						strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST_TIT ELSE 0 END)) AS GEST_TIT,"
						strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST_GENERAL ELSE 0 END)) AS GEST_GENERAL,"
						strSql= strSql & " MAX((CASE WHEN ED.GRUPO='ACTIVOS' THEN ID_ULT_GEST ELSE 0 END)) AS GEST_SCOR"

						strSql= strSql & " FROM GESTIONES G		INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION"
						strSql= strSql & " 						INNER JOIN CUOTA C ON GC.ID_CUOTA = C.ID_CUOTA"
						strSql= strSql & " 						INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO"
						strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA AND"
						strSql= strSql & " 																 G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA AND"
						strSql= strSql & " 																 G.COD_GESTION = GTG.COD_GESTION AND"
						strSql= strSql & " 																 G.COD_CLIENTE = GTG.COD_CLIENTE"
						strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON G.COD_CATEGORIA = GTC.COD_CATEGORIA"
						strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA AND"
						strSql= strSql & " 																	   G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA"

						strSql= strSql & " WHERE GTG.CATEGORIZACION IN (1,2) AND G.COD_CLIENTE IN (" & strCodCliente &")"

						If strEjeAsig <> "" then

						strSql = strSql & " AND  G.ID_USUARIO = " & strEjeAsig

						End If

						If strFechaInicio <> "" then

						strSql = strSql & " AND  G.FECHA_COMPROMISO > = '" & strFechaInicio & " 00:00:00'"

						End If

						If strFechaTermino <> "" then

						strSql = strSql & " AND G.FECHA_COMPROMISO < = '" & strFechaTermino & " 23:59:59'"

						End If

						If Trim(strCobranza) = "INTERNA" Then
							strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
						End if

						If Trim(strCobranza) = "EXTERNA" Then
							strSql = strSql & " AND C.CUSTODIO IS NULL"
						End if

						strSql= strSql & " GROUP BY G.ID_GESTION) AS PP ON G.ID_GESTION=PP.ID_GESTION"

						strSql= strSql & " GROUP BY U.LOGIN"

						'Response.write "<br>strSql=" & strSql

						set RsInf=Conn.execute(strSql)

						If not RsInf.eof then

							intEfectividadGes = ((intTTGestCC + intTTGestCPC) / intTTGestC)*100

							intEfectividadDoc = ((intTTDocPagadoCC + intTTDocPagadoCPC) / intTTDocC)*100

							intEfectividadMonto = ((intTTMontoPagadoCC + intTTMontoPagadoCPC) / intTTMontoC)*100

							do until RsInf.eof

							strUsuarioEfectividad = RsInf("LOGIN")

							intTTGestionesEfecE = RsInf("TOTAL_GESTIONES")
							intTTDocEfectE = RsInf("TOTAL_DOCUMENTOS")
							intTTMontoEfectE  = RsInf("TOTAL_MONTO")
							intTTGestPagadasEfecE = RsInf("GESTIONES_PAGADAS")
							intTTDocPagadoEfecE  = RsInf("TOTAL_DOC_PAGADOS")
							intTTMontoPagadoEfecE  = RsInf("TOTAL_MONTO_PAGADO")

							intEfectividadGesE = ((intTTGestPagadasEfecE) / intTTGestionesEfecE)*100

							intEfectividadDocE = ((intTTDocPagadoEfecE) / intTTDocEfectE)*100

							intEfectividadMontoE = ((intTTMontoPagadoEfecE) / intTTMontoEfectE)*100

						%>
							<TR bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
								<TD ALIGN="left"><%=(strUsuarioEfectividad)%></TD>
								<TD ALIGN="CENTER"><%=FN(intEfectividadGesE,1)%>&nbsp;%</TD>
								<TD ALIGN="CENTER"><%=FN(intEfectividadDocE,1)%>%&nbsp;</TD>
								<TD ALIGN="CENTER"><%=FN(intEfectividadMontoE,1)%>%&nbsp;</TD>
							</TR>
						<%

								RsInf.movenext
							loop
						end if
						RsInf.close
						set RsInf=nothing

			CerrarSCG()%>
				</tbody>
				<thead>
					<tr class="totales">
						<TD ALIGN="CENTER">General</TD>
						<TD ALIGN="CENTER"><%=FN(intEfectividadGes,1)%>%&nbsp;</TD>
						<TD ALIGN="CENTER"><%=FN(intEfectividadDoc,1)%>%&nbsp;</TD>
						<TD ALIGN="CENTER"><%=FN(intEfectividadMonto,1)%>%&nbsp;</TD>
					</tr>
				</thead>	
				</table>
			</td>
		</tr>
	</table>

</body>
</html>
<script language="JavaScript1.2">

function envia(){
		//datos.action='cargando.asp';
		datos.action='informe_compromisos.asp?intTipoInforme=<%=intTipoInforme%>&resp=si';
		datos.submit();
}
function MostrarFilas(Fila) {
var elementos = document.getElementsByName(Fila);
	for (i = 0; i< elementos.length; i++) {
		if(navigator.appName.indexOf("Microsoft") > -1){
			   var visible = 'block'
		} else {
			   var visible = 'table-row';
		}
elementos[i].style.display = visible;
		}
}

function OcultarFilas(Fila) {
	var elementos = document.getElementsByName(Fila);
	for (k = 0; k< elementos.length; k++) {
			   elementos[k].style.display = "none";
	}
}
function cajas1()
{
	MostrarFilas('divCNC');
	OcultarFilas('divCPC');
	OcultarFilas('divCC');
}
function cajas2()
{
	OcultarFilas('divCNC');
	MostrarFilas('divCPC');
	OcultarFilas('divCC');
}
function cajas3()
{
	OcultarFilas('divCNC');
	OcultarFilas('divCPC');
	MostrarFilas('divCC');
}
function cajas4()
{
	MostrarFilas('divGes');
	OcultarFilas('divDoc');
}
function cajas5()
{
	OcultarFilas('divGes');
	MostrarFilas('divDoc');
}

function CargaUsuarios(subCat,cat)
{
	//alert(subCat);
	//alert(cat);

	var comboBox = document.getElementById('CB_EJECUTIVO');
	switch (cat)
	{
		<%
		  AbrirSCG()
			strSql="SELECT COD_CLIENTE FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE"
			set rsGestCat=Conn.execute(strSql)
			Do While not rsGestCat.eof
		%>
		case '<%=rsGestCat("COD_CLIENTE")%>':

			comboBox.options.length = 0;

				if (subCat=='INTERNA') {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

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
					break;
				}

				if (subCat=='EXTERNA' && (<%=intVerEjecutivos%>=='1')) {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

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
					break;
				}
				else if ((subCat=='EXTERNA') && (<%=intVerEjecutivos%>=='0')) {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					break;
				}
				else {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

					strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
					''strSql = strSql & " AND U.PERFIL_EMP=0"

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
					break;
				}

		<%
		  	rsGestCat.movenext
		  	Loop
		  	rsGestCat.close
		  	set rsGestCat=nothing
			CerrarSCG()
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
CargaUsuarios('<%=strCobranza%>','<%=strCodCliente%>');
<%End If%>

<%If strEjeAsig <> "" then%>
datos.CB_EJECUTIVO.value='<%=strEjeAsig%>';
<%End If%>
</script>
