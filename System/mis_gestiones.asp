<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/comunes/rutinas/funcionesBD.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<script language="JavaScript">
function ventanaSecundaria (URL){
	window.open(URL,"DETALLE","width=200, height=200, scrollbars=no, menubar=no, location=no, resizable=no")
}

</script>

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	strFechaInicio = request("TX_FECINICIO")
	strFechaTermino = request("TX_FECTERMINO")
	intIdFocoAT=Request("CB_FOCOAT")
	intCodCampana = Request("CB_CAMPANA")
	strTramoVenc = Request("CB_TRVENC")
	strTramoMonto = Request("CB_TRMONTO")
	strSucursal = Request("CB_SUCURSAL")
	
	resp = request("resp")
	If resp="" then resp="1" end if	
	
	strCobranza = Request("CB_COBRANZA")
	intVerCob = "1"

	intCodEstadoCob=Trim(Request("CB_ESTADOCOB"))

	If Trim(intCodEstadoCob) = "" Then intCodEstadoCob = "0"
	
	If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

		strEjeAsig = Request("CB_EJECUTIVO")

	Else
	
		strEjeAsig =  session("session_idusuario")
	
	End If
	
	'Response.write "strSucursal=" & strSucursal

	AbrirSCG()

	if Trim(strFechaInicio) = "" Then
		strFechaInicio = TraeFechaActual(Conn)
	End If

	if Trim(strFechaTermino) = "" Then
		strFechaTermino = TraeFechaActual(Conn)

	Else strFechaTermino = strFechaTermino

	End If

	If resp = "1" then
		strColor1 = "boton_rojo"
	else
		strColor1 = ""
	End if

	If resp = "2" then
		strColor2 = "boton_rojo"
	else
		strColor2 = ""
	End if
	
	CerrarSCG()

	If Request("CB_CLIENTE") = "" then

		strCodCliente = session("ses_codcli")
	Else
		strCodCliente = Request("CB_CLIENTE")
	End If
	
	If intIdFocoAT = "" Then intIdFocoAT = 0
	If intCodCampana = "" Then intCodCampana = 0

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

End If

If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

	sinCbUsario="0"

End If

'---Fin codigo tipo de cobranza---'

%>
<title>MIS GESTIONES REALIZADAS</title>
<style type="text/css">
<!--
.Estilo37 {color: #FFFFFF}
-->

.uno a {
	text-align:center;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	
	color: #FFFFFF;
}
.uno a:hover {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	text-decoration: none;
	color: #FFFFFF;
}
</style>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

</head>
<body>
<form name="datos" method="post">
	<div class="titulo_informe">DETALLE GESTIONES</div>
<br>
<table width="90%" border="0" align="center">

  <tr height="20">
    <td style="vertical-align: top;">
		<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
		<thead>	
			<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

				<td>CLIENTE</td>
				<td align="left">FOCO</td>
				<td align="left">CAMPAÑA</td>
				<td align="left">TRAMO VENCIMIENTO</td>
				<td align="left">TRAMO MONTO</td>
				<td align="left">SEDE</td>
				<td>FECHA INICIO</td>
				<td>FECHA TERMINO</td>

			<% If sinCbUsario = "0" Then %>
				<td>EJECUTIVO</td>
			<%Else%>
				<td>&nbsp;</td>
			<%End If%>

			</tr>
		</thead>
			<tr bordercolor="#999999" class="Estilo8">

				<td>

				<SELECT NAME="CB_CLIENTE" id="CB_CLIENTE" onChange="CargaUsuarios(CB_COBRANZA.value,this.value);">

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
					<select name="CB_FOCOAT" id="CB_FOCOAT">
						<option value=0>TODOS</option>
						<%
						
						AbrirSCG()
							strSql="SELECT ID_FOCO,NOMBRE_FOCO FROM FOCOS WHERE TIPO_FOCO=1 ORDER BY ID_FOCO ASC"
							set rsFocos=Conn.execute(strSql)
							Do While not rsFocos.eof
								If Trim(intIdFocoAT)=Trim(rsFocos("ID_FOCO")) Then strSelFoco = "SELECTED" Else strSelFoco = ""
								%>
								<option value="<%=rsFocos("ID_FOCO")%>" <%=strSelFoco%>> <%=rsFocos("NOMBRE_FOCO")%></option>
								<%
								rsFocos.movenext
							Loop
							rsFocos.close
							set rsFocos=nothing
						''Response.End
						CerrarSCG()
						
						%>
						<option value=100 <%if Trim(intIdFocoAT)=100 then response.Write("Selected") end if%>>SIN FOCO</option>
					</select>
				</td>
				<td>
					<select name="CB_CAMPANA" id="CB_CAMPANA" onChange="buscar()">
						<option value=0>TODAS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_CAMPANA,NOMBRE FROM CAMPANA WHERE COD_CLIENTE IN ('" & strCodCliente & "')"
							set rsCampana=Conn.execute(strSql)
							Do While not rsCampana.eof
								If Trim(intCodCampana)=Trim(rsCampana("ID_CAMPANA")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsCampana("ID_CAMPANA")%>" <%=strSelCam%>> <%=rsCampana("NOMBRE")%></option>
								<%
								rsCampana.movenext
							Loop
							rsCampana.close
							set rsCampana=nothing
						CerrarSCG()
						''Response.End
						%>
						<option value=1 <%if Trim(intCodCampana)=1 then response.Write("Selected") end if%>>SIN CAMPAÑA</option>
					</select>
				</td>
				<td>
					<select name="CB_TRVENC" id="CB_TRVENC">
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_SVENC = ID,NOMBRE_SEGMENTO=LTRIM(NOMBRE) FROM SEGMENTACION_VENCIMIENTO WHERE COD_CLIENTE = '" & strCodCliente & "' AND GESTIONABLE=1 ORDER BY ORDEN ASC"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(strTramoVenc)=Trim(rsSel("ID_SVENC")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("ID_SVENC")%>" <%=strSelCam%>> <%=rsSel("NOMBRE_SEGMENTO")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>				
				<td>
					<select name="CB_TRMONTO" id="CB_TRMONTO" >
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_SMONTO = ID,NOMBRE_SEGMENTO=LTRIM(NOMBRE) FROM SEGMENTACION_MONTO WHERE COD_CLIENTE = ('" & strCodCliente & "') AND GESTIONABLE=1 ORDER BY ORDEN ASC"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(strTramoMonto)=Trim(rsSel("ID_SMONTO")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("ID_SMONTO")%>" <%=strSelCam%>> <%=rsSel("NOMBRE_SEGMENTO")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>
				<td>
					<select name="CB_SUCURSAL" id="CB_SUCURSAL" >
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT DISTINCT C.SUCURSAL FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"
							strSql= strSql & " WHERE C.COD_CLIENTE IN ('" & strCodCliente & "') AND ED.ACTIVO=1"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(strSucursal)=Trim(rsSel("SUCURSAL")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("SUCURSAL")%>" <%=strSelCam%>> <%=rsSel("SUCURSAL")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>	

				<td>	<input name="TX_FECINICIO" id="TX_FECINICIO" readonly="true" type="text" value="<%=strFechaInicio%>" size="10" maxlength="10">
				</td>

				<td>	<input name="TX_FECTERMINO" id="TX_FECTERMINO" readonly="true" type="text" value="<%=strFechaTermino%>" size="10" maxlength="10">
				</td>

			<% If sinCbUsario="0" Then %>

				<td>
					<select name="CB_EJECUTIVO" id="CB_EJECUTIVO">
					</select>
				</td>

			<% End If %>

			</tr>


			<tr >
				<td ALIGN="Left" Colspan="6">
					<input class="fondo_boton_100 <%=strColor1%>" name="Submit1"  style="font-size:11px;width:130px" type="button" onClick="envia('1');" value="Ver Gestiones">

					<input class="fondo_boton_100 <%=strColor2%>" name="Submit2" style="font-size:11px;width:130px" type="button" onClick="envia('2');" value="Ver Arbol General">
				</td>
			</tr>
			
		</table>
    </td>
   </tr>

	<table width="100%" border="1" bordercolor = "#<%=session("COLTABBG")%>" cellSpacing="0" cellPadding="2" class="intercalado">

    <%
	
	If resp="1" then
	
	AbrirSCG()

		strSql=" SELECT COUNT(DISTINCT G.ID_GESTION) AS TOTAL_GES,"
		strSql = strSql & " COUNT(G.ID_GESTION) AS TOTAL_DOC_GES, ISNULL(SUM(C.VALOR_CUOTA),0) AS TOTAL_MONTO_GES,"
		strSql = strSql & " G.COD_CATEGORIA AS COD_CA,G.COD_SUB_CATEGORIA AS COD_SUB_CA,G.COD_GESTION AS COD_GES, "
		strSql = strSql & " (CASE WHEN ROW_NUMBER() OVER(PARTITION BY G.COD_CATEGORIA ORDER BY G.COD_CATEGORIA,G.COD_SUB_CATEGORIA,G.COD_GESTION ASC) = 1"
		strSql = strSql & " 	  THEN CAST(G.COD_CATEGORIA AS VARCHAR(10))+ ' - ' + GTC.DESCRIPCION "
		strSql = strSql & " 	  ELSE '&nbsp;'"
		strSql = strSql & " END) AS DES1, "
		strSql = strSql & " (CASE WHEN ROW_NUMBER() OVER(PARTITION BY G.COD_CATEGORIA,G.COD_SUB_CATEGORIA ORDER BY G.COD_CATEGORIA,G.COD_SUB_CATEGORIA,G.COD_GESTION ASC) = 1 "
		strSql = strSql & " 	  THEN CAST(G.COD_SUB_CATEGORIA AS VARCHAR(10))+ ' - ' + GTSC.DESCRIPCION "
		strSql = strSql & " 	  ELSE '&nbsp;' "
		strSql = strSql & " END) AS DES2, "
		strSql = strSql & " GTG.DESCRIPCION AS DES3"

		strSql = strSql & " FROM GESTIONES G INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON G.COD_CATEGORIA = GTC.COD_CATEGORIA "
		strSql = strSql & " 				 INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA "
		strSql = strSql & " 																   AND G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA"
		strSql = strSql & " 				 INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
		strSql = strSql & " 														   AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA"
		strSql = strSql & " 														   AND G.COD_GESTION = GTG.COD_GESTION"
		strSql = strSql & " 														   AND GTG.COD_CLIENTE IN ('" & strCodCliente & "')"
		strSql = strSql & " 				 INNER JOIN DEUDOR DD ON G.COD_CLIENTE = DD.COD_CLIENTE AND G.RUT_DEUDOR = DD.RUT_DEUDOR"
		strSql = strSql & " 				 INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION "
		strSql = strSql & " 				 INNER JOIN CUOTA C ON C.ID_CUOTA = GC.ID_CUOTA AND GC.ID_GESTION = G.ID_GESTION "
		strSql = strSql & " 										AND C.COD_CLIENTE IN ('" & strCodCliente & "')"
		
		strSql = strSql & " WHERE G.FECHA_INGRESO BETWEEN CAST('" & strFechaInicio & "' AS DATETIME) AND CAST('" & strFechaTermino & "' AS DATETIME)"
		strSql = strSql & " AND G.COD_CLIENTE IN ('" & strCodCliente & "')"

		if Trim(strEjeAsig) <> "" Then
			strSql = strSql & " AND  G.ID_USUARIO = '" & strEjeAsig & "'"
		End if

		if Trim(intIdFocoAT) <> "" AND Trim(intIdFocoAT) <> "0" AND Trim(intIdFocoAT) <> "100" Then
			strSql = strSql & " AND  DD.ID_FOCO = " & intIdFocoAT
		End if

		if Trim(intIdFocoAT) = "100" Then
			strSql = strSql & " AND  DD.ID_FOCO IS NULL"
		End if

		if Trim(intCodCampana) <> "" AND Trim(intCodCampana) <> "0" AND  Trim(intCodCampana) <> "1" Then
			strSql = strSql & " AND  DD.ID_CAMPANA = " & intCodCampana
		End if

		if Trim(intCodCampana) = "1" Then
			strSql = strSql & " AND  DD.ID_CAMPANA IS NULL"
		End if
		
		if Trim(strTramoVenc) <> "" Then
			strSql = strSql & " AND  G.ID_SEGMENTO_VENC = " & strTramoVenc
		End if
		
		if Trim(strTramoMonto) <> "" Then
			strSql = strSql & " AND  G.ID_SEGMENTO_MONTO = " & strTramoMonto
		End if			
		
		strSql = strSql & " GROUP BY G.COD_CATEGORIA ,G.COD_SUB_CATEGORIA,G.COD_GESTION,GTSC.DESCRIPCION ,GTC.DESCRIPCION,GTG.DESCRIPCION"

		if Trim(strSucursal) <> "" Then
			strSql = strSql & " HAVING MAX(C.SUCURSAL)= '" & strSucursal & "'"
		End if	
		
		strSql = strSql & " ORDER BY COD_CA,COD_SUB_CA,COD_GES ASC" 
		
		intTotalGeneralGes = 0
		intTotalGeneralDocGes = 0
		intTotalGeneralMontoGes = 0

		intTotalGes = 0
		intTotalDocGes = 0
		intTotalGeneralMontoGes = 0

		'Response.write "<br>strSql = " & strSql
		
		set rsGES= Conn.execute(strSql)

		if not rsGES.eof then%>
		<thead>	
		  <tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="150">CATEGORIA</td>
			<td width="200">SUBCATEGORIA</td>
			<td>GESTION</td>
			<td width="70" align="right">CANT.G</td>
			<td width="70" align="right">DOC. ASOC.</td>
			<td width="80" align="right">MONTO ASOC.</td>
		  </tr>
		</thead>
	  	<tbody>
		<%
			intCodCaAnt = rsGES("COD_CA")

			Do while not rsGES.eof

			intCodCa = rsGES("COD_CA")

			intCodCategoria = rsGES("COD_CA")
			intCodSubCategoria = rsGES("COD_SUB_CA")
			intCodGestion = rsGES("COD_GES")

			strDes1 = rsGES("DES1")
			strDes2 = rsGES("DES2")
			strDes3 = Cstr(intCodGestion) + " - " + rsGES("DES3")

			intTotalDocGes = rsGES("TOTAL_DOC_GES")

			intTotalMontoGes = rsGES("TOTAL_MONTO_GES")

			intTotalGeneralDocGes = intTotalGeneralDocGes + intTotalDocGes

			intTotalGeneralMontoGes = intTotalGeneralMontoGes + intTotalMontoGes

			intTotalGes = rsGES("TOTAL_GES")

			intTotalGeneralGes = intTotalGeneralGes + intTotalGes
			
			If Cdbl(intCodCa) <> Cdbl(intCodCaAnt) then

			%>

			<tr>
				<td colspan = "6" bgcolor="#<%=session("COLTABBG2")%>">&nbsp;</td>
			</tr>

			<%End If

			intCodCaAnt = rsGES("COD_CA")%>

			<tr>
				<td>
					<%=strDes1%>
				</td>
				<td>
					<%=strDes2%>
				</td>
				<td>
					<%=strDes3%>
				</td>
				<td Align="right">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&intCodCategoria=<%=intCodCategoria%>&intCodSubCategoria=<%=intCodSubCategoria%>&intCodGestion=<%=intCodGestion%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
						<%=intTotalGes%>
				</td>
				<td Align="right">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&intCodCategoria=<%=intCodCategoria%>&intCodSubCategoria=<%=intCodSubCategoria%>&intCodGestion=<%=intCodGestion%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
					<%=intTotalDocGes%>
				</td>
				<td Align="right">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&intCodCategoria=<%=intCodCategoria%>&intCodSubCategoria=<%=intCodSubCategoria%>&intCodGestion=<%=intCodGestion%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
					<%=FN(Cdbl(intTotalMontoGes),0)%>
				</td>
			</tr>
			<%

				rsGES.movenext
			Loop
		rsGES.close
		set rsGES=nothing

		CerrarSCG()%>
		</tbody>
		<thead  class="totales">
			<td colspan = "4"align="right" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<div class="uno">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
				<%=FN(intTotalGeneralGes,0)%>
			<div>
			</td>

			<td align="right" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<div class="uno">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
				<%=FN(intTotalGeneralDocGes,0)%>
			<div>
			</td>

			<td align="right" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<div class="uno">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
				<%=FN(intTotalGeneralMontoGes,0)%>
			<div>
			</td>
			</thead>
		<%Else%>
		</thead>
		<tr class="estilo_columna_individual">
			<td Colspan = "1">&nbsp;</td>
		</tr>
			
		<tr >
			<td ALIGN="CENTER" Colspan = "1">NO HAY GESTIONES INGRESADAS SEGUN PARAMETROS DE BUSQUEDA</td>
		</tr>

		<tr class="estilo_columna_individual">
			<td Colspan = "1">&nbsp;</td>
		</tr>
		
		<%End If

	End If

	If resp="2" then%>
		<thead>
		  <tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="150">CATEGORIA</td>
			<td width="200">SUBCATEGORIA</td>
			<td>GESTION</td>
			<td width="70" align="right">CANT.G</td>
			<td width="70" align="right">DOC. ASOC.</td>
			<td width="80" align="right">MONTO ASOC.</td>
		  </tr>
	  	</thead>
	  	<tbody>
	<%AbrirSCG()

		strSql="SELECT GTG.COD_CATEGORIA AS COD_CA,GTG.COD_SUB_CATEGORIA AS COD_SUB_CA,GTG.COD_GESTION AS COD_GES,"
		strSql = strSql & " (CASE WHEN ROW_NUMBER() OVER(PARTITION BY GTG.COD_CATEGORIA ORDER BY GTG.COD_CATEGORIA,GTG.COD_SUB_CATEGORIA,GTG.COD_GESTION ASC) = 1 THEN CAST(GTG.COD_CATEGORIA AS VARCHAR(10))+ ' - ' + GTC.DESCRIPCION ELSE '&nbsp;' END) AS DES1,"
		strSql = strSql & " (CASE WHEN ROW_NUMBER() OVER(PARTITION BY GTG.COD_CATEGORIA,GTG.COD_SUB_CATEGORIA ORDER BY GTG.COD_CATEGORIA,GTG.COD_SUB_CATEGORIA,GTG.COD_GESTION ASC) = 1 THEN CAST(GTG.COD_SUB_CATEGORIA AS VARCHAR(10))+ ' - ' + GTSC.DESCRIPCION ELSE '&nbsp;' END) AS DES2,"
		strSql = strSql & "  GTG.DESCRIPCION AS DES3"

		strSql = strSql & " FROM GESTIONES_TIPO_GESTION GTG INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON GTSC.COD_CATEGORIA = GTG.COD_CATEGORIA AND GTSC.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA"
		strSql = strSql & " 								INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON GTC.COD_CATEGORIA = GTG.COD_CATEGORIA"
		strSql = strSql & " WHERE GTG.COD_CLIENTE = " & strCodCliente
		strSql = strSql & " ORDER BY GTG.COD_CATEGORIA,GTG.COD_SUB_CATEGORIA,GTG.COD_GESTION ASC"

		'Response.write "strSql = " & strSql

		intTotalGeneralGes = 0
		intTotalGeneralDocGes = 0
		intTotalGeneralMontoGes = 0

		intTotalGes = 0
		intTotalDocGes = 0
		intTotalGeneralMontoGes = 0

		set rsGES= Conn.execute(strSql)

			intCodCaAnt = rsGES("COD_CA")

			Do while not rsGES.eof

			intCodCa = rsGES("COD_CA")

			intCodCategoria = rsGES("COD_CA")
			intCodSubCategoria = rsGES("COD_SUB_CA")
			intCodGestion = rsGES("COD_GES")

			strDes1 = rsGES("DES1")
			strDes2 = rsGES("DES2")
			strDes3 = Cstr(intCodGestion) + " - " + rsGES("DES3")

			AbrirSCG2()

			strSql=" SELECT COUNT(G.ID_GESTION) AS TOTAL_DOC_GES, ISNULL(SUM(C.VALOR_CUOTA),0) AS TOTAL_MONTO_GES, COUNT(DISTINCT G.ID_GESTION) ASTOTAL_GES, COUNT(DISTINCT G.ID_GESTION) AS TOTAL_GES"
			strSql = strSql & " FROM GESTIONES G INNER JOIN DEUDOR DD ON G.COD_CLIENTE = DD.COD_CLIENTE AND G.RUT_DEUDOR = DD.RUT_DEUDOR "
			strSql = strSql & " 				 INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION"
			strSql = strSql & " 				 INNER JOIN CUOTA C ON C.ID_CUOTA = GC.ID_CUOTA AND GC.ID_GESTION = G.ID_GESTION AND C.COD_CLIENTE IN ('" & strCodCliente & "')"

			strSql = strSql & " WHERE G.FECHA_INGRESO BETWEEN CAST('" & strFechaInicio & "' AS DATETIME) AND CAST('" & strFechaTermino & "' AS DATETIME)"
			strSql = strSql & " AND G.COD_CLIENTE IN ('" & strCodCliente & "')"
			strSql = strSql & " AND G.COD_CATEGORIA = " & rsGES("COD_CA") & " AND G.COD_SUB_CATEGORIA = " & rsGES("COD_SUB_CA") & " AND G.COD_GESTION = " & rsGES("COD_GES")

			if Trim(strEjeAsig) <> "" Then
				strSql = strSql & " AND  G.ID_USUARIO = '" & strEjeAsig & "'"
			End if

			If Trim(strCobranza) = "INTERNA" Then
				strSql = strSql & " AND DD.CUSTODIO IS NOT NULL"
			End if

			If Trim(strCobranza) = "EXTERNA" Then
				strSql = strSql & " AND DD.CUSTODIO IS NULL"
			End if

			'Response.write "<br>strSql = " & strSql

			set rsGES2= Conn2.execute(strSql)

				Do while not rsGES2.eof

				intTotalDocGes = rsGES2("TOTAL_DOC_GES")

				intTotalMontoGes = rsGES2("TOTAL_MONTO_GES")

				intTotalGeneralDocGes = intTotalGeneralDocGes + intTotalDocGes

				intTotalGeneralMontoGes = intTotalGeneralMontoGes + intTotalMontoGes

				intTotalGes = rsGES2("TOTAL_GES")

				intTotalGeneralGes = intTotalGeneralGes + intTotalGes
				
					rsGES2.movenext
					Loop
				rsGES2.close
				set rsGES2=nothing

			CerrarSCG2()

			If Cdbl(intCodCa) <> Cdbl(intCodCaAnt) then

			%>

			<tr>
				<td colspan = "6" bgcolor="#<%=session("COLTABBG2")%>" >&nbsp;</td>
			</tr>

			<%End If

			intCodCaAnt = rsGES("COD_CA")%>

			<tr>
				<td>
					<%=strDes1%>
				</td>
				<td>
					<%=strDes2%>
				</td>
				<td>
					<%=strDes3%>
				</td>
				<td Align="right">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&intCodCategoria=<%=intCodCategoria%>&intCodSubCategoria=<%=intCodSubCategoria%>&intCodGestion=<%=intCodGestion%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
						<%=intTotalGes%>
				</td>
				<td Align="right">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&intCodCategoria=<%=intCodCategoria%>&intCodSubCategoria=<%=intCodSubCategoria%>&intCodGestion=<%=intCodGestion%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
					<%=intTotalDocGes%>
				</td>
				<td Align="right">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&intCodCategoria=<%=intCodCategoria%>&intCodSubCategoria=<%=intCodSubCategoria%>&intCodGestion=<%=intCodGestion%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
					<%=FN(Cdbl(intTotalMontoGes),0)%>
				</td>
			</tr>
			<%
			
				rsGES.movenext
			Loop
		rsGES.close
		set rsGES=nothing
		
	CerrarSCG()%>
			</tbody>
			<thead class="totales">
				<td colspan = "4"align="right" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<div class="uno">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
				<%=FN(intTotalGeneralGes,0)%>
			<div>
			</td>

			<td align="right" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<div class="uno">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
				<%=FN(intTotalGeneralDocGes,0)%>
			<div>
			</td>

			<td align="right" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<div class="uno">
						<A HREF="Detalle_informe_gestiones_2.asp?CB_TIPO_LISTADO=1&strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&CB_COBRANZA=<%=strCobranza%>&intIdFocoAT=<%=intIdFocoAT%>&intCodCampana=<%=intCodCampana%>&strTramoVenc=<%=strTramoVenc%>&strTramoMonto=<%=strTramoMonto%>&strSucursal=<%=strSucursal%>">
				<%=FN(intTotalGeneralMontoGes,0)%>
			<div>
			</td>
			</thead>
	<%End If%>
    </table>

	  </td>
  </tr>
</table>

</form>
</body>
</html>
<script type="text/javascript">
$(document).ready(function(){

	$('#TX_FECTERMINO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FECINICIO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
  
})
</script>
<script language="JavaScript1.2">

function envia(tipo){

	if(tipo=='1'){
	datos.Submit1.disabled = true;
	resp='1'
	document.datos.action = "mis_gestiones.asp?strBuscar=S&resp="+ resp +"";
	document.datos.submit();
	}
	else if(tipo=='2') {
	datos.Submit2.disabled = true;
	resp='2'
	document.datos.action = "mis_gestiones.asp?strBuscar=S&resp="+ resp +"";
	document.datos.submit();
	}
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

<%If sinCbUsario = "0" then%>
CargaUsuarios('<%=strCobranza%>','<%=strCodCliente%>');
<%End If%>

<%If strEjeAsig <> "" then%>
datos.CB_EJECUTIVO.value='<%=strEjeAsig%>';
<%End If%>
</script>
