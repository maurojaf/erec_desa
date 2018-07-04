<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib2.asp"-->

	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/Minimo.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->
	
<%
	Response.CodePage 	=65001
	Response.charset 	="utf-8"

	Dim PaginaActual ' en que pagina estamos
	Dim PaginasTotales ' cuantas paginas tenemos
	Dim TamPagina ' cuantos registros por pagina
	Dim CuantosRegistros ' para imprimir solo el n de registro por pagina que

	if Request("CB_CLIENTE") <> "" and Request("CB_CLIENTE") <> session("ses_codcli") then
	intCodCliente=Request("CB_CLIENTE")
	else
	intCodCliente=session("ses_codcli")
	end if
	
    strNombres= Request("TX_NOMBRES")
	strRut = Request("TX_RUT_DEUDOR")
	intCodCampana = Request("CB_CAMPANA")
	strEjeAsig = Request("CB_EJECUTIVO")
	strTipoInf = Request("CB_TIPOCARTERA")
	intEstadoCob = Request("CB_TIPOCOB")
	strCobranza = Request("CB_COBRANZA")
	intVerCob = "1"

	intGestionPrinc = Request("CB_TIPOGESTION_PRINC")
	intGestion = Request("CB_TIPOGESTION")
	strPrioridad = Request("CB_PRIORIDAD")
	dtmInicio = Request("TX_INICIO")
	dtmTermino = Request("TX_TERMINO")
	intTipoDoc = Request("CB_TIPODOC")
	strTipoAgend = Request("CB_TIPOAGEND")
	strTipoUbic = Request("CB_UBICABILIDAD")
	strRubro = Request("CB_RUBRO")

	'Response.write "strTipoUbic=" & strTipoUbic

	If Trim(strTipoInf) = "" Then strTipoInf = "GESTIONABLES"

	'--Calcula Objetos relacionados al tipo de cobranza (Interna, Externa), CB_COBRANZA Y CB_EJECUTIVO--'

	abrirscg()

			strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
			strSql = strSql & " FROM CLIENTE CL"
			strSql = strSql & " WHERE CL.COD_CLIENTE = '" & session("ses_codcli") & "'"
		
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

		perfilEjecutivo="0"

	End If

	'---Fin codigo tipo de cobranza---'

	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_Ejecutivo") = strEjeAsig
		session("Ftro_Campana") = intCodCampana
		session("Ftro_DtmInicio") = dtmInicio
		session("Ftro_DtmTermino") = dtmTermino
		session("Ftro_TipoCartera") = strTipoInf
		session("Ftro_TipoGPpal") = intGestionPrinc
		session("Ftro_TipoGTel") = intGestion
		session("Ftro_Cliente") = intCodCliente
		session("Ftro_TipoDoc") = intTipoDoc
		session("Ftro_TipoCob") = intEstadoCob
		session("Ftro_Pioridad") = strPrioridad
		session("Ftro_TipoAgend") = strTipoAgend
		session("Ftro_TipoUbic") = strTipoUbic
		session("FtroCA_Rubro") = strRubro
		session("FtroCB_Cobranza") = strCobranza
	End If

	'Response.write "Ftro_Cliente=" & session("Ftro_Cliente")

	If Trim(Request("strBuscar")) = "N" or Trim(Request("strLimpiar")) = "S" Then
		session("Ftro_Ejecutivo") = ""
		session("Ftro_Campana") = ""
		session("Ftro_DtmInicio") = ""
		session("Ftro_DtmTermino") = ""
		session("Ftro_TipoCartera") = ""
		session("Ftro_TipoGPpal") = ""
		session("Ftro_TipoGTel") = ""
		session("Ftro_Cliente") = ""
		session("Ftro_TipoDoc") = ""
		session("Ftro_TipoCob") = ""
		session("Ftro_Pioridad") = ""
		session("Ftro_TipoUbic") = ""
		session("Ftro_TipoAgend")  = ""
		session("FtroCA_Rubro") = ""
		session("FtroCB_Cobranza") = ""
	End If

	If Trim(Request("strLimpiar")) = "S" Then
		strEjeAsig = "0"
		intCodCampana = ""
		dtmInicio = ""
		dtmTermino = ""
		intGestionPrinc = ""
		intGestion = ""
		intTipoDoc = ""
		intEstadoCob = ""
		strPrioridad = ""
		strTipoUbic = ""
		strRubro = ""
		strTipoAgend = ""
		strCobranza = ""
		
		If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
		intCodCliente = ""
		End If
		
	End If
	
	if strEjeAsig = "" then
		strEjeAsig = "0"
	end if
	
	If strEjeAsig <> "0" Then strEjeAsig = session("Ftro_Ejecutivo")
	If intCodCampana = "" Then intCodCampana = session("Ftro_Campana")
	If dtmInicio = "" Then dtmInicio = session("Ftro_DtmInicio")
	If dtmTermino = "" Then dtmTermino = session("Ftro_DtmTermino")
	If intGestionPrinc = "" Then intGestionPrinc = session("Ftro_TipoGPpal")
	If intGestion = "" Then intGestion = session("Ftro_TipoGTel")
	If intTipoDoc = "" Then intTipoDoc = session("Ftro_TipoDoc")
	If intEstadoCob = "" Then intEstadoCob = session("Ftro_TipoCob")
	If strPrioridad = "" Then strPrioridad = session("Ftro_Pioridad")
	If strTipoUbic <> "0" Then strTipoUbic = session("Ftro_TipoUbic")
	If strRubro = "" Then strRubro = session("FtroCA_Rubro")
	If strCobranza = "" Then strCobranza = session("FtroCB_Cobranza")

	'Response.write "intCodCliente=" & intCodCliente
	'Response.write "<br>perfil_adm=" & session("perfil_adm")
	'Response.write "<br>perfil_sup=" & session("perfil_sup")
	
	If session("Ftro_TipoAgend") <> "" Then

	strTipoAgend = session("Ftro_TipoAgend")

	strPrioridad= replace(strPrioridad, ",", ".")

	Else

		AbrirSCG()
			strSql="SELECT ID_USUARIO FROM USUARIO WHERE ID_USUARIO = " & session("session_idusuario") & " AND GESTIONADOR_PREVENTIVO = 1"

			'Response.write "strSql=" & strSql

			set rsUusarioGestion=Conn.execute(strSql)

			If Not rsUusarioGestion.Eof Then
			    strTipoAgend = "2"
			ElseIf TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_sup")) <> "Si"Then
      			strTipoAgend = "0"
			End If

		CerrarSCG()

	End If

	If strTipoAgend = "2" Then

	strColorMosPrev = "ff6666"
	strLetrasColorModPrev = "F3F3F3"
	strMensajeModPrev = " del Listado de Casos Preventivos"

	Else

	strColorMosPrev = "F3F3F3"
	strLetrasColorModPrev = "FF0000"

	End If


	'MODIFICAR AQUI PARA CAMBIAR EL N? DE REGISTRO POR PAGINA
	TamPagina=100

	'Leemos qu? p?gina mostrar. La primera vez ser? la inicial
	if Request.Querystring("pagina")="" then
		PaginaActual=1
	else
		PaginaActual=CInt(Request.Querystring("pagina"))
	end if


%>
<title>MODULO DE AGENDAMIENTOS</title>

<%strTitulo="MI CARTERA"%>


<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
<link href="../css/style_generales_sistema.css" rel="stylesheet">

<script language='javascript' src="../javascripts/popcalendar.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script language="JavaScript1.2">
$(document).ready(function(){

		$("#table_tablesorter").tablesorter({dateFormat: "uk",
							headers: {	0: {sorter: false},
										10: {sorter: false},
										11: {sorter: false},
										4: {sorter: 'grades' }
	 									}
		}
        )
		
	$.prettyLoader();
	$('#TX_TERMINO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_INICIO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
 
 	
})

function buscar(){
	$.prettyLoader.show(2000);
	datos.Buscar.disabled = true;
	datos.action='modulo_agendamientos.asp?strBuscar=S';
	datos.submit();

}

function limpiar(){
	$.prettyLoader.show(2000);
	datos.Limpiar.disabled = true;
	datos.action='modulo_agendamientos.asp?strLimpiar=S';
	datos.submit();

}


function SetCustomer(codigoCliente) {

	var topFrame = window.parent.frames['topFrame'];
	
	$('#CB_CLIENTE', topFrame.document).val(codigoCliente);

}

</script>

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">

<div class="titulo_informe">MÓDULO DE AGENDAMIENTOS</div>
<br>
	<table width="90%" align="CENTER" border="0" bgcolor="#f6f6f6" class="estilo_columnas">
		<thead>
			<tr>		
				<td width="170">COBRANZA</td>
				  <td width="170">TIPO COBRANZA</td>
				  <td width="170">RUBRO</td>
				  <td width="170">TIPO DOC</td>
				  <td width="170">CAMPAÑA</td>

				  <% If perfilEjecutivo = "0" Then %>
				  	<td width="150">EJECUTIVO</td>
				  <% Else %>
				  	<td width="150">CLIENTE</td>
				  <% End If %>
			</tr>
		</thead>
			<tr>
				<td>
					<select name="CB_COBRANZA" id="CB_COBRANZA"  <%If perfilEjecutivo = "0" then%> onChange="CargaUsuarios(CB_COBRANZA.value);" <%End If%> >
					
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
					<select name="CB_TIPOCOB" id="CB_TIPOCOB">
						<option value="">TODOS</option-->
						<%
						abrirscg()
						ssql="SELECT COD_ESTADO_COBRANZA, NOM_ESTADO_COBRANZA FROM ESTADO_COBRANZA ORDER BY 1"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("COD_ESTADO_COBRANZA")%>"<%if Trim(intEstadoCob)=Trim(rsTemp("COD_ESTADO_COBRANZA")) then response.Write("Selected") End If%>><%=rsTemp("NOM_ESTADO_COBRANZA")%></option>
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
					<select name="CB_RUBRO" id="CB_RUBRO">
					<option value="" <%if Trim(strRubro)="" then response.Write("Selected") end if%>>SELECCIONE</option>
						<%
						abrirscg()
						ssql="SELECT DISTINCT ISNULL(ADIC_2,'OTRO') AS ADIC_2 FROM DEUDOR  WHERE COD_CLIENTE IN ('" & intCodCliente & "') ORDER BY ADIC_2"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("ADIC_2")%>"<%if strRubro=rsTemp("ADIC_2") then response.Write("Selected") End If%>><%=rsTemp("ADIC_2")%></option>
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
					<select name="CB_TIPODOC" id="CB_TIPODOC">
						<option value="">TODOS</option-->
						<%
						abrirscg()

						strSql="SELECT DISTINCT COD_TIPO_DOCUMENTO, NOM_TIPO_DOCUMENTO"
						strSql=strSql & " FROM CUOTA LEFT JOIN TIPO_DOCUMENTO ON TIPO_DOCUMENTO = COD_TIPO_DOCUMENTO"
						strSql=strSql & " WHERE CUOTA.COD_CLIENTE IN ('" & intCodCliente & "')"
						strSql=strSql & " ORDER BY NOM_TIPO_DOCUMENTO ASC"

						set rsTemp= Conn.execute(strSql)
						if not rsTemp.eof then
							do until rsTemp.eof
							%>
							<option value="<%=rsTemp("COD_TIPO_DOCUMENTO")%>"<%if Trim(intTipoDoc)=Trim(rsTemp("COD_TIPO_DOCUMENTO")) then response.Write("Selected") End If%>><%=rsTemp("NOM_TIPO_DOCUMENTO")%></option>
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
					<select name="CB_CAMPANA" id="CB_CAMPANA">
						<option value="">TODAS</option>
						<%
						AbrirSCG()
							strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE IN ('" & intCodCliente & "')"
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

			<% If perfilEjecutivo="0" Then %>
				<td>
					<select name="CB_EJECUTIVO" id="CB_EJECUTIVO">
					</select>
				</td>
			<% Else %>

				<td>
					<select name="CB_CLIENTE" id="CB_CLIENTE">
						<option value="0">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
							set rsCliente=Conn.execute(strSql)
							Do While not rsCliente.eof
								If session("Ftro_Cliente")=rsCliente("COD_CLIENTE") Then strSelCam = "SELECTED" Else strSelCam = "0"
								%>
								<option value="<%=rsCliente("COD_CLIENTE")%>" <%if session("Ftro_Cliente")=trim(rsCliente("COD_CLIENTE")) then response.Write " Selected " End If%>><%=rsCliente("RAZON_SOCIAL")%></option>
								<%
								rsCliente.movenext
							Loop
							rsCliente.close
							set rsCliente=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>			
			
			<% End If %>

			</tr>
	</table>

	<table width="90%" align="CENTER" class="estilo_columnas">
		<thead>
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="340">NOMBRE O RAZON SOCIAL</td>
			<td width="170">RUT</td>
			<td width="170">FEC.INICIO</td>
			<td width="170">FEC.TERMINO</td>
			<td width="150">&nbsp;</td>
		</tr>
		</thead>
		<tr bgcolor="#f6f6f6" class="Estilo8">

			<td><input name="TX_NOMBRES" type="text" value="" size="48" maxlength="77"></td>

			<td><input name="TX_RUT_DEUDOR" type="text" value="" size="12" maxlength="12"></td>
			<td>
				<input name="TX_INICIO" readonly="true" type="text" id="TX_INICIO" value="<%=dtmInicio%>" size="10">
			</td>
			<td>
				<input name="TX_TERMINO" readonly="true" type="text" id="TX_TERMINO" value="<%=dtmTermino%>" size="10">
			</td>

			<td align="right" width="107"><input style="width:63px" class="fondo_boton_100" name="Buscar" type="button" value="Buscar"  onClick="buscar();"></td>

		</tr>
	</table>

	<table width="90%" align="CENTER" class="estilo_columnas">
		<thead>
			<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td width ="340">FILTRO ULTIMA GESTION</td>
				<td width="167">UBICABILIDAD</td>
				<td Width = "167" Colspan =1>TIPO AGENDAMIENTO</td>
				<td Width = "167" Colspan =1>PRIORIDAD</td>
				<td Width = "148">&nbsp;</td>
			</tr>
		</thead>
			<tr bgcolor="#f6f6f6" class="Estilo8">
				<td>
					<select name="CB_TIPOGESTION" onchange="this.style.width=300">

						<option value="" <%if Trim(intGestion) = ""  Then Response.write "SELECTED" %>>TODAS</option>
						<option value="SIN GESTION" <%if Trim(intGestion) = "SIN GESTION" Then Response.write "SELECTED" %>>SIN GESTION DOC</option>
						<option value="SIN GESTION EFECTIVA" <%if Trim(intGestion) = "SIN GESTION EFECTIVA" Then Response.write "SELECTED" %>>SIN GESTION EFECTIVA DOC</option>
						<option value="SIN GESTION CASO" <%if Trim(intGestion) = "SIN GESTION CASO" Then Response.write "SELECTED" %>>SIN GESTION CASO</option>
						<option value="SIN GESTION TEL CASO" <%if Trim(intGestion) = "SIN GESTION TEL CASO" Then Response.write "SELECTED" %>>SIN GESTION TELEFONICA CASO</option>
						<option value="SIN GESTION MAIL CASO" <%if Trim(intGestion) = "SIN GESTION MAIL CASO" Then Response.write "SELECTED" %>>SIN GESTION MAIL CASO</option>

						<%
						abrirscg()
							strSql = "SELECT DISTINCT * FROM GESTIONES_TIPO_GESTION WHERE COD_CLIENTE IN ('" & intCodCliente & "')"
							strSql = strSql & " AND ISNULL(GESTIONES_TIPO_GESTION.VER_AGEND,1) = 1"

							set rsGest = Conn.execute(strSql)
							''strCodComPago = ""

							Do While not rsGest.eof

								strSql = "SELECT DESCRIPCION FROM GESTIONES_TIPO_CATEGORIA WHERE COD_CATEGORIA = " & rsGest("COD_CATEGORIA")
								set rsTemp = Conn.execute(strSql)
								If Not rsTemp.Eof Then
									strNomCategoria = rsTemp("DESCRIPCION")
								End if

								strSql = "SELECT DESCRIPCION FROM GESTIONES_TIPO_SUBCATEGORIA WHERE COD_CATEGORIA = " & rsGest("COD_CATEGORIA") & " AND COD_SUB_CATEGORIA = " & rsGest("COD_SUB_CATEGORIA")
								set rsTemp = Conn.execute(strSql)
								If Not rsTemp.Eof Then
									strNomSubCategoria = rsTemp("DESCRIPCION")
								End if

								strNombreGestion = rsGest("DESCRIPCION")
								strGestionTotal = strNomCategoria & "-" & strNomSubCategoria & "-" & strNombreGestion
								'strGestionTotal = strNomSubCategoria & "-" & strNombreGestion
								strCodigo = rsGest("COD_CATEGORIA") & "*" & rsGest("COD_SUB_CATEGORIA") & "*" & rsGest("COD_GESTION")

								if strCodigo = Trim(intGestion) Then strGestSel="SELECTED" Else strGestSel=""
							%>
								<option value="<%=Trim(strCodigo)%>" <%=strGestSel%>><%=strGestionTotal%></option>

							<%
								rsGest.movenext
							Loop

						cerrarscg()
						%>
					</select>
				</td>

				<td>
				<select name="CB_UBICABILIDAD">
					<option value="0" <%if Trim(strTipoUbic)="0" then response.Write("Selected") end if%>>TODOS</option>
					<option value="1" <%if Trim(strTipoUbic)="1" then response.Write("Selected") end if%>>CONTACTADO</option>
					<option value="2" <%if Trim(strTipoUbic)="2" then response.Write("Selected") end if%>>NO CONTACTADO</option>
				</select>
				</td>

				<td>
				<select name="CB_TIPOAGEND">
					<option value="0" <%if Trim(strTipoAgend)="0" then response.Write("Selected") end if%>>TODOS</option>
					<option value="1" <%if Trim(strTipoAgend)="1" then response.Write("Selected") end if%>>NORMAL</option>
					<option value="2" <%if Trim(strTipoAgend)="2" then response.Write("Selected") end if%>>PREVENTIVO</option>
					<option value="3" <%if Trim(strTipoAgend)="3" then response.Write("Selected") end if%>>FUTURO</option>
				</select>
				</td>

				<td>
					<select name="CB_PRIORIDAD" onchange="this.style.width=105">
					<option value="">TODAS</option>
						<%
						AbrirSCG()
							strSql="SELECT DISTINCT PRIORIDAD_CUOTA AS PRIORIDAD_CUOTA FROM CUOTA WHERE COD_CLIENTE IN ('" & intCodCliente & "') AND PRIORIDAD_CUOTA IS NOT NULL ORDER BY PRIORIDAD_CUOTA ASC"
							set rspcuota=Conn.execute(strSql)
							Do While not rspcuota.eof
								%>
								<option value="<%=rspcuota("PRIORIDAD_CUOTA")%>"<%if replace(strPrioridad, ".", ",")=Trim(rspcuota("PRIORIDAD_CUOTA")) then response.Write("Selected") End If%>><%=rspcuota("PRIORIDAD_CUOTA")%></option>
								<%
								rspcuota.movenext
							Loop
							rspcuota.close
							set rspcuota=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>

				<td align="right" Width = "105"><input name="Limpiar" class="fondo_boton_100" type="button" value="Limpiar"  onClick="limpiar();"></td>
			</tr>
	</table>

	<table width="1000" align="CENTER">
		<tr class="Estilo13">
			<td width="60%" align="center"><%=strMensaje%></td>
		</tr>
	</table>
<br/>
				<%
					AbrirSCG()					

					strSql = "SELECT"
					strSql = strSql & " RUT_DEUDOR = D.RUT_DEUDOR,"
					strSql = strSql & " COD_CLIENTE = CL.COD_CLIENTE,"
					strSql = strSql & " NOMBRE_DEUDOR = D.NOMBRE_DEUDOR,"
					strSql = strSql & " SALDO = SUM(SALDO),"
					strSql = strSql & " DOC = COUNT(C.ID_CUOTA),"
					strSql = strSql & " DIAVENC = MAX(DATEDIFF(DAY,FECHA_VENC,GETDATE())),"
					strSql = strSql & " PRIORIDAD_CUOTA = MIN(ISNULL(C.PRIORIDAD_CUOTA,11)),"
					strSql = strSql & " FEC_AGEND = CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),103),"
					strSql = strSql & " FEC_AGEND2 = MIN((ISNULL(C.FECHA_AGEND_ULT_GES,GETDATE()+300) + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),"

					strSql = strSql & " DIFMINUTOS = IsNull(datediff(minute,D.FECHA_CONF,IsNull(D.FECHA_UG_TITULAR,'01/01/1900')),0),"
					strSql = strSql & " OBSERVACIONES_CONF = D.OBSERVACIONES_CONF,"
					strSql = strSql & " FECHA_CONF = D.FECHA_CONF,"
					strSql = strSql & " USUARIO_CONF = D.USUARIO_CONF,"
					strSql = strSql & " PRIORIZACION = ISNULL(D.RESP_EMAIL,0),"
					strSql = strSql & " NOM_GEST = MAX(GTC.DESCRIPCION+'-'+GTSC.DESCRIPCION+'-'+GTG.DESCRIPCION),"
					strSql = strSql & " HORA_AGEND = "
					strSql = strSql & " (CASE WHEN CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),108) = '00:00:00'"
					strSql = strSql & " 	  THEN ''"
					strSql = strSql & " 	  WHEN SUBSTRING(CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),108),5,1)= ':'"
					strSql = strSql & " 	  THEN SUBSTRING(CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),108),1,4)"
					strSql = strSql & " 	  ELSE SUBSTRING(CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),108),1,5)"
					strSql = strSql & " END),"
					strSql = strSql & " UGT = MAX(C.COD_ULT_GEST),"
					strSql = strSql & " FONOS_VAL = (select count(ID_TELEFONO) from deudor_telefono where RUT_DEUDOR=D.rut_deudor and Estado in (1)),"
					strSql = strSql & " EMAIL_VAL = (select count(ID_EMAIL) from deudor_email where RUT_DEUDOR=D.rut_deudor and Estado in (1))"

					strSql = strSql & " FROM DEUDOR D	INNER JOIN CLIENTE	CL						ON D.COD_CLIENTE = CL.COD_CLIENTE"
					strSql = strSql & " 				INNER JOIN CUOTA C							ON D.RUT_DEUDOR = C.RUT_DEUDOR AND D.COD_CLIENTE = C.COD_CLIENTE"
					strSql = strSql & " 				INNER JOIN ESTADO_DEUDA ED					ON C.ESTADO_DEUDA = ED.CODIGO"
					strSql = strSql & " 				LEFT JOIN GESTIONES G						ON C.ID_ULT_GEST_GENERAL = G.ID_GESTION"
					strSql = strSql & " 				LEFT JOIN GESTIONES_TIPO_CATEGORIA GTC		ON G.COD_CATEGORIA = GTC.COD_CATEGORIA"
					strSql = strSql & " 				LEFT JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC  ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA"
					strSql = strSql & " 																AND G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA"
					strSql = strSql & " 				LEFT JOIN GESTIONES_TIPO_GESTION GTG		ON G.COD_CATEGORIA = GTG.COD_CATEGORIA"
					strSql = strSql & " 																AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA"
					strSql = strSql & " 																AND G.COD_GESTION = GTG.COD_GESTION"
					strSql = strSql & " 																AND G.COD_CLIENTE = GTG.COD_CLIENTE"					

					strSql = strSql & " WHERE ED.ACTIVO=1"
					strSql = strSql & " AND CL.ACTIVO = 1"
					
					If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then
					
						strSql = strSql & " AND D.COD_CLIENTE IN ('" & intCodCliente & "')"

					ElseIf session("Ftro_Cliente") <> "0" and session("Ftro_Cliente") <> "" then
					
						strSql = strSql & " AND D.COD_CLIENTE IN ('" & session("Ftro_Cliente") & "') and 1=1"
					
					End if

					If strPrioridad <> ""  Then

						strSql = strSql & " AND C.PRIORIDAD_CUOTA = " & strPrioridad

					End If

					If Trim(strCobranza) = "INTERNA" Then
						strSql = strSql & " AND D.CUSTODIO IS NOT NULL"
						strParametro = "1"
					End if

					If Trim(strCobranza) = "EXTERNA" Then
						strSql = strSql & " AND D.CUSTODIO IS NULL"
						strParametro = "1"
					End if

					If strTipoUbic = "0" or strTipoUbic = "" Then

						strSql = strSql & " AND (D.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0,1))"
						strSql = strSql & " OR D.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0,1)))"

					ElseIf strTipoUbic = "1" Then

						strSql = strSql & " AND (D.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (1))"
						strSql = strSql & " OR D.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (1)))"

					ElseIf strTipoUbic = "2" Then

						strSql = strSql & " AND (D.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0))"
						strSql = strSql & " OR D.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0)))"

						strSql = strSql & " AND (D.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (1))"
						strSql = strSql & " AND D.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (1)))"

					End If
					
					If strTipoAgend <> "3" Then

					strSql = strSql & " AND (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0"
					strSql = strSql & " AND ISNULL(GTG.VER_AGEND,1) = 1)"

					strSql = strSql & " AND C.ESTADO_DEUDA IN (SELECT ESTADO_DEUDA.CODIGO FROM ESTADO_DEUDA WHERE ESTADO_DEUDA.ACTIVO = 1)"
					
					Else

					strSql = strSql & " AND (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) < 0"
					strSql = strSql & " AND ISNULL(GTG.VER_AGEND,1) = 1)"					
					
					End If

					If strTipoAgend = "1" Then

						strSql = strSql & " AND D.RUT_DEUDOR NOT IN (SELECT RUT_DEUDOR"
						strSql = strSql & " FROM CUOTA LEFT JOIN GESTIONES_TIPO_GESTION 	  ON SUBSTRING(CUOTA.COD_ULT_GEST,1,1) = GESTIONES_TIPO_GESTION.COD_CATEGORIA"
						strSql = strSql & " 													 AND SUBSTRING(CUOTA.COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
						strSql = strSql & " 													 AND SUBSTRING(CUOTA.COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
						strSql = strSql & " 													 AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
						strSql = strSql & " WHERE CUOTA.COD_CLIENTE = 1100 AND ISNULL(GESTIONES_TIPO_GESTION.VER_AGEND,1) = 1"
						strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (SELECT ESTADO_DEUDA.CODIGO FROM ESTADO_DEUDA WHERE ESTADO_DEUDA.ACTIVO = 1)"
						strSql = strSql & " GROUP BY RUT_DEUDOR"
						strSql = strSql & " HAVING MAX((CAST((CAST(convert(varchar(10), getdate(),103) AS DATETIME)-FECHA_VENC) AS INT)))<-5)"

					ElseIf strTipoAgend = "2" Then

						strSql = strSql & " AND CUOTA.RUT_DEUDOR IN (SELECT RUT_DEUDOR"
						strSql = strSql & " FROM CUOTA LEFT JOIN GESTIONES_TIPO_GESTION 	  ON SUBSTRING(CUOTA.COD_ULT_GEST,1,1) = GESTIONES_TIPO_GESTION.COD_CATEGORIA"
						strSql = strSql & " 													 AND SUBSTRING(CUOTA.COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
						strSql = strSql & " 													 AND SUBSTRING(CUOTA.COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
						strSql = strSql & " 													 AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
						strSql = strSql & " WHERE CUOTA.COD_CLIENTE = 1100 AND ISNULL(GESTIONES_TIPO_GESTION.VER_AGEND,1) = 1"
						strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (SELECT ESTADO_DEUDA.CODIGO FROM ESTADO_DEUDA WHERE ESTADO_DEUDA.ACTIVO = 1)"
						strSql = strSql & " GROUP BY RUT_DEUDOR"
						strSql = strSql & " HAVING MAX((CAST((CAST(convert(varchar(10), getdate(),103) AS DATETIME)-FECHA_VENC) AS INT)))<-5)"

					End If

					If trim(strEjeAsig) = "0" OR trim(strEjeAsig) = "" Then
						If TraeSiNo(session("perfil_adm")) <> "Si" and TraeSiNo(session("perfil_sup")) <> "Si"Then
							strSql = strSql & " AND C.USUARIO_ASIG = " & session("session_idusuario")
						End If
					Else
						strSql = strSql & " AND C.USUARIO_ASIG = " & strEjeAsig
					End if

					strParametro = "0"

					If Trim(strNombres) <> "" Then
						strSql = strSql & " AND D.NOMBRE_DEUDOR  LIKE '%" & strNombres & "%'"
						strParametro = "1"
					End if

					If Trim(strRut) <> "" Then
						strSql = strSql & " AND C.RUT_DEUDOR  LIKE '" & strRut & "%'"
						strParametro = "1"
					End if

					If Trim(intEstadoCob) <> "0" and Trim(intEstadoCob) <> "" Then
						strSql = strSql & " AND D.ETAPA_COBRANZA = " & intEstadoCob
						strParametro = "1"
					End if

					If Trim(intTipoDoc) <> "0" and Trim(intTipoDoc) <> "" Then
						strSql = strSql & " AND C.TIPO_DOCUMENTO = '" & intTipoDoc & "'"
						strParametro = "1"
					End if

					If Trim(intCodCampana) <> "0" and Trim(intCodCampana) <> "" Then
						strSql = strSql & " AND C.RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE IN ('" & intCodCliente & "') AND ID_CAMPANA = " & intCodCampana & ")"
						strParametro = "1"
					End if

					If Trim(intGestion) = "SIN GESTION TEL CASO" Then
						strSql = strSql & " AND [dbo].[fun_trae_estatus_gestion] (D.COD_CLIENTE,D.RUT_DEUDOR,'TELEFONICA') = 0"

					ElseIf Trim(intGestion) = "SIN GESTION CASO" Then
						strSql = strSql & " AND [dbo].[fun_trae_estatus_gestion] (D.COD_CLIENTE,D.RUT_DEUDOR,'GENERAL') = 0"

					ElseIf Trim(intGestion) = "SIN GESTION MAIL CASO" Then
						strSql = strSql & " AND [dbo].[fun_trae_estatus_gestion] (D.COD_CLIENTE,D.RUT_DEUDOR,'MAIL') = 0"

					ElseIf Trim(intGestion) = "SIN GESTION" Then
						strSql = strSql & " AND (ID_ULT_GEST_GENERAL IS NULL OR ID_ULT_GEST_GENERAL=0)"

					ElseIf Trim(intGestion) = "SIN GESTION EFECTIVA" Then
						strSql = strSql & " AND (ID_ULT_GEST_EFE IS NULL OR ID_ULT_GEST_EFE=0)"

					ElseIf Trim(intGestion) <> "" Then
						strSql = strSql & " AND C.COD_ULT_GEST= '" & intGestion & "'"
					End If

					If Trim(dtmInicio) <> "" Then
						strSql = strSql & " AND FECHA_AGEND_ULT_GES >= '" & dtmInicio & " 00:00:00'"
					End If

					If Trim(dtmTermino) <> "" Then
						strSql = strSql & " AND FECHA_AGEND_ULT_GES <= '" & dtmTermino & " 23:58:59'"
					End If

					If Trim(strRubro) <> "" Then
						strSql = strSql & " AND D.ADIC_2 = '" & strRubro & "'"
					End If

					strSql = strSql & " GROUP BY D.FECHA_UG_TITULAR,D.OBSERVACIONES_CONF, D.FECHA_CONF, D.USUARIO_CONF,RESP_EMAIL, D.RUT_DEUDOR,D.NOMBRE_DEUDOR, CL.COD_CLIENTE"

					strSql = strSql & " ORDER BY CL.COD_CLIENTE, C.PRIORIDAD_CUOTA ASC,DIAVENC DESC, FEC_AGEND2 ASC, SUM(SALDO) DESC"					

					'RESPONSE.WRITE "strSql=" & strSql
					'RESPONSE.END

					set rsCuota=Server.CreateObject("ADODB.Recordset")
					rsCuota.Open strSql, Conn, 1, 2
					intTotalSaldo = 0
					intTotalRut = 0

					' Defino el tama?o de las p?ginas
					rsCuota.PageSize=TamPagina
					rsCuota.CacheSize=TamPagina
					PaginasTotales=rsCuota.PageCount
					''Response.write "PaginaActual=" & PaginasTotales

					'Compruebo que la pagina actual est? en el rango
					if PaginaActual < 1 then
						PaginaActual = 1
					end if
					if PaginaActual > PaginasTotales then
						PaginaActual = PaginasTotales
					end if

					'Por si la consulta no devuelve registros!
					if PaginasTotales=0 then
						strMensaje = "No hay Casos Agendados Para Gestionar"
					else
						rsCuota.AbsolutePage=PaginaActual
					End If

					sintPagina = PaginaActual
					sintTotalPaginas = PaginasTotales
					%>
								<!-- Paginado -->
								<TABLE width="90%" align="CENTER" >
									<TR BGCOLOR="#<%=strColorMosPrev%>">
										<TD WIDTH="20%" ALIGN=left>
											<%if PaginaActual > 1 then %>
											<INPUT TYPE=BUTTON NAME="Retroceder" VALUE="  &lt;  " onClick="IrPagina( 'Retroceder')">
											<% end if %>
										</TD>
										<TD WIDTH="60%" height = "20" ALIGN=center>
											<FONT FACE="verdana, Sans-Serif" Size=1 COLOR="#<%=strLetrasColorModPrev%>"><b>Página <%= sintPagina %> de <%= sintTotalPaginas %> <%= strMensajeModPrev%> </b></FONT>
										</TD>
										<TD WIDTH="20%" ALIGN=right>
											<%if PaginaActual < PaginasTotales then%>
											<INPUT TYPE=BUTTON NAME="Avanzar" VALUE="  &gt;  " onClick="IrPagina( 'Avanzar')">
											<% end if %>
										</TD>
									</TR>
								</TABLE>
					  <table  width="90%"  align="CENTER" class="tablesorter"  id="table_tablesorter">
					  	<thead>
						<tr >
							<th  width="50px">CONT. </th>
							<th  id="rut" align="center">RUT</th>
							<th  width="350">NOMBRE O RAZON SOCIAL </th>
							<th  align="center">DOC.</th>
							<th  id="SALDO" align="center">SALDO</th>
							<th  width="100px" align="center">ULT.GESTION </th>
							<th  align="center">F.AGEND.</th>
							<th  align="center">H.AGEND. </th>
							<th  align="center">PRIOR.</th>
							<th>ATENCION</th>
							<th width = "20">&nbsp;</th>
							<th>&nbsp</th>
						</tr>
						</thead>
				<tbody>
	
					<%
						'Response.write "valor_moneda=" & session("valor_moneda")
						'Response.write "SALDO=" & rsCuota("SALDO")

						'session("valor_moneda") = 22000
						'session("valor_moneda") = 1

						If Not rsCuota.eof Then
							totalventa=0

							Do while not rsCuota.eof and CuantosRegistros < TamPagina


								strNomGestion = rsCuota("NOM_GEST")
								strCodGestion = rsCuota("UGT")
								dtmFecAgend = rsCuota("FEC_AGEND")
								dtmHoraAgend = rsCuota("HORA_AGEND")

								intMinDif = rsCuota("DIFMINUTOS")
								strObsConf = rsCuota("OBSERVACIONES_CONF")
								strPriorizacion = rsCuota("PRIORIZACION")
								strFechaConf = rsCuota("FECHA_CONF")
								strUsuarioConf = rsCuota("USUARIO_CONF")
								strTextoConf=""
								If Trim(strFechaConf) <> "" and Trim(strUsuarioConf) <> "" then
									strTextoConf = "Fecha : " & strFechaConf & " , Usuario : " & strUsuarioConf & ", Obs : "
								End If

								intValorSaldo = Round(session("valor_moneda") * ValNulo(rsCuota("SALDO"),"N"),0)
								'intValorSaldo =  rsCuota("SALDO")
								'Response.write "SALDO=" & intValorSaldo
								intTotalSaldo = intTotalSaldo + intValorSaldo
								intValorDoc = rsCuota("DOC")
								intTotalDoc = intTotalDoc + intValorDoc
								intTotalRut = intTotalRut + 1


								if rsCuota("FONOS_VAL") > 0 OR rsCuota("EMAIL_VAL") > 0 Then
									strContactado = "tel_contactado.jpg"
								Else
									strContactado = "tel_nocontactado.jpg"
								End If

								AbrirSCG1()
									strSql = "SELECT PR.ID_PRIORIZACION,PR.OBSERVACION_PRIORIZACION, USUARIO.LOGIN, PR.FECHA_PRIORIZACION,TSP.NOM_TIPO_SOLICITUD"
									strSql= strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
									strSql= strSql & " 					 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
									strSql= strSql & " 					 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
									strSql= strSql & " 					 INNER JOIN USUARIO ON PR.ID_USUARIO_PRIORIZACION = USUARIO.ID_USUARIO"
									strSql= strSql & " 					 INNER JOIN TIPO_SOLICITUD_PRIORIZACION TSP ON TSP.ID_TIPO_SOLICITUD = PR.ID_TIPO_SOLICITUD"

									strSql= strSql & " WHERE PR.RUT_DEUDOR = '" & rsCuota("RUT_DEUDOR") & "' AND PRC.ESTADO_PRIORIZACION = 0 AND PR.COD_CLIENTE IN ('" & intCodCliente & "') AND ESTADO_DEUDA.ACTIVO = 1"
									strSql= strSql & " 					 GROUP BY PR.ID_PRIORIZACION,PR.OBSERVACION_PRIORIZACION,USUARIO.LOGIN, PR.FECHA_PRIORIZACION, TSP.NOM_TIPO_SOLICITUD"

									'Response.write "<br>strSql=" & strSql
									set RsPrio=Conn1.execute(strSql)

									strTextoPrioF = ""
									intEstadoPrior = 0

									If not RsPrio.eof then

										Do While Not RsPrio.Eof

										intEstadoPrior = 1
										strObsPrio = RsPrio("OBSERVACION_PRIORIZACION")
										strUsuarioPrio = RsPrio("LOGIN")
										strFechaPrio = RsPrio("FECHA_PRIORIZACION")
										strTipoSol = RsPrio("NOM_TIPO_SOLICITUD")

										If Trim(strFechaPrio) <> "" and Trim(strUsuarioPrio) <> "" then
											strTextoPrio = "Fecha: " & strFechaPrio & " , Usuario : " & strUsuarioPrio & chr(13) & "Tipo Sol: " & strTipoSol & chr(13) & "Obs : " & strObsPrio & chr(13) & strTotalDoc & chr(13) & chr(13)

											strTextoPrioF = strTextoPrioF & strTextoPrio
										End If

											RsPrio.movenext
										Loop
									End If

								CerrarSCG1()

								AbrirSCG1()

									strSql = " SELECT TOP 1"
									strSql = strSql & " (CASE WHEN   (DATEDIFF(MINUTE,(CONVERT(VARCHAR(10),GETDATE(),103) +' '+ convert(varchar(10),(CASE WHEN HORA_DESDE = '' THEN '22:00' ELSE ISNULL(HORA_DESDE,'22:00') END),108)),GETDATE())) >= 0"
									strSql = strSql & " 			  AND (DATEDIFF(MINUTE,(CONVERT(VARCHAR(10),GETDATE(),103) +' '+ convert(varchar(10),(CASE WHEN HORA_HASTA = '' THEN '22:00' ELSE ISNULL(HORA_HASTA,'22:00') END),108)),GETDATE())) < 0"
									strSql = strSql & " 			  AND ISNULL(DIAS_ATENCION,'') LIKE  '%' + SUBSTRING(DATENAME(weekday, GETDATE()),1,2) + '%'"
									strSql = strSql & " 	  THEN 3"
									strSql = strSql & " 	  WHEN ISNULL(DIAS_ATENCION,'') = ''"
									strSql = strSql & " 	  THEN 0"
									strSql = strSql & " 	  WHEN ISNULL(HORA_DESDE,'') <> ''"
									strSql = strSql & " 	  THEN 2"
									strSql = strSql & " 	  ELSE 1"
									strSql = strSql & "  END) AS ORDEN,"
									strSql = strSql & " CASE WHEN DIAS_ATENCION = SUBSTRING(DATENAME(weekday, GETDATE()),1,2) THEN"
									strSql = strSql & " 	 'Solo Hoy'"
									strSql = strSql & " 	 WHEN ISNULL(HORA_DESDE,'') = '' AND ISNULL(DIAS_ATENCION,'') LIKE  '%' + SUBSTRING(DATENAME(weekday, GETDATE()),1,2) + '%' THEN"
									strSql = strSql & " 	 'Hoy sin horario'"
									strSql = strSql & " 	 WHEN ISNULL(DIAS_ATENCION,'') LIKE  '%' + SUBSTRING(DATENAME(weekday, GETDATE()),1,2) + '%' THEN"
									strSql = strSql & " 	 'Hoy de ' + HORA_DESDE+ ' a '+ HORA_HASTA"
									strSql = strSql & " 	 WHEN ISNULL(DIAS_ATENCION,'')<> ''THEN"
									strSql = strSql & " 	 'No atiende hoy'"
									strSql = strSql & " 	 ELSE"
									strSql = strSql & " 	 'No definido'"
									strSql = strSql & " END AS ATENCION"
									strSql = strSql & " FROM DEUDOR_TELEFONO"
									strSql = strSql & " WHERE RUT_DEUDOR = '" & rsCuota("RUT_DEUDOR") & "'"
									strSql = strSql & " 	  AND ESTADO IN (1,0)"

									strSql = strSql & " ORDER BY "
									strSql = strSql & " (CASE WHEN   (DATEDIFF(MINUTE,(CONVERT(VARCHAR(10),GETDATE(),103) +' '+ convert(varchar(10),(CASE WHEN HORA_DESDE = '' THEN '22:00' ELSE ISNULL(HORA_DESDE,'22:00') END),108)),GETDATE())) >= 0"
									strSql = strSql & " 			AND (DATEDIFF(MINUTE,(CONVERT(VARCHAR(10),GETDATE(),103) +' '+ convert(varchar(10),(CASE WHEN HORA_HASTA = '' THEN '22:00' ELSE ISNULL(HORA_HASTA,'22:00') END),108)),GETDATE())) < 0"
									strSql = strSql & " 			AND ISNULL(DIAS_ATENCION,'') LIKE  '%' + SUBSTRING(DATENAME(weekday, GETDATE()),1,2) + '%'"
									strSql = strSql & " 	  THEN 3"
									strSql = strSql & " 	  WHEN ISNULL(DIAS_ATENCION,'') = ''"
									strSql = strSql & " 	  THEN 0"
									strSql = strSql & " 	  WHEN ISNULL(HORA_DESDE,'') <> ''"
									strSql = strSql & " 	  THEN 2"
									strSql = strSql & " 	  ELSE 1"
									strSql = strSql & " END) DESC"


									'RESPONSE.WRITE "strSql=" & strSql

									set rsFonos = Conn.execute(strSql)

									If Not rsFonos.Eof Then

									strAtencion = rsFonos("ATENCION")

									Else

									strAtencion = "No definido"

									End if

								CerrarSCG1()
								
								%>
				<tr bgcolor="<%=strbgcolor%>" class="Estilo8">

										<td ALIGN="center"><img src="../imagenes/<%=strContactado%>" border="0"></td>

										<td>
											<A HREF="principal.asp?TX_RUT=<%=rsCuota("RUT_DEUDOR")%>&cliente=<%=rsCuota("COD_CLIENTE")%>" onclick="javascript:SetCustomer('<%=rsCuota("COD_CLIENTE")%>');">
												<acronym title="Llevar a pantalla principal"><%=rsCuota("RUT_DEUDOR")%></acronym>
											</A>
										</td>										
										
										<td><%=Mid(rsCuota("NOMBRE_DEUDOR"),1,30)%></td>
										<td ALIGN="right"><%=FN(intValorDoc,0)%></td>
										<td ALIGN="right"><%=FN(intValorSaldo,0)%></td>
										<td ALIGN="center">
										<acronym title="<%=strNomGestion%>">
										<%=strCodGestion%>
										</acronym>
										</td>
										<td ALIGN="center"><%=dtmFecAgend%></td>
										<td ALIGN="center"><%=dtmHoraAgend%></td>
										<td ALIGN="center"><b><%=rsCuota("PRIORIDAD_CUOTA")%></b></td>
										<td ALIGN="LEFT"><%=strAtencion%></td>
										<td>

											<% If intEstadoPrior = "1" then %>
												<abbr title="<%=strTextoPrioF%>">
												<img src="../imagenes/priorizar_urgente.png" border="0">
												<abbr>
											<% End If %>

										</td>
										<td ALIGN="center">
											<a href="detalle_gestiones.asp?rut=<%=rsCuota("RUT_DEUDOR")%>&cliente=<%=rsCuota("COD_CLIENTE")%>" onclick="javascript:SetCustomer('<%=rsCuota("COD_CLIENTE")%>');">
												<acronym title="Llevar a pantalla de ingreso de gestión">Seleccionar</acronym>
											</A>
										</td>			
								<%
								CuantosRegistros=CuantosRegistros+1
								rsCuota.movenext
										Loop
									End If
								rsCuota.close
								set rsCuota=NOTHING
					%>
					 </tr>
			 </tbody>
					<tr class="totales">
						<td Colspan = "2" >Totales</td>
						<td Colspan = "2" align="right"  colspan=2>Documentos Agendados :<%=FN(intTotalDoc,0)%> </td>
						<td Colspan = "3" align="right"  colspan=2>Saldo Agendados : $<%=FN(intTotalSaldo,0)%></td>
						<td align="center" colspan=7>Total Rut : <%=intTotalRut%> </td>
					</tr>
				
				<!-- Paginado -->
				<TR>
						<TD COLSPAN=12>
							<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="100%">
								<TR BGCOLOR="#F3F3F3">
									<TD WIDTH="20%" ALIGN=left>
										<%if PaginaActual > 1 then %>
										<INPUT TYPE=BUTTON NAME="Retroceder" VALUE="  &lt;  " onClick="IrPagina( 'Retroceder')">
										<% end if %>
									</TD>
									<TD WIDTH="60%" ALIGN=center>
										<FONT FACE="verdana, Sans-Serif" Size=1 COLOR="#FF0000"><b>Página <%= sintPagina %> de <%= sintTotalPaginas %></b></FONT>
									</TD>
									<TD WIDTH="20%" ALIGN=right>
										<%if PaginaActual < PaginasTotales then%>
										<INPUT TYPE=BUTTON NAME="Avanzar" VALUE="  &gt;  " onClick="IrPagina( 'Avanzar')">
										<% end if %>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
			</table>
			

</form>
</body>
</html>

<script>

    $.tablesorter.addParser({
        // set a unique id 
        id: 'grades',
        is: function (s) {
            // return false so this parser is not auto detected 
            return false;
        },
        format: function (s) {
            // format your data for normalization 
            s = s.replace(/\./g, "");
            s = s.replace(/\,/g, ".");

            return s //$.tablesorter.formatFloat(s.replace(new RegExp(/[^0-9.]/g),""));
        },
        
        type: 'numeric'
    }); 

</script>


<script language="JavaScript1.2">

function IrPagina( sintAccion ) {
	$.prettyLoader.show();
	if (sintAccion == 'Retroceder') {
	
    	self.location.href = 'modulo_agendamientos.asp?pagina=<%=PaginaActual - 1%>&TX_NOMBRES=<%=strNombres%>&CB_REMESA=<%=intCodRemesa%>&CB_UBICABILIDAD=<%=strTipoUbic%>&CB_EJECUTIVO=<%=strEjeAsig%>&CB_CAMPANA=<%=intCodCampana%>&CB_TIPOCARTERA=<%=strTipoInf%>&TX_INICIO=<%=dtmInicio%>&TX_TERMINO=<%=dtmTermino%>&CB_TIPOGESTION=<%=intGestion%>&CB_PRIORIDAD=<%=strPrioridad%>&CB_TIPOGESTION_PRINC=<%=intGestionPrinc%>'
    }
    if (sintAccion == 'Avanzar') {
	    self.location.href = 'modulo_agendamientos.asp?pagina=<%=PaginaActual + 1%>&TX_NOMBRES=<%=strNombres%>&CB_REMESA=<%=intCodRemesa%>&CB_UBICABILIDAD=<%=strTipoUbic%>&CB_EJECUTIVO=<%=strEjeAsig%>&CB_CAMPANA=<%=intCodCampana%>&CB_TIPOCARTERA=<%=strTipoInf%>&TX_INICIO=<%=dtmInicio%>&TX_TERMINO=<%=dtmTermino%>&CB_TIPOGESTION=<%=intGestion%>&CB_PRIORIDAD=<%=strPrioridad%>&CB_TIPOGESTION_PRINC=<%=intGestionPrinc%>'
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
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE IN ('" & intCodCliente & "')"

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
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE IN ('" & intCodCliente & "')"

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
			var newOption = new Option('SIN USUARIO', '');
			comboBox.options[comboBox.options.length] = newOption;
						
		}
		else {
		
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;
			
			<%

			AbrirSCG2()

			strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE IN ('" & intCodCliente & "')"

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
			
			$(comboBox).val('');
		}
	
}

function InicializaInforme()
{
		var comboBox = document.getElementById('CB_EJECUTIVO');
		comboBox.options.length = 0;
		var newOption = new Option('TODOS','');
		comboBox.options[comboBox.options.length] = newOption;
}

<%If perfilEjecutivo = "0" then%>
CargaUsuarios('<%=strCobranza%>');
<%End If%>

<%If strEjeAsig <> "" then
	if strEjeAsig = "0" then
		strEjeAsig = ""
	end if
%>
datos.CB_EJECUTIVO.value='<%=strEjeAsig%>';
<%End If%>
</script>