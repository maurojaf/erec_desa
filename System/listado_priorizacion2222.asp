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
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/comunes/rutinas/chkFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/sondigitos.inc"-->
	<!--#include file="../lib/comunes/rutinas/formatoFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/validarFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	AbrirSCG()

	sucursal = request("cmb_sucursal")
	intTipoPago = request("CB_TIPOPAGO")

	strEjeAsig = request("CB_EJECUTIVO")
	usuario_sol = request("cmb_usuario_sol")
	Gestionado = request ("cmb_Gestionado")

	If request("cmb_tipoPriorizacion") = "" then
		strTipoPriorizacion = 1
	Else
		strTipoPriorizacion = request("cmb_tipoPriorizacion")
	End If

	'Response.write "strTipoPriorizacion=" & strTipoPriorizacion

	if usuario_sol = "" then usuario_sol = "0"
	if Gestionado = "" then Gestionado = "2"

	termino = request("termino")
	inicio = request("inicio")

	strCodCliente = session("ses_codcli")

	If Trim(Request("strBuscar")) = "S" Then
			session("Ftro_Gestionado") = Gestionado
	End If

	If Trim(Request("strBuscar")) = "N" Then
			session("Ftro_Gestionado") = ""
	End If

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
		intVerCobExt = "0"

	End If

	If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

		sinCbUsario="0"

	End If

	'---Fin codigo tipo de cobranza---'

%>
<title>CRM Cobros</title>
<style type="text/css">
<!--
.Estilo13 {color: #FFFFFF}
.Estilo27 {color: #FFFFFF}
-->
</style>

<script language="JavaScript" src="../javascripts/cal2.js"></script>
<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
<script language="JavaScript" src="../javascripts/validaciones.js"></script>
<script src="../javascripts/SelCombox.js"></script>
<script src="../javascripts/OpenWindow.js"></script>

<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<script language="JavaScript " type="text/JavaScript">
$(document).ready(function(){

	$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$(document).tooltip();
 
})
function envia()
{
	resp='si'
	document.datos.action = "listado_priorizacion.asp?resp="+ resp +"";
	document.datos.submit();
}

</script>


</head>
<body>
<form name="datos" method="post">


<div class="titulo_informe">INFORME DE CASOS PRIORIZADOS</div>
<br>
			<table width="90%" align="center" border="0" bordercolor="#999999" class="estilo_columnas">
				<thead>
				  <tr height="20">

				  	<td align="center">COBRANZA</td>
				  	<td align="center">TIPO GESTION</td>
					<td align="center">FECHA DESDE</td>
					<td align="center">FECHA HASTA</td>

					 <% If sinCbUsario = "0" Then %>
						<td align="center">SOLICITANTE</td>
						<td>EJECUTIVO</td>
					  <% End If %>
					<td>&nbsp;</td>
				  </tr>
				</thead>
				  <tr>

					<td>
						<select name="CB_COBRANZA" <%If sinCbUsario = "0" then%> onChange="CargaUsuarios(CB_COBRANZA.value);" <%End If%> >

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
						<SELECT NAME="cmb_tipoPriorizacion" id="cmb_tipoPriorizacion" onChange="envia();">
							<option value="0" <%If Trim(strTipoPriorizacion)="0" Then Response.write "SELECTED"%>>TODOS</option>
							<option value="1" <%If Trim(strTipoPriorizacion)="1" Then Response.write "SELECTED"%>>PRIORIZACION PENDIENTE</option>
							<option value="2" <%If Trim(strTipoPriorizacion)="2" Then Response.write "SELECTED"%>>PRIORIZACION AGENDADA</option>
							<option value="3" <%If Trim(strTipoPriorizacion)="3" Then Response.write "SELECTED"%>>PRIORIZACION CONFIRMADA</option>
						</SELECT>
					</td>

					<td><input name="inicio" readonly="true" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
					</td>

					<td><input name="termino" readonly="true" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
					</td>

				  <% If sinCbUsario="0" Then %>
					<td>
						<SELECT NAME="cmb_usuario_sol" id="cmb_usuario_sol" onChange="envia();">
							<option value="0">TODOS</option>
							<%
							abrirscg()

							strSql="SELECT USUARIO.ID_USUARIO,USUARIO.LOGIN "
							strSql=strSql & " FROM USUARIO INNER JOIN USUARIO_CLIENTE ON USUARIO.ID_USUARIO = USUARIO_CLIENTE.ID_USUARIO "
							strSql=strSql & " AND USUARIO_CLIENTE.COD_CLIENTE = '" & strCodCliente & "' WHERE ACTIVO = 1 AND (PERFIL_SUP = 1 or PERFIL_ADM = 1)"

							strSql=strSql & " ORDER BY PERFIL_SUP, PERFIL_ADM, LOGIN"
							'Response.write "strSql=" & strSql

							set rsUsu=Conn.execute(strSql)
							if not rsUsu.eof then
								do until rsUsu.eof
								%>
								<option value="<%=rsUsu("LOGIN")%>"
								<%if Trim(usuario_sol)=Trim(rsUsu("LOGIN")) then
									response.Write("Selected")
								end if%>
								><%=ucase(rsUsu("LOGIN"))%></option>

								<%rsUsu.movenext
								loop
							end if
							rsUsu.close
							set rsUsu=nothing

							cerrarscg()
							%>
						</SELECT>
					</td>

					<td>
						<select name="CB_EJECUTIVO" id="CB_EJECUTIVO">
						</select>
					</td>
					<% End If %>

					<td align="center">
						<input type="Button" name="Submit" class="fondo_boton_100" value="Ver" onClick="envia();">
					</td>

				  </tr>
			</table>


			<%If strTipoPriorizacion = 1 or strTipoPriorizacion = 0 then%>

				<table width="100%" border="0" class="intercalado">
				<thead>
					<tr>

						<td align="center" Height = "25" colspan = "11" class="subtitulo_informe">> PRIORIZACIONES NO CONFIRMADAS / PENDIENTES A GESTIONAR</td>
					</tr>
				</thead>

					<%AbrirSCG()

						strSql = "SELECT PR.RUT_DEUDOR, PR.ID_PRIORIZACION AS ID_PRIORIZACION,"
						strSql = strSql & " TSP.NOM_TIPO_SOLICITUD AS NOM_TIPO_SOLICITUD, TRP.NOM_TIPO_RECLAMO AS NOM_TIPO_RECLAMO, OBSERVACION_PRIORIZACION,"
						strSql = strSql & " U1.LOGIN AS USUARIO_PRIO, (CASE WHEN PR.SOLICITA_RESPUESTA = 1 THEN 'SI' ELSE 'NO' END) AS SOL_RESP, U2.LOGIN AS USUARIO_ASIG, UPPER(DEUDOR.NOMBRE_DEUDOR) AS NOM_DEUDOR,"
						strSql = strSql & " ISNULL((SUBSTRING(CONVERT(VARCHAR(11),FECHA_PRIORIZACION,6),1,7) + '/ ' + SUBSTRING(CONVERT(VARCHAR(10),FECHA_PRIORIZACION,108),1,5)),'SIN AGEND') AS FECHA"


						strSql = strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
						strSql= strSql & " 						 INNER JOIN TIPO_SOLICITUD_PRIORIZACION TSP ON TSP.ID_TIPO_SOLICITUD = PR.ID_TIPO_SOLICITUD"
						strSql = strSql & " 				  	 INNER JOIN TIPO_RECLAMO_PRIORIZACION TRP ON TRP.ID_TIPO_RECLAMO = PR.ID_TIPO_RECLAMO"
						strSql = strSql & " 				  	 INNER JOIN USUARIO U1 ON U1.ID_USUARIO = PR.ID_USUARIO_PRIORIZACION"
						strSql = strSql & " 				  	 INNER JOIN DEUDOR ON PR.COD_CLIENTE = DEUDOR.COD_CLIENTE AND PR.RUT_DEUDOR = DEUDOR.RUT_DEUDOR"
						strSql = strSql & " 				  	 LEFT JOIN USUARIO U2 ON DEUDOR.USUARIO_ASIG = U2.ID_USUARIO"
						strSql = strSql & " 				  	 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
						strSql = strSql & "			   			 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO "
						strSql = strSql & "			   			 LEFT JOIN GESTIONES ON CUOTA.ID_ULT_GEST_GENERAL = GESTIONES.ID_GESTION "


						strSql = strSql & " WHERE PR.COD_CLIENTE = '" & strCodCliente & "' AND PRC.ESTADO_PRIORIZACION = 0 AND ESTADO_DEUDA.ACTIVO = 1 AND ((DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0) OR (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < PR.FECHA_PRIORIZACION)"

						If Trim(strCobranza) = "INTERNA" Then
							strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
						End if

						If Trim(strCobranza) = "EXTERNA" Then
							strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
						End if

						If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
							strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
						Else
							if Trim(strEjeAsig) <> "" Then
							strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
							End if
						End if

						if Trim(usuario_sol) <> "0" Then
						strSql = strSql & " AND  U1.LOGIN = '" & usuario_sol & "'"
						End if

						If inicio <> "" then

						strSql = strSql & " AND  FECHA_PRIORIZACION > = '" & inicio & " 00:00:00'"

						End If

						If termino <> "" then

						strSql = strSql & " AND FECHA_PRIORIZACION < = '" & termino & " 23:59:59'"

						End If

						strSql = strSql & " GROUP BY PR.RUT_DEUDOR, PR.ID_PRIORIZACION, FECHA_PRIORIZACION, FECHA_PRIORIZACION, TSP.NOM_TIPO_SOLICITUD,"
						strSql = strSql & " TRP.NOM_TIPO_RECLAMO, OBSERVACION_PRIORIZACION, U1.LOGIN , PR.SOLICITA_RESPUESTA, U2.LOGIN, DEUDOR.NOMBRE_DEUDOR"

						strSql = strSql & " ORDER BY PR.ID_PRIORIZACION DESC,FECHA_PRIORIZACION DESC"

						'Response.write "strSql=" & strSql

						set rsPriorizacion=Conn.execute(strSql)

						If Not rsPriorizacion.Eof Then

						%>
							<thead>
							<tr >

								<td class="Estilo4">&nbsp;</td>
								<td width = "65" class="Estilo4">FECHA</td>
								<td class="Estilo4">RUT DEUDOR</td>
								<td class="Estilo4">NOMBRE DEUDOR</td>
								<td class="Estilo4">TIPO SOLICITUD</td>
								<td class="Estilo4">TIPO RECLAMO</td>
								<td class="Estilo4">SOL. RESP.</td>
								<td width = "30" class="Estilo4">OBS.</td>
								<td class="Estilo4">COBRADOR</td>
								<td class="Estilo4">SOLICITANTE</td>
								<td class="Estilo4">&nbsp;</td>

							</tr>
							</thead>
							<tbody>

						<%
							intCorr = 0

							Do While Not rsPriorizacion.Eof

							strRutDeudor = rsPriorizacion("RUT_DEUDOR")
							strObsPrio = rsPriorizacion("OBSERVACION_PRIORIZACION")
							strUsuarioPrio = rsPriorizacion("USUARIO_PRIO")
							strFechaPrio = rsPriorizacion("FECHA")
							strTipoSol = rsPriorizacion("NOM_TIPO_SOLICITUD")
							strUsuarioAsig = rsPriorizacion("USUARIO_ASIG")
							strTipoSol = rsPriorizacion("NOM_TIPO_SOLICITUD")
							strTipoReclamo = rsPriorizacion("NOM_TIPO_RECLAMO")
							strSolResp = rsPriorizacion("SOL_RESP")
							intIdPriorizacion = rsPriorizacion("ID_PRIORIZACION")
							strNomDeudor = rsPriorizacion("NOM_DEUDOR")

							intCorr = intCorr + 1
							strTotalDoc = ""

								AbrirSCG2()
											strSql = "SELECT CUOTA.NRO_DOC, (CASE WHEN (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0) OR (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < PR.FECHA_PRIORIZACION THEN 1 ELSE 0 END) AS AGEND_PRIO"
											strSql= strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
											strSql= strSql & " 					 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
											strSql= strSql & " 					 INNER JOIN USUARIO ON PR.ID_USUARIO_PRIORIZACION = USUARIO.ID_USUARIO"
											strSql = strSql & "			   		 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
											strSql = strSql & "			   		 LEFT JOIN GESTIONES ON CUOTA.ID_ULT_GEST_GENERAL = GESTIONES.ID_GESTION "

											strSql= strSql & " WHERE PRC.ID_PRIORIZACION = '" & intIdPriorizacion & "' AND PRC.ESTADO_PRIORIZACION = 0 AND ESTADO_DEUDA.ACTIVO = 1 AND ((DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0) OR (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < PR.FECHA_PRIORIZACION)"

											'Response.write "<br>strSql=" & strSql
											set RsPrioDoc=Conn2.execute(strSql)

											If not RsPrioDoc.eof then

											intAgendPrio = 0

												Do While Not RsPrioDoc.Eof

													strDoc = RsPrioDoc("NRO_DOC")
													strTotalDoc = strTotalDoc & "-" & strDoc
													intAgendPrio = intAgendPrio + RsPrioDoc("AGEND_PRIO")

													RsPrioDoc.movenext
												Loop
											End If

								CerrarSCG2()

								If Trim(strTotalDoc) <> "" Then
									strTotalDoc = "Doc: " & Mid(strTotalDoc,2,Len(strTotalDoc))
								End If

								If Trim(strFechaPrio) <> "" and Trim(strUsuarioPrio) <> "" then
									strTextoPrio = "Fecha: " & strFechaPrio & " , Usuario : " & strUsuarioPrio & chr(13) & "Tipo Sol: " & strTipoSol & chr(13) & "Obs : " & strObsPrio & chr(13) & strTotalDoc & chr(13) & chr(13)

									strTextoPrioF = strTextoPrioF & strTextoPrio
								End If

								If Trim(strObsPrio) = "" Then
									strObsPrio = "SIN OBSERVACION ADICIONAL"
								End If

								'Response.write "<br>strTextoPrioF=" & strTextoPrioF
								%>

								<tr >

									<td><%=intCorr%></td>
									<td><%=strFechaPrio%></td>
									<td>
										<A HREF="principal.asp?TX_RUT=<%=strRutDeudor%>">
											<acronym title="Llevar a pantalla principal"><%=strRutDeudor%></acronym>
										</A>
									</td>

									<td class="Estilo4" title="<%=strNomDeudor%>"><%=mid(strNomDeudor,1,30)%>

									<td class="Estilo4" title="<%=strTipoSol%>"><%=mid(strTipoSol,1,25)%>

									<td class="Estilo4" title="<%=strTipoReclamo%>"><%=mid(strTipoReclamo,1,25)%>

									<td align = "center"><%=strSolResp%></td>

									<td align = "center" class="Estilo4" onClick="priorizar_caso('<%=strRutDeudor%>');" value="Priorizar caso" title="<%=strObsPrio%>"><img src="../imagenes/priorizar_urgente.png" border="0"></td>

									<td><%=strUsuarioAsig%></td>
									<td><%=strUsuarioPrio%></td>
									<td class="Estilo4" title="<%=strTotalDoc%>">
										<img src="../imagenes/carpeta1.png" border="0">

								</tr>

								<%
								rsPriorizacion.movenext
							Loop
						CerrarSCG()%>
								</tbody>
								<thead>
								<tr >
									<td class="totales" colspan = "11">&nbsp;</td>
								</tr>
								</thead>

					 <%Else%>
					 	<thead>
							<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
								<td colspan = "11">&nbsp;</td>
							</tr>

							<tr >
								<td height="30" Align="CENTER" Colspan = "11" class="estilo_columna_individual">NO HAY PRIORIZACIONES PENDIENTES A GESTIONAR</td>
							</tr>

							<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
								<td colspan = "11">&nbsp;</td>
							</tr>
						</thead>
					 <%End If%>

				</table>

			<%End If%>

			<%If strTipoPriorizacion = 2 or strTipoPriorizacion = 0 then%>

				<table width="100%" border="0" class="intercalado">
				<thead>
					<%If strTipoPriorizacion = 0 then%>
					<tr>
						<td Height = "50">&nbsp;</td>
					</tr>
					<%End If%>

					<tr>

						<td colspan = "11" class="subtitulo_informe">> PRIORIZACIONES NO CONFIRMADAS / AGENDADAS</td>
					</tr>
				</thead>

					<%AbrirSCG()

						strSql = "SELECT PR.RUT_DEUDOR, PR.ID_PRIORIZACION AS ID_PRIORIZACION, SUBSTRING(CONVERT(VARCHAR(38),FECHA_PRIORIZACION,121),12,5) AS HORA,"
						strSql = strSql & " TSP.NOM_TIPO_SOLICITUD AS NOM_TIPO_SOLICITUD, TRP.NOM_TIPO_RECLAMO AS NOM_TIPO_RECLAMO, OBSERVACION_PRIORIZACION,"
						strSql = strSql & " U1.LOGIN AS USUARIO_PRIO, (CASE WHEN PR.SOLICITA_RESPUESTA = 1 THEN 'SI' ELSE 'NO' END) AS SOL_RESP, U2.LOGIN AS USUARIO_ASIG, UPPER(DEUDOR.NOMBRE_DEUDOR) AS NOM_DEUDOR,"
						strSql = strSql & " ISNULL((SUBSTRING(CONVERT(VARCHAR(11),FECHA_PRIORIZACION,6),1,7) + '/ ' + SUBSTRING(CONVERT(VARCHAR(10),FECHA_PRIORIZACION,108),1,5)),'SIN AGEND') AS FECHA"


						strSql = strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
						strSql= strSql & " 						 INNER JOIN TIPO_SOLICITUD_PRIORIZACION TSP ON TSP.ID_TIPO_SOLICITUD = PR.ID_TIPO_SOLICITUD"
						strSql = strSql & " 				  	 INNER JOIN TIPO_RECLAMO_PRIORIZACION TRP ON TRP.ID_TIPO_RECLAMO = PR.ID_TIPO_RECLAMO"
						strSql = strSql & " 				  	 INNER JOIN USUARIO U1 ON U1.ID_USUARIO = PR.ID_USUARIO_PRIORIZACION"
						strSql = strSql & " 				  	 INNER JOIN DEUDOR ON PR.COD_CLIENTE = DEUDOR.COD_CLIENTE AND PR.RUT_DEUDOR = DEUDOR.RUT_DEUDOR"
						strSql = strSql & " 				  	 LEFT JOIN USUARIO U2 ON DEUDOR.USUARIO_ASIG = U2.ID_USUARIO"
						strSql = strSql & " 				  	 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
						strSql = strSql & "			   			 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO "
						strSql = strSql & "			   			 LEFT JOIN GESTIONES ON CUOTA.ID_ULT_GEST_GENERAL = GESTIONES.ID_GESTION "


						strSql = strSql & " WHERE PR.COD_CLIENTE = '" & strCodCliente & "' AND PRC.ESTADO_PRIORIZACION = 0 AND ESTADO_DEUDA.ACTIVO = 1 AND ((DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) < 0) OR (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > PR.FECHA_PRIORIZACION)"

						If Trim(strCobranza) = "INTERNA" Then
							strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
						End if

						If Trim(strCobranza) = "EXTERNA" Then
							strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
						End if

						If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
							strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
						Else
							if Trim(strEjeAsig) <> "" Then
							strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
							End if
						End if

						if Trim(usuario_sol) <> "0" Then
						strSql = strSql & " AND  U1.LOGIN = '" & usuario_sol & "'"
						End if

						strSql = strSql & " GROUP BY PR.RUT_DEUDOR, PR.ID_PRIORIZACION, FECHA_PRIORIZACION, FECHA_PRIORIZACION, TSP.NOM_TIPO_SOLICITUD,"
						strSql = strSql & " TRP.NOM_TIPO_RECLAMO, OBSERVACION_PRIORIZACION, U1.LOGIN , PR.SOLICITA_RESPUESTA, U2.LOGIN, DEUDOR.NOMBRE_DEUDOR"

						strSql = strSql & " ORDER BY ID_PRIORIZACION DESC,FECHA_PRIORIZACION DESC"

						'Response.write "strSql=" & strSql

						set rsPriorizacion=Conn.execute(strSql)

						If Not rsPriorizacion.Eof Then

						%>
						<thead>
							<tr >

								<td class="Estilo4">&nbsp;</td>
								<td width = "65" class="Estilo4">FECHA</td>
								<td class="Estilo4">RUT DEUDOR</td>
								<td class="Estilo4">NOMBRE DEUDOR</td>
								<td class="Estilo4">TIPO SOLICITUD</td>
								<td class="Estilo4">TIPO RECLAMO</td>
								<td class="Estilo4">SOL. RESP.</td>
								<td width = "30" class="Estilo4">OBS.</td>
								<td class="Estilo4">COBRADOR</td>
								<td class="Estilo4">SOLICITANTE</td>
								<td class="Estilo4">&nbsp;</td>

							</tr>
						</thead>
						<tbody>

						<%
							intCorr = 0

							Do While Not rsPriorizacion.Eof

							strRutDeudor = rsPriorizacion("RUT_DEUDOR")
							strObsPrio = rsPriorizacion("OBSERVACION_PRIORIZACION")
							strUsuarioPrio = rsPriorizacion("USUARIO_PRIO")
							strFechaPrio = rsPriorizacion("FECHA")
							strTipoSol = rsPriorizacion("NOM_TIPO_SOLICITUD")
							strUsuarioAsig = rsPriorizacion("USUARIO_ASIG")
							strTipoSol = rsPriorizacion("NOM_TIPO_SOLICITUD")
							strTipoReclamo = rsPriorizacion("NOM_TIPO_RECLAMO")
							strSolResp = rsPriorizacion("SOL_RESP")
							intIdPriorizacion = rsPriorizacion("ID_PRIORIZACION")
							strNomDeudor = rsPriorizacion("NOM_DEUDOR")

							intCorr = intCorr + 1
							strTotalDoc = ""

								AbrirSCG2()
											strSql = "SELECT CUOTA.NRO_DOC, (CASE WHEN (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0) OR (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < PR.FECHA_PRIORIZACION THEN 1 ELSE 0 END) AS AGEND_PRIO"
											strSql= strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
											strSql= strSql & " 					 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
											strSql= strSql & " 					 INNER JOIN USUARIO ON PR.ID_USUARIO_PRIORIZACION = USUARIO.ID_USUARIO"
											strSql = strSql & "			   		 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO "
											strSql = strSql & "			   		 LEFT JOIN GESTIONES ON CUOTA.ID_ULT_GEST_GENERAL = GESTIONES.ID_GESTION "

											strSql= strSql & " WHERE PRC.ID_PRIORIZACION = '" & intIdPriorizacion & "' AND PRC.ESTADO_PRIORIZACION = 0 AND ESTADO_DEUDA.ACTIVO = 1 AND ((DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) < 0) OR (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > PR.FECHA_PRIORIZACION)"

											'Response.write "<br>strSql=" & strSql
											set RsPrioDoc=Conn2.execute(strSql)

											If not RsPrioDoc.eof then

											intAgendPrio = 0

												Do While Not RsPrioDoc.Eof

													strDoc = RsPrioDoc("NRO_DOC")
													strTotalDoc = strTotalDoc & "-" & strDoc
													intAgendPrio = intAgendPrio + RsPrioDoc("AGEND_PRIO")

													RsPrioDoc.movenext
												Loop
											End If

								CerrarSCG2()

								If Trim(strTotalDoc) <> "" Then
									strTotalDoc = "Doc: " & Mid(strTotalDoc,2,Len(strTotalDoc))
								End If

								If Trim(strFechaPrio) <> "" and Trim(strUsuarioPrio) <> "" then
									strTextoPrio = "Fecha: " & strFechaPrio & " , Usuario : " & strUsuarioPrio & chr(13) & "Tipo Sol: " & strTipoSol & chr(13) & "Obs : " & strObsPrio & chr(13) & strTotalDoc & chr(13) & chr(13)

									strTextoPrioF = strTextoPrioF & strTextoPrio
								End If

								If Trim(strObsPrio) = "" Then
									strObsPrio = "SIN OBSERVACION ADICIONAL"
								End If

								'Response.write "<br>strTextoPrioF=" & strTextoPrioF
								%>

								<tr >

									<td><%=intCorr%></td>
									<td><%=strFechaPrio%></td>
									<td>
										<A HREF="principal.asp?TX_RUT=<%=strRutDeudor%>">
											<acronym title="Llevar a pantalla principal"><%=strRutDeudor%></acronym>
										</A>
									</td>

									<td class="Estilo4" title="<%=strNomDeudor%>"><%=mid(strNomDeudor,1,30)%>

									<td class="Estilo4" title="<%=strTipoSol%>"><%=mid(strTipoSol,1,25)%>

									<td class="Estilo4" title="<%=strTipoReclamo%>"><%=mid(strTipoReclamo,1,25)%>

									<td align = "center"><%=strSolResp%></td>

									<td align = "center" class="Estilo4" onClick="priorizar_caso('<%=strRutDeudor%>');" value="Priorizar caso" title="<%=strObsPrio%>"><img src="../imagenes/priorizar_urgente.png" border="0"></td>

									<td><%=strUsuarioAsig%></td>
									<td><%=strUsuarioPrio%></td>
									<td class="Estilo4" title="<%=strTotalDoc%>">
										<img src="../imagenes/carpeta1.png" border="0">

								</tr>

								<%
								rsPriorizacion.movenext
							Loop
						CerrarSCG()%>
							</tbody>
							<thead>
								<tr >
									<td class="totales" colspan = "11">&nbsp;</td>
								</tr>
							</thead>

					 <%Else%>
					 		<thead>
							<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
								<td Colspan = "11">&nbsp;</td>
							</tr>

							<tr >
								<td height="30" Align="CENTER" Colspan = "11" class="estilo_columna_individual">NO HAY PRIORIZACIONES AGENDADAS</td>
							</tr>

							<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
								<td Colspan = "11">&nbsp;</td>
							</tr>
							</thead>

					 <%End If%>

				</table>

			<%End If%>

			<%If strTipoPriorizacion = 3 or strTipoPriorizacion = 0 then%>

				<table width="100%" border="0" bordercolor="#000000" class="intercalado">
				<thead>
					<%If strTipoPriorizacion = 0 then%>
					<tr>
						<td Height = "50">&nbsp;</td>
					</tr>
					<%End If%>

					<tr>

						<td colspan = "13" class="subtitulo_informe">> ULTIMAS 50 PRIORIZACIONES CONFIRMADAS</td>
					</tr>
				</thead>
					<%AbrirSCG()

						strSql = "SELECT TOP 50 PR.RUT_DEUDOR, PR.ID_PRIORIZACION AS ID_PRIORIZACION,ISNULL((SUBSTRING(CONVERT(VARCHAR(11),PRC.FECHA_ESTADO,6),1,7) + '/ ' + SUBSTRING(CONVERT(VARCHAR(10),PRC.FECHA_ESTADO,108),1,5)),'SIN AGEND') AS FECHA_CONF,"
						strSql = strSql & " TSP.NOM_TIPO_SOLICITUD AS NOM_TIPO_SOLICITUD, TRP.NOM_TIPO_RECLAMO AS NOM_TIPO_RECLAMO, OBSERVACION_PRIORIZACION,"
						strSql = strSql & " U1.LOGIN AS USUARIO_PRIO, (CASE WHEN PR.SOLICITA_RESPUESTA = 1 THEN 'SI' ELSE 'NO' END) AS SOL_RESP, U2.LOGIN AS USUARIO_ASIG, UPPER(DEUDOR.NOMBRE_DEUDOR) AS NOM_DEUDOR, U3.LOGIN AS USUARIO_CONF,PRC.OBSERVACION_CONF_PRIORIZACION AS OBS_CONF,"
						strSql = strSql & " ISNULL((SUBSTRING(CONVERT(VARCHAR(11),FECHA_PRIORIZACION,6),1,7) + '/ ' + SUBSTRING(CONVERT(VARCHAR(10),FECHA_PRIORIZACION,108),1,5)),'SIN AGEND') AS FECHA"


						strSql = strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
						strSql= strSql & " 						 INNER JOIN TIPO_SOLICITUD_PRIORIZACION TSP ON TSP.ID_TIPO_SOLICITUD = PR.ID_TIPO_SOLICITUD"
						strSql = strSql & " 				  	 INNER JOIN TIPO_RECLAMO_PRIORIZACION TRP ON TRP.ID_TIPO_RECLAMO = PR.ID_TIPO_RECLAMO"
						strSql = strSql & " 				  	 INNER JOIN USUARIO U1 ON U1.ID_USUARIO = PR.ID_USUARIO_PRIORIZACION"
						strSql = strSql & " 				  	 INNER JOIN DEUDOR ON PR.COD_CLIENTE = DEUDOR.COD_CLIENTE AND PR.RUT_DEUDOR = DEUDOR.RUT_DEUDOR"
						strSql = strSql & " 				  	 LEFT JOIN USUARIO U2 ON DEUDOR.USUARIO_ASIG = U2.ID_USUARIO"
						strSql = strSql & " 				  	 LEFT JOIN USUARIO U3 ON PRC.ID_USUARIO_ESTADO = U3.ID_USUARIO"
						strSql = strSql & " 				  	 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
						strSql = strSql & "			   			 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO "

						strSql = strSql & " WHERE PR.COD_CLIENTE = '" & strCodCliente & "' AND PRC.ESTADO_PRIORIZACION = 1"

						If Trim(strCobranza) = "INTERNA" Then
							strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
						End if

						If Trim(strCobranza) = "EXTERNA" Then
							strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
						End if

						If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
							strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
						Else
							if Trim(strEjeAsig) <> "" Then
							strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
							End if
						End if

						if Trim(usuario_sol) <> "0" Then
						strSql = strSql & " AND  U1.LOGIN = '" & usuario_sol & "'"
						End if


						strSql = strSql & " GROUP BY PR.RUT_DEUDOR, PR.ID_PRIORIZACION, FECHA_PRIORIZACION, FECHA_PRIORIZACION, TSP.NOM_TIPO_SOLICITUD,"
						strSql = strSql & " TRP.NOM_TIPO_RECLAMO, OBSERVACION_PRIORIZACION, U1.LOGIN , PR.SOLICITA_RESPUESTA, U2.LOGIN, DEUDOR.NOMBRE_DEUDOR,PRC.FECHA_ESTADO, U3.LOGIN, PRC.OBSERVACION_CONF_PRIORIZACION"

						strSql = strSql & " ORDER BY FECHA_CONF DESC,FECHA_PRIORIZACION DESC"

						'Response.write "strSql=" & strSql

						set rsPriorizacion=Conn.execute(strSql)

						If Not rsPriorizacion.Eof Then

						%>
						<thead>
							<tr >

								<td class="Estilo4">&nbsp;</td>
								<td width = "65" class="Estilo4">FECHA</td>
								<td class="Estilo4">RUT DEUDOR</td>
								<td class="Estilo4">NOMBRE DEUDOR</td>
								<td class="Estilo4">TIPO SOLICITUD</td>
								<td class="Estilo4">TIPO RECLAMO</td>
								<td width = "30" class="Estilo4">OBS.</td>
								<td class="Estilo4">COBRADOR</td>
								<td class="Estilo4">SOLICITANTE</td>
								<td class="Estilo4">FECHA CONF.</td>
								<td class="Estilo4">USU. CONF.</td>
								<td class="Estilo4">CONF.</td>
								<td class="Estilo4">&nbsp;</td>

							</tr>
						</thead>
						<tbody>
						<%
							intCorr = 0

							Do While Not rsPriorizacion.Eof

							strRutDeudor = rsPriorizacion("RUT_DEUDOR")
							strObsPrio = rsPriorizacion("OBSERVACION_PRIORIZACION")
							strUsuarioPrio = rsPriorizacion("USUARIO_PRIO")
							strFechaPrio = rsPriorizacion("FECHA")
							strTipoSol = rsPriorizacion("NOM_TIPO_SOLICITUD")
							strUsuarioAsig = rsPriorizacion("USUARIO_ASIG")
							strTipoSol = rsPriorizacion("NOM_TIPO_SOLICITUD")
							strTipoReclamo = rsPriorizacion("NOM_TIPO_RECLAMO")
							strSolResp = rsPriorizacion("SOL_RESP")
							intIdPriorizacion = rsPriorizacion("ID_PRIORIZACION")
							strNomDeudor = rsPriorizacion("NOM_DEUDOR")
							strFechaConf = rsPriorizacion("FECHA_CONF")
							strUusarioConf = rsPriorizacion("USUARIO_CONF")
							strObsConf = rsPriorizacion("OBS_CONF")


							intCorr = intCorr + 1
							strTotalDoc = ""

								AbrirSCG2()
											strSql = "SELECT CUOTA.NRO_DOC, (CASE WHEN (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0) OR (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < PR.FECHA_PRIORIZACION THEN 1 ELSE 0 END) AS AGEND_PRIO"
											strSql= strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
											strSql= strSql & " 					 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
											strSql= strSql & " 					 INNER JOIN USUARIO ON PR.ID_USUARIO_PRIORIZACION = USUARIO.ID_USUARIO"
											strSql = strSql & "			   		 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
											strSql = strSql & "			   		 LEFT JOIN GESTIONES ON CUOTA.ID_ULT_GEST_GENERAL = GESTIONES.ID_GESTION "

											strSql= strSql & " WHERE PRC.ID_PRIORIZACION = '" & intIdPriorizacion & "' AND PRC.ESTADO_PRIORIZACION = 1"

											'Response.write "<br>strSql=" & strSql
											set RsPrioDoc=Conn2.execute(strSql)

											If not RsPrioDoc.eof then

											intAgendPrio = 0

												Do While Not RsPrioDoc.Eof

													strDoc = RsPrioDoc("NRO_DOC")
													strTotalDoc = strTotalDoc & "-" & strDoc
													intAgendPrio = intAgendPrio + RsPrioDoc("AGEND_PRIO")

													RsPrioDoc.movenext
												Loop
											End If

								CerrarSCG2()

								If Trim(strTotalDoc) <> "" Then
									strTotalDoc = "Doc: " & Mid(strTotalDoc,2,Len(strTotalDoc))
								End If

								If Trim(strFechaPrio) <> "" and Trim(strUsuarioPrio) <> "" then
									strTextoPrio = "Fecha: " & strFechaPrio & " , Usuario : " & strUsuarioPrio & chr(13) & "Tipo Sol: " & strTipoSol & chr(13) & "Obs : " & strObsPrio & chr(13) & strTotalDoc & chr(13) & chr(13)

									strTextoPrioF = strTextoPrioF & strTextoPrio
								End If

								If Trim(strObsConf) = "" Then
									strObsConf = "SIN OBSERVACION ADICIONAL"
								End If

								If Trim(strObsPrio) = "" Then
									strObsPrio = "SIN OBSERVACION ADICIONAL"
								End If

								'Response.write "<br>strTextoPrioF=" & strTextoPrioF
								%>

								<tr >

									<td><%=intCorr%></td>
									<td><%=strFechaPrio%></td>
									<td>
										<A HREF="principal.asp?TX_RUT=<%=strRutDeudor%>">
											<acronym title="Llevar a pantalla principal"><%=strRutDeudor%></acronym>
										</A>
									</td>

									<td class="Estilo4" title="<%=strNomDeudor%>"><%=mid(strNomDeudor,1,25)%>

									<td class="Estilo4" title="<%=strTipoSol%>"><%=mid(strTipoSol,1,20)%>

									<td class="Estilo4" title="<%=strTipoReclamo%>"><%=mid(strTipoReclamo,1,25)%>

									<td align = "center" class="Estilo4" title="<%=strObsPrio%>"><img src="../imagenes/priorizar_normal.png" border="0"></td>

									<td><%=strUsuarioAsig%></td>
									<td><%=strUsuarioPrio%></td>
									<td><%=strFechaConf%></td>
									<td><%=strUusarioConf%></td>

									<td align = "center" class="Estilo4" onClick="priorizar_caso('<%=strRutDeudor%>');" value="Priorizar caso" title="<%=strObsPrio%>"><img src="../imagenes/priorizar_urgente.png" border="0"></td>

									<td class="Estilo4" title="<%=strTotalDoc%>">
										<img src="../imagenes/carpeta1.png" border="0">

								</tr>

								<%
								rsPriorizacion.movenext
							Loop
						CerrarSCG()%>
							</tbody>
							<thead>
								<tr >
									<td class="totales" colspan = "13">&nbsp;</td>
								</tr>
							</thead>

						 <%Else%>
						 	<thead>
								<tr class="totales">
									<td>&nbsp;</td>
								</tr>

								<tr >
									<td height="30" Align="CENTER" Colspan = "13" class="estilo_columna_individual">NO HAY PRIORIZACIONES CONFIRMADAS</td>
								</tr>

								<tr class="totales">
									<td>&nbsp;</td>
								</tr>
							</thead>
						 <%End If%>

				</table>

			<%End If%>
</form>
</body>
</html>

<script language="JavaScript" type="text/JavaScript">

function priorizar_caso(strRutDeudor){
	datos.action='priorizar_caso.asp?strRut=' + strRutDeudor;
	datos.submit();
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
