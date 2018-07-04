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
	'cod_caja=110
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	strCodCliente = session("ses_codcli")

	AbrirSCG()

	strBuscar = Request("strBuscar")
	
	'response.end
	
	intIdUsuario = request("cmb_usuario")
	strTipoGestion = request("cmb_tipogestion")
	strEstadoRevision = request("cmb_estadoRevision")
	strRegion = request("cmb_region")
	termino = request("termino")
	inicio = request("inicio")

	If Trim(strBuscar) = "S" Then
		session("Ftro_TipoGestionRyR") = strTipoGestion
		session("Ftro_EstadoRevision") = strEstadoRevision
		session("Ftro_EjecAsigRyR") = intIdUsuario
		session("Ftro_Region") = strRegion
		
		strTipoGestion = session("Ftro_TipoGestionRyR")
		strEstadoRevision = session("Ftro_EstadoRevision")
		intIdUsuario = session("Ftro_EjecAsigRyR")
		strRegion = session("Ftro_Region")
	End If

	If Trim(strBuscar) = "N" Then
		session("Ftro_TipoGestionRyR") = ""
		session("Ftro_EstadoRevision") = ""
		session("Ftro_EjecAsigRyR") = ""
		session("Ftro_Region") = ""
		intIdUsuario = ""
		strTipoGestion = ""
		strRegion = "" 
	End If

	If Trim(strBuscar) = "" Then
		strTipoGestion = session("Ftro_TipoGestionRyR")
		strEstadoRevision = session("Ftro_EstadoRevision")
		intIdUsuario = session("Ftro_EjecAsigRyR")
		strRegion = session("Ftro_Region")
	End If
	
	if intIdUsuario = "" then intIdUsuario = "0"
	if strTipoGestion = "" then strTipoGestion = "0"
	if strEstadoRevision = "" then strEstadoRevision = "0"
	if strRegion = "" then strRegion = "0"

	'response.write " strEstadoRevision = " & strEstadoRevision
%>

<title>CRM FACTORING</title>


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
<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script language="JavaScript " type="text/JavaScript">
$(document).ready(function(){

	$("#table_tablesorter").tablesorter({dateFormat: "uk"}); 

	$.prettyLoader();
	
	$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
 	$(document).tooltip();
	
})

function envia()
{
	$.prettyLoader.show(2000);
	document.datos.action = "listado_Revision_Ruta.asp?strBuscar=S";
	document.datos.submit();
}

function limpiar() {
	$.prettyLoader.show(2000);
	document.datos.action = "listado_Revision_Ruta.asp?strBuscar=N";
	document.datos.submit();
}
</script>


</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">


<form name="datos" method="post">
<div class="titulo_informe">LISTADO REVISION RUTA</div>
<br>
	<table width="90%" border="0" class="estilo_columnas" align="center">
	<thead>		
	      <tr height="20">
			<td>TIPO GESTIÓN</td>
			<td>ESTADO REVISIÓN</td>
			<td>REGIÓN</td>
			<td>FECHA COMPROMISO DESDE</td>
			<td>FECHA COMPROMISO HASTA</td>
	      	<% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
			<td>EJEC.ASIG.</td>
			<% End If %>
			<td></td>
			<td></td>
	      </tr>
	</thead>
		  <tr >
			<td>
				<SELECT NAME="cmb_tipogestion" id="cmb_tipogestion" onChange="envia();">
					<option value="0" <%If Trim(strTipoGestion)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>RECAUDACION</option>
					<option value="2" <%If Trim(strTipoGestion)="2" Then Response.write "SELECTED"%>>VERIFICACION</option>

				</SELECT>
			</td>
			<td>
				<SELECT NAME="cmb_estadoRevision" id="cmb_estadoRevision" onChange="envia();">
					<option value="0" <%If Trim(strEstadoRevision)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strEstadoRevision)="1" Then Response.write "SELECTED"%>>POR REVISAR</option>
					<option value="2" <%If Trim(strEstadoRevision)="2" Then Response.write "SELECTED"%>>PENDIENTE</option>

				</SELECT>
			</td>
			<td>
				<SELECT NAME="cmb_region" id="cmb_region" onChange="envia();">
					<option value="0" <%If Trim(strRegion)="0" Then Response.write "SELECTED"%>>TODAS</option>
					<option value="1" <%If Trim(strRegion)="1" Then Response.write "SELECTED"%>>METROPOLITANA</option>
					<option value="2" <%If Trim(strRegion)="2" Then Response.write "SELECTED"%>>OTRAS REGIONES</option>

				</SELECT>
			</td>
			<td><input name="inicio" type="text" readonly="true" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
			</td>
			<td><input name="termino" type="text" readonly="true" id="termino" value="<%=termino%>" size="10" maxlength="10">
			
			</td>

		  <% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
			<td>
				<SELECT NAME="cmb_usuario" id="cmb_usuario" onChange="envia();">
					<option value="0">TODOS</option>
					<%
					sql_usuario ="SELECT DISTINCT U.ID_USUARIO, LOGIN = UPPER(LOGIN) "
					sql_usuario = sql_usuario & " FROM USUARIO U "
					sql_usuario = sql_usuario & " INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO "
					sql_usuario = sql_usuario & " AND UC.COD_CLIENTE = '"&trim(strCodCliente)&"' "
					sql_usuario = sql_usuario & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1 "
					sql_usuario = sql_usuario & " AND U.PERFIL_EMP=0 " 

					set rsUsu = conn.execute(sql_usuario)		
					
					if not rsUsu.eof then
						do until rsUsu.eof
						%>
						<option value="<%=rsUsu("ID_USUARIO")%>"
						<%if Trim(intIdUsuario)=Trim(rsUsu("ID_USUARIO")) then
							response.Write("Selected")
						end if%>
						><%=ucase(rsUsu("LOGIN"))%></option>

						<%rsUsu.movenext
						loop
					end if
					rsUsu.close
					set rsUsu=nothing
					%>
				</SELECT>
			</td>
			<% End If %>
			
			<td><input type="Button" class="fondo_boton_100" name="LimpiarButton" value="Limpiar" onClick="limpiar();"></td>
			<td><input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();"></td>
			
	      </tr>
    </table>

	<input type="hidden" id="TXT_CAMBIA" value='<%=strBuscar%>'/>

	<table width="100%" border="0" class="intercalado" align="center">
	
	<thead>
		<tr>
			<td>&nbsp;</td>
			<td>GESTION</td>
			<td>ESTADO</td>
			<td>FECHA</td>
			<td>RUT DEUDOR</td>
			<td width="250">NOMBRE DEUDOR</td>
			<td>MONTO</td>
			<td width="120">FORMA PAGO</td>
			<td>LUGAR PAGO</td>
			<td>COMUNA</td>
			<td>Nº DOC</td>
			<td>EJEC.ASIG.</td>
			<td width="45">OBS.</td>

		</tr>
	</thead>
	<tbody>
	<%

	strSql = "SELECT DEUDOR.NOMBRE_DEUDOR, USUARIO.LOGIN AS EJEC_ASIG, "
	strSql = strSql & " CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103) as 'FECHA_INGRESO',"

	strSql = strSql & " GESTIONSOLA = CASE WHEN ( GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11)"
	strSql = strSql & " 				   		  AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO <> 'TR' AND CAJA_FORMA_PAGO.ID_FORMA_PAGO <> 'DP')"
	strSql = strSql & " 				   THEN 'RECAUDACIÓN'"
	strSql = strSql & "					   WHEN ( GESTIONES_TIPO_GESTION.GESTION_MODULOS = 13)"
	strSql = strSql & " 				   		  AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO <> 'TR' AND CAJA_FORMA_PAGO.ID_FORMA_PAGO <> 'DP')"
	strSql = strSql & " 				   THEN 'VERIFICACION'"	
	strSql = strSql & " 				   ELSE 'NO DEFINIDO'"
	strSql = strSql & "					   END,"

	strSql = strSql & " ESTADO_REVISION = CASE WHEN GESTIONES.FECHA_COMPROMISO = CAST(CONVERT(VARCHAR(10),(CASE WHEN DATENAME(dw,GETDATE())='LUNES' THEN GETDATE()- 3 ELSE GETDATE()-1 END),103) AS DATETIME)"
	strSql = strSql & " 				   THEN 'POR REVISAR'"
	strSql = strSql & " 				   ELSE 'PENDIENTE'"
	strSql = strSql & "					   END,"
	
	strSql = strSql & " CUOTA.RUT_DEUDOR AS RUT_DEUDOR,"
	strSql = strSql & " CUOTA.NRO_DOC,"
	strSql = strSql & " ISNULL(GESTIONES.MONTO_CANCELADO,0) AS SALDO,"
	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES,CHAR(13),' '),CHAR(10),' ') as OBSERVACIONES,"


	strSql = strSql & " (CASE WHEN REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')<>'' THEN "
	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')+'							'+'/HORARIO RETIRO: '+ GESTIONES.HORA_DESDE+' - '+ GESTIONES.HORA_HASTA"
	strSql = strSql & " ELSE 'SIN OBSERVACIÓN' +'							'+'/HORARIO RETIRO: '+ GESTIONES.HORA_DESDE+' - '+ GESTIONES.HORA_HASTA"
	strSql = strSql & " END) as OBSERVACIONES_CAMPO,"


	strSql = strSql & " ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'') AS FECHA_NORMALIZACION,"

	strSql = strSql & " LUGAR_PAGO = UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO))),"
	strSql = strSql & " COMUNA = UPPER(ISNULL(DD.COMUNA,'NO DEFINIDA')),"
	
	strSql = strSql & " ISNULL(CAJA_FORMA_PAGO.DESC_FORMA_PAGO,'NO ESPEC.') AS 'FORMA_PAGO',"

	strSql = strSql & " CASE WHEN ISNULL(GESTIONES.MONTO_CANCELADO,'')= 0"
	strSql = strSql & " THEN ''"
	strSql = strSql & " ELSE ISNULL(GESTIONES.MONTO_CANCELADO,'NO ESPEC.')"
	strSql = strSql & " END AS 'MONTO REGULARIZADO'"


	strSql = strSql & " FROM GESTIONES			INNER JOIN GESTIONES_CUOTA ON GESTIONES.ID_GESTION = GESTIONES_CUOTA.ID_GESTION"
	strSql = strSql & "                         INNER JOIN CUOTA ON GESTIONES_CUOTA.ID_CUOTA = CUOTA.ID_CUOTA"
	strSql = strSql & "                         INNER JOIN DEUDOR ON GESTIONES.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND GESTIONES.COD_CLIENTE = DEUDOR.COD_CLIENTE"

	strSql = strSql & " 						LEFT JOIN FORMA_RECAUDACION RE ON RE.ID_FORMA_RECAUDACION= GESTIONES.ID_FORMA_RECAUDACION "
	strSql = strSql & " 						LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION= GESTIONES.ID_DIRECCION_COBRO_DEUDOR "
	
	strSql = strSql & "                         INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql & "                         LEFT JOIN CAJA_FORMA_PAGO ON GESTIONES.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO"
	strSql = strSql & "                         INNER JOIN GESTIONES_TIPO_GESTION ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_GESTION = GESTIONES_TIPO_GESTION.COD_GESTION"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_CATEGORIA.COD_CATEGORIA"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_CATEGORIA AND"
	strSql = strSql & " 																	GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_SUB_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
	strSql = strSql & " 						LEFT JOIN USUARIO ON CUOTA.USUARIO_ASIG = USUARIO.ID_USUARIO"
	strSql = strSql & " 						LEFT JOIN COMUNA CO ON LTRIM(RTRIM(DD.COMUNA)) = LTRIM(RTRIM(CO.NOMBRE_COMUNA))"


	strSql = strSql & " WHERE   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (11,13) AND CUOTA.COD_ULT_GEST <> '5*3*5'"
	strSql = strSql & " 		AND GESTIONES.FECHA_COMPROMISO <= GETDATE() - 1  AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO <> 'TR' AND CAJA_FORMA_PAGO.ID_FORMA_PAGO <> 'DP'))"

	strSql = strSql & " 		AND ESTADO_DEUDA.ACTIVO = 1 "
	strSql = strSql & " 		AND CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION "

	If inicio <> "" then

	strSql = strSql & "         AND GESTIONES.FECHA_COMPROMISO> = '" & inicio & " 00:00:00'"

	End If

	If termino <> "" then

	strSql = strSql & "         AND GESTIONES.FECHA_COMPROMISO< = '" & termino & " 23:59:59'"

	End If

	strSql = strSql & "			AND CUOTA.COD_CLIENTE = '" & strCodCliente & "'"


	if Trim(strTipoGestion) = "1" Then
		strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) "
	End If

	if Trim(strTipoGestion) = "2" Then
		strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 13)"
	End If

	if Trim(strEstadoRevision) = "1" Then
		strSql = strSql & " AND	  GESTIONES.FECHA_COMPROMISO = CAST(CONVERT(VARCHAR(10),(CASE WHEN DATENAME(dw,GETDATE())='LUNES' THEN GETDATE()- 3 ELSE GETDATE()-1 END),103) AS DATETIME)"
	End If

	if Trim(strEstadoRevision) = "2" Then
		strSql = strSql & " AND	  GESTIONES.FECHA_COMPROMISO <= CAST(CONVERT(VARCHAR(10),(CASE WHEN DATENAME(dw,GETDATE())='LUNES' THEN GETDATE()- 4 ELSE GETDATE()-2 END),103) AS DATETIME)"
	End If

	
	if Trim(strRegion) = "1" Then
		strSql = strSql & " AND (CO.CODIGO_REGION=13 OR DD.COMUNA IS NULL)"
	End If

	if Trim(strRegion) = "2" Then
		strSql = strSql & " AND CO.CODIGO_REGION<>13"
	End If

	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
		strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
	Else
		if Trim(intIdUsuario) <> "0" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & intIdUsuario & "'"
		End if
	End if


	strSql = strSql & " ORDER BY  GESTIONES.FECHA_INGRESO,CUOTA.RUT_DEUDOR,USUARIO.login"

	'response.write " strSql = " & strSql
	'response.end
	
	if strSql <> "" then
		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			intReg = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>
				<tr>
					<td><%=intReg%></td>
					<td><%=Mid(rsDet("GESTIONSOLA"),1,18)%></td>
					<td><%=Mid(rsDet("ESTADO_REVISION"),1,18)%></td>
					<td><%=rsDet("FECHA_NORMALIZACION")%></td>
					<td>
						<A HREF="principal.asp?TX_RUT=<%=rsDet("RUT_DEUDOR")%>">
						<acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym>
						</A>
					</td>
					<td title="<%=rsDet("NOMBRE_DEUDOR")%>"><%=Mid(rsDet("NOMBRE_DEUDOR"),1,35)%></td>
					
					<td ALIGN="RIGHT"><%=FN(rsDet("SALDO"),0)%></td>
					
					<td ALIGN="CENTER" title="<%=rsDet("FORMA_PAGO")%>"><%=Mid(rsDet("FORMA_PAGO"),1,15)%></td>
					
					<td title="<%=rsDet("LUGAR_PAGO")%>"><%=Mid(rsDet("LUGAR_PAGO"),1,25)%></td>
					<td title="<%=rsDet("COMUNA")%>"><%=Mid(rsDet("COMUNA"),1,20)%></td>

					<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
					<td ALIGN="RIGHT"><%=rsDet("EJEC_ASIG")%></td>
					<td ALIGN="CENTER" title="<%=rsDet("OBSERVACIONES_CAMPO")%>">
						<img src="../imagenes/priorizar_normal.png" border="0">
					</td>
				</tr>
				<%
				rsDet.movenext
			loop%>

			</tbody>
			
			<thead>
				<tr >
					<td colspan = "13">&nbsp;</td>
				</tr>
			</thead>
					
		<%Else%>

		<thead>
			<tr >
				<td colspan = "13">&nbsp;</td>
			</tr>
		</thead>
		
		<tr class="estilo_columnas">
			<td ALIGN="CENTER" Colspan = "13">NO EXISTEN RESULTADOS SEGUN PARAMETROS DE BUSQUEDA</td>
		</tr>
		
		<thead>
			<tr >
				<td colspan = "13">&nbsp;</td>
			</tr>
		</thead>
		
		<%end if
	end if%>
	</tbody>
  </table>
<br>
<br>
</form>


</body>
</html>


