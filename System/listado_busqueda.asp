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

	strEjeAsig = request("CB_EJECUTIVO")
	strTipoGestion = request("cmb_tipogestion")

	if strTipoGestion = "" then strTipoGestion = "0"

	termino = request("termino")
	inicio = request("inicio")

	strCodCliente = session("ses_codcli")

	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_TipoGestionBusq") = strTipoGestion
		session("Ftro_EjecAsigBusq") = strEjeAsig
	End If

	If Trim(Request("strBuscar")) = "N" Then
		session("Ftro_TipoGestionBusq") = ""
		session("Ftro_EjecAsigBusq") = ""
	End If

	If strEjeAsig = "" Then strEjeAsig = session("Ftro_EjecAsigBusq")
	If strTipoGestion = "0" Then strTipoGestion = session("Ftro_TipoGestionBusq")

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

	'---Fin codigo tipo de cobranza---'
%>


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
	datos.Submit.disabled = true;
	datos.SubmitButton.disabled = true;
	document.datos.action = "listado_busqueda.asp?strBuscar=S";
	document.datos.submit();
}
function exportar()
{
	datos.SubmitButton.disabled = true;
	document.datos.action = "exp_Busqueda.asp";
	document.datos.submit();
}

</script>


</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">


<form name="datos" method="post">
<div class="titulo_informe">MÓDULO DE UBICABILIDAD</div>	
<br>

<table width="90%" border="0" class="estilo_columnas" align="center">
<thead>
  <tr height="20" >  	
	<td>COBRANZA</td>
	<td>TIPO GESTIÓN</td>
	<td>F. SOLICITUD DESDE</td>
	<td>F. SOLICITUD HASTA</td>

  <% If sinCbUsario = "0" Then %>
	<td>EJECUTIVO</td>
  <% End If %>

	<td>&nbsp;</td>

  </tr>
 </thead>
  <tr >

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
		<SELECT NAME="cmb_tipogestion" id="cmb_tipogestion" onChange="envia();">
			<option value="0" <%If Trim(strTipoGestion)="0" Then Response.write "SELECTED"%>>TODOS</option>
			<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>SOLICITUD BUSQUEDA</option>
			<option value="2" <%If Trim(strTipoGestion)="2" Then Response.write "SELECTED"%>>SOLICITUD NO RESPONDIDA</option>
			<option value="3" <%If Trim(strTipoGestion)="3" Then Response.write "SELECTED"%>>BUSQUEDA SIN RESULTADOS</option>
			<option value="4" <%If Trim(strTipoGestion)="4" Then Response.write "SELECTED"%>>INUBICABLE TERMINAL</option>
			<option value="5" <%If Trim(strTipoGestion)="5" Then Response.write "SELECTED"%>>INUBICABLE</option>
		</SELECT>
	</td>

	<td><input name="inicio" type="text" readonly="true" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
	</td>

	<td><input name="termino" type="text" readonly="true" id="termino" value="<%=termino%>" size="10" maxlength="10">
	</td>

<% If sinCbUsario="0" Then %>
	<td>
		<select name="CB_EJECUTIVO" id="CB_EJECUTIVO" >
		</select>
	</td>
<% End If %>

	<td align="CENTER">
		<input type="Button" name="Submit" value="Ver" class="fondo_boton_100" onClick="envia();">
		<input Name="SubmitButton" Value="Exportar" Type="BUTTON" class="fondo_boton_100" onClick="exportar();">
	</td>
  </tr>
</table>

<table width="100%" border="0" class="intercalado" align="center">
	<thead>
	<tr >
		<td Width = "440">ESTADO</td>
		<td>TOTAL CASOS</td>
		<td>TOTAL MONTO</td>
	</tr>
	</thead>
	<tbody>
	<%
AbrirSCG()

			strSql = " SELECT GESTION, COUNT(TOTAL_CASOS) AS TOTAL_CASOS,SUM(TOTAL_SALDO) AS TOTAL_SALDO FROM"

			strSql = strSql & "	(SELECT CUOTA.RUT_DEUDOR,MAX(CASE WHEN GESTIONES_TIPO_GESTION.GESTION_MODULOS = 9"
			strSql = strSql & "	THEN 'SOLICITUD BUSQUEDA'"
			strSql = strSql & "	WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 5)"
			strSql = strSql & "	THEN 'SOLICITUD NO RESP.'"
			strSql = strSql & "	WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4 AND (ISNULL(DEUDOR_TELEFONO.RUT_DEUDOR,'0')<>'0' OR ISNULL(DEUDOR_EMAIL.RUT_DEUDOR,'0')<>'0'))"
			strSql = strSql & " THEN 'BUSQUEDA SIN RESULTADOS'"
			strSql = strSql & "	WHEN GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4"
			strSql = strSql & " THEN 'INUBICABLE TERMINAL'"
			strSql = strSql & "	ELSE 'INUBICABLE' END) AS GESTION,"
			strSql = strSql & " COUNT(*) AS TOTAL_CASOS,"
			strSql = strSql & " SUM(SALDO) AS TOTAL_SALDO"

			strSql = strSql & " FROM CUOTA INNER JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE "
			strSql = strSql & "			   LEFT JOIN GESTIONES ON CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION"
			strSql = strSql & "			   INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
			strSql = strSql & "			   LEFT JOIN GESTIONES_TIPO_GESTION ON SUBSTRING(COD_ULT_GEST,1,1) = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND SUBSTRING(COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
			strSql = strSql & "												AND SUBSTRING(COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
			strSql = strSql & "												AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
			strSql = strSql & "			   LEFT JOIN USUARIO ON DEUDOR.USUARIO_ASIG = USUARIO.ID_USUARIO"
			strSql = strSql & "			   LEFT JOIN DEUDOR_TELEFONO ON DEUDOR_TELEFONO.RUT_DEUDOR = CUOTA.RUT_DEUDOR AND DEUDOR_TELEFONO.ESTADO IN (0,1)"
			strSql = strSql & "			   LEFT JOIN DEUDOR_EMAIL ON DEUDOR_EMAIL.RUT_DEUDOR = CUOTA.RUT_DEUDOR AND DEUDOR_EMAIL.ESTADO IN (0,1)"


			strSql = strSql & " WHERE (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 9 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4"
			strSql = strSql & " OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 5 AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"

			strSql = strSql & " OR ((DEUDOR.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0,1))"
			strSql = strSql & " AND DEUDOR.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0,1))) AND ISNULL(GESTIONES_TIPO_GESTION.GESTION_MODULOS,0) NOT IN (5,10)))"

			If Trim(strCobranza) = "INTERNA" Then
				strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
				strParametro = "1"
			End if

			If Trim(strCobranza) = "EXTERNA" Then
				strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
				strParametro = "1"
			End if

			If inicio <> "" then

			strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) > = '" & inicio & " 00:00:00'"

			End If

			If termino <> "" then

			strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) < = '" & termino & " 23:59:59'"

			End If

			strSql = strSql & " AND ESTADO_DEUDA.ACTIVO = 1"
			strSql = strSql & " AND CUOTA.COD_CLIENTE = '" & strCodCliente & "'"

			If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then

			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"

			Else
				if Trim(strEjeAsig) <> "" Then
				strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
				End if
			End if

			strSql = strSql & "	GROUP BY CUOTA.RUT_DEUDOR) AS PP"

			strSql = strSql & "	GROUP BY GESTION"

			''Response.write "strSql = " & strSql

			if strSql <> "" then
				set rsDet=Conn.execute(strSql)

				if not rsDet.eof then

					TotalCasos = 0
					TotalMonto = 0
					do while not rsDet.eof

						TotalCasos = TotalCasos + rsDet("TOTAL_CASOS")
						TotalMonto = TotalMonto + rsDet("TOTAL_SALDO")

						%>
						<tr >
							<td><%=Mid(rsDet("GESTION"),1,28)%></td>
							<td ALIGN="RIGHT"><%=Cdbl(rsDet("TOTAL_CASOS"))%></td>
							<td ALIGN="RIGHT"><%=FN(Cdbl(rsDet("TOTAL_SALDO")),0)%></td>

						</tr>
						<%
						Response.flush()
						rsDet.movenext
					loop
				rsDet.close
				set rsDet=nothing
				end if%>
			</tbody>
			<thead>
				<%if TotalCasos = 0 then%>

						<tr class="estilo_columnas">
							<td ALIGN="CENTER" Colspan = "3">NO EXISTEN RESULTADOS SEGUN PARAMETROS DE BUSQUEDA</td>
						</tr>

				<%Else%>

						<tr class="totales">
							<td>TOTALES</td>
							<td ALIGN="RIGHT"><%=FN(TotalCasos,0)%></td>
							<td ALIGN="RIGHT"><%=FN(TotalMonto,0)%></td>
						</tr>

				<%end if
			end if
cerrarscg()%>
			</thead>
</table>

<table width="100%" border="0" class="intercalado" align="center">

<%
			
AbrirSCG()
			strSql = " 			SELECT * FROM"
			strSql = strSql & "	(SELECT MIN(ISNULL(CUOTA.PRIORIDAD_CUOTA,99)) AS PRIORIDAD,"
			strSql = strSql & "	MAX((CASE WHEN GESTIONES_TIPO_GESTION.GESTION_MODULOS = 9 AND '" & strTipoGestion & "' IN (0,1,5)"
			strSql = strSql & "	THEN 'SOLICITUD BUSQUEDA'"
			strSql = strSql & "	WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 5 AND '" & strTipoGestion & "' IN (0,2,5) )"
			strSql = strSql & "	THEN 'SOLICITUD NO RESP.'"
			strSql = strSql & "	WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4 AND (ISNULL(DEUDOR_TELEFONO.RUT_DEUDOR,'0') <> '0' OR ISNULL(DEUDOR_EMAIL.RUT_DEUDOR,'0')<>'0') AND '" & strTipoGestion & "' IN (0,3,5) )"
			strSql = strSql & " THEN 'BUSQUEDA SIN RESULTADOS'"
			strSql = strSql & "	WHEN GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4  AND '" & strTipoGestion & "' IN (0,4,5)"
			strSql = strSql & " THEN 'INUBICABLE TERMINAL'"
			strSql = strSql & "	ELSE 'INUBICABLE' END)) as GESTION,"

			strSql = strSql & " NOMCLIENTE = MAX(CUOTA.NOMBRE_SUBCLIENTE),"
			strSql = strSql & " RUT_DEUDOR = CUOTA.RUT_DEUDOR,"
			strSql = strSql & " NOMBRE_DEUDOR = DEUDOR.NOMBRE_DEUDOR,"
			strSql = strSql & " SALDO = SUM (CUOTA.SALDO),"
			strSql = strSql & " DOC = COUNT(CUOTA.ID_CUOTA),"
			strSql = strSql & " DM = MAX(DATEDIFF(DAY,FECHA_VENC,GETDATE())),"
			strSql = strSql & "	USUARIO.LOGIN AS USUARIO,"
			strSql = strSql & "	(CASE WHEN ((CUOTA.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0,1))"
			strSql = strSql & "		  		 AND CUOTA.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0,1)))) THEN 'SIN DATOS'"
			strSql = strSql & "		  WHEN ((CUOTA.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (1))"
			strSql = strSql & "				 OR CUOTA.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (1)))) THEN 'PERD. CONTACT.'"
			strSql = strSql & "		  ELSE 'NO CONTACTADO'"
			strSql = strSql & "	 END) AS UBICABILIDAD"

			strSql = strSql & " FROM CUOTA INNER JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE "
			strSql = strSql & "			   LEFT JOIN GESTIONES ON CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION"
			strSql = strSql & "			   INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
			strSql = strSql & "			   LEFT JOIN GESTIONES_TIPO_GESTION ON SUBSTRING(COD_ULT_GEST,1,1) = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND SUBSTRING(COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
			strSql = strSql & "												AND SUBSTRING(COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
			strSql = strSql & "												AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
			strSql = strSql & "			   LEFT JOIN USUARIO ON DEUDOR.USUARIO_ASIG = USUARIO.ID_USUARIO"
			strSql = strSql & "			   LEFT JOIN DEUDOR_TELEFONO ON DEUDOR_TELEFONO.RUT_DEUDOR = CUOTA.RUT_DEUDOR AND DEUDOR_TELEFONO.ESTADO IN (0,1)"
			strSql = strSql & "			   LEFT JOIN DEUDOR_EMAIL ON DEUDOR_EMAIL.RUT_DEUDOR = CUOTA.RUT_DEUDOR AND DEUDOR_EMAIL.ESTADO IN (0,1)"

			strSql = strSql & " WHERE (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 9 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4"
			strSql = strSql & " OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 5 AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"

			strSql = strSql & " OR ((DEUDOR.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0,1))"
			strSql = strSql & " AND DEUDOR.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0,1))) AND ISNULL(GESTIONES_TIPO_GESTION.GESTION_MODULOS,0) NOT IN (5,10)))"

			If Trim(strCobranza) = "INTERNA" Then
				strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
				strParametro = "1"
			End if

			If Trim(strCobranza) = "EXTERNA" Then
				strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
				strParametro = "1"
			End if

			If inicio <> "" then

			strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) > = '" & inicio & " 00:00:00'"

			End If

			If termino <> "" then

			strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) < = '" & termino & " 23:59:59'"

			End If

			strSql = strSql & " AND ESTADO_DEUDA.ACTIVO = 1"
			strSql = strSql & " AND CUOTA.COD_CLIENTE = '" & strCodCliente & "'"

			if Trim(strTipoGestion) = "1" Then

			strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 9"

			End If

			if Trim(strTipoGestion) = "2" Then

			strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 5"

			End If

			if Trim(strTipoGestion) = "3" Then

			strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4 "
			strSql = strSql & " AND (DEUDOR.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0,1))"
			strSql = strSql & " OR DEUDOR.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0,1)))"

			End If

			if Trim(strTipoGestion) = "4" Then

			strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 4"
			strSql = strSql & " AND (DEUDOR.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0,1))"
			strSql = strSql & " AND DEUDOR.RUT_DEUDOR NOT IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0,1)))"

			End If

			If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then

			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"

			Else
				if Trim(strEjeAsig) <> "" Then
				strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
				End if
			End if


			strSql = strSql & " GROUP BY CUOTA.RUT_DEUDOR,DEUDOR.NOMBRE_DEUDOR,USUARIO.LOGIN) AS PP"

			if Trim(strTipoGestion) = "5" Then

			strSql = strSql & " WHERE GESTION = 'INUBICABLE'"

			End If

			strSql = strSql & " ORDER BY GESTION ASC, PRIORIDAD ASC, DM DESC, SALDO DESC"

			TotalCasos = 0

			''Response.write "strSql = " & strSql
			''Response.End

				if strSql <> "" then
					set rsDet=Conn.execute(strSql)

					if not rsDet.eof then%>
					<thead>
					<tr>
						<td>&nbsp;</td>
						<td>PRIO.</td>
						<td>GESTION</td>
						<td>UBICABILIDAD</td>
						<td>NOMBRE CLIENTE</td>
						<td>RUT DEUDOR</td>
						<td>NOMBRE DEUDOR</td>
						<td>SALDO</td>
						<td ALIGN="CENTER">DOC</td>
						<td ALIGN="CENTER">DM</td>
						<td>EJEC.ASIG.</td>
					</tr>					
					</thead>
					<tbody>
					<%
						do while not rsDet.eof
							intReg = intReg + 1

							TotalCasos = TotalCasos + 1

							%>
							<tr >
								<td><%=intReg%></td>
								<td ALIGN="CENTER"><%=FN(rsDet("PRIORIDAD"),0)%></td>
								<td><%=rsDet("GESTION")%></td>
								<td><%=rsDet("UBICABILIDAD")%></td>

								<td ALIGN="LEFT" title="<%=rsDet("NOMCLIENTE")%>">
								<%=Mid(rsDet("NOMCLIENTE"),1,23)%>

								<td>
									<A HREF="principal.asp?TX_RUT=<%=rsDet("RUT_DEUDOR")%>">
									<acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym>
									</A>
								</td>

								<td ALIGN="LEFT" title="<%=rsDet("NOMBRE_DEUDOR")%>">
								<%=Mid(rsDet("NOMBRE_DEUDOR"),1,23)%>

								<td ALIGN="RIGHT"><%=FN(rsDet("SALDO"),0)%></td>
								<td ALIGN="RIGHT"><%=FN(rsDet("DOC"),0)%></td>
								<td ALIGN="RIGHT"><%=FN(rsDet("DM"),0)%></td>
								<td><%=Mid(rsDet("USUARIO"),1,15)%></td>
							</tr>
							<%
							Response.flush()
							rsDet.movenext
						loop
					end if%>

				<tr class="totales">
					<td ALIGN="CENTER" Colspan = "11">&nbsp;</td>
				</tr>
			</tbody>
			<%end if
cerrarscg()%>

</table>

</form>
</body>
</html>


<script type="text/javascript">

function CargaUsuarios(subCat)
{

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



<%If sinCbUsario = "0" then%>
CargaUsuarios('<%=strCobranza%>');
<%End If%>

<%If strEjeAsig <> "" then%>
datos.CB_EJECUTIVO.value='<%=strEjeAsig%>';
<%End If%>
</script>
