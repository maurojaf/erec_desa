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
	
	strEjeAsig = request("CB_EJECUTIVO")
	strTipoGestion = request("cmb_tipogestion")
	strEstadoNorm = request("CMB_ESTADO_NORM")

	if strTipoGestion = "" then strTipoGestion = "0"
	if strEstadoNorm = "" then strEstadoNorm = "2"

	termino = request("termino")
	inicio = request("inicio")

	strCodCliente = session("ses_codcli")

	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_TipoGestionNorm") = strTipoGestion
		session("Ftro_EjecAsigNorm") = strEjeAsig
		session("FtroCB_ESTADO_NORM") = strEstadoNorm
	End If

	If Trim(Request("strBuscar")) = "N" Then
		session("Ftro_TipoGestionNorm") = ""
		session("Ftro_EjecAsigNorm") = ""
		session("FtroCB_ESTADO_NORM") =""
	End If

	If strEjeAsig = "" Then strEjeAsig = session("Ftro_EjecAsigNorm")
	If strTipoGestion = "0" Then strTipoGestion = session("Ftro_TipoGestionNorm")
	If strEstadoNorm = "" Then strEstadoNorm = session("FtroCB_ESTADO_NORM")

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



<script>
$(document).ready(function(){

	$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
 	$(document).tooltip();
})
</script>
<script language="JavaScript " type="text/JavaScript">

function envia()
{
	//datos.TX_RUT.value='';
	//datos.TX_PAGO.value='';
	document.datos.action = "listado_normalizacion.asp?strBuscar=S";
	document.datos.submit();
}

function exportar()
{
	document.datos.action = "exp_Normalizacion.asp";
	document.datos.submit();
}

</script>

</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<form name="datos" method="post">
<div class="titulo_informe">MÓDULO NORMALIZACIÓN</div>

<br>
	<table width="90%" align="center" class="estilo_columnas">
		<thead>
	      <tr height="20" >
			<td>COBRANZA</td>
			<td>TIPO GESTIÓN</td>
			<td>ESTADO</td>
			<td>FECHA DESDE</td>
			<td>FECHA HASTA</td>

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
					<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>INDICA QUE PAGO</option>
					<option value="2" <%If Trim(strTipoGestion)="2" Then Response.write "SELECTED"%>>COMPROMISO D & T</option>
					<option value="3" <%If Trim(strTipoGestion)="3" Then Response.write "SELECTED"%>>PAGO NO APLICADO</option>
					<option value="5" <%If Trim(strTipoGestion)="5" Then Response.write "SELECTED"%>>INDICA PAGO NO RESPONDIDO</option>
					<option value="6" <%If Trim(strTipoGestion)="6" Then Response.write "SELECTED"%>>REITERA INDICA PAGO</option>
				</SELECT>
			</td>

			<td>
				<SELECT NAME="CMB_ESTADO_NORM" id="CMB_ESTADO_NORM">
					<option value="0" <%If Trim(strEstadoNorm)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strEstadoNorm)="1" Then Response.write "SELECTED"%>>EN CONSULTA</option>
					<option value="2" <%If Trim(strEstadoNorm)="2" Then Response.write "SELECTED"%>>NO PROCESADO</option>
				</SELECT>
			</td>

			<td><input name="inicio" type="text" id="inicio" readonly="true" value="<%=inicio%>" size="10" maxlength="10">

			</td>
			<td><input name="termino" type="text" id="termino" readonly="true" value="<%=termino%>" size="10" maxlength="10">

			</td>

		<% If sinCbUsario="0" Then %>
			<td>
				<select name="CB_EJECUTIVO" ID="CB_EJECUTIVO"  s>
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
			<td Width = "440">TIPO_GESTIÓN</td>
			<td>TOTAL DOCUMENTOS</td>
			<td>TOTAL MONTO</td>
		</tr>
		</thead>
		<tbody>

	<%
	abrirscg()

	strSql = " SELECT GESTIONSOLA = (CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL)"
	strSql = strSql & " THEN 'INDICA QUE PAGO' "
	strSql = strSql & " WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) "
	strSql = strSql & " AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') "
	strSql = strSql & " THEN 'COMPROMISO D & T' "
	strSql = strSql & " WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) "
	strSql = strSql & " AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ()) "
	strSql = strSql & " THEN 'INDICA PAGO EN CONSULTA' "
	strSql = strSql & " WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) "
	strSql = strSql & " AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ()) "
	strSql = strSql & " THEN 'INDICA PAGO NO RESP.'"
	strSql = strSql & " WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME)) "
	strSql = strSql & " THEN 'REITERA INDICA PAGO' ELSE 'PAGO NO APLICADO' END),"
	strSql = strSql & " COUNT (CUOTA.ID_CUOTA) AS TOTAL,SUM(ISNULL(CAST(GESTIONES.MONTO_CANCELADO AS BIGINT),0)) AS MONTO "


	strSql = strSql & " FROM GESTIONES			INNER JOIN GESTIONES_CUOTA ON GESTIONES.ID_GESTION = GESTIONES_CUOTA.ID_GESTION"
	strSql = strSql & "                         INNER JOIN CUOTA ON GESTIONES_CUOTA.ID_CUOTA = CUOTA.ID_CUOTA AND GESTIONES_CUOTA.ID_GESTION = GESTIONES.ID_GESTION"
	strSql = strSql & "                         INNER JOIN DEUDOR ON GESTIONES.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND GESTIONES.COD_CLIENTE = DEUDOR.COD_CLIENTE"
	strSql = strSql & "                         INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql & "                         LEFT JOIN CAJA_FORMA_PAGO ON GESTIONES.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO"
	strSql = strSql & "                         INNER JOIN GESTIONES_TIPO_GESTION ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_GESTION = GESTIONES_TIPO_GESTION.COD_GESTION AND"
	strSql = strSql & "                                                                             GESTIONES.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"

	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_CATEGORIA.COD_CATEGORIA"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_CATEGORIA AND"
	strSql = strSql & " 																	GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_SUB_CATEGORIA"
	strSql = strSql & " 						LEFT JOIN USUARIO ON DEUDOR.USUARIO_ASIG = USUARIO.ID_USUARIO"


	strSql = strSql & " WHERE   ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND '" & strTipoGestion & "' IN (0,1))"
	strSql = strSql & "			  OR ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') AND '" & strTipoGestion & "' IN (0,2) AND GESTIONES.FECHA_COMPROMISO <= (GETDATE()))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 6 AND '" & strTipoGestion & "' IN (0,3) AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE () AND '" & strTipoGestion & "' IN (0,5))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND FECHA_CONSULTA_NORM IS NOT NULL AND '" & strTipoGestion & "' IN (0,6)))"

	strSql = strSql & " 		  AND ESTADO_DEUDA.ACTIVO = 1 "

	If Trim(strCobranza) = "INTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
	End if

	If Trim(strCobranza) = "EXTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
	End if

	If inicio <> "" then

	strSql = strSql & " AND (CASE WHEN ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11))"
	strSql = strSql & " 	 THEN ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'')"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (2)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (6)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & "		 ELSE ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103),'')"
	strSql = strSql & " 	 END) > = CAST('" & inicio & " 00:00:00'AS DATETIME)"

	End If

	If termino <> "" then

	strSql = strSql & " AND (CASE WHEN ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11))"
	strSql = strSql & " 	 THEN ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'')"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (2)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (6)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & "		 ELSE ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103),'')"
	strSql = strSql & " 	 END) < = CAST('" & termino & " 23:59:59' AS DATETIME)"

	End If

	strSql = strSql & " AND CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION "
	strSql = strSql & "	AND CUOTA.COD_CLIENTE = '" & strCodCliente & "'"


	if Trim(strTipoGestion) = "1" Then
		strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL)"
	End If

	if Trim(strTipoGestion) = "2" Then
		strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) "
		strSql = strSql & " 	AND  (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') AND GESTIONES.FECHA_COMPROMISO <= (GETDATE())"
	End If

	if Trim(strTipoGestion) = "3" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 6) "
	End If

	if Trim(strTipoGestion) = "4" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
	End If

	if Trim(strTipoGestion) = "5" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
	End If

	if Trim(strTipoGestion) = "6" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND FECHA_CONSULTA_NORM IS NOT NULL )"
	End If


	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
		strSql = strSql & " 	AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
	Else
		if Trim(strEjeAsig) <> "" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
		End if
	End if

	strSql = strSql & " GROUP BY "
	strSql = strSql & " 	 	(CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL) "
	strSql = strSql & " 	 	THEN 'INDICA QUE PAGO' "
	strSql = strSql & " 	 	WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) "
	strSql = strSql & " 	 	AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') "
	strSql = strSql & " 	 	THEN 'COMPROMISO D & T' "
	strSql = strSql & " 	 	WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) "
	strSql = strSql & " 	 	AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ()) "
	strSql = strSql & " 	 	THEN 'INDICA PAGO EN CONSULTA' "
	strSql = strSql & " 	 	WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) "
	strSql = strSql & " 	 	AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ()) "
	strSql = strSql & " 	 	THEN 'INDICA PAGO NO RESP.' "
	strSql = strSql & " 	 	WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME)) "
	strSql = strSql & " 	 	THEN 'REITERA INDICA PAGO' ELSE 'PAGO NO APLICADO' END) "

	'Response.write "strSql = " & strSql
	''Response.write "strSql = " & strTipoGestion

	''Response.write "strSql = " & resp
		'Response.End
	if strSql <> "" then
		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			intReg = 0
			TotalCasos = 0
			TotalMonto = 0
			do while not rsDet.eof
				intReg = intReg + 1

				TotalCasos = TotalCasos + Cdbl(rsDet("TOTAL"))
				TotalMonto = TotalMonto + Cdbl(rsDet("MONTO"))

				%>
				<tr >
					<td><%=Mid(rsDet("GESTIONSOLA"),1,28)%></td>
					<td ALIGN="RIGHT"><%=Cdbl(rsDet("TOTAL"))%></td>
					<td ALIGN="RIGHT"><%=FN(Cdbl(rsDet("MONTO")),0)%></td>

				</tr>
				<%
				Response.flush()
				rsDet.movenext
			loop
		rsDet.close
		set rsDet=nothing
		end if

	cerrarscg()%>
	</tbody>
		<%if TotalCasos = 0 then%>

				<tr >
					<td ALIGN="CENTER" Colspan = "3" class="estilo_columna_individual">NO EXISTEN RESULTADOS SEGUN PARAMETROS DE BUSQUEDA</td>
				</tr>

		<%Else%>

				<tr class="totales" >
					<td>TOTALES</td>
					<td ALIGN="RIGHT"><%=FN(TotalCasos,0)%></td>
					<td ALIGN="RIGHT"><%=FN(TotalMonto,0)%></td>
				</tr>

		<%end if%>

	<%end if%>
	
	</table>
	<br>
	<table width="100%" border="0" class="intercalado" align="center">
	<thead>
		<%if TotalCasos = 0 then%>

			<tr >
				<td class="estilo_columna_individual" ALIGN="CENTER" Colspan = "3">&nbsp;</td>
			</tr>

		<%Else%>

		<tr>
			<td>&nbsp;</td>
			<td>GESTION</td>
			<td>FECHA</td>
			<td>RUT DEUDOR</td>
			<td>NOMBRE DEUDOR</td>
			<td>NRO_DOC</td>
			<td>MONTO</td>
			<td>FORMA PAGO</td>
			<td>LUGAR PAGO</td>
			<td>NRO CP</td>
			<td>EJEC.ASIG.</td>
			<td>DM</td>
			<td>OBS.</td>
			<td>&nbsp;</td>

		</tr>

		<%end if%>
	</thead>
	<tbody>
	<%
abrirscg()

	strSql = "SELECT DEUDOR.NOMBRE_DEUDOR, USUARIO.LOGIN AS EJEC_ASIG, CAST(GETDATE()-CUOTA.FECHA_VENC AS INT) AS DM, "
	strSql = strSql & " CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103) as 'FECHA_INGRESO', (CASE WHEN GESTIONES.NRO_DOC_PAGO = '' THEN 'NO ESPEC'ELSE GESTIONES.NRO_DOC_PAGO END) AS NRO_DOC_PAGO,"
	strSql = strSql & " 'FECHA CONSULTA: ' + ISNULL(CONVERT(VARCHAR(10),CUOTA.FECHA_CONSULTA_NORM,103),'NO CONSULTADO') AS FECHA_CONSULTA,"
	strSql = strSql & " GESTIONSOLA = CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL)"
	strSql = strSql & " 				   THEN 'INDICA QUE PAGO'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP')"
	strSql = strSql & "                    THEN 'COMPROMISO D & T'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
	strSql = strSql & "                    THEN 'INDICA PAGO EN CONSULTA'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
	strSql = strSql & "                    THEN 'INDICA PAGO NO RESP.'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME))"
	strSql = strSql & "                    THEN 'REITERA INDICA PAGO'"
	strSql = strSql & " 				   ELSE 'PAGO NO APLICADO'"
	strSql = strSql & "					   END,"


	strSql = strSql & " CUOTA.RUT_DEUDOR AS RUT_DEUDOR,"
	strSql = strSql & " CUOTA.NRO_DOC,"
	strSql = strSql & " ISNULL(GESTIONES.MONTO_CANCELADO,0) AS SALDO,"

	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES,CHAR(13),' '),CHAR(10),' ') as OBSERVACIONES,"

	strSql = strSql & " (CASE WHEN REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')<>'' THEN "
	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')+'							'+ (CASE WHEN GESTIONES.HORA_DESDE <> '' THEN'/HORARIO: '+ GESTIONES.HORA_DESDE+' - '+ GESTIONES.HORA_HASTA ELSE '' END)"
	strSql = strSql & " ELSE 'SIN OBSERVACIÓN' +'							'+ (CASE WHEN GESTIONES.HORA_DESDE <> '' THEN'/HORARIO: '+ GESTIONES.HORA_DESDE+' - '+ GESTIONES.HORA_HASTA ELSE '' END)"
	strSql = strSql & " END) as OBSERVACIONES_CAMPO,"

	strSql = strSql & " CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (1,11) )"
	strSql = strSql & " 	 THEN ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'')"
	strSql = strSql & " 	 WHEN ( GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (2) AND GESTIONES.FECHA_PAGO IS NOT NULL)"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103)"
	strSql = strSql & " 	 WHEN ( GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (6) AND GESTIONES.FECHA_PAGO IS NOT NULL)"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & "		 ELSE 'NO ESPEC'"
	strSql = strSql & " 	 END AS FECHA_NORMALIZACION,"

	strSql = strSql & " LUGAR_PAGO = UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))),"

	strSql = strSql & " ISNULL(CAJA_FORMA_PAGO.DESC_FORMA_PAGO,'NO ESPEC.') AS 'FORMA PAGO',"

	strSql = strSql & " CASE WHEN ISNULL(GESTIONES.MONTO_CANCELADO,'')= 0"
	strSql = strSql & " THEN ''"
	strSql = strSql & " ELSE ISNULL(GESTIONES.MONTO_CANCELADO,'NO ESPEC')"
	strSql = strSql & " END AS 'MONTO_REGULARIZADO', CUOTA.ID_CUOTA "


	strSql = strSql & " ,( "
	strSql = strSql & " SELECT COUNT(*) "
	strSql = strSql & " FROM CARGA_ARCHIVOS_CUOTA car "
	strSql = strSql & " WHERE CAR.ID_CUOTA =CUOTA.ID_CUOTA AND car.activo=1 "
	strSql = strSql & " ) CANTIDAD_DOCUMENTOS "


	strSql = strSql & " FROM GESTIONES			INNER JOIN GESTIONES_CUOTA ON GESTIONES.ID_GESTION = GESTIONES_CUOTA.ID_GESTION"
	strSql = strSql & "                         INNER JOIN CUOTA ON GESTIONES_CUOTA.ID_CUOTA = CUOTA.ID_CUOTA"
	strSql = strSql & "                         INNER JOIN DEUDOR ON GESTIONES.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND GESTIONES.COD_CLIENTE = DEUDOR.COD_CLIENTE"
	strSql = strSql & "                         INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql & "                         LEFT JOIN CAJA_FORMA_PAGO ON GESTIONES.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO"
	strSql = strSql & "                         INNER JOIN GESTIONES_TIPO_GESTION ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_GESTION = GESTIONES_TIPO_GESTION.COD_GESTION AND"
	strSql = strSql & "                                                                             GESTIONES.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"

	strSql = strSql & " LEFT JOIN FORMA_RECAUDACION RE ON RE.ID_FORMA_RECAUDACION= GESTIONES.ID_FORMA_RECAUDACION "
	strSql = strSql & " LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION= GESTIONES.ID_DIRECCION_COBRO_DEUDOR "

	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_CATEGORIA.COD_CATEGORIA"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_CATEGORIA AND"
	strSql = strSql & " 																	GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_SUB_CATEGORIA"
	strSql = strSql & " 						LEFT JOIN USUARIO ON DEUDOR.USUARIO_ASIG = USUARIO.ID_USUARIO"


	strSql = strSql & " WHERE   (	 ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') AND '" & strTipoGestion & "' IN (0,2) AND GESTIONES.FECHA_COMPROMISO <= (GETDATE()))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 6 AND '" & strTipoGestion & "' IN (0,3) AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE () AND '" & strTipoGestion & "' IN (0,5))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND FECHA_CONSULTA_NORM IS NOT NULL AND '" & strTipoGestion & "' IN (0,6))"

	If Trim(strEstadoNorm) = "0" or Trim(strEstadoNorm) = "2" Then
		strSql = strSql & " OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2  AND FECHA_CONSULTA_NORM IS NULL AND '" & strTipoGestion & "' IN (0,1)))"
	Else
		strSql = strSql & " OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2  AND '" & strTipoGestion & "' IN (0,1)))"
	End If

	strSql = strSql & " 		  AND ESTADO_DEUDA.ACTIVO = 1 "

	If Trim(strCobranza) = "INTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
	End if

	If Trim(strCobranza) = "EXTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
	End if

	If Trim(strEstadoNorm) = "1" Then
		strSql = strSql & "		AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
	End If

	If inicio <> "" then

	strSql = strSql & " AND (CASE WHEN ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11))"
	strSql = strSql & " 	 THEN ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'')"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (2)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (6)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & "		 ELSE ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103),'')"
	strSql = strSql & " 	 END) > = CAST('" & inicio & " 00:00:00'AS DATETIME)"

	End If

	If termino <> "" then

	strSql = strSql & " AND (CASE WHEN ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11))"
	strSql = strSql & " 	 THEN ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'')"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (2)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & " 	 WHEN (	   (GESTIONES_TIPO_GESTION.GESTION_MODULOS IN (6)))"
	strSql = strSql & " 	 THEN CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103)"
	strSql = strSql & "		 ELSE ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103),'')"
	strSql = strSql & " 	 END) < = CAST('" & termino & " 23:59:59' AS DATETIME)"

	End If

	strSql = strSql & " AND CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION "
	strSql = strSql & "	AND CUOTA.COD_CLIENTE = '" & strCodCliente & "'"


	if Trim(strTipoGestion) = "1" Then
		strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL)"
	End If

	if Trim(strTipoGestion) = "2" Then
		strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) "
		strSql = strSql & " 	AND  (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') AND GESTIONES.FECHA_COMPROMISO <= (GETDATE())"
	End If

	if Trim(strTipoGestion) = "3" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 6) "
	End If

	if Trim(strTipoGestion) = "4" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
	End If

	if Trim(strTipoGestion) = "5" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
	End If

	if Trim(strTipoGestion) = "6" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND FECHA_CONSULTA_NORM IS NOT NULL )"
	End If


	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
		strSql = strSql & " 	AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
	Else
		if Trim(strEjeAsig) <> "" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
		End if
	End if

	strSql = strSql & " ORDER BY  GESTIONSOLA, CUOTA.RUT_DEUDOR, FECHA_NORMALIZACION"

	if strSql <> "" then
		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			intReg = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>

				<tr >
					<td><%=intReg%></td>
					<td ALIGN="CENTER" title="<%=rsDet("FECHA_CONSULTA")%>">
					<%=Mid(rsDet("GESTIONSOLA"),1,25)%>
					<td><%=rsDet("FECHA_NORMALIZACION")%></td>
					<td>
											<A HREF="principal.asp?TX_RUT=<%=rsDet("RUT_DEUDOR")%>">
											<acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym>
											</A>
					</td>

					<td Align="LEFT" title="<%=rsDet("NOMBRE_DEUDOR")%>">
						<%=Mid(rsDet("NOMBRE_DEUDOR"),1,25)%></td>

					<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("MONTO_REGULARIZADO"),0)%></td>
					<td><%=Mid(rsDet("FORMA PAGO"),1,15)%></td>
					<td><%=Mid(rsDet("LUGAR_PAGO"),1,20)%></td>
					<td ALIGN="RIGHT"><%=rsDet("NRO_DOC_PAGO")%></td>
					<td ALIGN="LEFT"><%=rsDet("EJEC_ASIG")%></td>
					<td ALIGN="CENTER"><%=rsDet("DM")%></td>
					<td ALIGN="CENTER" title="<%=rsDet("OBSERVACIONES_CAMPO")%>">
						<img src="../imagenes/priorizar_normal.png" border="0">
					</td>
					<td>



							<%IF trim(rsDet("CANTIDAD_DOCUMENTOS"))>0 then%>
								<img src="../Imagenes/48px-Crystal_Clear_filesystem_folder_yellow.png" width="20" height="20" style="cursor:pointer;" alt="Historial documentos adjuntos" onclick="bt_ver_historial('<%=trim(rsDet("ID_CUOTA"))%>')">
							<%else%>
								<img src="../Imagenes/48px-Crystal_Clear_filesystem_folder_red.png" width="20" height="20" style="cursor:pointer;" alt="Sin Historial documentos adjuntos" onclick="bt_ver_historial('<%=trim(rsDet("ID_CUOTA"))%>')">
							<%end if%>

					</td>
				</tr>

				<%
				rsDet.movenext
			loop
		rsDet.close
		set rsDet=nothing
		end if

	cerrarscg()

		if TotalCasos > 0 then%>

			<tr class="totales">
				<td ALIGN="CENTER" Colspan = "14">&nbsp;</td>
			</tr>

		<%end if

	end if%>
	</tbody>
	</table>

</form>


</body>
</html>


<script type="text/javascript">
function bt_ver_historial(ID_CUOTA)
{

	window.open('historial_documentos_biblioteca_deudor.asp?ID_CUOTA='+ID_CUOTA,"_new","width=900, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
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
