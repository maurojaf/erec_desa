<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>

<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
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


<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	'cod_caja=110

	cod_caja=Session("intCodUsuario")

	AbrirSCG()

	sucursal = request("cmb_sucursal")
	intTipoPago = request("CB_TIPOPAGO")

	'response.write(perfil)
	intCodPago = request("TX_PAGO")
	strRut=request("TX_RUT")
	if sucursal="" then sucursal="0"
	'response.write(sucursal)
	usuario = request("cmb_usuario")
	strTipoGestion = request("cmb_tipogestion")

	if usuario = "" then usuario = "0"
	if strTipoGestion = "" then strTipoGestion = "0"


	'response.write "<br>strTipoGestion=" & strTipoGestion

	''response.write "<br>strBuscar=" & Request("strBuscar")



	termino = request("termino")
	inicio = request("inicio")
	resp = request("resp")

	CLIENTE = REQUEST("CLIENTE")
	'hoy=date

	intCOD_CLIENTE = session("ses_codcli")

	''response.write "<br>Ftro_TipoGestionNorm=" & session("Ftro_TipoGestionNorm")

	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_TipoGestionNorm") = strTipoGestion
		session("Ftro_EjecAsigNorm") = usuario
	End If

	If Trim(Request("strBuscar")) = "N" Then
		session("Ftro_TipoGestionNorm") = ""
		session("Ftro_EjecAsigNorm") = ""
	End If

	If usuario = "0" Then usuario = session("Ftro_EjecAsigNorm")
	If strTipoGestion = "0" Then strTipoGestion = session("Ftro_TipoGestionNorm")

	''response.write "<br>strTipoGestion1=" & strTipoGestion
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<LINK href="../css/isk_style.css" type=text/css rel=stylesheet>
<title>CRM FACTORINg</title>


<style type="text/css">
<!--
.Estilo13 {color: #FFFFFF}
.Estilo27 {color: #FFFFFF}
-->
</style>

<style type=text/css>

#dhtmltooltip{
position: absolute;
width: 150px;
border: 2px solid black;
padding: 2px;
background-color: lightyellow;
visibility: hidden;
z-index: 100;
/*Remove below line to remove shadow. Below line should always appear last within this CSS*/
filter: progid:DXImageTransform.Microsoft.Shadow(color=gray,direction=135);
}

</style>


<script language="JavaScript" src="../javascripts/cal2.js"></script>
<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
<script language="JavaScript" src="../javascripts/validaciones.js"></script>
<script src="../javascripts/SelCombox.js"></script>
<script src="../javascripts/OpenWindow.js"></script>




<script language="JavaScript " type="text/JavaScript">

function Refrescar()
{
	resp='Si'
	datos.action = "listado_normalizacion.asp?resp="+ resp +"";
	datos.submit();
}

function enviar(){
			datos.action = "man_Export.asp?archivo=1&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&CB_ASIGNACION=" + document.datos.CB_ASIGNACION.value + "&CH_ACTIVO=" + document.datos.CH_ACTIVO.checked;
			datos.submit()
}

function Ingresa()
{
	with( document.datos )
	{
		action = "listado_normalizacion.asp";
		submit();
	}
}

function Reversar(cod_pago)
{
	with( document.datos )
	{
		//alert("Opción deshabilitada");


	if (confirm("¿ Está seguro de reversar el pago ? El pago se eliminará completamente y la deuda será reversada, volviendo a su estado original antes del pago."))
		{
			action = "reversar_pago.asp?cod_pago=" + cod_pago;
			submit();
		}
	else
		alert("Reverso del pago cancelado");
	}
}

function Modificar(cod_pago)
{
	with( document.datos )
	{
		action = "modif_caja_web2.asp?strOrigen=listado_normalizacion.asp&cod_pago=" + cod_pago;
		submit();
	}
}

function envia()
{
	//datos.TX_RUT.value='';
	//datos.TX_PAGO.value='';
	resp='si'
	document.datos.action = "listado_normalizacion.asp?strBuscar=S&resp="+ resp +"";
	document.datos.submit();
}

function exportar()
{
	document.datos.action = "exp_Normalizacion.asp";
	document.datos.submit();
}

function imprimir()
{
	datos.action = "imprime_comprobantes.asp";
	datos.submit();
}


function envia_excel(URL){

window.open(URL,"INFORMACION","width=200, height=200, scrollbars=yes, menubar=yes, location=yes, resizable=yes")
}
</script>


<link href="../css/style.css" rel="Stylesheet">

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">


<form name="datos" method="post">

<table width="100%" border="0" bordercolor="#999999">
  <tr>
    <td height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" ALIGN="CENTER">MÓDULO NORMALIZACION</td>
  </tr>
</table>

  <tr>
    <td valign="top">
	<table width="100%" border="0" bordercolor="#999999">

	      <tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
	      	<% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
			<td>EJEC.ASIG.</td>
			<% End If %>
			<td>TIPO GESTIÓN</td>
			<td>F. INGRESO GESTION DESDE</td>
			<td>F. INGRESO GESTION HASTA</td>
			<td>&nbsp;</td>

	      </tr>
		  <tr bordercolor="#999999" class="Estilo8">
		  <% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
			<td>
				<SELECT NAME="cmb_usuario" id="cmb_usuario" onChange="envia();">
					<option value="0">TODOS</option>
					<%
					stsSql="SELECT DISTINCT ID_USUARIO, LOGIN FROM USUARIO WHERE ACTIVO = 1 AND (PERFIL_COB = 1 or PERFIL_SUP = 1) AND PERFIL_EMP <>1 AND PERFIL_ADM <>1"
					set rsUsu=Conn.execute(stsSql)
					if not rsUsu.eof then
						do until rsUsu.eof
						%>
						<option value="<%=rsUsu("ID_USUARIO")%>"
						<%if Trim(usuario)=Trim(rsUsu("ID_USUARIO")) then
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
			<td>
				<SELECT NAME="cmb_tipogestion" id="cmb_tipogestion" onChange="envia();">
					<option value="0" <%If Trim(strTipoGestion)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>INDICA QUE PAGO</option>
					<option value="2" <%If Trim(strTipoGestion)="2" Then Response.write "SELECTED"%>>COMPROMISO D & T</option>
					<option value="3" <%If Trim(strTipoGestion)="3" Then Response.write "SELECTED"%>>PAGO NO APLICADO</option>
					<option value="4" <%If Trim(strTipoGestion)="4" Then Response.write "SELECTED"%>>INDICA PAGO EN CONSULTA</option>
					<option value="5" <%If Trim(strTipoGestion)="5" Then Response.write "SELECTED"%>>INDICA PAGO NO RESPONDIDO</option>
					<option value="6" <%If Trim(strTipoGestion)="6" Then Response.write "SELECTED"%>>REITERA INDICA PAGO</option>

				</SELECT>
			</td>
			<td><input name="inicio" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
		<a href="javascript:showCal('Calendar7');"><img src="../imagenes/calendario.gif" border="0"></a>
			</td>
			<td><input name="termino" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
          		<a href="javascript:showCal('Calendar6');"><img src="../imagenes/calendario.gif" border="0"></a>
			</td>
			<td align="CENTER">
				<input type="Button" name="Submit" value="Ver" onClick="envia();">
				<input Name="SubmitButton" Value="Exportar" Type="BUTTON" onClick="exportar();">
			</td>
	      </tr>
    </table>

	<table width="100%" border="0" bordercolor="#000000">
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td Width = "440">TIPO_GESTIÓN</td>
			<td>TOTAL CASOS</td>
			<td>TOTAL MONTO</td>
		</tr>

	<%


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
	strSql = strSql & "                         INNER JOIN CUOTA ON GESTIONES_CUOTA.ID_CUOTA = CUOTA.ID_CUOTA"
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


	strSql = strSql & " WHERE   ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL AND '" & strTipoGestion & "' IN (0,1))"
	strSql = strSql & "			  OR ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') AND '" & strTipoGestion & "' IN (0,2) AND GESTIONES.FECHA_COMPROMISO <= (GETDATE()))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 6 AND '" & strTipoGestion & "' IN (0,3) AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ()) AND '" & strTipoGestion & "' IN (0,4)"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE () AND '" & strTipoGestion & "' IN (0,5))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND FECHA_CONSULTA_NORM IS NOT NULL AND '" & strTipoGestion & "' IN (0,6)))"

	strSql = strSql & " 		  AND ESTADO_DEUDA.ACTIVO = 1 "

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
	strSql = strSql & "	AND CUOTA.COD_CLIENTE = '" & intCOD_CLIENTE & "'"


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
		if Trim(usuario) <> "0" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & usuario & "'"
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



	''Response.write "strSql = " & strSql
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
				<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
					<td><%=Mid(rsDet("GESTIONSOLA"),1,28)%></td>
					<td ALIGN="RIGHT"><%=Cdbl(rsDet("TOTAL"))%></td>
					<td ALIGN="RIGHT"><%=FN(Cdbl(rsDet("MONTO")),0)%></td>

				</tr>
				<%
				rsDet.movenext
			loop
		end if%>

		<%if TotalCasos = 0 then%>

				<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
					<td ALIGN="CENTER" Colspan = "3">NO EXISTEN CASOS PENDIENTES A NORMALIZAR</td>
				</tr>

		<%Else%>

				<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					<td>TOTALES</td>
					<td ALIGN="RIGHT"><%=FN(TotalCasos,0)%></td>
					<td ALIGN="RIGHT"><%=FN(TotalMonto,0)%></td>
				</tr>

		<%end if%>

	<%end if%>


	</table>



	<table width="100%" border="0" bordercolor="#000000">

		<%if TotalCasos = 0 then%>

			<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td ALIGN="CENTER" Colspan = "3">&nbsp;</td>
			</tr>

		<%Else%>

		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
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

		</tr>

		<%end if%>

	<%


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
	strSql = strSql & " END AS 'MONTO_REGULARIZADO'"


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


	strSql = strSql & " WHERE   (	 (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND FECHA_CONSULTA_NORM IS NULL AND '" & strTipoGestion & "' IN (0,1))"
	strSql = strSql & "			  OR ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 11) AND (CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'TR' OR CAJA_FORMA_PAGO.ID_FORMA_PAGO = 'DP') AND '" & strTipoGestion & "' IN (0,2) AND GESTIONES.FECHA_COMPROMISO <= (GETDATE()))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 6 AND '" & strTipoGestion & "' IN (0,3) AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ()) AND '" & strTipoGestion & "' IN (4)"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE () AND '" & strTipoGestion & "' IN (0,5))"
	strSql = strSql & "			  OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 2 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND FECHA_CONSULTA_NORM IS NOT NULL AND '" & strTipoGestion & "' IN (0,6)))"

	strSql = strSql & " 		  AND ESTADO_DEUDA.ACTIVO = 1 "

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
	strSql = strSql & "	AND CUOTA.COD_CLIENTE = '" & intCOD_CLIENTE & "'"


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
		if Trim(usuario) <> "0" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & usuario & "'"
		End if
	End if

	strSql = strSql & " ORDER BY  GESTIONSOLA, CUOTA.RUT_DEUDOR, FECHA_NORMALIZACION"

	''Response.write "strSql = " & strSql
	''Response.write "strSql = " & strTipoGestion


		''Response.write "strSql = " & strSql
		'Response.End
	if strSql <> "" then
		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			intReg = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>

				<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
					<td><%=intReg%></td>
					<td ALIGN="CENTER" onMouseover="ddrivetip('<%=rsDet("FECHA_CONSULTA")%>', '#EFEFEF',300)"; onMouseout="hideddrivetip()">
					<%=Mid(rsDet("GESTIONSOLA"),1,25)%>
					<td><%=rsDet("FECHA_NORMALIZACION")%></td>
					<td>
											<A HREF="principal.asp?TX_RUT=<%=rsDet("RUT_DEUDOR")%>">
											<acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym>
											</A>
					</td>
					<td><%=Mid(rsDet("NOMBRE_DEUDOR"),1,28)%></td>
					<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("MONTO_REGULARIZADO"),0)%></td>
					<td><%=Mid(rsDet("FORMA PAGO"),1,15)%></td>
					<td><%=Mid(rsDet("LUGAR_PAGO"),1,20)%></td>
					<td ALIGN="RIGHT"><%=rsDet("NRO_DOC_PAGO")%></td>
					<td ALIGN="LEFT"><%=rsDet("EJEC_ASIG")%></td>
					<td ALIGN="CENTER"><%=rsDet("DM")%></td>
					<td ALIGN="CENTER" onMouseover="ddrivetip('<%=rsDet("OBSERVACIONES_CAMPO")%>', '#EFEFEF',300)"; onMouseout="hideddrivetip()">
						<img src="../imagenes/priorizar_normal.png" border="0">
					</td>
				</tr>

				<%
				rsDet.movenext
			loop
		end if

		if TotalCasos > 0 then%>

			<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td ALIGN="CENTER" Colspan = "13">&nbsp;</td>
			</tr>

		<%end if

	end if%>
	</table>
	</td>
   </tr>

</form>


</body>
</html>

<div id="dhtmltooltip"></div>

<script type="text/javascript">

/***********************************************
* Cool DHTML tooltip script- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/

var offsetxpoint=-60 //Customize x offset of tooltip
var offsetypoint=20 //Customize y offset of tooltip
var ie=document.all
var ns6=document.getElementById && !document.all
var enabletip=false
if (ie||ns6)
var tipobj=document.all? document.all["dhtmltooltip"] : document.getElementById? document.getElementById("dhtmltooltip") : ""

function ietruebody(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function ddrivetip(thetext, thecolor, thewidth){
if (ns6||ie){
if (typeof thewidth!="undefined") tipobj.style.width=thewidth+"px"
if (typeof thecolor!="undefined" && thecolor!="") tipobj.style.backgroundColor=thecolor
tipobj.innerHTML=thetext
enabletip=true
return false
}
}

function positiontip(e){
if (enabletip){
var curX=(ns6)?e.pageX : event.clientX+ietruebody().scrollLeft;
var curY=(ns6)?e.pageY : event.clientY+ietruebody().scrollTop;
//Find out how close the mouse is to the corner of the window
var rightedge=ie&&!window.opera? ietruebody().clientWidth-event.clientX-offsetxpoint : window.innerWidth-e.clientX-offsetxpoint-20
var bottomedge=ie&&!window.opera? ietruebody().clientHeight-event.clientY-offsetypoint : window.innerHeight-e.clientY-offsetypoint-20

var leftedge=(offsetxpoint<0)? offsetxpoint*(-1) : -1000

//if the horizontal distance isn't enough to accomodate the width of the context menu
if (rightedge<tipobj.offsetWidth)
//move the horizontal position of the menu to the left by it's width
tipobj.style.left=ie? ietruebody().scrollLeft+event.clientX-tipobj.offsetWidth+"px" : window.pageXOffset+e.clientX-tipobj.offsetWidth+"px"
else if (curX<leftedge)
tipobj.style.left="5px"
else
//position the horizontal position of the menu where the mouse is positioned
tipobj.style.left=curX+offsetxpoint+"px"

//same concept with the vertical position
if (bottomedge<tipobj.offsetHeight)
tipobj.style.top=ie? ietruebody().scrollTop+event.clientY-tipobj.offsetHeight-offsetypoint+"px" : window.pageYOffset+e.clientY-tipobj.offsetHeight-offsetypoint+"px"
else
tipobj.style.top=curY+offsetypoint+"px"
tipobj.style.visibility="visible"
}
}

function hideddrivetip(){
if (ns6||ie){
enabletip=false
tipobj.style.visibility="hidden"
tipobj.style.left="-1000px"
tipobj.style.backgroundColor=''
tipobj.style.width=''
}
}


document.onmousemove=positiontip

</script>
