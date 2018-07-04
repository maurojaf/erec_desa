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

	<!--#include file="../lib/lib.asp"-->



	<!--#include file="../lib/comunes/rutinas/chkFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/sondigitos.inc"-->
	<!--#include file="../lib/comunes/rutinas/formatoFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/validarFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	cod_caja=Session("intCodUsuario")

	AbrirSCG()
	intCodConvenio = request("TX_pago")
	strRut=request("TX_RUT")
	usuario = request("cmb_usuario")
	if usuario = "" then usuario = "0"
	termino = request("termino")
	inicio = request("inicio")
	resp = request("resp")
	resp="si"
	if Trim(inicio) = "" Then
		inicio = TraeFechaActual(Conn)
	End If
	if Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If
	CLIENTE = REQUEST("CLIENTE")

	strAgrupado=Request("strAgrupado")

	'Response.write "CLIENTE=" & CLIENTE
	'hoy=date

	'response.write(hoy)
%>
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
	 
	})
	function Refrescar()
	{
		resp='no'
		datos.action = "convenios_vencidos.asp?resp="+ resp +"";
		datos.submit();
	}

	function envia()
	{
		resp='si'
		datos.action = "convenios_vencidos.asp?resp="+ resp +"";
		datos.submit();
	}

	function enviaA()
	{
		resp='si'
		datos.action = "convenios_vencidos.asp?strAgrupado=S&resp="+ resp +"";
		datos.submit();
	}

	function envia_excel(URL){

	window.open(URL,"INFORMACION","width=200, height=200, scrollbars=yes, menubar=yes, location=yes, resizable=yes")
	}
	</script>


</head>
<body>
<form name="datos" method="post">
<div class="titulo_informe">LISTADO DE CONVENIOS VENCIDOS</div>	
<br>
<table width="90%" height="500" border="0" align="center">
  <tr>
    <td valign="top">
	<table width="100%" border="0" bordercolor="#999999">
		<tr height="20" class="Estilo8">
	        <td>RUT: </td>
			<td><INPUT TYPE="text" NAME="TX_RUT" value="<%=strRut%>" onchange="envia();"></td>
		</tr>
		<tr height="20" class="Estilo8">
	        <td>CODIGO CONVENIO: </td>
			<td ><INPUT TYPE="text" NAME="TX_pago" value="<%=intCodConvenio%>" onchange="envia();"></td>
		</tr>
	</table>

	<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
		<thead>
	      <tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
		  	<td>CLIENTE</td>
		 	<td>DESDE</td>
		 	<td>HASTA</td>
		 	<td>&nbsp;</td>
			</tr>
		</thead>
		  <tr bordercolor="#999999" class="Estilo8">
		  <td>

		<select name="CLIENTE" ID = "CLIENTE" width="15" onchange="tipopago()">
		<option value="">SELECCIONAR</option>
		<%
		ssql="SELECT * FROM CLIENTE WHERE ACTIVO=1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ")"
		set rsCLI=Conn.execute(ssql)
		if not rsCLI.eof then
			do until rsCLI.eof
			%>
			<option value="<%=rsCLI("COD_CLIENTE")%>"
			<%if Trim(cliente)=Trim(rsCLI("COD_CLIENTE")) then
				response.Write("Selected")
			end if%>
			><%=ucase(rsCLI("DESCRIPCION"))%></option>

			<%rsCLI.movenext
			loop
		end if
		rsCLI.close
		set rsCLI=nothing
		%>
        </select>
        </td>
		  <td><input name="inicio" readonly="true" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">

			</td>
		<td><input name="termino" readonly="true" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">

		</td>
		<td>
			<input type="Button" class="fondo_boton_100" name="Submit" value="Ver Detalle" onClick="envia();">
			<input type="Button" class="fondo_boton_100" name="Submit" value="Ver Agrupado" onClick="enviaA();">
		</td>
			 </tr>
    </table>
	<table width="100%" border="0" bordercolor="#000000" class="intercalado" style="width:100%;">
		<thead>
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td>COD. CONV</td>
			<td>CUOTA</td>
			<td>MONTO</td>
			<td>FECHA</td>
			<td>MANDANTE</td>
			<td>RUT</td>
			<td>CONVENIO</td>
		</tr>
		</thead>
		<tbody>
	<% If resp="si" then

			If trim(strAgrupado)="S" Then
				strSql = "SELECT D.ID_CONVENIO, COUNT(*) AS CUOTA, sum(TOTAL_CUOTA) AS TOTAL_CUOTA, CONVERT(VARCHAR(10),min(FECHA_PAGO),103) as FECHA_PAGO, E.COD_CLIENTE,E.RUT_DEUDOR , IsNull(datediff(d,mIN(FECHA_PAGO),getdate()),0) as ANTIGUEDAD FROM CONVENIO_DET D, CONVENIO_ENC E WHERE D.ID_CONVENIO = E.ID_CONVENIO AND D.FECHA_PAGO >= '" & inicio & "' AND D.FECHA_PAGO <= '" & termino & "' AND D.PAGADA IS NULL"
			Else
				strSql = "SELECT D.ID_CONVENIO, CUOTA, TOTAL_CUOTA, CONVERT(VARCHAR(10),FECHA_PAGO,103) as FECHA_PAGO, E.COD_CLIENTE,E.RUT_DEUDOR , IsNull(datediff(d,FECHA_PAGO,getdate()),0) as ANTIGUEDAD"
				strSql = strSql & " FROM CONVENIO_DET D, CONVENIO_ENC E WHERE D.ID_CONVENIO = E.ID_CONVENIO "
				strSql = strSql & " AND D.FECHA_PAGO >= '" & inicio & "' AND D.FECHA_PAGO <= '" & termino & "' AND D.PAGADA IS NULL"
			End If

			strSql = strSql & " AND E.COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ")"

			If cliente <> "" Then
				strSql = strSql & " AND E.COD_CLIENTE = '" & CLIENTE & "'"
			End if

			If strRut <> "" Then
				strSql = strSql & " AND E.RUT_DEUDOR = '" & strRut & "'"
			End if

			If intCodConvenio <> "" Then
				strSql = strSql & " AND E.ID_CONVENIO = " & intCodConvenio
			End if

			If trim(strAgrupado)="S" Then
				strSql = strSql & " GROUP BY D.ID_CONVENIO, E.COD_CLIENTE, E.RUT_DEUDOR"
			End if
		End if
		'response.write(strSql)
		'response.end
	if strSql <> "" then
		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			do while not rsDet.eof
				intTotalCuota = ValNulo(rsDet("TOTAL_CUOTA"),"N")
				intCuota = ValNulo(rsDet("CUOTA"),"N")
				If Trim(intCuota)= "0" Then
					strCuota="PIE"
				Else
					strCuota=intCuota
				End if
				intTotalCuota = ValNulo(rsDet("TOTAL_CUOTA"),"N")
				intTotalTotalCuota = intTotalTotalCuota + ValNulo(rsDet("TOTAL_CUOTA"),"N")

			%>
			<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
				<td><%=rsDet("ID_CONVENIO")%></td>
				<td><%=strCuota%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalCuota,0)%></td>
				<td><%=rsDet("FECHA_PAGO")%></td>
				<td><%=rsDet("COD_CLIENTE")%></td>
				<td>
					<A HREF="principal.asp?rut=<%=rsDet("RUT_DEUDOR")%>"><acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym></A>
				</td>
				<td><A HREF="det_convenio.asp?id_convenio=<%=rsDet("ID_CONVENIO")%>"><acronym title="VER CONVENIO">Ver </acronym></A>&nbsp;<%=rsDet("ID_CONVENIO")%></td>
				<!--td><A HREF="caja\caja_web.asp?CB_TIPOPAGO=CO&id_convenio=<%=rsDet("ID_CONVENIO")%>&rut=<%=rsDet("RUT_DEUDOR")%>"><acronym title="PAGO DE CUOTAS DEL CONVENIO">Pagar Cuota</acronym></A></td-->
				</tr>
			<%
			rsDet.movenext
			loop
		end if
		%>

	<%end if%>
	</tbody>
	<thead>
		<tr bgcolor="#<%=session("COLTABBG2")%>" class="totales">
			<td>&nbsp</td>
			<td>&nbsp</td>
			<td ALIGN="RIGHT"><%=FN(intTotalTotalCuota,0)%></td>
			<td>&nbsp</td>
			<td>&nbsp</td>
			<td>&nbsp</td>
			<td>&nbsp</td>
		</tr>
	</thead>
	</table>
	</td>
   </tr>
  </table>

</form>
</body>
</html>
