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

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	intCodUsuario=session("session_idusuario")

	AbrirSCG()

	sucursal = request("cmb_sucursal")
	intTipoPago = request("CB_TIPOPAGO")

	intFormaPago = request("CB_FORMA_PAGO")

	intUsuario = request("CB_USUARIO")


	intDestino = request("CB_DESTINO")
	'response.write(perfil)
	intCodPago = request("TX_PAGO")
	strRut=request("TX_RUT")
	if sucursal="" then sucursal="0"
	'response.write(sucursal)
	usuario = request("cmb_usuario")
	if usuario = "" then usuario = "0"
	termino = request("TERMINO")
	inicio = request("INICIO")
	resp = request("resp")
	if Trim(inicio) = "" Then
		inicio = TraeFechaMesActual(Conn,-1)
		inicio = "01" & Mid(inicio,3,10)
		inicio = TraeFechaActual(Conn)
	End If
	if Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If
	CLIENTE = REQUEST("CLIENTE")
	If Trim(CLIENTE) = "" Then
		CLIENTE=session("ses_codcli")
	End if
	'hoy=date

	'response.write(hoy)
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


	<script language="JavaScript " type="text/JavaScript">

	$(document).ready(function(){
			$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
			$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	})
	function Refrescar()
	{
		resp='no'
		datos.action = "detalle_cuadratura.asp?resp="+ resp +"";
		datos.submit();
	}

	function Ingresa()
	{
		with( document.datos )
		{
			action = "detalle_cuadratura.asp";
			submit();
		}
	}

	function Modificar(cod_pago)
	{
		with( document.datos )
		{
			action = "modif_caja_web2.asp?strOrigen=detalle_caja.asp&cod_pago=" + cod_pago;
			submit();
		}
	}


	function envia()
	{
		resp='si'
		datos.action = "detalle_cuadratura.asp?resp="+ resp +"";
		datos.submit();
	}

	function envia_excel(URL){

	window.open(URL,"INFORMACION","width=200, height=200, scrollbars=yes, menubar=yes, location=yes, resizable=yes")
	}
	</script>



</head>
<body>
<form name="datos" method="post">
	<div class="titulo_informe">LISTADO CUADRATURA DE CAJA
		
	</div>
<table width="90%" height="500" border="0" align="center">
  <tr>
    <td valign="top">
	<table width="100%" border="0" bordercolor="#999999">
		<tr height="20" class="Estilo8">
	        <td>RUT: </td>
			<td><INPUT TYPE="text" NAME="TX_RUT" value="<%=strRut%>" onchange=""></td>
		    <td>FORMA PAGO: </td>
			  <td>
				<select name="CB_FORMA_PAGO">
				<option value="">TODOS</option>
					<%
					strSql="SELECT * FROM CAJA_FORMA_PAGO"
					set rsFPago=Conn.execute(strSql)
					if Not rsFPago.eof then
						do until rsFPago.eof
						%>
						<option value="<%=rsFPago("ID_FORMA_PAGO")%>"
						<%if Trim(intFormaPago)=Trim(rsFPago("ID_FORMA_PAGO")) then Response.Write("SELECTED") end if%> WIDTH="10"
						><%=MID(rsFPago("DESC_FORMA_PAGO"),1,19)%></option>
						<%rsFPago.movenext
						loop
					end if
					rsFPago.close
					set rsFPago=nothing
					%>
				</select>
			</td>
			<td>EJECUTIVO: </td>
			<td>
			<select name="CB_USUARIO">
			<%If Trim(TraeSiNo(session("perfil_sup"))="Si") or Trim(TraeSiNo(session("perfil_emp"))="Si") Then%>
				<option value="">TODOS</option>
			<%End If%>
			<%
				strSql="SELECT * FROM USUARIO WHERE PERFIL_CAJA = 1 AND ACTIVO = 1"
				If Trim(TraeSiNo(session("perfil_sup"))="No") Then
					strSql=strSql & " AND ID_USUARIO = " & intCodUsuario
				End if
				set rsUsuario=Conn.execute(strSql)
				if Not rsUsuario.eof then
					do until rsUsuario.eof
					%>
					<option value="<%=rsUsuario("ID_USUARIO")%>"
					<%if Trim(intUsuario)=Trim(rsUsuario("ID_USUARIO")) then Response.Write("SELECTED") end if%> WIDTH="10"
					><%=MID(rsUsuario("LOGIN"),1,19)%></option>
					<%rsUsuario.movenext
					loop
				end if
				rsUsuario.close
				set rsUsuario=nothing
				%>
			</select>
			</td>
		</tr>

	</table>
	<table width="100%" border="0" class="estilo_columnas">
		<thead>
	      <tr height="20">
		  <td>CLIENTE</td>
		  <%if perfil="caja_modif" or perfil = "caja_listado" then%>
	        <td>SUCURSAL</td>
			<%end if%>
			<td>TIPO DE PAGO</td>
			<td>DESTINO</td>
			<td>DESDE</td>
			<td>HASTA</td>
	      </tr>
	    </thead>
		  <tr bordercolor="#999999" class="Estilo8">
		  <td>

		<select name="CLIENTE" width="15">
		<% If TraeSiNo(session("perfil_emp")) <> "Si" Then %>
			<option value="0">SELECCIONAR</option>
		<% End If%>
		<%
		ssql="SELECT * FROM CLIENTE WHERE ACTIVO=1 AND COD_CLIENTE <> '999'"
		If TraeSiNo(session("perfil_emp")) = "Si" Then
			ssql = ssql & "AND COD_CLIENTE = " & session("ses_codcli")
		End If
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

        <td>
			<select name="CB_TIPOPAGO">
			<option value="">TODOS</option>
				<%
				ssql="SELECT * FROM CAJA_TIPO_PAGO	"
				If Trim(intTipoPago)="CO" Then
					'strTipoCompArch = "comp_pago_convenio.asp"
					'ssql = ssql & " WHERE ID_TIPO_PAGO = 'CO'"
				ElseIf Trim(intTipoPago)="RP" Then
					'strTipoCompArch = "comp_pago_repactacion.asp"
					'ssql = ssql & " WHERE ID_TIPO_PAGO = 'RP'"
				Else
					'strTipoCompArch = "comp_pago.asp"
					'ssql = ssql & " WHERE ID_TIPO_PAGO NOT IN ('CO','RP')"
				%>
					<option value="">SELECCIONAR</option>
				<%
				End If
				set rsCLI=Conn.execute(ssql)
				if not rsCLI.eof then
					do until rsCLI.eof
					%>
					<option value="<%=rsCLI("ID_TIPO_PAGO")%>"
					<%if Trim(intTipoPago)=Trim(rsCLI("ID_TIPO_PAGO")) then Response.Write("SELECTED") end if%> WIDTH="10"
					><%=MID(rsCLI("DESC_TIPO_PAGO"),1,19)%></option>
					<%rsCLI.movenext
					loop
				end if
				rsCLI.close
				set rsCLI=nothing
				%>
			</select>
		</td>
		<td>
			<SELECT NAME="CB_DESTINO">
				<option value="" <%if Trim(intDestino)="" then Response.Write("SELECTED") end if%>>TODOS</option>
				<option value="0" <%if Trim(intDestino)="0" then Response.Write("SELECTED") end if%>>CLIENTE</option>
				<option value="1" <%if Trim(intDestino)="1" then Response.Write("SELECTED") end if%>>EMPRESA</option>
			</SELECT>
		</td>
			<td>
				<input name="inicio" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
		
			</td>
			<td>
				<input name="termino" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
          	
			<input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();">
			<input type="button" class="fondo_boton_100" name="Submit" value="Excel" onClick="javascript:envia_excel('detalle_cuadratura_excel.asp?strRut=<%=strRut%>&intCodPago=<%=intCodPago%>&sucursal=<%=sucursal%>&termino=<%=termino%>&inicio=<%=inicio%>&usuario=<%=usuario%>&CLIENTE=<%=CLIENTE%>')">
			</td>
	      </tr>

    </table>
	<table width="100%" class="intercalado" style="width:100%;">
		<thead>
		<tr>
			<td>FECHA</td>
			<td>USUARIO</td>
			<td>NRO.COMP.</td>
			<td>BOLETA</td>
			<td>CLIENTE</td>
			<td>RUT</td>
			<td>TIPO PAGO</td>
			<td>ASEG</td>
			<td>ETAPA</td>
			<td>FORMA PAGO</td>
			<td>MONTO</td>
			<td>DESTINO</td>
			<td>ID.PAGO</td>
			<td>&nbsp;</td>
		</tr>
		</thead>
		<tbody>
	<%

		strSql = "SELECT U.LOGIN, CONVERT(VARCHAR(10),CWC.FECHA_PAGO,103) AS FECHA_PAGO,CWC.TIPO_PAGO AS CONCEPTO, CWC.COD_CLIENTE, CWDP.ID_PAGO,CWC.RUT_DEUDOR, COMP_INGRESO,NRO_BOLETA,CAST(SUM(MONTO) AS INT) AS MONTO, CWDP.FORMA_PAGO,CWDP.TIPO_PAGO "
		strSql = strSql & " FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, USUARIO U "
		strSql = strSql & " WHERE CWC.USR_INGRESO = U.ID_USUARIO AND CWC.ID_PAGO = CWDP.ID_PAGO AND CWDP.FORMA_PAGO IN ('EF','DP','CF','CD','CU','OC') "
		strSql = strSql & " AND CWDP.TIPO_PAGO IN (0,1) "
		strSql = strSql & " AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(DAY,FECHA_PAGO,'" & TERMINO & "')>=0 "

		IF CLIENTE <> "0" THEN
			strSql = strSql & "and CWC.COD_CLIENTE = '" & CLIENTE & "'"
		END IF

		If Trim(TraeSiNo(session("perfil_sup"))="No") and Trim(TraeSiNo(session("perfil_emp"))="No") Then
			strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
		End if
		If Trim(TraeSiNo(session("perfil_sup"))="Si") and intUsuario <> "" Then
			strSql = strSql & "AND  CWC.USR_INGRESO= " & intUsuario
		END IF
		If Trim(strRut) <> "" Then
			strSql = strSql & " AND CWC.RUT_DEUDOR = '" & strRut & "'"
		End If
		If intTipoPago <> "" THEN
			strSql = strSql & " AND CWC.TIPO_PAGO = '" & intTipoPago & "' "
		End if

		If intFormaPago <> "" THEN
			strSql = strSql & " AND CWDP.FORMA_PAGO = '" & intFormaPago & "' "
		End if


		IF intCodPago <> "" THEN
			'strSql = strSql & " AND ID_PAGO = " & intCodPago & " "
		END IF

		IF intDestino <> "" THEN
			strSql = strSql & " AND CWDP.TIPO_PAGO = " & intDestino
		END IF



		strSql = strSql & " GROUP BY U.LOGIN,FECHA_PAGO,CWC.TIPO_PAGO,CWC.COD_CLIENTE,CWDP.ID_PAGO,COMP_INGRESO,NRO_BOLETA,CWDP.FORMA_PAGO,CWDP.TIPO_PAGO,CWC.RUT_DEUDOR"
		strSql = strSql & " ORDER BY COMP_INGRESO,CWC.RUT_DEUDOR "

		'Response.write "strSql = " & strSql


		'Response.write "strSql = " & strSql
		'Response.End
	If strSql <> "" then
		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			Do while not rsDet.eof
				ssql="SELECT * FROM CAJA_FORMA_PAGO WHERE ID_FORMA_PAGO = '" & rsDet("FORMA_PAGO") & "'"
				set rsFormaPago=Conn.execute(ssql)
				If not rsFormaPago.eof then
					strFormaPago= rsFormaPago("DESC_FORMA_PAGO")
				Else
					strFormaPago= ""
				End if
				strTipoPago = rsDet("CONCEPTO")
				If Trim(rsDet("TIPO_PAGO")) = "0" Then strDestino = "CLIENTE"
				If Trim(rsDet("TIPO_PAGO")) = "1" Then strDestino = "EMPRESA"

				If Trim(strTipoPago)="CO" Then
					strTipoCompArch = "comp_pago_convenio.asp"
				ElseIf Trim(intTipoPago)="RP" Then
					strTipoCompArch = "comp_pago_repactacion.asp"
				Else
					strTipoCompArch = "comp_pago.asp"
				End If




				strSql = "SELECT ID_TIPO_ASEG FROM CONVENIO_DET WHERE ID_PAGO IN (SELECT ID_PAGO FROM CAJA_WEB_EMP WHERE COMP_INGRESO = " & rsDet("COMP_INGRESO") & ")"
				set rsAseg=Conn.execute(strSql)
				If not rsAseg.eof then
					intAseg = rsAseg("ID_TIPO_ASEG")
				Else
					intAseg = ""
				End If
				strLogin = rsDet("LOGIN")


				strSql = "SELECT ETAPA_COBRANZA FROM DEUDOR WHERE COD_CLIENTE = '" & rsDet("COD_CLIENTE") & "' AND RUT_DEUDOR = '" & rsDet("RUT_DEUDOR") & "'"
				set rsAseg=Conn.execute(strSql)
				If not rsAseg.eof then
					intEtapa = rsAseg("ETAPA_COBRANZA")
				Else
					intEtapa = ""
				End If


			%>
			<tr >
				<td><%=rsDet("FECHA_PAGO")%></td>
				<td><%=strLogin%></td>
				<td><A HREF="<%=strTipoCompArch%>?strImprime=S&intNroComp=<%=rsDet("COMP_INGRESO")%>"><%=rsDet("COMP_INGRESO")%></A></td>
				<td><%=rsDet("NRO_BOLETA")%></td>
				<td><%=rsDet("COD_CLIENTE")%></td>
				<td>
					<A HREF="principal.asp?rut=<%=rsDet("RUT_DEUDOR")%>"><acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym></A>
				</td>
				<td><%=strTipoPago%></td>
				<td><%=intAseg%></td>
				<td><%=intEtapa%></td>
				<td><%=strFormaPago%></td>
				<td ALIGN="right"><%=rsDet("MONTO")%></td>
				<td><%=strDestino%></td>
				<td><%=rsDet("ID_PAGO")%></td>
				<td>
				<% If Trim(TraeSiNo(session("perfil_emp")) <> "Si") Then %>
				<A HREF="#" onClick="Modificar(<%=rsDet("ID_PAGO")%>)";>Modificar</A>
				<% End If %>
				</td>
			</tr>
			<%
				intTotal = intTotal + rsDet("MONTO")
			rsDet.movenext
			loop
		end if
	End if%>
</tbody>
	<thead class="totales">

		<tr >
			<td colspan=5>TOTALES</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td ALIGN="right"><%=intTotal%></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
	</thead>
	</table>
	</td>
   </tr>
  </table>

</form>
</body>
</html>