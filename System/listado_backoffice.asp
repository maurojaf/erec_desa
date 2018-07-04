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

<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/style_generales_sistema.css" rel="stylesheet">

<%

	Response.CodePage=65001
	Response.charset ="utf-8"

	AbrirSCG()
	
	strCodCliente = session("ses_codcli")	
	strBuscar =  request("strBuscar")

	strTipoGestion = request("cmb_tipogestion")	
	intGestionado = request("cmb_estadoGestion")
	inicio = request("inicio")	
	termino = request("termino")
	strRutDeudor = request("txtRutDeudor")
	intIdUsuario = request("cmb_usuario")
		  
	'response.write " intIdUsuario = " & intIdUsuario	
	
	if intGestionado = "" then intGestionado = "0" end if
	
	if intIdUsuario = "" then intIdUsuario = null end if
	
	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
	
	intIdUsuario = session("session_idusuario")
	
	end if
	'--Calcula Objetos relacionados al tipo de cobranza (Interna, Externa), CB_COBRANZA Y CB_EJECUTIVO--'

	strCobranza = Request("CB_COBRANZA")

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

AbrirSCG()
	
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
<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script language="JavaScript " type="text/JavaScript">
$(document).ready(function(){

	$("#table_tablesorter").tablesorter({dateFormat: "uk"}); 
		
	/*$("#table_tablesorter").tablesorter();*/
	
	$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'});
	$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'});
	$(document).tooltip();
	$.prettyLoader();
	
	var cambiaestado = $('#TXT_CAMBIA').val();
	if (cambiaestado == '') {
		envia();
	}
});

function envia()
{
	$.prettyLoader.show();
	datos.SubmitButton.disabled = true;
	document.datos.action = "listado_backoffice.asp?strBuscar=S";
	document.datos.submit();
}

function limpiar() {
        self.location.href = 'listado_backoffice.asp'
}

function enviaorigen(destino,rutdeudor,idGestion) {

	if(destino =='1'){
	
	datos.action = "confirmar_cp.asp?id_gestion=" + idGestion + "&Rut=" + rutdeudor;
	datos.submit();
	
	}
	else if (destino =='2'){
	
	datos.action="enviar_backoffice.asp?intOrigen=IP&Rut=" + rutdeudor;
	datos.submit();
	
	}
	else if (destino =='3'){
	
	datos.action="detalle_gestiones.asp?rut=" + rutdeudor + "&cliente=<%=strCodCliente%>";
	datos.submit();
	}
	else if (destino =='4'){
	
	datos.action="auditar_doc.asp?rut=" + rutdeudor;
	datos.submit();
	
	}
}

</script>

</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<form name="datos" method="post">
<div class="titulo_informe">LISTADO BACKOFFICE</div>
<br>
<table width="90%"  border="0" align="center">
  <tr>
    <td>
	<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
		<thead>

	      <tr height="20">
			<td>TIPO COBRANZA</td>
			<td>TIPO GESTION</td>
			<td>GESTIONADO</td>
			<td>FECHA SOLICITUD DESDE</td>
			<td>FECHA SOLICITUD HASTA</td>
			<td>RUT DEUDOR</td>
	      <% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
			<td>EJECUTIVO ASIGNADO</td>
		  <% End If %>
			<td colspan = "2">&nbsp;</td>
		
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
				<SELECT NAME="cmb_tipogestion" id="cmb_tipogestion">
					<option value= "Null" <%If Trim(strTipoGestion)= "Null" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>COMPROMISO NO CONFIRMADO</option>
					<option value="2" <%If Trim(strTipoGestion)="2" Then Response.write "SELECTED"%>>SOLICITUD BACKOFFICE</option>
					<option value="3" <%If Trim(strTipoGestion)="3" Then Response.write "SELECTED"%>>SOLICITUD A CLIENTE</option>
					<option value="4" <%If Trim(strTipoGestion)="4" Then Response.write "SELECTED"%>>NOTIFICACION O FACTURA NO RECEPCIONADA</option>
				</SELECT>
			</td>		  

			<td>
				<SELECT NAME="cmb_estadoGestion" id="cmb_estadoGestion" style=" width: 75px">
					<option value="null" <%If Trim(intGestionado)="null" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(intGestionado)="1" Then Response.write "SELECTED"%>>SI</option>
					<option value="0" <%If Trim(intGestionado)="0" Then Response.write "SELECTED"%>>NO</option>
				</SELECT>
			</td>	
			
			<td><input name="inicio" type="text" readonly="true" id="inicio" value="<%=inicio%>" size="10" maxlength="10"></td>

			<td><input name="termino" type="text" readonly="true" id="termino" value="<%=termino%>" size="10" maxlength="10"></td>
			
			<td>
				<input name="txtRutDeudor" type="text" id="txtRutDeudor" <%If strRutDeudor = "" then %> value="" <% Else %> value="<%=strRutDeudor%>" <%end if%> size="12" maxlength="11">
			</td>

		  <% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
			<td>
				<SELECT NAME="cmb_usuario" id="cmb_usuario" style=" width: 125px">
					<option value="Null">TODOS</option>
					<%
					strSql=" SELECT U.ID_USUARIO, LOGIN = UPPER(U.LOGIN)" 
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO"
					strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1 AND PERFIL_EMP <> 1 AND UC.COD_CLIENTE = '" & strCodCliente & "'"
					set rsUsu=Conn.execute(strSql)
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
			
			<td align="right"><input type="Button" class="fondo_boton_100" name="LimpiarButton" value="Limpiar" onClick="limpiar();"></td>
			
			<td align="right"><input type="Button" class="fondo_boton_100" name="SubmitButton" value="Ver" onClick="envia();"></td>
			
	      </tr>
    </table>

	<input type="hidden" id="TXT_CAMBIA" value='<%=strBuscar%>'/>

	<table width="100%" border="0" Id="table_tablesorter" class="tablesorter" style="width:100%;">
		
		<%if strBuscar = "S" then
		
		strSql = " EXEC uspInformeModuloBackOfficeSelect '" & strCodCliente & "', '" & strCobranza & "', " & strTipoGestion & ", " & intGestionado & ", '" & inicio & "', '" & termino & "','" & strRutDeudor & "', " & intIdUsuario 
		
		'response.write " strSql = " & strSql

			set rsDet=Conn.execute(strSql)

					if not rsDet.eof then%>
					
						<thead>
						<tr >
							<td>&nbsp;</td>
							<th>TIPO GESTION</th>
							<th>SOLICITUD</th>
							<th>AGENDAMIENTO</th>
							<th>RUT DEUDOR</th>
							<th>NOMBRE DEUDOR</th>
							<th>SALDO</th>
							<th>TOTAL DOC.</th>
							<th>DIA MORA</th>
							<th>EJECUTIVO ASIG.</th>
							<td width="30" align="center">OBS</td>
						</tr>
						</thead>
					<tbody>
			<%
					do while not rsDet.eof
					
						strFechaCompromiso = ""
						intReg = intReg + 1

						TotalCasos = TotalCasos + 1
						intIdTipoGestion= rsDet("ID_TIPO_GESTION")
						strRutDeudor = rsDet("RUT_DEUDOR")
						intIdGestion = rsDet("ID_GESTION")
						FechaCompBase = rsDet("MIN_FECHA_COMPROMISO")
						strObservacion = rsDet("OBS_TIPO_GESTION")
						strHoraSolicitud = "Hora Solicitud: " & rsDet("MIN_HORA_SOLICITUD")
						
						If FechaCompBase <> "" then
						strFechaCompromiso = "Fecha Compromiso: " &FechaCompBase
						End If

						%>
						<tr >
							<td><%=intReg%></td>
							
							<td ALIGN="LEFT" title="<%=strFechaCompromiso%>">
							<%=rsDet("DESCRIPT_TIPO_GESTION")%>
							
							<td ALIGN="LEFT" title="<%=strHoraSolicitud%>">
							<%=rsDet("MIN_FECHA_SOLICITUD")%>
							
							<td><%=rsDet("MIN_FECHA_AGEND")%></td>

							<td>
								<A HREF="#" onClick="enviaorigen('<%=intIdTipoGestion%>','<%=strRutDeudor%>','<%=intIdGestion%>');">
								<acronym title="Llevar a pantalla de origen"><%=strRutDeudor%></acronym>
								</A>
							</td>
							
							<td ALIGN="LEFT" title="<%=rsDet("NOMBRE_DEUDOR")%>">
							<%=mid(rsDet("NOMBRE_DEUDOR"),1,25)%>
							
							<td><%=FN(rsDet("SALDO"),0)%></td>
							<td align="center"><%=FN(rsDet("TOTAL_DOC"),0)%></td>
							<td align="center"><%=FN(rsDet("DIA_MORA"),0)%></td>
							<td><%=rsDet("COBRADOR_ASIGNADO")%></td>
							
							<td width="30" align="center" title="<%=strObservacion%>">
								<img src="../imagenes/priorizar_normal.png" border="0">
							</td>							
							
						</tr>
						<%
						Response.flush()
						rsDet.movenext
					loop%>
					
					</tbody>
					
					<thead>
						<tr >
							<td colspan = "11">&nbsp;</td>
						</tr>
					</thead>
					
					<%Else%>

					<thead>
						<tr >
							<td colspan = "10">&nbsp;</td>
						</tr>
					</thead>
					
					<tr class="estilo_columnas">
						<td ALIGN="CENTER" Colspan = "9">NO EXISTEN RESULTADOS SEGUN PARAMETROS DE BUSQUEDA</td>
					</tr>
					
					<thead>
						<tr >
							<td colspan = "10">&nbsp;</td>
						</tr>
					</thead>
					
					<%end if						

		end if
		
CerrarScg()

%>
	
	</table>
	</td>
	
   </tr>
 </table>

</form>

</body>
</html>

