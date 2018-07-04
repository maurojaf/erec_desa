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
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link rel="stylesheet" href="../css/style_generales_sistema.css">
	
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	strRutDeudor = request("rut")
	
	'response.write " strRut = " & strRutDeudor	
	
	if trim(strRutDeudor) = "" Then
		strRutDeudor = session("session_RUT_DEUDOR") 
	End if

	session("session_RUT_DEUDOR") = strRutDeudor	

	strGraba = request("strGraba")
	strModificaObservacion = request("strModificaObservacion")
	strTipo = request("strTipo")

	strConfirmarTarea = request("strConfirmarTarea")
	strDesConfirmarTarea = request("strDesConfirmarTarea")
	strCodCliente = session("ses_codcli")
	
	'response.write " strModificaObservacion = " & strModificaObservacion	

	usuario=session("session_idusuario")

	AbrirSCG()

	If Trim(request("strGraba")) = "SI" Then

		strSql = "UPDATE DEUDOR SET OBSERVACIONES_BACKOFFICE = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "', FECHA_BACKOFFICE = GETDATE(), USUARIO_BACKOFFICE = '" & session("session_login") & "', CONFIRMACION_BACKOFFICE = 0"
		strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)

		Response.Redirect "principal.asp?strRut=" & strRutDeudor
		%>
		<SCRIPT>
			IrAPrincipal();
		</SCRIPT>
		<%
	End If
	
	If Trim(strModificaObservacion) = "SI" Then

		strSql = "UPDATE DEUDOR SET OBSERVACIONES_BACKOFFICE = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "'"
		strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)

	End If	

	If Trim(request("strLimpiar")) = "SI" Then

		strSql = "UPDATE DEUDOR SET OBSERVACIONES_BACKOFFICE = ''"
		'strSql = strSql & " ,HORA_AGEND_BO = NULL, FECHA_AGEND_BO = NULL" 
		strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)
		%>

		<SCRIPT>
			IrAPrincipal();
		</SCRIPT>
			<%
	End If


	If Trim(request("strConfirmarTarea")) = "SI" Then

		strSql = "UPDATE DEUDOR SET CONFIRMACION_BACKOFFICE = 1"
		strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)
		%>

		<SCRIPT>
			IrAPrincipal();
		</SCRIPT>
			<%
	End If

	If Trim(request("strDesConfirmarTarea")) = "SI" Then

		strSql = "UPDATE DEUDOR SET CONFIRMACION_BACKOFFICE = 0"
		strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)
		%>

		<SCRIPT>
			IrAPrincipal()
		</SCRIPT>
		<%
	End If

	If Trim(request("strAgendar")) = "SI" Then

		dtmFecAgend = Request("TX_FEC_AGEND")
		strHoraAgend = Request("TX_HORA_AGEND")
		If trim(strHoraAgend)="" Then strHoraAgend = "08:00"

		'Response.write "dtmFecAgend=" & dtmFecAgend

		If dtmFecAgend <> "" Then
			dtmFecAgend = dtmFecAgend & " " &  strHoraAgend & ":00"
		End If

		If dtmFecAgend = "" Then
			strSql = "UPDATE DEUDOR SET OBSERVACIONES_BACKOFFICE = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "', FECHA_AGEND_BO = NULL, HORA_AGEND_BO = NULL"
			strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		Else
			strSql = "UPDATE DEUDOR SET OBSERVACIONES_BACKOFFICE = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "', FECHA_AGEND_BO = '" & dtmFecAgend & "', HORA_AGEND_BO = '" & strHoraAgend & "'"
			strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		End If
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)
		%>

		<SCRIPT>
			IrAPrincipal()
		</SCRIPT>
		<%
	End If



%>
<title>Empresa</title>
<style type="text/css">
<!--
.Estilo13 {color: #FFFFFF}
.Estilo27 {color: #FFFFFF}
.Estilo1 {
	color: #FF0000;
	font-weight: bold;
	font-family: Arial, Helvetica, sans-serif; }
-->
</style>

<script language="JavaScript" src="../javascripts/cal2.js"></script>
<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
<script language="JavaScript" src="../javascripts/validaciones.js"></script>

<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

</head>
<body>
<form name="datos" method="post">
<INPUT TYPE="hidden" NAME="intOrigen" value="<%=intOrigen%>">

<%

	strSql = "SELECT CONVERT(VARCHAR(10),FECHA_AGEND_BO,103) as FECHA_AGEND_BO, HORA_AGEND_BO, IsNull(CONFIRMACION_BACKOFFICE,1) as CONFIRMACION_BACKOFFICE, OBSERVACIONES_BACKOFFICE, FECHA_BACKOFFICE, USUARIO_BACKOFFICE FROM DEUDOR WHERE COD_CLIENTE='" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
	
	'response.write "strSql = " & strSql
	
	set rsGestion=Conn.execute(strSql)
	if Not rsGestion.eof Then
		strObsNueva = Trim(rsGestion("OBSERVACIONES_BACKOFFICE"))
		strFechaObs = Trim(rsGestion("FECHA_BACKOFFICE"))
		strUsuarioObs = Trim(rsGestion("USUARIO_BACKOFFICE"))
		dtmFechaAgend = Trim(rsGestion("FECHA_AGEND_BO"))
		dtmHoraAgend = Trim(rsGestion("HORA_AGEND_BO"))

		If Trim(rsGestion("CONFIRMACION_BACKOFFICE")) = "1" Then
			strTareaConfirmada = "SI"
		Else
			strTareaConfirmada = "NO"
		End If


		If Trim(strFechaObs) <> "" and Trim(strUsuarioObs) <> "" then
			strTexto = "Fecha : " & strFechaObs & " , Usuario : " & strUsuarioObs & " , Solicitud Cerrada : " & strTareaConfirmada
		End If
	End If


%>

<% If Trim(strRutDeudor) <> "" then %>
	<div class="titulo_informe">SOLICITUD BACKOFFICE</div>
	<br>

    <tr>
    <td>

	  <%

	If strRutDeudor <> "" then
		strNombreDeudor = TraeNombreDeudor(Conn,strRutDeudor)
	Else
		strNombreDeudor=""
	End if

	%>

	<table width="90%" border="0" bordercolor="#FFFFFF"  class="estilo_columnas" align = "center">
		<thead>
		<tr >
			<td>RUT DEUDOR</td>
			<td>NOMBRE O RAZON SOCIAL</td>
			<td>USUARIO</td>
			<td>FECHA</td>
		</tr>
		</thead>
	      <tr bgcolor="#FFFFFF" class="Estilo8">

			<td>
				<A HREF="principal.asp?TX_RUT=<%=strRutDeudor%>">
				<acronym title="Llevar a pantalla de selección"><%=strRutDeudor%></acronym>
				</A>
			</td>
								
			<td><%=strNombreDeudor%><INPUT TYPE="hidden" NAME="rut" value="<%=strRutDeudor%>"> </td>

			<td ALIGN="Left"><%=session("nombre_user")	%></td>

	        <td><%=DATE%></td>
	      </tr>
    </table>
	
	<td>
	<tr>
	
	<table width="90%" border="0" ALIGN="CENTER">
	  <tr>
	    <td valign="top">

			<table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">
			<thead>
			<tr >
				<td>OBSERVACIONES (Max. 300 Caract.)</td>
			</tr>
			</thead>
			<tr>
			<td align="left">
			<TEXTAREA NAME="TX_OBSERVACIONES" ROWS=4 COLS=65><%=strObsNueva%></TEXTAREA>
			</td>
			</tr>
			<tr>
			<td align="LEFT">
			<%=strTexto%>
			</td>
			</tr>

		 </table>

	    </td>
	  </tr>

		<tr>
			<TD>
				<table width="100%" border="0" bordercolor="#FFFFFF">
					<tr bordercolor="#999999" class="Estilo8">
					<td align="LEFT">
					
						<% If strTareaConfirmada = "NO" Then %>
						<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Modificar Observación" value="Modificar Observación" onClick="envia('<%=strTareaConfirmada%>');" class="Estilo8">
						<%else%>
						<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Enviar Solicitud" value="Enviar Solicitud" onClick="envia('<%=strTareaConfirmada%>');" class="Estilo8">
						<%end if%>


						
						<% If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then %>
							<INPUT TYPE="BUTTON" class="fondo_boton_100"  NAME="Cerrar Solicitud" value="Cerrar Solicitud" onClick="ConfirmarTarea();" class="Estilo8">
							<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Reabrir Solicitud" value="Reabrir Solicitud" onClick="DesConfirmarTarea();" class="Estilo8">
						<%end if%>

					</td>

					<td align="right">
					
							<INPUT TYPE="BUTTON" NAME="Ver Gestiones" class="fondo_boton_100" value="Ver Gestiones" onClick="ir_detalle_gestiones();" class="Estilo8">
							<INPUT TYPE="BUTTON" NAME="Limpiar" class="fondo_boton_100" value="Limpiar" onClick="LimpiarDatos();" class="Estilo8">
							<INPUT TYPE="BUTTON" NAME="Volver" class="fondo_boton_100" value="Volver" onClick="history.back();" class="Estilo8">

					</td>
					
					</tr>
				</table>
			</TD>
		</tr>


		<tr>
			<TD ALIGN="CENTER">
				<table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">
					<thead>
					<tr>
						<td>FECHA AGENDAMIENTO</td>
						<td colspan = "2" align="left">HORA AGEND.</td>
					</tr>
				</thead>
					<tr>
						 <td width = "200" >
							<input name="TX_FEC_AGEND" type="text" id="TX_FEC_AGEND" value="<%=dtmFechaAgend%>" size="10" maxlength="10">
	
						 </td>
						 <td width="100" align="left">
							<input name="TX_HORA_AGEND" type="text" id="TX_HORA_AGEND" value="<%=dtmHoraAgend%>" size="5" maxlength="5" onChange="return ValidaHora(this,this.value)">
						</td>
					<% If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then %>
						<td align="left" >
							<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="BT_AGENDAR" value="Agendar" onClick="Agendar();" class="Estilo8">
						</td>
					<% Else%>
						<td align="left" >&nbsp;</td>						
					<% End If %>
					</tr>
				</table>
			</TD>
		</tr>

	</table>
<% End If %>


<INPUT TYPE="hidden" NAME="strLimpiar" value="">
<INPUT TYPE="hidden" NAME="strModificaObservacion" value="">
<INPUT TYPE="hidden" NAME="strGraba" value="">
<INPUT TYPE="hidden" NAME="strTipo" value="">
<INPUT TYPE="hidden" NAME="strDesConfirmarTarea" value="">
<INPUT TYPE="hidden" NAME="strConfirmarTarea" value="">
<INPUT TYPE="hidden" NAME="strAgendar" value="">


</form>
</body>
</html>


<script language="JavaScript" type="text/JavaScript">
	$(document).ready(function(){

		$('#TX_FEC_AGEND').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

	})

	function envia() {

		if( '<%=strTareaConfirmada%>' == 'NO' )
		{
			alert("La solicitud ya se encuentra cursada por lo que solo se modifico la observación, si requiere cerrar la solicitud favor comuníquese con Backoffice.");
			datos.strModificaObservacion.value='SI';
			datos.submit();
		}
		else if (confirm("¿ Está seguro que desea solicitar ayuda a backoffice ? "))
		{
			datos.strGraba.value='SI';
			datos.submit();
			
		}
	}
	
	function LimpiarDatos() {

		if (confirm("¿ Está seguro de Limpiar la observación de la solicitud de ayuda a backoffice ? "))
		{
				datos.strLimpiar.value='SI';
				datos.submit();
		}

	}

function ir_detalle_gestiones(){

	datos.action="detalle_gestiones.asp?rut=<%=strRutDeudor%>&cliente=<%=strCodCliente%>";
	datos.submit();
}

	function ConfirmarTarea() {

		if (confirm("¿ Está seguro de cerrar la solicitud en curso ? "))
		{
				datos.strConfirmarTarea.value='SI';
				datos.submit();
		}

	}

	function Agendar() {

		if (confirm("¿ Está seguro de agendar ? "))
		{
				datos.strAgendar.value='SI';
				datos.submit();
		}

	}

	function DesConfirmarTarea() {

		if (confirm("¿ Está seguro que desea reabrir la solicitud cerrada ? "))
		{
				datos.strDesConfirmarTarea.value='SI';
				datos.submit();
		}

	}

	function IrAPrincipal() {
		datos.action='principal.asp?strRut=<%=strRutDeudor%>';
		datos.submit();
	}

	function ValidaHora( ObjIng, strHora )
	{
	        var er_fh = /^(00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23)\:([0-5]0|[0-5][1-9])$/
	        if( strHora == "" )
	        {
	                alert("Introduzca la hora.")
	                return false
	        }
	        if ( !(er_fh.test( strHora )) )
	        {
	                alert("El dato en el campo hora no es válido.");
	                ObjIng.value = '';
	                ObjIng.focus();
	                return false
	        }

	        //alert("¡Campo de hora correcto!")
	        return true
	}


</script>


















