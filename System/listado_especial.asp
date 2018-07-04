<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<link href="../css/style.css" rel="Stylesheet">
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

<%

	Response.CodePage=65001
	Response.charset ="utf-8"

	strEjeAsig = request("CB_EJECUTIVO")
	strTipoGestion = request("cmb_tipogestion")
	strEstadoCC = request("CMB_ESTADO_CC")

	termino = request("termino")
	inicio = request("inicio")
	
	if strTipoGestion = "" then strTipoGestion = "0"

	strCodCliente = session("ses_codcli")

	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_TipoGestionEspecial") = strTipoGestion
		session("Ftro_EjecAsigEspecial") = strEjeAsig
		session("FtroCB_ESTADO_CC") = strEstadoCC
	End If

	If Trim(Request("strBuscar")) = "N" Then
		session("Ftro_TipoGestionEspecial") = ""
		session("Ftro_EjecAsigEspecial") = ""
		session("FtroCB_ESTADO_CC") = ""
	End If

	If strEjeAsig = "0" Then strEjeAsig = session("Ftro_EjecAsigEspecial")
	If strTipoGestion = "0" Then strTipoGestion = session("Ftro_TipoGestionEspecial")
	If strEstadoCC = "" Then strEstadoCC = session("FtroCB_ESTADO_CC")

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

	$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$(document).tooltip();
	 

})

function envia()
{
	//datos.TX_RUT.value='';
	//datos.TX_PAGO.value='';
	resp='si'
	document.datos.action = "listado_especial.asp?strBuscar=S";
	document.datos.submit();
}

</script>

</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">


<form name="datos" method="post">
<div class="titulo_informe">MÓDULO CASOS COMPLEJOS</div>
<br>
<table width="90%" height="500" border="0" align="center">
  <tr>
    <td valign="top">
	<table width="100%" border="0" class="estilo_columnas">
		<thead>
	      <tr height="20" >

			<td>COBRANZA</td>
			<td>TIPO GESTIÓN</td>
			<td>ESTADO</td>
			<td>F. INGRESO GESTION DESDE</td>
			<td>F. INGRESO GESTION HASTA</td>

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
				<SELECT NAME="cmb_tipogestion" id="cmb_tipogestion">
					<option value="0" <%If Trim(strTipoGestion)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>DEUDOR SIN RECURSOS</option>
					<option value="2" <%If Trim(strTipoGestion)="2" Then Response.write "SELECTED"%>>DEUDOR COMPLEJO</option>
				</SELECT>
			</td>

			<td>
				<SELECT NAME="CMB_ESTADO_CC" id="CMB_ESTADO_CC">
					<option value="0" <%If Trim(strEstadoCC)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strEstadoCC)="1" Then Response.write "SELECTED"%>>AGENDADO</option>
					<option value="2" <%If Trim(strEstadoCC)="2" Then Response.write "SELECTED"%>>NO GESTIONADO</option>
				</SELECT>
			</td>

			<td><input name="inicio" readonly="true" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
				<!--<a href="javascript:showCal('Calendar7');"><img src="../imagenes/calendario.gif" border="0"></a>-->
			</td>

			<td><input name="termino" readonly="true" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
         		 <!--<a href="javascript:showCal('Calendar6');"><img src="../imagenes/calendario.gif" border="0"></a>-->
			</td>

		<% If sinCbUsario="0" Then %>
			<td>
				<select name="CB_EJECUTIVO" id="CB_EJECUTIVO">
				</select>
			</td>
		<% End If %>

			<td Align="center">
				<input type="Button" name="Submit" value="Ver" class="fondo_boton_100" onClick="envia();">
			</td>

	      </tr>
    </table>

	<table width="100%" border="0" class="intercalado" style="width:100%;">
		<thead>
		<tr >
			<td>&nbsp;</td>
			<td>PRIOR.</td>
			<td>GESTION</td>
			<td>FECHA GESTION</td>
			<td>RUT DEUDOR</td>
			<td>NOMBRE DEUDOR</td>
			<td>SALDO</td>
			<td>DOC</td>
			<td>DM</td>
			<td>EJEC.ASIG.</td>
			<td>OBS.</td>
		</tr>
		</thead>
		<tbody>
	<%

	abrirscg()
	
		strSql = "SELECT MIN(CUOTA.PRIORIDAD_CUOTA) AS PRIORIDAD,"

		strSql = strSql & "	(CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 8) AND '" & strTipoGestion & "' IN (0,1) "
		strSql = strSql & "	THEN 'DEUDOR SIN RECURSOS'"
		strSql = strSql & " WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 7) AND '" & strTipoGestion & "' IN (0,2)"
		strSql = strSql & "	THEN 'DEUDOR COMPLEJO' ELSE 'OTRO' END) as GESTION,"

		strSql = strSql & "	DM = MAX(DATEDIFF(DAY,FECHA_VENC,GETDATE())),"
		strSql = strSql & " RUT_DEUDOR = DEUDOR.RUT_DEUDOR,"
		strSql = strSql & " NOMBRE_DEUDOR = DEUDOR.NOMBRE_DEUDOR, "
		strSql = strSql & " FECHA_GESTION = CAST(CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103) AS DATETIME),"
		strSql = strSql & " SALDO = SUM (CUOTA.SALDO),"
		strSql = strSql & " TOTAL_DOC = COUNT (CUOTA.NRO_DOC),"
		strSql = strSql & " OBS = GESTIONES.OBSERVACIONES,"
		strSql = strSql & " USUARIO = USUARIO.LOGIN"


		strSql = strSql & " FROM DEUDOR INNER JOIN CUOTA 						ON DEUDOR.RUT_DEUDOR = CUOTA.RUT_DEUDOR AND DEUDOR.COD_CLIENTE = CUOTA.COD_CLIENTE"
		strSql = strSql & " 			LEFT JOIN USUARIO						ON DEUDOR.USUARIO_ASIG = USUARIO.ID_USUARIO"
		strSql = strSql & " 			INNER JOIN GESTIONES					ON CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION"
		strSql = strSql & " 			INNER JOIN GESTIONES_TIPO_GESTION		ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_GESTION.COD_CATEGORIA"
		strSql = strSql & " 													   AND GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
		strSql = strSql & " 													   AND GESTIONES.COD_GESTION = GESTIONES_TIPO_GESTION.COD_GESTION"
		strSql = strSql & " 													   AND GESTIONES.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"

		strSql = strSql & " WHERE (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 7 OR GESTIONES_TIPO_GESTION.GESTION_MODULOS = 8)"
		strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (SELECT ESTADO_DEUDA.CODIGO FROM ESTADO_DEUDA WHERE ESTADO_DEUDA.ACTIVO = 1)"
		strSql = strSql & " AND CUOTA.COD_CLIENTE = '" & strCodCliente & "'"
		strSql = strSql & " AND CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION"


		if Trim(strTipoGestion) = "1" Then

		strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 8"

		End If

		if Trim(strTipoGestion) = "2" Then

		strSql = strSql & " AND GESTIONES_TIPO_GESTION.GESTION_MODULOS = 7 "

		End If

		If Trim(strCobranza) = "INTERNA" Then
			strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
			strParametro = "1"
		End if

		If Trim(strCobranza) = "EXTERNA" Then
			strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
			strParametro = "1"
		End if

		if Trim(strEstadoCC) = "1" Then
			strSql = strSql & " AND (DATEDIFF(MINUTE,(ISNULL(CUOTA.FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(CUOTA.HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) < 0)"
		End If

		if Trim(strEstadoCC) = "2" Then
			strSql = strSql & " AND (DATEDIFF(MINUTE,(ISNULL(CUOTA.FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(CUOTA.HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"
		End If

		If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then

		strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"

		Else
			if Trim(strEjeAsig) <> "" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
			End if
		End if

		strSql = strSql & " GROUP BY DEUDOR.RUT_DEUDOR,DEUDOR.NOMBRE_DEUDOR,USUARIO.LOGIN,GESTIONES.FECHA_INGRESO,GESTIONES.OBSERVACIONES,GESTIONES_TIPO_GESTION.GESTION_MODULOS"
		
		strSql = strSql & " ORDER BY GESTION DESC, FECHA_GESTION ASC, PRIORIDAD, DEUDOR.RUT_DEUDOR, SALDO DESC"

		'Response.write "strSql = " & strSql
		'Response.End

		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			do while not rsDet.eof
				intReg = intReg + 1

				%>
				<tr >
					<td><%=intReg%></td>
					<td ALIGN="CENTER"><%=FN(rsDet("PRIORIDAD"),0)%></td>
					<td><%=rsDet("GESTION")%></td>
					<td><%=rsDet("FECHA_GESTION")%></td>
					<td>
						<A HREF="principal.asp?TX_RUT=<%=rsDet("RUT_DEUDOR")%>">
						<acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym>
						</A>
					</td>
					<td Align="LEFT" title="<%=rsDet("NOMBRE_DEUDOR")%>">					
						<%=Mid(rsDet("NOMBRE_DEUDOR"),1,30)%></td>
						
					<td ALIGN="RIGHT"><%=FN(rsDet("SALDO"),0)%></td>
					<td ALIGN="CENTER"><%=FN(rsDet("TOTAL_DOC"),0)%></td>
					<td ALIGN="CENTER"><%=FN(rsDet("DM"),0)%></td>
					<td><%=Mid(rsDet("USUARIO"),1,15)%></td>
					<td Align="center" title="<%=rsDet("OBS")%>">
						<img src="../imagenes/priorizar_normal.png" border="0">
					</td>
				<%
				rsDet.movenext
			loop
		Else
		
		%>
				<tr >
					<td class="estilo_columna_individual" HEIGHT = "20"  ALIGN="CENTER" Colspan = "11">NO EXISTEN RESULTADOS SEGUN PARAMETROS DE BUSQUEDA</td>
				</tr>
		<%
		rsDet.close
		set rsDet=nothing
		end if
		
cerrarscg()%>
	</tbody>
	<thead>
		<tr class="totales">
			<td Colspan = "11">&nbsp;</td>
		</tr>
	</thead>
	</table>
	</td>
   </tr>
  </table>

</form>


</body>
</html>


<script type="text/javascript">

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
