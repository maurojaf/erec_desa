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
<LINK href="../css/style.css" type="text/css" rel="stylesheet">
<link href="../css/style_generales_sistema.css" rel="stylesheet">

<%

	Response.CodePage=65001
	Response.charset ="utf-8"

	strEjeAsig = request("CB_EJECUTIVO")
	strTipoGestion = request("cmb_tipogestion")
	strEstadoCP = request("CMB_ESTADO_CP")

	termino = request("termino")
	inicio = request("inicio")
	strCobranza = Request("CB_COBRANZA")

	if strTipoGestion = "" then strTipoGestion = "0"
	if strEstadoCP = "" then strEstadoCP = "2"
	strCodCliente = session("ses_codcli")

	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_TipoGestionRyR") = strTipoGestion
		session("Ftro_EjecAsigRyR") = strEjeAsig
	End If

	If Trim(Request("strBuscar")) = "N" Then
		session("Ftro_TipoGestionRyR") = ""
		session("Ftro_EjecAsigRyR") = ""
	End If

	If strEjeAsig = "" Then strEjeAsig = session("Ftro_EjecAsigRyR")
	If strTipoGestion = "0" Then strTipoGestion = session("Ftro_TipoGestionRyR")


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
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script language="JavaScript " type="text/JavaScript">

$(document).ready(function(){

	$("#table_tablesorter").tablesorter({dateFormat: "uk"});

	$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
 	$(document).tooltip();

})
function envia()
{
	resp='si'
	document.datos.action = "listado_Revision_CP.asp?strBuscar=S";
	document.datos.submit();
}

</script>


</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<form name="datos" method="post">
<div class="titulo_informe">MÓDULO REVISIÓN DE COMPROMISOS</div>	
<br>
<table width="90%" height="500" border="0" align="center">
  <tr>
    <td valign="top">
	<table width="100%" border="0"class="estilo_columnas">
		<thead>
	      <tr height="20">

			<td>COBRANZA</td>
			<td>TIPO GESTIÓN</td>
			<td>ESTADO</td>
			<td>FECHA COMPROMISO DESDE</td>
			<td>FECHA COMPROMISO HASTA</td>

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
					<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>COMPROMISO</option>
					<option value="2" <%If Trim(strTipoGestion)="2" Then Response.write "SELECTED"%>>COMPROMISO PENDIENTE</option>
				</SELECT>
			</td>

			<td>
				<SELECT NAME="CMB_ESTADO_CP" id="CMB_ESTADO_CP">
					<option value="0" <%If Trim(strEstadoCP)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strEstadoCP)="1" Then Response.write "SELECTED"%>>AGENDADO</option>
					<option value="2" <%If Trim(strEstadoCP)="2" Then Response.write "SELECTED"%>>NO GESTIONADO</option>
				</SELECT>
			</td>

			<td><input name="inicio" readonly="true" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
			</td>

			<td><input name="termino" readonly="true" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
			</td>

		<% If sinCbUsario="0" Then %>
			<td>
				<select name="CB_EJECUTIVO" id="CB_EJECUTIVO">
				</select>
			</td>
		<% End If %>
			
			<td Align="center">
				<input type="Button" name="Submit" class="fondo_boton_100" value="Ver" onClick="envia();">
			</td>

	      </tr>
    </table>



	<table width="100%" border="0" id="table_tablesorter" class="tablesorter intercalado" style="width:100%;">
		<thead>
		<tr >
			<td>&nbsp;</td>
			<th>GESTION</th>
			<th>FECHA</th>
			<th>RUT DEUDOR</th>
			<th>NOMBRE DEUDOR</th>
			<th>MONTO</th>
			<th>FORMA PAGO</th>
			<th>LUGAR PAGO</th>
			<th>EJEC.ASIG.</th>
			<td>OBS.</td>

		</tr>
		</thead>
		<tbody>
	<%
abrirscg()

	strSql = "SELECT DEUDOR.NOMBRE_DEUDOR,"
	strSql = strSql & " USUARIO.LOGIN AS EJEC_ASIG,"
	strSql = strSql & " CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103) as 'FECHA_INGRESO',"

	strSql = strSql & " GESTIONSOLA = CASE WHEN ( GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1)"
	strSql = strSql & " 						  AND GESTIONES.FECHA_COMPROMISO <  CAST(CONVERT(VARCHAR(10),GETDATE(),103) AS DATETIME)"
	strSql = strSql & " 				   THEN 'COMPROMISO PENDIENTE'"
	strSql = strSql & "					   WHEN ( GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1)"
	strSql = strSql & " 				   THEN 'COMPROMISO'"
	strSql = strSql & " 				   ELSE 'NO DEFINIDO'"
	strSql = strSql & "					   END,"


	strSql = strSql & " CUOTA.RUT_DEUDOR AS RUT_DEUDOR,"
	strSql = strSql & " ISNULL(GESTIONES.MONTO_CANCELADO,0) AS SALDO,"
	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES,CHAR(13),' '),CHAR(10),' ') as OBSERVACIONES,"


	strSql = strSql & " (CASE WHEN REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')<>'' THEN "
	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')+'							'+'/HORARIO RETIRO: '+ GESTIONES.HORA_DESDE+' - '+ GESTIONES.HORA_HASTA"
	strSql = strSql & " ELSE 'SIN OBSERVACIÓN' +'							'+'/HORARIO RETIRO: '+ GESTIONES.HORA_DESDE+' - '+ GESTIONES.HORA_HASTA"
	strSql = strSql & " END) as OBSERVACIONES_CAMPO,"


	strSql = strSql & " ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'') AS FECHA_NORMALIZACION,"

	strSql = strSql & " LUGAR_PAGO = UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))),"

	strSql = strSql & " ISNULL(CAJA_FORMA_PAGO.DESC_FORMA_PAGO,'NO ESPEC.') AS 'FORMA PAGO',"

	strSql = strSql & " CASE WHEN ISNULL(GESTIONES.MONTO_CANCELADO,'')= 0"
	strSql = strSql & " THEN ''"
	strSql = strSql & " ELSE ISNULL(GESTIONES.MONTO_CANCELADO,'NO ESPEC.')"
	strSql = strSql & " END AS 'MONTO REGULARIZADO'"


	strSql = strSql & " FROM GESTIONES			INNER JOIN GESTIONES_CUOTA ON GESTIONES.ID_GESTION = GESTIONES_CUOTA.ID_GESTION"
	strSql = strSql & "                         INNER JOIN CUOTA ON GESTIONES_CUOTA.ID_CUOTA = CUOTA.ID_CUOTA"
	strSql = strSql & "                         INNER JOIN DEUDOR ON GESTIONES.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND GESTIONES.COD_CLIENTE = DEUDOR.COD_CLIENTE"
	strSql = strSql & "                         INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
	strSql = strSql & "                         LEFT JOIN CAJA_FORMA_PAGO ON GESTIONES.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO"

	strSql = strSql & " LEFT JOIN FORMA_RECAUDACION RE ON RE.ID_FORMA_RECAUDACION= GESTIONES.ID_FORMA_RECAUDACION "
	strSql = strSql & " LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION= GESTIONES.ID_DIRECCION_COBRO_DEUDOR "


	strSql = strSql & "                         INNER JOIN GESTIONES_TIPO_GESTION ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_GESTION = GESTIONES_TIPO_GESTION.COD_GESTION"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_CATEGORIA.COD_CATEGORIA"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_CATEGORIA AND"
	strSql = strSql & " 																	GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_SUB_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
	strSql = strSql & " 						LEFT JOIN USUARIO ON DEUDOR.USUARIO_ASIG = USUARIO.ID_USUARIO"


	strSql = strSql & " WHERE   (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1"
	strSql = strSql & " 		AND GESTIONES.FECHA_COMPROMISO <= CAST(CONVERT(VARCHAR(10),GETDATE(),103) AS DATETIME))"

	strSql = strSql & " 		AND ESTADO_DEUDA.ACTIVO = 1 "
	strSql = strSql & " 		AND CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION "

	strSql = strSql & "			AND CUOTA.COD_CLIENTE = '" & strCodCliente & "'"

	strSql = strSql & " AND (DEUDOR.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0,1))"
	strSql = strSql & " OR DEUDOR.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0,1)))"

	If Trim(strCobranza) = "INTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
		strParametro = "1"
	End if

	If Trim(strCobranza) = "EXTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
		strParametro = "1"
	End if

	if Trim(strTipoGestion) = "1" Then
		strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1) "
		strSql = strSql & " AND	  GESTIONES.FECHA_COMPROMISO = CAST(CONVERT(VARCHAR(10),GETDATE(),103) AS DATETIME)"
	End If

	if Trim(strTipoGestion) = "2" Then
		strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 1) "
		strSql = strSql & " AND	  GESTIONES.FECHA_COMPROMISO < CAST(CONVERT(VARCHAR(10),GETDATE(),103) AS DATETIME)"
	End If

	if Trim(strEstadoCP) = "1" Then
		strSql = strSql & " AND (DATEDIFF(MINUTE,(ISNULL(CUOTA.FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(CUOTA.HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) < 0)"
	End If

	if Trim(strEstadoCP) = "2" Then
		strSql = strSql & " AND (DATEDIFF(MINUTE,(ISNULL(CUOTA.FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(CUOTA.HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0)"
	End If

	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then

		strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
	Else
		if Trim(strEjeAsig) <> "" Then
		strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
		End if
	End if

	strSql = strSql & " GROUP BY NOMBRE_DEUDOR, USUARIO.LOGIN, GESTIONES.FECHA_INGRESO, GESTIONES_TIPO_GESTION.GESTION_MODULOS,GESTIONES.OBSERVACIONES,GESTIONES.OBSERVACIONES_CAMPO,"
	strSql = strSql & " GESTIONES.FECHA_COMPROMISO, CUOTA.RUT_DEUDOR,GESTIONES.HORA_DESDE, GESTIONES.HORA_HASTA, RE.NOMBRE, RE.UBICACION, DD.CALLE,DD.NUMERO,DD.RESTO,DD.comuna, CAJA_FORMA_PAGO.DESC_FORMA_PAGO, GESTIONES.MONTO_CANCELADO"

	strSql = strSql & " ORDER BY  USUARIO.LOGIN,GESTIONES.FECHA_INGRESO"


		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			intReg = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>
				<tr >
					<td><%=intReg%></td>
					<td><%=Mid(rsDet("GESTIONSOLA"),1,18)%></td>
					<td><%=rsDet("FECHA_NORMALIZACION")%></td>
					<td>
											<A HREF="principal.asp?TX_RUT=<%=rsDet("RUT_DEUDOR")%>">
											<acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym>
											</A>
					</td>
					<td><%=Mid(rsDet("NOMBRE_DEUDOR"),1,28)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("SALDO"),0)%></td>
					<td><%=Mid(rsDet("FORMA PAGO"),1,8)%></td>
					<td><%=Mid(rsDet("LUGAR_PAGO"),1,20)%></td>

					<td ALIGN="RIGHT"><%=rsDet("EJEC_ASIG")%></td>
					<td ALIGN="CENTER" title="<%=rsDet("OBSERVACIONES_CAMPO")%>">
						<img src="../imagenes/priorizar_normal.png" border="0">
					</td>
				</tr>

				<%
				rsDet.movenext
			loop
		Else
		
		%>
				<tr >
					<td HEIGHT = "20"  ALIGN="CENTER" Colspan = "10" class="estilo_columna_individual">NO EXISTEN RESULTADOS SEGUN PARAMETROS DE BUSQUEDA</td>
				</tr>
		<%
		rsDet.close
		set rsDet=nothing
		end if
		cerrarscg()%>

	</tbody>
	</table>
	</td>
   </tr>
  </table>

</form>

<br />
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
			var newOption = new Option('SIN USUARIO', '');
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
