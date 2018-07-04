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

<!--#include file="../lib/comunes/js_css/top_tooltip.inc" -->

<%

	Response.CodePage=65001
	Response.charset ="utf-8"
	
	AbrirSCG()

	strEjeAsig = request("CB_EJECUTIVO")
	strTipoGestion = request("cmb_tipogestion")
	strEstadoObj = request("CMB_ESTADO_OBJ")

	if strTipoGestion = "" then strTipoGestion = "0"
	if strEstadoObj = "" then strEstadoObj = "2"

	termino = request("termino")
	inicio = request("inicio")
	
	strCodCliente = session("ses_codcli")

	''response.write "<br>Ftro_TipoGestionNorm=" & session("Ftro_TipoGestionNorm")

	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_TipoGestionExp") = strTipoGestion
		session("Ftro_EjecAsigExp") = strEjeAsig
	End If

	If Trim(Request("strBuscar")) = "N" Then
		session("Ftro_TipoGestionExp") = ""
		session("Ftro_EjecAsigExp") = ""
	End If

	If strEjeAsig = "" Then strEjeAsig = session("Ftro_EjecAsigExp")
	If strTipoGestion = "0" Then strTipoGestion = session("Ftro_TipoGestionExp")

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

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />


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
function envia()
{
	//datos.TX_RUT.value='';
	//datos.TX_PAGO.value='';
	resp='si'
	document.datos.action = "Listado_Expone_requerimientos.asp?strBuscar=S&resp="+ resp +"";
	document.datos.submit();
}

function exportar()
{
	document.datos.action = "exp_Expone.asp";
	document.datos.submit();
}

</script>

<link href="../css/style.css" rel="Stylesheet">

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">


<form name="datos" method="post">
<table width="100%" border="0">
  <tr>
    <td height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" ALIGN="CENTER">MÓDULO CASOS OBJETADOS</td>
  </tr>
  <tr>
    <td valign="top">
	<table width="100%" border="0" bordercolor="#999999">

	      <tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

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
		  <tr bordercolor="#999999" class="Estilo8">

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
					<option value="1" <%If Trim(strTipoGestion)="1" Then Response.write "SELECTED"%>>CASO OBJETADO</option>
					<option value="3" <%If Trim(strTipoGestion)="3" Then Response.write "SELECTED"%>>CASO OBJETADO NO RESP.</option>

				</SELECT>
			</td>

			<td>
				<SELECT NAME="CMB_ESTADO_OBJ" id="CMB_ESTADO_OBJ">
					<option value="0" <%If Trim(strEstadoObj)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(strEstadoObj)="1" Then Response.write "SELECTED"%>>EN CONSULTA</option>
					<option value="2" <%If Trim(strEstadoObj)="2" Then Response.write "SELECTED"%>>NO PROCESADO</option>
				</SELECT>
			</td>

			<td><input name="inicio" type="text" id="inicio" readonly="true" value="<%=inicio%>" size="10" maxlength="10">
				<!--<a href="javascript:showCal('Calendar7');"><img src="../imagenes/calendario.gif" border="0"></a>-->
			</td>

			<td><input name="termino" type="text" id="termino" readonly="true" value="<%=termino%>" size="10" maxlength="10">
          		<!--<a href="javascript:showCal('Calendar6');"><img src="../imagenes/calendario.gif" border="0"></a>-->
			</td>

		<% If sinCbUsario="0" Then %>
			<td>
				<select name="CB_EJECUTIVO" style="width:130px;border:1px pgsolid #04467E;background-color:#FFFFFF;color:#000000;font-size:12px">
				</select>
			</td>
		<% End If %>

			<td align="center">
				<input type="Button" name="Submit" value="Ver" onClick="envia();">
				<input Name="SubmitButton" Value="Exportar" Type="BUTTON" onClick="exportar();">
			</td>

	      </tr>
    </table>

	<table width="100%" border="0" bordercolor="#000000">
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td>&nbsp;</td>
			<td>GESTION</td>
			<td>FECHA_ING</td>
			<td>RUT DEUDOR</td>
			<td>NOMBRE DEUDOR</td>
			<td>NRO_DOC</td>
			<td>FECHA GESTION</td>
			<td>$ ASOCIADO</td>
			<td>FORMA</td>
			<td>LUGAR GESTION</td>
			<td>NRO CP</td>
			<td>EJEC.ASIG.</td>
			<td>OBS.</td>

		</tr>
	<%
AbrirSCG()
	
	strSql = "SELECT DEUDOR.NOMBRE_DEUDOR, USUARIO.LOGIN AS EJEC_ASIG, "
	strSql = strSql & " CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103) as 'FECHA_INGRESO', (CASE WHEN GESTIONES.NRO_DOC_PAGO = '' THEN 'NO ESPEC'ELSE GESTIONES.NRO_DOC_PAGO END) AS NRO_DOC_PAGO,ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_PAGO,103),'NO ESPEC') AS FECHA_NORMALIZACION,"
	strSql = strSql & " 'FECHA CONSULTA: ' + ISNULL(CONVERT(VARCHAR(10),CUOTA.FECHA_CONSULTA_NORM,103),'NO CONSULTADO') AS FECHA_CONSULTA,"
	strSql = strSql & " GESTIONSOLA = CASE WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND FECHA_CONSULTA_NORM IS NULL)"
	strSql = strSql & " 				   THEN 'CASO OBJETADO'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ())"
	strSql = strSql & "                    THEN 'CASO OBJETADO EN CONSULTA'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
	strSql = strSql & "                    THEN 'CASO OBJETADO NO RESPONDIDO'"
	strSql = strSql & " 				   WHEN (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) > CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME))"
	strSql = strSql & "                    THEN 'REITERA CASO OBJETADO'"
	strSql = strSql & " 				   ELSE 'OTRO'"
	strSql = strSql & "					   END,"

	strSql = strSql & " CUOTA.RUT_DEUDOR AS RUT_DEUDOR,"
	strSql = strSql & " CUOTA.NRO_DOC,"
	strSql = strSql & " ISNULL(GESTIONES.MONTO_CANCELADO,0) AS SALDO,"

	strSql = strSql & " (CASE WHEN REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')<>'' THEN "
	strSql = strSql & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES_CAMPO,CHAR(13),' '),CHAR(10),' ')"
	strSql = strSql & " ELSE 'SIN OBSERVACIÓN'"
	strSql = strSql & " END) as OBSERVACIONES_CAMPO,"


	strSql = strSql & " CAST(GESTIONES.FECHA_INGRESO AS DATETIME) AS FECHA_INGRESO,"

	strSql = strSql & " ISNULL(UPPER(ISNULL(UPPER(RE.NOMBRE+' '+RE.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))), 'NO ESPEC') AS LUGAR_PAGO,"

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

	strSql = strSql & " LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION=GESTIONES.ID_DIRECCION_COBRO_DEUDOR "
	strSql = strSql & " LEFT JOIN FORMA_RECAUDACION RE ON RE.ID_FORMA_RECAUDACION=GESTIONES.ID_FORMA_RECAUDACION "

	strSql = strSql & "                         INNER JOIN GESTIONES_TIPO_GESTION ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA AND"
	strSql = strSql & "                                                                             GESTIONES.COD_GESTION = GESTIONES_TIPO_GESTION.COD_GESTION AND"
	strSql = strSql & "                                                                             GESTIONES.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_CATEGORIA.COD_CATEGORIA"
	strSql = strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_CATEGORIA AND"
	strSql = strSql & " 																	GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_SUBCATEGORIA.COD_SUB_CATEGORIA"
	strSql = strSql & " 						LEFT JOIN USUARIO ON CUOTA.USUARIO_ASIG = USUARIO.ID_USUARIO"


	strSql = strSql & " WHERE   ((GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ()) AND '" & strTipoGestion & "' IN (0,3)"

	If Trim(strEstadoObj) = "0" or Trim(strEstadoObj) = "2" Then
		strSql = strSql & " OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND FECHA_CONSULTA_NORM IS NULL AND '" & strTipoGestion & "' IN (0,1)))"
	Else
		strSql = strSql & " OR (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND '" & strTipoGestion & "' IN (0,1)))"
	End If

	strSql = strSql & " AND ESTADO_DEUDA.ACTIVO = 1 "

	If inicio <> "" then

	strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) > = '" & inicio & " 00:00:00'"

	End If

	If termino <> "" then

	strSql = strSql & " AND CAST(GESTIONES.FECHA_INGRESO AS DATETIME) < = '" & termino & " 23:59:59'"

	End If

	If Trim(strCobranza) = "INTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NOT NULL"
		strParametro = "1"
	End if

	If Trim(strCobranza) = "EXTERNA" Then
		strSql = strSql & " AND DEUDOR.CUSTODIO IS NULL"
		strParametro = "1"
	End if

	strSql = strSql & " 		AND CUOTA.ID_ULT_GEST = GESTIONES.ID_GESTION "
	strSql = strSql & "			AND CUOTA.COD_CLIENTE = '" & strCodCliente & "'"


	if Trim(strTipoGestion) = "1" and Trim(strEstadoObj) <> "1" Then
		strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND FECHA_CONSULTA_NORM IS NULL)"
	End If

	if Trim(strEstadoObj) = "1" Then
		strSql = strSql & " 	AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ()) "
	End If

	if Trim(strTipoGestion) = "3" Then
			strSql = strSql & " AND (GESTIONES_TIPO_GESTION.GESTION_MODULOS = 3 AND (GESTIONES.FECHA_INGRESO + GESTIONES.HORA_INGRESO) < CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ())"
	End If


	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
		strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & session("session_idusuario") & "'"
	Else
		if Trim(strEjeAsig) <> "" Then
			strSql = strSql & " AND  DEUDOR.USUARIO_ASIG = '" & strEjeAsig & "'"
		End if
	End if

	strSql = strSql & " ORDER BY  USUARIO.login,GESTIONES.FECHA_INGRESO"



		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			intReg = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>
				<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
					<td><%=intReg%></td>
					<td ALIGN="LEFT" onMouseover="ddrivetip('<%=rsDet("FECHA_CONSULTA")%>', '#EFEFEF',300)"; onMouseout="hideddrivetip()">
					<%=Mid(rsDet("GESTIONSOLA"),1,25)%>
					<td><%=rsDet("FECHA_INGRESO")%></td>
					<td>
											<A HREF="principal.asp?TX_RUT=<%=rsDet("RUT_DEUDOR")%>">
											<acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym>
											</A>
					</td>
					<td ALIGN="LEFT" onMouseover="ddrivetip('<%=rsDet("NOMBRE_DEUDOR")%>', '#EFEFEF',300)"; onMouseout="hideddrivetip()"><%=Mid(rsDet("NOMBRE_DEUDOR"),1,20)%>

					<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
					<td><%=rsDet("FECHA_NORMALIZACION")%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("MONTO REGULARIZADO"),0)%></td>
					<td><%=Mid(rsDet("FORMA PAGO"),1,8)%></td>
					<td><%=Mid(rsDet("LUGAR_PAGO"),1,26)%></td>
					<td><%=Mid(rsDet("NRO_DOC_PAGO"),1,20)%></td>
					<td ALIGN="LEFT"><%=rsDet("EJEC_ASIG")%></td>
					<td ALIGN="CENTER" onMouseover="ddrivetip('<%=rsDet("OBSERVACIONES_CAMPO")%>', '#EFEFEF',300)"; onMouseout="hideddrivetip()">
						<img src="../imagenes/priorizar_normal.png" border="0">
					</td>
					</td>
				</tr>
				<%
				rsDet.movenext
			loop
		rsDet.close
		set rsDet=nothing

		Else
		
		%>
				<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
					<td HEIGHT = "20"  ALIGN="CENTER" Colspan = "13">NO EXISTEN RESULTADOS SEGUN PARAMETROS DE BUSQUEDA</td>
				</tr>
		<%
		
		end if

	cerrarscg()%>

		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td ALIGN="CENTER" Colspan = "13">&nbsp;</td>
		</tr>

	</table>
	</td>
   </tr>
  </table>

</form>


</body>
</html>

<!--#include file="../lib/comunes/js_css/bottom_tooltip.inc" -->

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
