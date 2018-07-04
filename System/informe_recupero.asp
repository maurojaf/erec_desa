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
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../Componentes/FC/FusionCharts_Gen.asp"-->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<script language="JavaScript">
function ventanaSecundaria (URL){
	window.open(URL,"DETALLE","width=200, height=200, scrollbars=no, menubar=no, location=no, resizable=no")
}

</script>

<%
abrirscg()

Response.CodePage=65001
Response.charset ="utf-8"

inicio= request("inicio")
termino= request("termino")

intFechaIni = inicio
intFechaFin = termino


	If Trim(inicio) = "" Then
		inicio = TraeFechaActual(Conn)
		inicio = "01/" & Mid(TraeFechaActual(Conn),4,10)
	End If

	If Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If
cerrarscg()

strCodCliente=session("ses_codcli")
intOrigen = request("CB_ORIGEN")
intOrigen2 = request("CB_ESTADO_DEUDA")
intCodRemesa = request("CB_REMESA")
strEjeAsig = request("CB_EJECUTIVO")

If Trim(intOrigen) = "" Then intOrigen = "T"
If Trim(intOrigen2) = "" Then intOrigen2 = "T"

'Response.write "<BR>strEjeAsig=" & strEjeAsig
'Response.write "<BR>intOrigen=" & intOrigen

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

If Trim(strCodCliente) = "" Then strCodCliente = "1000"
%>
<title>INFORME RECUPERACIÓN</title>
<SCRIPT LANGUAGE="Javascript" SRC="../Componentes/FC/FusionCharts.js"></SCRIPT>

<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<style type="text/css">
<!--
.Estilo37 {color: #FFFFFF}
-->
</style>
</head>
<body>
<DIV class="titulo_informe">RECUPERACIÓN DE PAGOS</DIV>
<br>
<table width="90%" align="CENTER" border="0">
  <tr>
    <td valign="top">
	<BR>
	<form name="datos" method="post">
	<table width="100%" border="0" class="estilo_columnas">
		<thead>
			<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td>CLIENTE</td>
				<td>COBRANZA</td>
				<td>ESTADO DEUDA</td>
				<td>ORIGEN PAGO</td>
				<td>F.INICIO</td>
				<td>F.TERMINO</td>

				<% If sinCbUsario = "0" Then %>
				<td>EJECUTIVO</td>
				<% End If %>

				<td width="10%">&nbsp</td>
			</tr>
		</thead>
			<tr>
				<td>
					<select name="CB_CLIENTE" onChange="refrescar();">
						<%
						abrirscg()
						ssql="SELECT COD_CLIENTE,DESCRIPCION FROM CLIENTE WHERE COD_CLIENTE = '" & strCodCliente & "'"
						
						abrirscg()
						set rsCLI= Conn.execute(ssql)
						if not rsCLI.eof then
							Do until rsCLI.eof%>
							<option value="<%=rsCLI("COD_CLIENTE")%>" <%if Trim(strCodCliente)=rsCLI("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsCLI("descripcion")%></option>
						<%
							rsCLI.movenext
							Loop
							end if
							rsCLI.close
							set rsCLI=nothing
						cerrarscg()
						%>
					</select>
				</td>

				<td>
					<select name="CB_COBRANZA" <%If sinCbUsario = "0" then%> onChange="CargaUsuarios(this.value,CB_CLIENTE.value);" <%End If%> >

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
					<%If Trim(intOrigen2) = "T" Then strSelTodos2 = "SELECTED"
					If Trim(intOrigen2) = "E" Then strSelEmpresa2 = "SELECTED"
					If Trim(intOrigen2) = "C" Then strSelMandante2 = "SELECTED"
					If Trim(intOrigen2) = "CR" Then strSelConvenioRepactacion = "SELECTED"
					If Trim(intOrigen2) = "AE" Then strSelAbonoEmpresa = "SELECTED"
					If Trim(intOrigen2) = "AC" Then strSelAbonoCliente = "SELECTED"
					%>
					<select name="CB_ESTADO_DEUDA">
						<option value="T" <%=strSelTodos2%>>TODOS</option>
						<option value="E" <%=strSelEmpresa2%>>PAGO EN EMPRESA</option>
						<option value="C" <%=strSelMandante2%>>PAGO EN CLIENTE</option>
						<option value="CR" <%=strSelConvenioRepactacion%>>CONVENIO Y REPACTACION</option>
						<option value="AE" <%=strSelAbonoEmpresa%>>ABONO EMPRESA</option>
						<option value="AC" <%=strSelAbonoCliente%>>ABONO CLIENTE</option>
					</select>
				</td>
				<td>
					<%
					If Trim(intOrigen) = "T" Then strSelTodos = "SELECTED"
					If Trim(intOrigen) = "E" Then strSelEmpresa = "SELECTED"
					If Trim(intOrigen) = "C" Then strSelCliente = "SELECTED"
					%>
					<select name="CB_ORIGEN">
						<option value="T" <%=strSelTodos%>>TODOS</option>
						<option value="E" <%=strSelEmpresa%>>SUCURSALES LLACRUZ</option>
						<option value="C" <%=strSelCliente%>>CLIENTE</option>
					</select>
				</td>
				<td>
					<input name="inicio" readonly="true" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
				</td>
				<td>
					<input name="termino" readonly="true" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
				</td>

				<% If sinCbUsario="0" Then %>
				<td>
					<select name="CB_EJECUTIVO" id="CB_EJECUTIVO" >
					</select>
				</td>
				<% End If %>

				<td>
					<input type="button" class="fondo_boton_100" name="Submit" value="Aceptar" onClick="envia();">
				</td>
			</tr>
    </table>
</form>

<table border="0" class="intercalado" style="width:100%;">
<thead>
  <tr bgcolor="#<%=session("COLTABBG")%>">
		<td><span class="Estilo37">FECHA</span></td>
  		<td><span class="Estilo37">DIA</span></td>
  		<td><span class="Estilo37">CASOS</span></td>
  		<td><span class="Estilo37">MONTO</span></td>
  		<td><span class="Estilo37">DOCS</span></td>
  		<td><span class="Estilo37">ACUM CASOS</span></td>
  		<td><span class="Estilo37">ACUM MONTO</span></td>
  		<td><span class="Estilo37">ACUM DOCS</span></td>
  	</tr>
</thead>
<tbody>
    <%
'QUERY QUE RETORNA LOS DATOS ASOCIADOS A LAS GESTIONES
	strSql = "SELECT "
	strSql = strSql & " DIA   = DATENAME (dw, FECHA_ESTADO), "
	strSql = strSql & " FECHA = CONVERT(VARCHAR(10),FECHA_ESTADO,103), "
	strSql = strSql & " CASOS = COUNT(DISTINCT(RUT_DEUDOR)), "
	strSql = strSql & " DOC = IsNull(COUNT(NRO_DOC),0), "
	strSql = strSql & " MONTO = IsNull(SUM(VALOR_CUOTA),0) "
	strSql = strSql & " FROM CUOTA "
	strSql = strSql & " WHERE COD_CLIENTE = '"& strCodCliente &"' "
	strSql = strSql & " AND SALDO = 0 "
	strSql = strSql & " AND FECHA_ESTADO BETWEEN '"&inicio&" 00:00:00' AND '"&termino&" 23:59:59' "
	
	strSql = strSql & " AND ESTADO_DEUDA IN "
	
	If Trim(intOrigen) = "T" Then
	 strSql = strSql & " (3,4,7,8,10,11)"
	End if
	If Trim(intOrigen) = "E" Then
	 strSql = strSql & " (4,8,10,11)"
	End if
	If Trim(intOrigen) = "C" Then
	 strSql = strSql & " (3,7)"
	End if

	strSql = strSql & " AND ESTADO_DEUDA IN "

	If Trim(intOrigen2) = "T" Then
	 strSql = strSql & " (3,4,7,8,10,11)"
	End if
	If Trim(intOrigen2) = "E" Then
	 strSql = strSql & " (4)"
	End if
	If Trim(intOrigen2) = "C" Then
	 strSql = strSql & " (3)"
	End if
	If Trim(intOrigen2) = "CR" Then
	 strSql = strSql & " (10, 11)"
	End if
	If Trim(intOrigen2) = "AC" Then
	 strSql = strSql & " (7)"
	End if
	If Trim(intOrigen2) = "AE" Then
	 strSql = strSql & " (8)"
	End if
	
	If sinCbUsario = "" Then
	 strSql = strSql & " AND USUARIO_ASIG = " & session("session_idusuario")
	End If
	
	If Trim(strEjeAsig) <> "" Then
	 strSql = strSql & " AND USUARIO_ASIG = " & strEjeAsig
	End If

	If Trim(strCobranza) = "INTERNA" Then
	 strSql = strSql & " AND CUOTA.CUSTODIO IS NOT NULL"
	End if

	If Trim(strCobranza) = "EXTERNA" Then
	 strSql = strSql & " AND CUOTA.CUSTODIO IS NULL"
	End if
	
	strSql = strSql & " GROUP BY CONVERT(VARCHAR(10),FECHA_ESTADO,103),DATENAME (dw, FECHA_ESTADO),CAST(FECHA_ESTADO AS DATE) "
	strSql = strSql & " ORDER BY CAST(FECHA_ESTADO AS DATE) ASC "

	'Response.write ("<br>" & strSql & "<br>")
	'Response.end
	
	AbrirSCG()
	
		set rsDatos = Conn.execute(strSql)
		
		intNumReg=0
		
		Do while not rsDatos.eof
		
		intNumReg= intNumReg + 1
		strNomDiaFecha  = rsDatos("DIA")
		intFecha		= rsDatos("FECHA")
		intCasos        = rsDatos("CASOS")
		intMonto		= rsDatos("MONTO")
		intDocs			= rsDatos("DOC")
		
		intAcumCasos	= intAcumCasos + intCasos
		intAcumMonto	= intAcumMonto + intMonto
		intAcumDocs		= intAcumDocs + intDocs

				%>
				<tr>
						<TD WIDTH="10%" ALIGN="LEFT">
							<A HREF="detalle_recuperacion.asp?intFechaIni=<%=intFechaIni%>&intFechaFin=<%=intFechaFin%>&intFecha=<%=intFecha%>&intCliente=<%=strCodCliente%>&intOrigen=<%=intOrigen%>&intOrigen2=<%=intOrigen2%>&intCodRemesa=<%=intCodRemesa%>&intCodUsuario=<%=strEjeAsig%>">
								<%=intFecha%>
							</A>
						</td>
						<TD WIDTH="10%" ALIGN="LEFT">
							<A HREF="detalle_recuperacion.asp?intFechaIni=<%=intFechaIni%>&intFechaFin=<%=intFechaFin%>&intFecha=<%=intFecha%>&intCliente=<%=strCodCliente%>&intOrigen=<%=intOrigen%>&intOrigen2=<%=intOrigen2%>&intCodRemesa=<%=intCodRemesa%>&intCodUsuario=<%=strEjeAsig%>">
								<%=strNomDiaFecha%>
							</A>
						</td>
						
						<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intCasos,0)%></td>
						<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intMonto,0)%></td>
						<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intDocs,0)%></td>
						<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intAcumCasos,0)%></td>
						<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intAcumMonto,0)%></td>
						<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intAcumDocs,0)%></td>
				</tr>
				<%	
			rsDatos.movenext
			Loop
		rsDatos.close
		set rsDatos=nothing	
		
		If intNumReg=0 then				
			
			%>
			<tr bgcolor="<%=strbgcolor%>" class="Estilo8">																					
				<td colspan="8" align = "center"><h3>No Existen Pagos según Parámetros Definidos</h3></td>	
			</tr>
<%		end if

	CerrarSCG()

	%>
	</tbody>
	<thead>
	  <tr class="totales">
	  		<td colspan="2" ALIGN="LEFT">TOTALES</td>
	  		<td ALIGN="RIGHT"><span><%=FN(intAcumCasos,0)%></span></td>
			<td ALIGN="RIGHT"><span><%=FN(intAcumMonto,0)%></span></td>
	  		<td ALIGN="RIGHT"><span><%=FN(intAcumDocs,0)%></span></td>
			<td colspan="3"><span>&nbsp</span></td>
  	</tr>
  </thead>

</table>

<br>
<br>

</body>
</html>

<script type="text/javascript">
$(document).ready(function(){

	$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
 
})
</script>

<script language="JavaScript1.2">
function envia(){
		if (datos.CB_CLIENTE.value=='0'){
			alert('DEBE SELECCIONAR UN CLIENTE');
		}else if(datos.inicio.value==''){
			alert('DEBE SELECCIONAR FECHA DE INICIO');
		}else if(datos.termino.value==''){
			alert('DEBES SELECCIONAR FECHA DE TERMINO');
		}else{
		//datos.action='cargando.asp';
		datos.action='informe_recupero.asp';
		datos.submit();
	}
}


function refrescar(){
		if (datos.CB_CLIENTE.value=='0'){
			alert('DEBE SELECCIONAR UN CLIENTE');
		}else
		{
		datos.action='informe_recupero.asp';
		datos.submit();
	}
}

function CargaUsuarios(subCat,cat)
{
	//alert(subCat);
	//alert(cat);

	var comboBox = document.getElementById('CB_EJECUTIVO');
	switch (cat)
	{
		<%
		  AbrirSCG()
			strSql="SELECT COD_CLIENTE FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE"
			set rsGestCat=Conn.execute(strSql)
			Do While not rsGestCat.eof
		%>
		case '<%=rsGestCat("COD_CLIENTE")%>':

			comboBox.options.length = 0;

				if (subCat=='INTERNA') {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

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
					break;
				}

				if (subCat=='EXTERNA' && (<%=intVerEjecutivos%>=='1')) {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

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
					break;
				}
				else if ((subCat=='EXTERNA') && (<%=intVerEjecutivos%>=='0')) {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					break;
				}
				else {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

					strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
					''strSql = strSql & " AND U.PERFIL_EMP=0"


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
					break;
				}

		<%
		  	rsGestCat.movenext
		  	Loop
		  	rsGestCat.close
		  	set rsGestCat=nothing
			CerrarSCG()
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
CargaUsuarios('<%=strCobranza%>','<%=strCodCliente%>');
<%End If%>

<%If strEjeAsig <> "" then%>
datos.CB_EJECUTIVO.value='<%=strEjeAsig%>';
<%End If%>
</script>

