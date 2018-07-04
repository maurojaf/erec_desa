<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>

<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/lib.asp"-->
<!--#include file="../lib/freeaspupload.asp" -->
<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->

<html xmlns="http:www.w3.org/1999/xhtml">
<head>
<LINK rel="stylesheet" TYPE="text/css" HREF="../css/style.css">
</head>


<%

Response.CodePage=65001
Response.charset ="utf-8"

strCliente=Trim(Request("CB_CLIENTE"))

intUsuario = Trim(request("CB_USUARIO"))

strCobranza=Trim(Request("CB_CUSTODIO"))

strTipoProceso=Request("CB_TIPOPROCESO")

strFechaArchivoGestion =Request("CB_GES_CARGA")


inicioFecEstado=Request("inicioFecEstado")
terminoFecEstado=Request("terminoFecEstado")

inicioFecGest=Request("inicioFecGest")
terminoFecGest=Request("terminoFecGest")

if Request("archivo")<>"" then
	archivo=Request("archivo")
End if

abrirscg()
	If Trim(inicio) = "" Then
		inicioFecGest = TraeFechaActual(Conn)
		inicioFecGest = "01/" & Mid(TraeFechaActual(Conn),4,10)
	End If

	If Trim(termino) = "" Then
		terminoFecGest = TraeFechaActual(Conn)
	End If
cerrarscg()


	Dim DestinationPath
	DestinationPath = Server.mapPath("../Archivo/UploadFolder")

	Dim uploadsDirVar
	uploadsDirVar = DestinationPath

	function SaveFiles


			If Trim(strTipoProceso) = "DEUDA" Then
		  		Response.Redirect "exp_Deuda.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&Fecha=" + Fecha +"&CB_CUSTODIO=" + strCobranza + "&CB_USUARIO=" + intUsuario + "&dtmInicio=" + inicioFecEstado + "&dtmTermino=" + terminoFecEstado + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_ACTIVO=" + strActivo
		  	End If
		  	If Trim(strTipoProceso) = "DEUDA_AGRUP" Then
				Response.Redirect "exp_DeudaAgrupada.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&Fecha=" + Fecha +"&CB_CUSTODIO=" + strCobranza + "&CB_ASIGNACION=" + strAsignacion + "&CB_USUARIO=" + intUsuario + "&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoProceso) = "GESTIONES" Then
				Response.Redirect "exp_Gestiones.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha + "&CB_CUSTODIO=" + strCobranza + "&CB_USUARIO=" + intUsuario + "&dtmInicio=" + inicioFecGest + "&dtmTermino=" + terminoFecGest + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_EFECTIVA=" + CH_EFECTIVA +"&CH_ACTIVO=" + strActivo
		  	End If
		  	If Trim(strTipoProceso) = "TELEFONOS" Then
				Response.Redirect "exp_Telefonos.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha +"&CB_ASIGNACION=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_ACTIVO=" + strActivo
		  	End If
		  	If Trim(strTipoProceso) = "DIRECCIONES" Then
				Response.Redirect "exp_Direcciones.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha +"&CB_ASIGNACION=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_ACTIVO=" + strActivo
		  	End If
		  	If Trim(strTipoProceso) = "EMAIL" Then
				Response.Redirect "exp_Email.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha +"&CB_ASIGNACION=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_ACTIVO=" + strActivo
		  	End If
		  	If Trim(strTipoProceso) = "BASE_ESTADO" and strCliente = 1070 Then
				Response.Redirect "exp_Base_estado_UMA.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha + "&CB_USUARIO=" + intUsuario + "&dtmInicio=" + inicioFecEstado + "&dtmTermino=" + terminoFecEstado + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_EFECTIVA=" + CH_EFECTIVA +"&CH_ACTIVO=" + strActivo
		  	End If
		  	If Trim(strTipoProceso) = "BASE_ESTADO" and strCliente = 1500 Then
				Response.Redirect "exp_Base_estado_UMA_Fact.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha + "&CB_USUARIO=" + intUsuario + "&dtmInicio=" + inicioFecEstado + "&dtmTermino=" + terminoFecEstado + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_EFECTIVA=" + CH_EFECTIVA +"&CH_ACTIVO=" + strActivo
		  	End If
			If Trim(strTipoProceso) = "BASE_ESTADO" and strCliente = 1200 Then
				Response.Redirect "exp_Base_estado_upv.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha + "&CB_USUARIO=" + intUsuario + "&dtmTermino=" + dtmTermino + "&dtmInicio=" + dtmInicio + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_EFECTIVA=" + CH_EFECTIVA +"&CH_ACTIVO=" + strActivo
		  	End If

	End function

if Request.ServerVariables("REQUEST_METHOD") = "POST" and archivo = "1" then

	response.write SaveFiles()

End if


if Request.ServerVariables("REQUEST_METHOD") = "POST" and archivo = "100" then

	response.write DownloadFile("../Archivo/Otros/1100/Gestiones/"&trim(strFechaArchivoGestion))

End if


%>
<title>MODULO DE EXPORTACION DE DATOS</title>

<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
</head>
<%strTitulo="MI CARTERA"%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<TABLE WIDTH="1000" align = "CENTER" border=0 cellspacing=0>
   <TR HEIGHT="20" VALIGN="MIDDLE" BGCOLOR="#EEEEEE">
		<TD ALIGN=CENTER>
			<B>MODULO DE EXPORTES</B>
		</TD>
    </TR>
</TABLE>

<form name="datos" id="datos" onSubmit="return enviar(this)"  method="POST" action="man_Export.asp">

	<table width="1000" border="0" ALIGN="CENTER">
		<tr height="25" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="33%">Cliente:</td>
			<td width="77%">Tipo Exporte:</td>
		</tr>
		  	<tr height="50" BGCOLOR="#EEEEEE">
		  			<td>
						<select name="CB_CLIENTE" id="CB_CLIENTE" onChange="CargaTipoExporte(CB_CLIENTE.value,'');">
							<option value="Seleccionar">SELECCIONAR</option>
							<%
							AbrirSCG()
							ssql="SELECT COD_CLIENTE, DESCRIPCION FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") ORDER BY DESCRIPCION "
							set rsTemp= Conn.execute(ssql)
							if not rsTemp.eof then
								do until rsTemp.eof%>
									<option value="<%=rsTemp("COD_CLIENTE")%>"<%if strCliente=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("DESCRIPCION")%></option>
								<%
								rsTemp.movenext
								loop
							end if
							rsTemp.close
							set rsTemp=nothing
							CerrarSCG()
							%>
						</select>
					</td>
					<td>
						<select name="CB_TIPOPROCESO" id="CB_TIPOPROCESO" style="width:130px;border:1px pgsolid #04467E;background-color:#FFFFFF;color:#000000;font-size:12px" onChange="cajas();">
						</select>
					</td>
			</tr>
	</table>


	<div name="divCustodio" id="divCustodio" style="display:none" >
		<table width="1000" width="100%" border="0" ALIGN="CENTER">
			<tr>
				<td colspan = "2" height="25" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">Filtros:</td>
			</tr>
			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="325">Tipo Cobranza :</td>
				<td>
					<select name="CB_CUSTODIO" style="width:130px;border:1px pgsolid #04467E;background-color:#FFFFFF;color:#000000;font-size:12px" onChange="CargaUsuarios(this.value,CB_CLIENTE.value);">
					</select>
				</td>
			</tr>
		</table>

	</div>

	<div name="divGestionCarga" id="divGestionCarga" style="display:none" >
		<table width="1000" width="100%" border="0" ALIGN="CENTER">
			<tr>
				<td colspan = "2" height="25" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">Filtros:</td>
			</tr>
			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="325">Fecha Archivo :</td>
				<td>
					<select name="CB_GES_CARGA">
						<%
						abrirscg()
						ssql ="SELECT TOP 20 id_archivo, nombre_archivo, cod_cliente, rut  " & _
								"FROM CARGA_ARCHIVOS " & _
								"WHERE activo =1 AND cod_cliente=1100 "  & _
								" AND origen = 6 " & _
								" order by fecha_carga desc "
						set rsTemp= Conn.execute(ssql)

						response.write ssql
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=Trim(rsTemp("NOMBRE_ARCHIVO"))%>"<%if Trim(strFechaArchivoGestion)=Trim(rsTemp("NOMBRE_ARCHIVO")) then response.Write("Selected") End If%>><%=Trim(rsTemp("NOMBRE_ARCHIVO"))%></option>
							<%
							rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing
						cerrarscg()
						%>
					</select>
				</td>
			</tr>
		</table>

	</div>

	<div name="divEjecutivo" id="divEjecutivo" style="display:none" >

		<table width="1000" width="100%" border="0" ALIGN="CENTER">
			<tr height="50" BGCOLOR="#EEEEEE">

				<td width="325">Ejecutivo :</td>
				<td>
					<select name="CB_USUARIO" style="width:130px;border:1px pgsolid #04467E;background-color:#FFFFFF;color:#000000;font-size:12px">
					</select>
				</td>
			</tr>
	</table>
	</div>

	<div name="divFecGestion" id="divFecGestion" style="display:none" >
		<table width="1000" border="0" ALIGN="CENTER">
			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="325">Fecha Ingreso Gestión:</td>
				<td>
					<b>Desde :&nbsp;<input name="inicioFecGest" type="text" id="inicioFecGest" value="<%=inicioFecGest%>" size="10" maxlength="10">
					<a href="javascript:showCal('inicioFecGest');"><img src="../Imagenes/calendario.gif" border="0"></a>
					Hasta :&nbsp;<input name="terminoFecGest" type="text" id="terminoFecGest" value="<%=terminoFecGest%>" size="10" maxlength="10">
					<a href="javascript:showCal('terminoFecGest');"><img src="../Imagenes/calendario.gif" border="0"></a>
					</b>
				</td>
			</tr>
		</table>
	</div>

	<div name="divFecEstado" id="divFecEstado" style="display:none" >
		<table width="1000" border="0" ALIGN="CENTER">
			<tr>
				<td colspan = "2" height="25" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">Si desea añadir documentos regularizados al exporte seleccione fechas, de lo contrario solo saldrán documentos activos. </td>
			</tr>

			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="152">Fecha regularización:</td>
				<td>
					<b>Desde :&nbsp;<input name="inicioFecEstado" type="text" id="inicioFecEstado" size="10" maxlength="10">
					<a href="javascript:showCal('inicioFecEstado');"><img src="../Imagenes/calendario.gif" border="0"></a>
					Hasta :&nbsp;<input name="terminoFecEstado" type="text" id="terminoFecEstado"  size="10" maxlength="10">
					<a href="javascript:showCal('terminoFecEstado');"><img src="../Imagenes/calendario.gif" border="0"></a>
					</b>
				</td>
			</tr>
		</table>
	</div>

	<!--div name="divExportar" id="divExportar" style="display:none" -->
		<table width="1000" width="100%" border="0" ALIGN="CENTER">
				<tr height="50" BGCOLOR="#EEEEEE">
					<td widht="500" align="RIGHT">
						<input Name="btProcesar" Value="Exportar" Type="BUTTON" onClick="enviar();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						
					</td>
				</tr>
		</table>
	<!--/div-->

</FORM>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script language="JavaScript1.2">

function MostrarFilas(Fila) {
var elementos = document.getElementsByName(Fila);
	for (i = 0; i< elementos.length; i++) {
		if(navigator.appName.indexOf("Microsoft") > -1){
			   var visible = 'block'
		} else {
			   var visible = 'table-row';
		}
elementos[i].style.display = visible;
		}
}

function OcultarFilas(Fila) {
	var elementos = document.getElementsByName(Fila);
	for (k = 0; k< elementos.length; k++) {
			   elementos[k].style.display = "none";
	}
}

function cajas()
{
	if (datos.CB_TIPOPROCESO.value == 'GESTIONES')
		{
			MostrarFilas('divCustodio');
			MostrarFilas('divEjecutivo');
			MostrarFilas('divFecGestion');
			OcultarFilas('divFecEstado');
			OcultarFilas('divGestionCarga');
		}
	else if (datos.CB_TIPOPROCESO.value == 'DEUDA')
		{
			MostrarFilas('divEjecutivo');
			MostrarFilas('divCustodio');
			MostrarFilas('divFecEstado');
			OcultarFilas('divFecGestion');
			OcultarFilas('divGestionCarga');
		}
	else if (datos.CB_TIPOPROCESO.value == 'DEUDA_AGRUP')
		{
			OcultarFilas('divFecGestion');
			OcultarFilas('divFecEstado');
			MostrarFilas('divCustodio');
			MostrarFilas('divEjecutivo');
			OcultarFilas('divGestionCarga');
		}
	else if (datos.CB_TIPOPROCESO.value == 'BASE_ESTADO' && datos.CB_CLIENTE.value != '1200')
		{
			OcultarFilas('divFecGestion');
			MostrarFilas('divFecEstado');
			OcultarFilas('divCustodio');
			OcultarFilas('divEjecutivo');
			OcultarFilas('divGestionCarga');
		}
	else if (datos.CB_TIPOPROCESO.value == 'CARGA_GESTIONES' && datos.CB_CLIENTE.value == '1100')
		{
			MostrarFilas('divGestionCarga');
			OcultarFilas('divEjecutivo');
			OcultarFilas('divFecGestion');
			OcultarFilas('divCustodio');
			OcultarFilas('divFecEstado');
		}
	else
		{
			OcultarFilas('divEjecutivo');
			OcultarFilas('divFecGestion');
			OcultarFilas('divCustodio');
			OcultarFilas('divFecEstado');
			OcultarFilas('divGestionCarga');
		}

	CargaUsuarios(document.datos.CB_CUSTODIO.value,document.datos.CB_CLIENTE.value)
}

function enviar()
{

	var CB_CLIENTE 		=$('#CB_CLIENTE').val()
	var CB_TIPOPROCESO 	=$('#CB_TIPOPROCESO').val()


	if(CB_CLIENTE==1100 && CB_TIPOPROCESO=="CARGA_GESTIONES")
	{

		
		if(document.datos.CB_CLIENTE.value =='Seleccionar'){
			alert('Debe seleccionar el cliente');
			return false;
		}else if(document.datos.CB_TIPOPROCESO.value ==''){
			alert('Debe seleccionar el tipo de proceso');
			return false;
		}else{

			datos.action = "man_Export.asp?archivo=100&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&CB_CUSTODIO=" + document.datos.CB_CUSTODIO.value + "&CB_USUARIO=" + document.datos.CB_USUARIO.value;
			datos.submit();
		}

	}
	else{

		

		if(document.datos.CB_CLIENTE.value =='Seleccionar'){
			alert('Debe seleccionar el cliente');
			return false;
		}else if(document.datos.CB_TIPOPROCESO.value ==''){
			alert('Debe seleccionar el tipo de proceso');
			return false;
		}else{
			datos.btProcesar.disabled = true;
			datos.action = "man_Export.asp?archivo=1&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&CB_CUSTODIO=" + document.datos.CB_CUSTODIO.value + "&CB_USUARIO=" + document.datos.CB_USUARIO.value;
			datos.submit();
		}

	}
	
	

	
}

function baja_txt()
{
		if(document.datos.CB_CLIENTE.value =='Seleccionar'){
			alert('Debe seleccionar el cliente');
			return false;
		}else if(document.datos.CB_TIPOPROCESO.value ==''){
			alert('Debe seleccionar el tipo de proceso');
			return false;
		}else{

			datos.action = "man_Export.asp?archivo=100&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&CB_CUSTODIO=" + document.datos.CB_CUSTODIO.value + "&CB_USUARIO=" + document.datos.CB_USUARIO.value;
			datos.submit();
		}
		
}

function CargaTipoExporte(subCat,datos)
{
		//alert(subCat);

		var comboBox = document.getElementById('CB_TIPOPROCESO');
		comboBox.options.length = 0;

		var newOption = new Option('SELECCIONAR');
		comboBox.options[comboBox.options.length] = newOption;

		var newOption = new Option('DEUDA', 'DEUDA');
		comboBox.options[comboBox.options.length] = newOption;

		var newOption = new Option('DEUDA AGRUPADA', 'DEUDA_AGRUP');
		comboBox.options[comboBox.options.length] = newOption;

		var newOption = new Option('GESTIONES', 'GESTIONES');
		comboBox.options[comboBox.options.length] = newOption;

		var newOption = new Option('TELEFONOS', 'TELEFONOS');
		comboBox.options[comboBox.options.length] = newOption;

		var newOption = new Option('DIRECCIONES', 'DIRECCIONES');
		comboBox.options[comboBox.options.length] = newOption;

		var newOption = new Option('EMAIL', 'EMAIL');
		comboBox.options[comboBox.options.length] = newOption;


	if (subCat =='1070' || subCat =='1500' || subCat =='1200')
		{
			var newOption = new Option('BASE ESTADO', 'BASE_ESTADO');
			comboBox.options[comboBox.options.length] = newOption;
		}

	if (subCat =='1100' || datos == 'CARGA_GESTIONES' )
		{
			var newOption = new Option('CARGA GESTIONES', 'CARGA_GESTIONES');
			comboBox.options[comboBox.options.length] = newOption;
		}

		CargaCustodio(subCat)
		CargaUsuarios(document.datos.CB_CUSTODIO.value,document.datos.CB_CLIENTE.value)
		OcultarFilas('divEjecutivo');
		OcultarFilas('divFecGestion');
		OcultarFilas('divCustodio');
		OcultarFilas('divFecEstado');

}

function CargaCustodio(subCat)
{
	var comboBox = document.getElementById('CB_CUSTODIO');
	comboBox.options.length = 0;

	if (subCat !='1200')
		{
			var newOption = new Option('TODOS','0');
			comboBox.options[comboBox.options.length] = newOption;

			var newOption = new Option('INTERNA', 'INTERNA');
			comboBox.options[comboBox.options.length] = newOption;

			var newOption = new Option('EXTERNA', 'EXTERNA');
			comboBox.options[comboBox.options.length] = newOption;

		}
	else
		{
			var newOption = new Option('TODOS','0');
			comboBox.options[comboBox.options.length] = newOption;

			var newOption = new Option('INTERNA', 'INTERNA');
			comboBox.options[comboBox.options.length] = newOption;

			var newOption = new Option('EXTERNA', 'EXTERNA');
			comboBox.options[comboBox.options.length] = newOption;
		}

}


function InicializaComboTipoExporte()
{
		var comboBox = document.getElementById('CB_TIPOPROCESO');
		comboBox.options.length = 0;
		var newOption = new Option('SELECCIONAR');
		comboBox.options[comboBox.options.length] = newOption;
}


function CargaUsuarios(subCat,cat)
{
	//alert(subCat);
	//alert(cat);

	var comboBox = document.getElementById('CB_USUARIO');
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

				if (subCat=='EXTERNA') {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

					strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
					strSql = strSql & " AND U.PERFIL_EMP=0"



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


				if (subCat=='0') {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

					strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"


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

<%If Trim(strTipoProceso) = "CARGA_GESTIONES" Then%>
CargaTipoExporte('<%=strCliente%>','<%=strTipoProceso%>')
MostrarFilas('divGestionCarga');
<%Else%>
InicializaComboTipoExporte();
<%End If%>

</script>

</body>


