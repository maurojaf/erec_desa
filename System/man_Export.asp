<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="sesion.asp"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/freeaspupload.asp" -->
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

</head>


<%

Response.CodePage=65001
Response.charset ="utf-8"

strCliente 				=Trim(Request("CB_CLIENTE"))
intUsuario 				=Trim(request("CB_USUARIO"))
strCobranza 			=Trim(Request("CB_CUSTODIO"))
strTipoProceso 			=Trim(Request("CB_TIPOPROCESO"))
strFechaArchivoGestion 	=Trim(Request("CB_GES_CARGA"))
inicioFecEstado 		=Trim(Request("inicioFecEstado"))
terminoFecEstado 		=Trim(Request("terminoFecEstado"))
inicioFecGest			=Trim(Request("inicioFecGest"))
terminoFecGest 			=Trim(Request("terminoFecGest"))

rut_especifico 			=Trim(Request("rut_especifico"))
opcion_rut 				=Trim(Request("opcion_rut"))
rut_especifico_masivo 	=Trim(Request("rut_especifico_masivo"))

inicio 					=trim(request("inicio"))
termino 				=trim(request("termino"))
'response.write  rut_especifico&"<br>"&sin_guion&"<br>"&sin_dv&"<br>"&sin_dv&"<br>"&rut_especifico_masivo


if Request("archivo")<>"" then
	archivo=Request("archivo")
End if

abrirscg()

cerrarscg()


	Dim DestinationPath
	DestinationPath = Server.mapPath("../Archivo/UploadFolder")

	Dim uploadsDirVar
	uploadsDirVar = DestinationPath

	function SaveFiles

			inicio 	="NO"
			termino ="NO"
			If Trim(strTipoProceso) = "DEUDA" Then
		  		Response.Redirect "exp_Deuda.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&Fecha=" + Fecha +"&CB_CUSTODIO=" + strCobranza + "&CB_USUARIO=" + intUsuario + "&dtmInicio=" + inicioFecEstado + "&dtmTermino=" + terminoFecEstado + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_ACTIVO=" + strActivo + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo

		  	End If
		  	If Trim(strTipoProceso) = "DEUDA_AGRUP" Then
				Response.Redirect "exp_DeudaAgrupada.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&Fecha=" + Fecha +"&CB_CUSTODIO=" + strCobranza + "&CB_ASIGNACION=" + strAsignacion + "&CB_USUARIO=" + intUsuario + "&archivo=" + archivo +"&opAc=" + iOpAc + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo

		  	End If
		  	If Trim(strTipoProceso) = "GESTIONES" Then
				Response.Redirect "exp_Gestiones.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha + "&CB_CUSTODIO=" + strCobranza + "&CB_USUARIO=" + intUsuario + "&dtmInicio=" + inicioFecGest + "&dtmTermino=" + terminoFecGest + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_EFECTIVA=" + CH_EFECTIVA +"&CH_ACTIVO=" + strActivo + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo 

		  	End If
		  	If Trim(strTipoProceso) = "TELEFONOS" Then
				Response.Redirect "exp_Telefonos.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha +"&CB_ASIGNACION=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_ACTIVO=" + strActivo + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo

		  	End If
		  	If Trim(strTipoProceso) = "DIRECCIONES" Then
				Response.Redirect "exp_Direcciones.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha +"&CB_ASIGNACION=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_ACTIVO=" + strActivo + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo

		  	End If
		  	If Trim(strTipoProceso) = "EMAIL" Then
				Response.Redirect "exp_Email.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha +"&CB_ASIGNACION=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_ACTIVO=" + strActivo + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo

		  	End If
		  	If Trim(strTipoProceso) = "BASE_ESTADO" and strCliente = 1070 Then
				Response.Redirect "exp_Base_estado_UMA.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha + "&CB_USUARIO=" + intUsuario + "&dtmInicio=" + inicioFecEstado + "&dtmTermino=" + terminoFecEstado + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_EFECTIVA=" + CH_EFECTIVA +"&CH_ACTIVO=" + strActivo + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo

		  	End If
		  	If Trim(strTipoProceso) = "BASE_ESTADO" and strCliente = 1500 Then
				Response.Redirect "exp_Base_estado_UMA_Fact.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha + "&CB_USUARIO=" + intUsuario + "&dtmInicio=" + inicioFecEstado + "&dtmTermino=" + terminoFecEstado + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_EFECTIVA=" + CH_EFECTIVA +"&CH_ACTIVO=" + strActivo + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo

		  	End If
			If Trim(strTipoProceso) = "BASE_ESTADO" and strCliente = 1200 Then
				Response.Redirect "exp_Base_estado_upv.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"Fecha=" + Fecha + "&CB_USUARIO=" + intUsuario + "&dtmTermino=" + dtmTermino + "&dtmInicio=" + dtmInicio + "&archivo=" + archivo +"&opAc=" + iOpAc +"&CH_EFECTIVA=" + CH_EFECTIVA +"&CH_ACTIVO=" + strActivo + "&rut_especifico=" +rut_especifico+ "&opcion_rut=" +opcion_rut+ "&rut_especifico_masivo=" +rut_especifico_masivo

		  	End If

	End function

if Request.ServerVariables("REQUEST_METHOD") = "POST" and archivo = "1" then
	response.write SaveFiles()
End if

if Request.ServerVariables("REQUEST_METHOD") = "POST" and archivo = "100" then
	response.write DownloadFile("../Archivo/Otros/1100/Gestiones/"&trim(strFechaArchivoGestion))
    response.Flush
	response.end
End if


%>
<title>MODULO DE EXPORTACION DE DATOS</title>

<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
</head>
<%strTitulo="MI CARTERA"%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<div class="titulo_informe">MODULO DE EXPORTES</div>
<br>
<form name="datos" id="datos" onSubmit="return enviar(this)"  method="POST" action="man_Export.asp">
<input type="hidden" name="inicio" 	id="inicio" 	value="<%=inicio%>">
<input type="hidden" name="termino" id="termino" 	value="<%=termino%>">

	<table width="90%" border="0" ALIGN="CENTER" class="estilo_columnas">
		<thead>
		<tr height="25">
			<td width="33%">Cliente:</td>
			<td width="77%">Tipo Exporte:</td>
		</tr>
		</thead>
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
						<select name="CB_TIPOPROCESO" id="CB_TIPOPROCESO" onChange="cajas();">
						</select>
					</td>
			</tr>
	
		<tr>
		<td colspan="2">
		<div name="divCustodio" id="divCustodio" style="display:none" >
		<table width="100%"border="0" ALIGN="CENTER">
			<tr>
				<td colspan = "2" height="25" class="estilo_columna_individual">Filtros:</td>
			</tr>
			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="325">Tipo rut :</td>
				<td>
					<input type="checkbox" name="rut_especifico" id="rut_especifico" value="1">Filtrar rut específico&nbsp;&nbsp;&nbsp;
					<input type="radio" name="opcion_rut" id="opcion_rut" value="1">Sin guión&nbsp;&nbsp;&nbsp; 
					<input type="radio" name="opcion_rut" id="opcion_rut" value="2">Sin dígito verificador&nbsp;&nbsp;&nbsp;
					<img src="../Imagenes/limpia_campo.png" style="cursor:pointer" alt="Limpia campos" width="18"  height="18" onclick="bt_limpia_campos()">
				</td>
			</tr>
			<tr id="textarea_rut_especifico" height="50" BGCOLOR="#EEEEEE">
				<td width="325"></td>
				<td>
					<textarea rows="5" id="rut_especifico_masivo" name="rut_especifico_masivo" cols="20"></textarea>
				</td>
			</tr>


			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="325">Tipo Cobranza :</td>
				<td>
					<select name="CB_CUSTODIO" id="CB_CUSTODIO" onChange="CargaUsuarios(this.value,CB_CLIENTE.value);">
					</select>
				</td>
			</tr>
		</table>
	</div>
		
		</td>
		</tr>
		
		<tr>
		<td colspan="2">
		<div name="divGestionCarga" id="divGestionCarga" style="display:none;width:100%"  >
		<table width="100%" border="0" ALIGN="CENTER">
			<tr>
				<td colspan = "2" height="25" class="estilo_columna_individual">Filtros:</td>
			</tr>
			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="325">Fecha Archivo :</td>
				<td>
					<select name="CB_GES_CARGA" id="CB_GES_CARGA" style="width:200px;">
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

	</div></td>
		</tr>
		
		<tr>
		<td colspan="2">
		<div name="divEjecutivo" id="divEjecutivo" style="display:none;width:100%" >

		<table width="100%" border="0" ALIGN="CENTER">
			<tr height="50" BGCOLOR="#EEEEEE">

				<td width="325">Ejecutivo :</td>
				<td>
					<select name="CB_USUARIO" id="CB_USUARIO">
					</select>
				</td>
			</tr>
	</table>
	</div></td>
		</tr>
		
		<tr>
		<td colspan="2">
		<div name="divFecGestion" id="divFecGestion" style="display:none;width:100%"  >
		<table width="100%" border="0" ALIGN="CENTER">
			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="325">Fecha Ingreso Gestión:</td>
				<td>
					<b>Desde :&nbsp;<input name="inicioFecGest" id="inicioFecGest" type="text" value="<%= "01/" & Mid(date(),4,10)%>" size="10" maxlength="10">
					Hasta :&nbsp;<input name="terminoFecGest" id="terminoFecGest" type="text" value="<%=date()%>" size="10" maxlength="10">
					</b>
				</td>
			</tr>
		</table>
	</div>
	</td>
		</tr>
		<tr>
		<td colspan="2">
		<div name="divFecEstado" id="divFecEstado" style="display:none;width:100%"  >
		<table width="100%" border="0" ALIGN="CENTER">
			<tr>
				<td colspan = "2" height="25" class="estilo_columna_individual">Si desea añadir documentos regularizados al exporte seleccione fechas, de lo contrario solo saldrán documentos activos. </td>
			</tr>

			<tr height="50" BGCOLOR="#EEEEEE">
				<td width="152">Fecha regularización:</td>
				<td>
					<b>Desde :&nbsp;<input name="inicioFecEstado" type="text" id="inicioFecEstado" size="10" maxlength="10">
					Hasta :&nbsp;<input name="terminoFecEstado" type="text" id="terminoFecEstado"  size="10" maxlength="10">
					</b>
				</td>
			</tr>
		</table>
	</div>
		</td>
		</tr>
		
	</table>


	
	
	
	

	

	

	

	<!--div name="divExportar" id="divExportar" style="display:none" -->
	
	<table width="90%" border="0" ALIGN="CENTER">
				<tr height="50" BGCOLOR="#EEEEEE">
					<td widht="500" align="RIGHT">
						<input Name="btProcesar" class="fondo_boton_100" Value="Exportar" Type="BUTTON" onClick="enviar();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						
					</td>
				</tr>
		</table>
	

</FORM>

<script type="text/javascript">
	$(document).ready(function(){

		$('#textarea_rut_especifico').css('display','none')

		$('input[id="rut_especifico"]').click(function(){
			if($(this).is(':checked')){
				$('#textarea_rut_especifico').css('display','block')
			}else{
				$('#textarea_rut_especifico').css('display','none')
				$('#rut_especifico_masivo').val("")
			}
		})

		$('#inicioFecEstado').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#terminoFecEstado').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

		$('#inicioFecGest').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#terminoFecGest').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})


	})

	function bt_limpia_campos(){
		$('input[id="opcion_rut"]').removeAttr("checked")
	}
</script>
<script language="JavaScript1.2">


function MostrarFilas(Fila) 
{
//alert(Fila);
var elementos = document.getElementsByName(Fila);

	for (i = 0; i< elementos.length; i++) {
		if(navigator.appName.indexOf("Microsoft") > -1){
			//   alert('asdas');
			   var visible = 'block'
		} else {
		 //alert('asdassss')
			   var visible = 'block' //'table-row';
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


