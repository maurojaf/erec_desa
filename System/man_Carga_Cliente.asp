<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/freeASPUpload.asp" -->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">



<%
Response.CodePage=65001
Response.charset ="utf-8"

'Stores only files with size less than MaxFileSize

if Request("CB_CLIENTE")<>"" then
	strCliente=Request("CB_CLIENTE")
End if

strTipoCarga 	=Request("CB_TIPOCARGA")
strTipoProceso 	=Request("CB_TIPOPROCESO")
archivo 		=Request("archivo")

sFechaHoy 		=right("00"&Day(DATE()), 2) & "/" &right("00"&(Month(DATE())), 2) & "/" & Year(DATE())

if Request("opAc")<>"" then
	iOpAc=Request("opAc")
End if



	Dim DestinationPath
	DestinationPath = Server.mapPath("../Archivo/CargaActualizaciones") & "\" & strCliente  

	' crear una instancia
	set Obj_FSO = createobject("scripting.filesystemobject")

	If not Obj_FSO.FolderExists(Server.mapPath("../Archivo/CargaActualizaciones") & "\" & strCliente) = True Then ' verifica la existencia del archivo
		Obj_FSO.CreateFolder(Server.mapPath("../Archivo/CargaActualizaciones") & "\" & strCliente) 

	End if



	Dim uploadsDirVar
	uploadsDirVar = DestinationPath

	function SaveFiles
		Dim Upload, fileName, fileSize, ks, i, fileKey, resumen
		Set Upload = New FreeASPUpload
		Upload.Save(uploadsDirVar)
		If Err.Number <> 0 then Exit function
		SaveFiles = ""
		ks = Upload.UploadedFiles.keys
		if (UBound(ks) <> -1) then
			resumen = "<B>Archivos subidos:</B> "
			for each fileKey in Upload.UploadedFiles.keys
				resumen = resumen & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
			archivo = Upload.UploadedFiles(fileKey).FileName
			next

		else
		End if


		''Response.End

			If Trim(strTipoCarga) = "DEUDOR" Then
		  		Response.Redirect "Man_UploadDeudor.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If

		  	If Trim(strTipoCarga) = "DEUDA" Then

				If Trim(strTipoProceso) = "CARGA" Then
					Response.Redirect "Man_UploadDeuda.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
				End If
				If Trim(strTipoProceso) = "ACTUALIZACION" Then
					Response.Redirect "Man_UploadActDeuda.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
				End If
		  	End If

		  	If Trim(strTipoCarga) = "UBICABILIDAD" Then
				Response.Redirect "Man_UploadUbicabilidad.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "UF" Then
				Response.Redirect "Man_UploadUF.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "GESTIONES" Then
				Response.Redirect "Man_UploadGestiones.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "DOC_BCI" Then
				Response.Redirect "Man_UploadDocBCI.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "DOC_UMA" and Trim(strCliente) = 1070 Then
				Response.Redirect "Man_UploadDoc_UMA.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "ACTUALIZACION_DEUDA" and Trim(strCliente) = 1070 Then
				Response.Redirect "Man_Actualiza_UMA.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "ACTUALIZACION_DATOS" and Trim(strCliente) = 1070 Then
				Response.Redirect "Man_Actualiza_Datos_UMA.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "DOC_UMA" and Trim(strCliente) = 1500 Then
				Response.Redirect "Man_UploadFacturas_UMA.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "ACTUALIZACION_DEUDA" and Trim(strCliente) = 1500 Then
				Response.Redirect "Man_Actualiza_UMA_FACT.asp?CB_CLIENTE=" + strCliente +"&strTipoProceso=" + strTipoProceso +"&strTipoCarga=" + strTipoCarga +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
	End function


%>

<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	response.write SaveFiles()

	AbrirSCG()
	strSql = "EXEC Proc_Audita_Archivo 1, 3, "&trim(session("session_idusuario"))&",null, '"&trim(strCliente)&"', '"&trim(archivo)&"', '',0 "
	'response.write strSql
	Conn.execute(strSql)
	CerrarSCG()	

End if

'******************************
'*	INICIO CODIGO PARTICULAR  *
''******************************
%>

<title>MODULO DE CARGAS</title>


</head>
<%strTitulo="MI CARTERA"%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="titulo_informe">MODULO DE CARGA (DEUDA - UBICABILIDAD - GESTIONES) Y ACTUALIZACION DE DEUDA</div>	
<br>

	<table width="90%" border = "0" align="center">

			<tr>
			  <td valign="top">
			  <form name="frmSend" id="frmSend" onSubmit="return enviar(this)"  method="POST" enctype="multipart/form-data" action="man_Carga_Cliente.asp">
			  <INPUT TYPE=HIDDEN NAME="FechaHoy" id="FechaHoy" value="<%=sFechaHoy%>">

			  <table width="100%" border="0" class="estilo_columnas" >
			  	<thead>
				<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					<td>Cliente</td>
					<td>Tipo Proceso</td>
					<td>Detalle Proceso</td>
				</tr>
				</thead>
				<tr>
		  		<td>
					<select name="CB_CLIENTE" id="CB_CLIENTE">
						<option value="Seleccionar">SELECCIONE</option>
						<% If session("perfil_emp") = true then %>
							<option value="" <%if strCliente="" then response.Write("Selected") End If%>>Todos</option>
						<% End If %>
						<%
						AbrirSCG()
						ssql="SELECT COD_CLIENTE, DESCRIPCION FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") ORDER BY DESCRIPCION "
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
								<option value="<%=rsTemp("COD_CLIENTE")%>" <%if strCliente=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("COD_CLIENTE") & "-" & rsTemp("DESCRIPCION")%></option>
							<%
							rsTemp.movenext
							loop
						end if
						CerrarSCG()
						%>
					</select>
			
				</td>

				<td>
					<select name="CB_TIPOCARGA" id="CB_TIPOCARGA" onChange="CargaFechas(this.value);">
						<option value="Seleccionar" >SELECCIONE</option>
						<option value="DOC_UMA">CARGA DOCUMENTOS</option>
						<option value="UBICABILIDAD">CARGA UBICABILIDAD</option>
						<option value="GESTIONES">CARGA GESTIONES</option>
						<option value="ACTUALIZACION_DEUDA">ACTUALIZACION DEUDA</option>
						<option value="ACTUALIZACION_DATOS">ACTUALIZACION DATOS</option>
					</select>
				</td>
				<td>
					<select name="CB_TIPOPROCESO" id="CB_TIPOPROCESO">
					</select>
				</td>

			</tr>
			<tr bordercolor="#999999">
				<td colspan="5" class="estilo_columna_individual">
					Archivo de Carga
				</td>
	   		</tr>
	   		<tr>
	   		<td colspan=2>
	   			<input name="File1" type="file" VALUE="<%=File1%>" size="80">
	     	 </td>
	     	<td align="left">
				<input type="hidden" name="ckbAc" value="ckbAc">
				<input Name="SubmitButton" class="fondo_boton_100" Value="Procesar" Type="BUTTON" onClick="enviar();">
		     </td>

			</tr>

		</FORM>

		</table>


</td>
</tr>


<tr>

<td>

 <table width="100%" border="1" class="estilo_columnas">
 	<thead>
 	<tr bordercolor="#999999">
		<td colspan="10"><b>FORMATO DE ARCHIVOS DE CARGA (CSV (MS-DOS))</b></td>
	</tr>
	</thead>
	<tr>
		<td>Deuda : </td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_DEUDA_UMA.CSV" target='Contenido'>Descargar</a></td>
		<td>Facturas : </td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_DEUDA_UMA_FACT.CSV" target='Contenido'>Descargar</a></td>
		<td>Ubicabilidad:</td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_UBICABILIDAD.CSV" target='Contenido'>Descargar</a></td>
		<td>Gestiones : </td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_GESTIONES.CSV" target='Contenido'>Descargar</a></td>
		<td>Actualización Deuda : </td>
		<td><a href="../Archivo/Formatos/FORMATO_ACTUALIZACION_DEUDA_UMA.CSV" target='Contenido'>Descargar</a></td>
	</tr>

<br>


<script language="JavaScript1.2">

function enviar(){

		if(document.frmSend.CB_CLIENTE.value =='Seleccionar'){
			alert('Debe seleccionar el cliente');
			return false;
		}else if(document.frmSend.CB_TIPOCARGA.value =='Seleccionar'){
			alert('Debe seleccionar el tipo de carga');
			return false;
		}else if(document.frmSend.CB_TIPOPROCESO.value =='Seleccionar'){
			alert('Debe seleccionar el tipo de proceso');
			return false;
		}else if(document.frmSend.File1.value ==''){
			alert('Debe ingresar la direccion de la carpeta del documento para cargarlo');
			return false;
		}else{
			if(document.frmSend.ckbAc.checked == true){
				chek = 1;
			}else{
				chek = 0;
			}

			frmSend.action = "man_Carga_Cliente.asp?CB_CLIENTE=" + document.frmSend.CB_CLIENTE.value + "&CB_TIPOCARGA=" + document.frmSend.CB_TIPOCARGA.value + "&CB_TIPOPROCESO=" + document.frmSend.CB_TIPOPROCESO.value + "&opAc=" + chek;
			frmSend.submit();
		}
}

function Refrescar(){

		if(document.frmSend.CB_CLIENTE.value =='Seleccionar'){
			alert('Debe seleccionar el cliente');
			return false;
		}
		else
		{
			frmSend.action = "man_Carga_Cliente.asp?CB_CLIENTE=" + document.frmSend.CB_CLIENTE.value + "&CB_TIPOCARGA=" + document.frmSend.CB_TIPOCARGA.value + "&CB_TIPOPROCESO=" + document.frmSend.CB_TIPOPROCESO.value;
			frmSend.submit();
		}
}

function CargaFechas(subCat)
{
	var comboBox = document.getElementById('CB_TIPOPROCESO');

				comboBox.options.length = 0;

				if (subCat=='DOC_UMA') {

					var newOption = new Option('SELECCIONE', 'Seleccionar');comboBox.options[comboBox.options.length] = newOption;

					var newOption = new Option('CARGA INTERNA', 'CARGA_INTERNA');comboBox.options[comboBox.options.length] = newOption;

					var newOption = new Option('CARGA EXTERNA', 'CARGA_EXTERNA');comboBox.options[comboBox.options.length] = newOption;


				}
				else if (subCat=='ACTUALIZACION_DEUDA') {
					var newOption = new Option('SELECCIONE', 'Seleccionar');
					comboBox.options[comboBox.options.length] = newOption;

					var newOption = new Option('DEUDA', 'ACTUA_DEUDA');comboBox.options[comboBox.options.length] = newOption;


				}

				else {
					var newOption = new Option('SELECCIONE', 'Seleccionar');comboBox.options[comboBox.options.length] = newOption;

					var newOption = new Option('TODOS', '01/01/1900');comboBox.options[comboBox.options.length] = newOption;

				}

}

function InicializaComboTipoProceso()
{
		var comboBox = document.getElementById('CB_TIPOPROCESO');
		comboBox.options.length = 0;
		var newOption = new Option('SELECCIONE');
		comboBox.options[comboBox.options.length] = newOption;
}

InicializaComboTipoProceso();

</script>

</body>
</html>


