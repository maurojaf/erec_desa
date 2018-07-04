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
	<!--#include file="../lib/freeaspupload.asp" -->

	<LINK rel="stylesheet" TYPE="text/css" HREF="../css/style.css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
<%

Response.CodePage=65001
Response.charset ="utf-8"

'Stores only files with size less than MaxFileSize

if Request("CB_CLIENTE")<>"" then
	strCliente=Request("CB_CLIENTE")

End if

strTipoCarga 		=Request("CB_TIPOCARGA")
strTipoProceso 		=Request("CB_TIPOPROCESO")
dtmFechaCreacion 	=Request("TX_FECHACREACION")
archivo 			=Request("archivo")

sFechaHoy = right("00"&Day(DATE()), 2) & "/" &right("00"&(Month(DATE())), 2) & "/" & Year(DATE())

if Request("Fecha")<>"" then
	FechaR=Request("Fecha")
End if

if Request("Asignacion")<>"" then
	strAsignacion=Request("Asignacion")
End if

if Request("opAc")<>"" then
	iOpAc=Request("opAc")
End if


'Response.write "<br>strTipoCarga=" & strTipoCarga
'Response.write "<br>CB_TIPOCARGA=" & Request("CB_TIPOCARGA")
'Response.write "<br>CB_TIPOPROCESO=" & Request("CB_TIPOPROCESO")
'Response.write "<br>CB_CLIENTE=" & Request("CB_CLIENTE")

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


			AbrirSCG()
			strSql = "EXEC Proc_Audita_Archivo 1, 3, "&trim(session("session_idusuario"))&",null, '"&trim(strCliente)&"', '"&trim(archivo)&"', '',0 "
			response.write strSql
			Conn.execute(strSql)
			CerrarSCG()	

			
			next

		else
		End if


		''Response.End

			If Trim(strTipoCarga) = "DEUDOR" Then
		  		Response.Redirect "Man_UploadDeudor.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"&Fecha=" + Fecha +"&Asignacion=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "DEUDA" Then

				If Trim(strTipoProceso) = "CARGA" Then
					Response.Redirect "Man_UploadDeuda.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"Fecha=" + Fecha +"&Asignacion=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc +"&TX_FECHACREACION=" + dtmFechaCreacion
				End If
				If Trim(strTipoProceso) = "ACTUALIZACION" Then
					Response.Redirect "Man_UploadActDeuda.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"Fecha=" + Fecha +"&Asignacion=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc
				End If

		  	End If
		  	If Trim(strTipoCarga) = "UBICABILIDAD" Then
				Response.Redirect "Man_UploadUbicabilidad.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"Fecha=" + Fecha +"&Asignacion=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "UF" Then
				Response.Redirect "Man_UploadUF.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"Fecha=" + Fecha +"&Asignacion=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "GESTIONES" Then
				'Response.Redirect "Man_UploadGestiones.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"Fecha=" + Fecha +"&Asignacion=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If
		  	If Trim(strTipoCarga) = "DOC_BCI" Then
				Response.Redirect "Man_UploadDocBCI.asp?CB_CLIENTE=" + strCliente +"&strTipoCarga=" + strTipoCarga +"Fecha=" + Fecha +"&Asignacion=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc
		  	End If

	End function


%>

<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	response.write SaveFiles()




End if

'******************************
'*	INICIO CODIGO PARTICULAR  *
''******************************
%>

<title>MODULO DE CARGAS</title>


</head>
<%strTitulo="MI CARTERA"%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="titulo_informe">MODULO DE CARGAS - DEUDOR - DEUDA - UBICABILIDAD - UF</div>	
<br>
	<table width="90%" align="center" class="">
			<tr>
			  <td valign="top">
			  <form name="frmSend" id="frmSend" onSubmit="return enviar(this)"  method="POST" enctype="multipart/form-data" action="man_Carga.asp">
			  <INPUT TYPE="HIDDEN" NAME="FechaHoy" id="FechaHoy" value="<%=sFechaHoy%>">

			  <table width="100%" border="0" class="estilo_columnas">
			  	<thead>
				<tr>
					<td>Cliente</td>
					<td>Dato a Cargar</td>
					<td>Tipo Proceso</td>
					<td>Asignación (cartera)</td>
					<td>Fecha Creación</td>
				</tr>
				</thead>


				<td>
					<select name="CB_CLIENTE" onchange="Refrescar()">
						<option value="Seleccionar">Seleccionar</option>
						<%
						AbrirSCG()
						ssql="SELECT COD_CLIENTE, DESCRIPCION FROM CLIENTE ORDER BY COD_CLIENTE "
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
								<option value="<%=rsTemp("COD_CLIENTE")%>"<%if strCliente=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("COD_CLIENTE") & "-" & rsTemp("DESCRIPCION")%></option>
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
					<select name="CB_TIPOCARGA">
						<option value="Seleccionar" >Seleccionar</option>
						<option value="DEUDOR">DEUDOR</option>
						<option value="DEUDA">DEUDA</option>
						<option value="UBICABILIDAD">UBICABILIDAD</option>
						<!--<option value="GESTIONES">GESTIONES</option>-->
						<option value="UF">U.F.</option>
						<option value="DOC_BCI">DOCUMENTOS BCI</option>
					</select>
				</td>
				<td>
					<select name="CB_TIPOPROCESO">
						<option value="Seleccionar">Seleccionar</option>
						<option value="CARGA">CARGA</option>
					</select>
				</td>
				<td>
				<input name="Fecha" type="TEXT" VALUE="<%=sFechaHoy%>" size="14" maxlength="14">
				<select name="Asignacion" onchange="RefrescaPagina()" >
				<option value="Seleccionar">Seleccionar</option>
				<%

						AbrirSCG()
						ssql="select COD_REMESA, CONVERT(VARCHAR(10),FECHA_CARGA,103) AS FECHA_CARGA FROM REMESA WHERE COD_CLIENTE = '" & strCliente & "' AND COD_REMESA > 0 ORDER BY COD_REMESA DESC "
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
								<option value="<%=rsTemp("COD_REMESA")%>"><%=rsTemp("COD_REMESA")& "-" & rsTemp("fecha_carga")%></option>
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
				<input name="TX_FECHACREACION" type="TEXT" VALUE="<%=sFechaHoy%>" size="10" maxlength="10">
				</td>
			</tr>
			<tr class="estilo_columna_individual">
				<td colspan="5">
					Archivo de Carga
				</td>
	   		</tr>
	   		<tr>
	   		<td colspan="4">
	   			<input name="File1" type="file" VALUE="<%= File1%>" size="80">
	     	 </td>
	     	<td align="right">
				<input type="hidden" name="ckbAc" value="ckbAc">
				<input Name="SubmitButton" class="fondo_boton_100" Value="Cargar" Type="BUTTON" onClick="enviar();">
		     </td>
			</tr>

		</FORM>

		</table>


</td>
</tr>


<tr>

<td>

 <table class="intercalado" style="width:100%;">
 	<thead>
 	<tr >
		<td colspan="6">FORMATO DE ARCHIVOS DE CARGA (CSV (MS-DOS))</td>
	</tr>
	</thead>
	<tbody>
	<tr>
		<td>Deudor : </td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_DEUDOR.CSV" target='Contenido'>Descargar</a></td>
		<td>Deuda : </td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_DEUDA.CSV" target='Contenido'>Descargar</a></td>
		<td>Ubicabilidad:</td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_UBICABILIDAD.CSV" target='Contenido'>Descargar</a></td>
	</tr>
	<tr>

		<td>Gestiones : </td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_GESTIONES.CSV" target='Contenido'>Descargar</a></td>
		<td>U.F.: </td>
		<td><a href="../Archivo/Formatos/FORMATO_CARGA_UF.CSV" target='Contenido'>Descargar</a></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	</tbody>
</table>

<br>

<table width="100%" border="0" align="center">
<tr>
	<td style="vertical-align: top;">
		<table width="100%" class="intercalado" style="width:100%;">
		<thead>
		<tr >
			<td colspan=2><b>TIPO DOCUMENTO</b></td>
		</tr>
		<tr>
			<td><b>Cod</b></td>
			<td><b>Nombre</b></td>
		</tr>
		</thead>
		<tbody>
		<%
			AbrirSCG()
				ssql="SELECT * FROM TIPO_DOCUMENTO ORDER BY COD_TIPO_DOCUMENTO "
				set rsTemp= Conn.execute(ssql)
				if not rsTemp.eof then
					do until rsTemp.eof%>
					<tr>
							<td><%=rsTemp("COD_TIPO_DOCUMENTO")%></td>
							<td><%=rsTemp("NOM_TIPO_DOCUMENTO")%></td>
					</tr>

					<%
					rsTemp.movenext
					loop
				end if
				rsTemp.close
				set rsTemp=nothing
			CerrarSCG()

		%>
		</tbody>
		</table>
	</td>

	<td style="vertical-align: top;">
		<table width="100%" class="intercalado" style="width:100%;">
		<thead>
		<tr>
			<td colspan=2><b>ETAPA COBRANZA</b></td>
		</tr>
		<tr>
			<td><b>Cod</b></td>
			<td><b>Nombre</b></td>
		</tr>
		</thead>
		<tbody>
		<%
			AbrirSCG()
				ssql="SELECT * FROM ESTADO_COBRANZA ORDER BY COD_ESTADO_COBRANZA "
				set rsTemp= Conn.execute(ssql)
				if not rsTemp.eof then
					do until rsTemp.eof%>
					<tr>
							<td><%=rsTemp("COD_ESTADO_COBRANZA")%></td>
							<td><%=rsTemp("NOM_ESTADO_COBRANZA")%></td>
					</tr>

					<%
					rsTemp.movenext
					loop
				end if
				rsTemp.close
				set rsTemp=nothing
			CerrarSCG()

		%>
		</tbody>

		</table>
</td>
</tr>
</table>

</td>
</tr>

</table>

</body>
</html>
<script language="JavaScript1.2">

function RefrescaPagina() {

		 var texto;
    	 var indice = document.frmSend.Asignacion.selectedIndex;
		 var fechahoy = document.frmSend.FechaHoy.value;

		if(document.frmSend.Asignacion.value =='Seleccionar'){

			document.frmSend.Fecha.value = document.frmSend.FechaHoy.value;

		}else{
			document.frmSend.Fecha.value = document.frmSend.Asignacion.options[indice].text;
		}

}

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
		}else if(document.frmSend.Fecha.value ==''){
			alert('Debe ingresar la fecha de ingreso');
			return false;
		}else if(document.frmSend.TX_FECHACREACION.value ==''){
			alert('Debe ingresar la fecha de creación');
			return false;
		}else if(document.frmSend.File1.value ==''){
			alert('Debe ingresar la direccion del documento para cargarlo');
			return false;
		}else{
			if(document.frmSend.ckbAc.checked == true){
				chek = 1;
			}else{
				chek = 0;
			}

			frmSend.action = "man_Carga.asp?CB_CLIENTE=" + document.frmSend.CB_CLIENTE.value + "&CB_TIPOCARGA=" + document.frmSend.CB_TIPOCARGA.value + "&TX_FECHACREACION=" + document.frmSend.TX_FECHACREACION.value + "&CB_TIPOPROCESO=" + document.frmSend.CB_TIPOPROCESO.value + "&Fecha=" + document.frmSend.Fecha.value  + "&Asignacion=" + document.frmSend.Asignacion.value + "&opAc=" + chek;
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
			frmSend.action = "man_Carga.asp?CB_CLIENTE=" + document.frmSend.CB_CLIENTE.value + "&CB_TIPOCARGA=" + document.frmSend.CB_TIPOCARGA.value + "&TX_FECHACREACION=" + document.frmSend.TX_FECHACREACION.value + "&CB_TIPOPROCESO=" + document.frmSend.CB_TIPOPROCESO.value + "&Fecha=" + document.frmSend.Fecha.value  + "&Asignacion=" + document.frmSend.Asignacion.value;
			frmSend.submit();
		}
}
</script>




