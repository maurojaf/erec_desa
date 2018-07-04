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

strCodCliente				=Request("CB_CLIENTE")
intTipoProceso 				=Request("CB_TIPOPROCESO")
dtmFechaProceso 			=Request("TX_FECHA_PROCESO")
archivo 					=Request("archivo")
intEtapaProceso				=Request("intEtapaProceso")
intIdUsario					=session("session_idusuario")
intTipoCargaDatosContacto   =Request("CB_TIPO_CARGA_DATOS_CONTACTO")
intProveedorFuente			=Request("CB_PROVEEDOR_CONTACTO")

If (intTipoProceso = 3 and intEtapaProceso = 2) or intTipoCargaDatosContacto = 2 then

	%>

	<script>
	
	alert("Proceso terminado exitosamente");
	location.href='Carga_Actualizacion_General.asp'
	
	</script>
	
	<%
	
End If

'Response.write "<br>intTipoProceso=" & intTipoProceso
'Response.write "<br>intEtapaProceso=" & intEtapaProceso

sFechaHoy = right("00"&Day(DATE()), 2) & "-" &right("00"&(Month(DATE())), 2) & "-" & Year(DATE())

If intProveedorFuente = "" then intProveedorFuente = 14 End If

If intTipoProceso <> 1 and intEtapaProceso = 1 then 

	strNomComboDinamico = "Fecha Pago"

ElseIf intTipoProceso = 1 or intEtapaProceso = 2 then 

	strNomComboDinamico = "Fecha Carga"

End If

	if intEtapaProceso = 1 then

		Dim DestinationPath
		DestinationPath = Server.mapPath("../Archivo/CargaActualizacion_20")  

		' crear una instancia
		set Obj_FSO = createobject("scripting.filesystemobject")

		Function TraeExtension(strArchivoFn)
			strArchivo 		= Mid(strArchivoFn,Len(strArchivoFn)-6,len(strArchivoFn))
			intPos 			= Instr(strArchivoFn,".")
			strExtension 	= Mid(strArchivoFn,intPos,len(strArchivoFn))
			TraeExtension 	= strExtension
		End Function
		
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
			End if

			''If Trim(strTipoCarga) = "DOC_BCI" Then
			''	Response.Redirect "Man_UploadDocBCI.asp?CB_CLIENTE=" + strCodCliente +"&strTipoCarga=" + strTipoCarga +"Fecha=" + Fecha +"&Asignacion=" + strAsignacion +"&archivo=" + archivo +"&opAc=" + iOpAc
			''End If

		End function

	End If

	if Request.ServerVariables("REQUEST_METHOD") = "POST" and intEtapaProceso = 1 then

		if intTipoProceso = 1 then
		
		   strNomTipoProceso = "CARGA"
		   
		ElseIf intTipoProceso = 2 then
		
			strNomTipoProceso = "CARGA_ACTUALIZACION"

		ElseIf intTipoProceso = 3 then
		
			strNomTipoProceso = "ACTUALIZACION"

		ElseIf intTipoProceso = 4 then
		
			strNomTipoProceso = "CARGA_CONTACTABILIDAD"
			
		End If

		response.write SaveFiles()

		Dim FSO, Fich , NombreAnterior, NombreNuevo

		NombreAnterior = archivo
		strExtension = TraeExtension(NombreAnterior)
		NombreNuevo = strCodCliente & "_" & strNomTipoProceso & strExtension

		'Response.write "<br>NombreAnterior=" & NombreAnterior
		'Response.write "<br>NombreNuevo=" & NombreNuevo

		' Instanciamos el objeto
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		' Asignamos el fichero a renombrar a la variable fich
		
		strRutaArchAntiguo = DestinationPath & "\" & NombreAnterior
		strRutaArchNuevo = DestinationPath & "\" & NombreNuevo

		Set Fich = FSO.GetFile(strRutaArchAntiguo)
		Call Fich.Copy(strRutaArchNuevo)

		if NombreAnterior <> NombreNuevo then 
			Call Fich.Delete()
		end if

		Set Fich = Nothing
		Set FSO = Nothing
		
		AbrirSCG()
		strSql = "EXEC proc_carga_data_carga_actualizacion "&TRIM(strCodCliente)&","&TRIM(intTipoProceso)
		'response.write strSql
		Conn.execute(strSql)
		CerrarSCG()	
		
		If intTipoProceso <> 4 then

		AbrirSCG()
		strSql = "EXEC proc_Validacion_Archivo_Carga_Actualizacion "&TRIM(strCodCliente)&","&TRIM(intTipoProceso)
		'response.write strSql
		Conn.execute(strSql)
		CerrarSCG()	
		
		End If
		
		AbrirSCG()
		strSql2 = "EXEC proc_Actualizacion_Data_Cuota "&TRIM(strCodCliente)
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
		
		If intTipoProceso = 4 then

		AbrirSCG()
		strSql2 = "EXEC proc_inserta_datos_contactabilidad "&TRIM(strCodCliente)&",0,1,'/'"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
		
		AbrirSCG()
		strSql2 = "EXEC proc_inserta_datos_contactabilidad "&TRIM(strCodCliente)&",0,2,'/'"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()
		
		AbrirSCG()
		strSql2 = "EXEC proc_Validacion_Archivo_Carga_Ubicabilidad_Masiva 1"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
		
		AbrirSCG()
		strSql2 = "EXEC proc_Validacion_Archivo_Carga_Ubicabilidad_Masiva 2"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()
		
		End If
		
	%>
	<script>
	
	alert("Archivo Subido con éxito");
	
	</script>
	<%
	
	ElseIf intEtapaProceso = 2 then
	
		AbrirSCG()
		strSql2 = "EXEC proc_Actualizacion_Estado_Deuda_Clientes "&TRIM(strCodCliente)&",'''"&TRIM(dtmFechaProceso)&"'''"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
	
	ElseIf intEtapaProceso = 3 then
	
		AbrirSCG()
		strSql2 = "EXEC proc_Carga_Deudor_Cuota "&TRIM(strCodCliente)&",'"&TRIM(dtmFechaProceso)&"',"&TRIM(intIdUsario)
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()		
		
		AbrirSCG()
		strSql2 = "EXEC proc_inserta_datos_contactabilidad "&TRIM(strCodCliente)&","&TRIM(intTipoCargaDatosContacto)&",1,'/'"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
		
		AbrirSCG()
		strSql2 = "EXEC proc_inserta_datos_contactabilidad "&TRIM(strCodCliente)&","&TRIM(intTipoCargaDatosContacto)&",2,'/'"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
		
		AbrirSCG()
		strSql2 = "EXEC proc_Validacion_Archivo_Carga_Ubicabilidad_Masiva 1"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
		
		AbrirSCG()
		strSql2 = "EXEC proc_Validacion_Archivo_Carga_Ubicabilidad_Masiva 2"
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
		
	ElseIf intEtapaProceso = 4 then
		
		AbrirSCG()
		strSql2 = "EXEC proc_Carga_Ubicabilidad 1,"&TRIM(intIdUsario)&","&intProveedorFuente
		'response.write strSql2
		'response.end
		Conn.execute(strSql2)
		CerrarSCG()	
		
		AbrirSCG()
		strSql2 = "EXEC proc_Carga_Ubicabilidad 2,"&TRIM(intIdUsario)&","&intProveedorFuente
		'response.write strSql2
		Conn.execute(strSql2)
		CerrarSCG()	
		
	%>

	<script>
	
	alert("Proceso terminado exitosamente");
	location.href='Carga_Actualizacion_General.asp'
	
	</script>
	
	<%
	
	End If
	
	'Response.write "<br>intEtapaProceso" = intEtapaProceso
	
	%>

<title>MODULO DE CARGAS</title>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" id="datos" method="POST" enctype="multipart/form-data">

<div class="titulo_informe">MÓDULO CARGA ACTUALIZACIÓN GENERAL DE DATOS</div>

<br>
	<table width="90%" align="center" class="">
		<tr>
		<td valign="top">
			<table width="100%" border="0" class="estilo_columnas">
					<thead>
					<tr>
						<td width="200">Cliente</td>
						<td width="200">Tipo Proceso</td>
						
						<% If intEtapaProceso = "" then %>
						<td colspan="2" width="620">Archivo de Carga</td>
						<% End If%>

						<% If (intTipoProceso = 1 and intEtapaProceso = 1) or (intTipoProceso = 2 and intEtapaProceso = 2) then %>						
						<td align="left"><%=strNomComboDinamico%></td>
						<% End If%>
						
						<% If (intTipoProceso = 2 and intEtapaProceso = 1) or (intTipoProceso = 3 and intEtapaProceso = 1) then %>						
						<td align="left" colspan="2"><%=strNomComboDinamico%></td>
						<% End If%>
						
						<% If (intTipoProceso = 1 and intEtapaProceso = 1) or (intTipoProceso = 2 and intEtapaProceso = 2) then %>						
						<td align="left" colspan="2">TIPO CARGA CONTACTOS</td>
						<% End If%>
		
						<% If  (intTipoProceso = 1 and intEtapaProceso = 3) or (intTipoProceso = 2 and intEtapaProceso = 3) or (intTipoProceso = 4 and intEtapaProceso = 1) then %>						
						<td colspan="2">PROVEEDOR CONTACTO</td>
						<% End If%>
						
					</tr>
					</thead>
					<td>
						<select name="CB_CLIENTE" id="CB_CLIENTE" onChange="CargaRegistros(CB_CLIENTE.value);">
							<option value="">SELECCIONAR</option>
							<option value=0>TODOS</option>
							<%
							AbrirSCG()
							ssql="SELECT COD_CLIENTE, NOM_CLIENTE = UPPER(NOMBRE_FANTASIA) FROM CLIENTE WHERE ACTIVO=1 AND COD_CLIENTE IN (1200,1500,1070,2000,2008,2010,2011,2013,2014) ORDER BY COD_CLIENTE"
							set rsTemp= Conn.execute(ssql)
							if not rsTemp.eof then
								do until rsTemp.eof%>
									<option value="<%=rsTemp("COD_CLIENTE")%>"<%if strCodCliente=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("NOM_CLIENTE")%></option>
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
						<select name="CB_TIPOPROCESO" ID="CB_TIPOPROCESO">
						</select>
					</td>
					
					<% If intEtapaProceso = "" then %>
					
					<td>
						<input name="File1" type="file" VALUE="<%=File1%>" size="80">    (máximo 100 MB)
					</td>
					
					<%End If%>
					
					<% If (intTipoProceso = 1 and intEtapaProceso = 1) or (intTipoProceso = 2 and intEtapaProceso = 2) or (intTipoProceso = 2 and intEtapaProceso = 1) or (intTipoProceso = 3 and intEtapaProceso = 1) then %>	
					
					<td>
						<input name="TX_FECHA_PROCESO" readonly="true" type="text" id="TX_FECHA_PROCESO" value="<%=sFechaHoy%>" size="10">
					</td>
					
					<%End If%>
					
					<% If (intTipoProceso = 1 and intEtapaProceso <> 3 ) or intEtapaProceso = 2 then %>	
					
					<td>
						<select name="CB_TIPO_CARGA_DATOS_CONTACTO">
							<option value="1" <%if Trim(intTipoCargaDatosContacto)="1" then response.Write("Selected") end if%>>NUEVA ASIGNACIÓN</option>
							<option value="0" <%if Trim(intTipoCargaDatosContacto)="0" then response.Write("Selected") end if%>>BASE COMPLETA </option>
							<option value="2" <%if Trim(intTipoCargaDatosContacto)="2" then response.Write("Selected") end if%>>NO CARGAR</option>
						</select>
					</td>
					
					<%End If%>
					
					<% If (intTipoProceso = 1 and intEtapaProceso = 3) or (intTipoProceso = 2 and intEtapaProceso = 3) or (intTipoProceso = 4 and intEtapaProceso = 1)  then %>
					
					<td>
						<select name="CB_PROVEEDOR_CONTACTO" id="CB_PROVEEDOR_CONTACTO">
							<%
							AbrirSCG()
							ssql="SELECT COD_FUENTE,NOM_FUENTE FROM FUENTE_UBICABILIDAD WHERE ESTADO=1"
							set rsTemp= Conn.execute(ssql)
							if not rsTemp.eof then
								do until rsTemp.eof%>
									<option value="<%=rsTemp("COD_FUENTE")%>"<%if intProveedorFuente=rsTemp("COD_FUENTE") then response.Write("Selected") End If%>><%=rsTemp("NOM_FUENTE")%></option>
								<%
								rsTemp.movenext
								loop
							end if
							rsTemp.close
							set rsTemp=nothing
							CerrarSCG()
							%>
							<option value="0">CONTACTABILIDAD GENERAL</option>
						</select>
					</td>
					
					<%End If%>
					
					<% If intEtapaProceso = "" then %>
					
					<td align="right">
						<input Name="SubmitButton" class="fondo_boton_100" Value="Procesar" Type="BUTTON" onClick="procesar();">
					</td>
				
					<% ElseIf intEtapaProceso = 1 and intTipoProceso <> 1 and intTipoProceso <> 4 then %>
					
					<td align="right">
						<input Name="SubmitButton" class="fondo_boton_100" Value="Actualizar" Type="BUTTON" onClick="actualizar();">
					</td>
					
					<% ElseIf (intTipoProceso = 1 and intEtapaProceso <> 3) or intEtapaProceso = 2 then %>
					
					<td align="right">
						<input Name="SubmitButton" class="fondo_boton_100" Value="Cargar Deuda" Type="BUTTON" onClick="cargarDeuda();">
					</td>

					
					<% ElseIf intEtapaProceso = 3 or intTipoProceso = 4 then %>
					
					<td align="right">
						<input Name="SubmitButton" class="fondo_boton_100" Value="Cargar Datos Contacto" Type="BUTTON" onClick="cargarDatosContacto();">
					</td>
					
					<%End If%>
					
				</tr>
				</FORM>
			</table>
			
			&nbsp;
			&nbsp;
			
			<table width="100%" border="0" class="estilo_columnas">
			
				<% If intEtapaProceso = "" then %>
				
				<thead>
					<tr>
						<td>Formatos</td>
					</tr>
				</thead>
				
					<table class="intercalado" style="width:100%;">
						<tbody>
							<tr >
								<td colspan="2" class="Estilo10" bgcolor="#C9DEF2">FORMATO DE ARCHIVOS DE CARGA             (CSV delimitado por comas)</td>
							</tr>
							<tr>
								<td width="20">UMA</td>
								<td width="50" align="left" ><a href="../Archivo/Formatos/CARGA_UMA.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">UPV</td>
								<td width="50" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">Contactabilidad General</td>
								<td width="50" align="left" ><a href="../Archivo/Formatos/Carga_Contactabilidad_General.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							
							<tr >
								<td colspan="2" class="Estilo10" bgcolor="#C9DEF2">FORMATO DE ARCHIVOS DE CARGA Y ACTUALIZACIÓN             (CSV delimitado por comas)</td>
							</tr>
							<tr>
								<td width="20">UMA FACTURAS</td>
								<td width="50" align="left" ><a href="../Archivo/Formatos/Carga_Actualizacion_UMA_Facturas.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">LEONISA</td>
								<td width="50" align="left" ><a href="../Archivo/Formatos/Carga_Actualizacion_Leonisa.txt" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td>CPECH</td>
								<td><a href="../Archivo/Formatos/Carga_Actualizacion_CPECH.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td>COLEGIOS TERRAUSTRAL</td>
								<td><a href="../Archivo/Formatos/Carga_Actualizacion_Terraustral.CSV" target='Contenido'>Descargar</a></td>
							</tr>

							<tr >
								<td colspan="2" class="Estilo10" bgcolor="#C9DEF2">FORMATO DE ARCHIVOS DE ACTUALIZACIÓN             (CSV delimitado por comas)</td>
							</tr>
							<tr>
								<td width="20">UMA</td>
								<td width="50" align="left" ><a href="../Archivo/Formatos/Actualizacion_UMA.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">UPV</td>
								<td width="50" align="left" ><a href="../Archivo/Formatos/Actualizacion_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
						</tbody>
					</table>
				
				<% ElseIf intEtapaProceso = 1 then %>
				
				<thead>
					<tr>
						<td>Informe</td>
					</tr>
				</thead>
				
					<table class="intercalado" style="width:100%;">
						<tbody>
							<tr>
								<td width="20">Errores</td>
								<td width="20"><%=intTotalRegistrosConError%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">Total Documentos a Actualizar</td>
								<td width="20"><%=intTotalRegistrosActualizacion%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/Carga_Contactabilidad_General.CSV" target='Contenido'>Descargar</a></td>
							</tr>
						</tbody>
					</table>

				<% ElseIf intEtapaProceso = 2 then %>
				
				<thead>
					<tr>
						<td>Informe</td>
					</tr>
				</thead>
				
					<table class="intercalado" style="width:100%;">
						<tbody>
							<tr>
								<td width="20">Total Deudores a Cargar</td>
								<td width="20"><%=intTotalDeudoresCarga%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">Total Documentos a Cargar</td>
								<td width="20"><%=intTotalDocumentosCarga%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/Carga_Contactabilidad_General.CSV" target='Contenido'>Descargar</a></td>
							</tr>
						</tbody>
					</table>
					
				<% ElseIf intEtapaProceso = 3 then %>
				
				<thead>
					<tr>
						<td>Informe</td>
					</tr>
				</thead>
				
					<table class="intercalado" style="width:100%;">
						<tbody>
							<tr >
								<td colspan="3" class="Estilo10" bgcolor="#C9DEF2">FONOS</td>
							</tr>
							<tr>
								<td width="20">Total Fonos a Cargar</td>
								<td width="20"><%=intTotalFonosCarga%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">Total Fonos Incorrectos</td>
								<td width="20"><%=intTotalFonosIncorrectos%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">Total Fonos Cargados</td>
								<td width="20"><%=intTotalFonosCargados%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>							
							<tr >
								<td colspan="3" class="Estilo10" bgcolor="#C9DEF2">EMAIL</td>
							</tr>
							<tr>
								<td width="20">Total email a Cargar</td>
								<td width="20"><%=intTotalemailCarga%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">Total email Incorrectos</td>
								<td width="20"><%=intTotalemailIncorrectos%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
							<tr>
								<td width="20">Total email Cargados</td>
								<td width="20"><%=intTotalemailCargados%></td>
								<td width="500" align="left" ><a href="../Archivo/Formatos/CARGA_UPV.CSV" target='Contenido'>Descargar</a></td>
							</tr>
						</tbody>
					</table>
					
				<% End If %>
					
				<tr class="totales">
					<td >&nbsp;</td>
				</tr>
				
			</table>			
			
			</td>
		</tr>
	</table>
<br/>

</form>
</body>
</html>

<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
<link href="../css/style_generales_sistema.css" rel="stylesheet">

<script language='javascript' src="../javascripts/popcalendar.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 


<script language="JavaScript1.2">

$(document).ready(function(){
	$.prettyLoader();
	$('#TX_FECHA_PROCESO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd-mm-yy'})	
})

function procesar(){

		var comboBox = document.getElementById('CB_TIPOPROCESO');

		if(datos.CB_CLIENTE.value ==''){
			alert('Debe seleccionar el cliente');
			return false;
		}else if(comboBox.value =='0'){
			alert('Debe seleccionar el tipo de proceso');
			return false;
		}else if(datos.File1.value ==''){
			alert("Debe seleccionar un archivo");
			return false;	
		}else{
		
			//alert(datos.CB_TIPOPROCESO.value )
			
			$.prettyLoader.show(200000);
			datos.action="Carga_Actualizacion_General.asp?intEtapaProceso=1&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value;
			datos.submit();		
		}
}

function actualizar(){
	
			//alert(datos.CB_TIPOPROCESO.value )
			
			$.prettyLoader.show(200000);
			datos.action="Carga_Actualizacion_General.asp?intEtapaProceso=2&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&TX_FECHA_PROCESO=" + document.datos.TX_FECHA_PROCESO.value;
			datos.submit();		
}

function cargarDeuda(){
	
			//alert(datos.CB_TIPOPROCESO.value )
			
			$.prettyLoader.show(200000);
			datos.action="Carga_Actualizacion_General.asp?intEtapaProceso=3&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&TX_FECHA_PROCESO=" + document.datos.TX_FECHA_PROCESO.value + "&CB_TIPO_CARGA_DATOS_CONTACTO=" + document.datos.CB_TIPO_CARGA_DATOS_CONTACTO.value;
			datos.submit();		
}

function cargarDatosContacto(){
	
			//alert(datos.CB_TIPOPROCESO.value )
			
			$.prettyLoader.show(200000);
			datos.action="Carga_Actualizacion_General.asp?intEtapaProceso=4&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&CB_PROVEEDOR_CONTACTO=" + document.datos.CB_PROVEEDOR_CONTACTO.value;
			datos.submit();		
}

function CargaRegistros(codCliente,tipoProceso)
{
	var comboBox = document.getElementById('CB_TIPOPROCESO');
	comboBox.options.length = 0;

		if (codCliente=='1200') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA', '1');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ACTUALIZACIÓN', '3');
			comboBox.options[comboBox.options.length] = newOption;			
		}
		else if (codCliente=='1500') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA', '1');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ACTUALIZACIÓN', '3');
			comboBox.options[comboBox.options.length] = newOption;	
		}
		else if (codCliente=='1070') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA', '1');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ACTUALIZACIÓN', '3');
			comboBox.options[comboBox.options.length] = newOption;	
		}
		else if (codCliente=='2000') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA Y ACTUALIZACIÓN', '2');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (codCliente=='2008') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA', '1');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('ACTUALIZACIÓN', '3');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (codCliente=='2010') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA Y ACTUALIZACIÓN', '2');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (codCliente=='2011') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA Y ACTUALIZACIÓN', '2');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (codCliente=='2013') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA Y ACTUALIZACIÓN', '2');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (codCliente=='0') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA CONTACTABILIDAD', '4');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else if (codCliente=='2014') {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('CARGA', '1');
			comboBox.options[comboBox.options.length] = newOption;
		}
		else {
			var newOption = new Option('SELECCIONAR', '0');
			comboBox.options[comboBox.options.length] = newOption;
		}
		ActualizaComboRegistro(comboBox,tipoProceso)
}

function InicializaInforme()
{
		var comboBox = document.getElementById('CB_TIPOPROCESO');
		comboBox.options.length = 0;
		var newOption = new Option('SELECCIONAR','0');
		comboBox.options[comboBox.options.length] = newOption;
}
function ActualizaComboRegistro(comboRegistro,tipoProceso)
{
		for (var i=0; i< comboRegistro.options.length; i ++)
		{
		if (comboRegistro.options[i].value == tipoProceso)
			comboRegistro.options[i].selected = true;
			}
}

<%If intTipoProceso = "" then%>
InicializaInforme()
<%End If%>

<%If intTipoProceso <> "" then%>
CargaRegistros('<%=strCodCliente%>','<%=intTipoProceso%>');
<%End If%>

</script>




