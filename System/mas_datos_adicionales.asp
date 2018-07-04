<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/freeaspupload.asp" -->
	<link rel="stylesheet" href="../css/style_generales_sistema.css">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	NRO_DOC 		= request("NRO_DOC")
	cliente 		= request("cliente")
	strRutDeudor 	= request("strRUT_DEUDOR")
	ID_CUOTA 		= request("ID_CUOTA")

	strNroDoc 		=Request("strNroDoc")
	strNroCuota 	=Request("strNroCuota")
	strSucursal 	=Request("strSucursal")
	strCodRemesa 	=Request("strCodRemesa")
	ruta 			=Request("ruta")
	AbrirSCG()

    archivo 		= Request("archivo")
	
	
	Dim DestinationPath

	DestinationPath = Server.mapPath("../Archivo/BibliotecaImagenesCuota") & "\" & cliente
	'Response.write "<br>DestinationPath=" & DestinationPath

	Set Obj_FSO = createobject("scripting.filesystemobject")

	If not Obj_FSO.FolderExists(Server.mapPath("../Archivo/BibliotecaImagenesCuota") & "\" & cliente) = True Then ' verifica la existencia del archivo
		Obj_FSO.CreateFolder(Server.mapPath("../Archivo/BibliotecaImagenesCuota") & "\" & cliente) 
	End if

	Function TraeExtension(strArchivoFn)
		strArchivo 		= Mid(strArchivoFn,Len(strArchivoFn)-6,len(strArchivoFn))
		intPos 			= Instr(strArchivoFn,".")
		strExtension 	= Mid(strArchivoFn,intPos,len(strArchivoFn))
		TraeExtension 	= strExtension
	End Function

	If Trim(archivo) = "1" Then

		'Response.write "<br>DestinationPath=" & DestinationPath
		Dim uploadsDirVar
		uploadsDirVar = DestinationPath

		function SaveFiles

			Dim Upload, fileName, fileSize, ks, i, fileKey, resumen
			Set Upload = New FreeASPUpload
			Upload.Save(uploadsDirVar)
			If Err.Number <> 0 then Exit function
			SaveFiles = ""
			ks = Upload.UploadedFiles.keys
			If (UBound(ks) <> -1) Then
				resumen = "<B>Archivos subidos:</B> "
				for each fileKey in Upload.UploadedFiles.keys
					resumen = resumen & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
				archivo = Upload.UploadedFiles(fileKey).FileName
				next
			Else
			End if
		End function



		If Request.ServerVariables("REQUEST_METHOD") = "POST"  and archivo = "1" then
			Response.write SaveFiles()

			'response.write "ok"
			'Previamente el fichero  Anterior.txt
			'ha de existir en nuestra carpeta.

			'Declaracion de variables
			Dim FSO, Fich , NombreAnterior, NombreNuevo
			'Inicialización

			'Response.write "<br>archivo=" & archivo
			NombreAnterior = archivo
			strExtension = TraeExtension(NombreAnterior)
			NombreNuevo = ID_CUOTA & strExtension

			'Response.write "<br>NombreAnterior=" & NombreAnterior
			'Response.write "<br>NombreNuevo=" & NombreNuevo

			' Instanciamos el objeto
			Set FSO = Server.CreateObject("Scripting.FileSystemObject")
			' Asignamos el fichero a renombrar a la variable fich

			''DestinationPath = Server.mapPath("UploadFolder") & "\" & IntId  & "\" & strRutDeudor

			strRutaArchAntiguo = DestinationPath & "\" & NombreAnterior
			strRutaArchNuevo = DestinationPath & "\" & NombreNuevo

			'Response.write "<br>strRutaArchAntiguo=" & strRutaArchAntiguo
			'Set Fich = FSO.GetFile(Server.MapPath("\" & NombreAnterior))
			Set Fich = FSO.GetFile(strRutaArchAntiguo)
			' llamamos a la funcion copiar,
			'y duplicamos el archivo pero con otro nombre
			''Call Fich.Copy(Server.MapPath("\" & NombreNuevo))
			Call Fich.Copy(strRutaArchNuevo)

			if NombreAnterior <> NombreNuevo then 
				' finalmente borramos el fichero original
				Call Fich.Delete()
			end if

			Set Fich = Nothing
			Set FSO = Nothing

			''strSql = "UPDATE CUOTA SET IMAGEN_DOC = '" & archivo & "' WHERE ID_CUOTA = " & ID_CUOTA
			strSql = "UPDATE CUOTA SET IMAGEN_DOC = '" & NombreNuevo & "' WHERE ID_CUOTA = " & ID_CUOTA
			'Response.write "<br>strSql=" & strSql
			'Response.write "<br>archivo=" & archivo
			''response.end
			Conn.execute(strSql)
 			


		End if

	End if

	if Request.ServerVariables("REQUEST_METHOD") = "POST" and archivo = "100" then

		response.write DownloadFile(ruta)
		response.write ruta

	End if	



strSql="SELECT ISNULL(ADIC_1,'NO INGRESADO') AS ADIC_1, ISNULL(ADIC_2,'NO INGRESADO') AS ADIC_2, ISNULL(ADIC_3,'NO INGRESADO') AS ADIC_3, ISNULL(NRO_CLIENTE_DOC,'NO INGRESADO') AS NRO_CLIENTE_DOC, ISNULL(ADIC_4,'NO INGRESADO') AS ADIC_4, ISNULL(ADIC_5,'NO INGRESADO') AS ADIC_5, ISNULL(ADIC_91,'NO INGRESADO') AS ADIC_91, ISNULL(ADIC_92,'NO INGRESADO') AS ADIC_92, ISNULL(ADIC_93,'NO INGRESADO') AS ADIC_93, ISNULL(ADIC_94,'NO INGRESADO') AS ADIC_94, ISNULL(ADIC_95,'NO INGRESADO') AS ADIC_95, ISNULL(ADIC_96,'NO INGRESADO') AS ADIC_96, ISNULL(ADIC_97,'NO INGRESADO') AS ADIC_97, ISNULL(ADIC_98,'NO INGRESADO') AS ADIC_98, ISNULL(ADIC_99,'NO INGRESADO') AS ADIC_99, ISNULL(ADIC_100,'NO INGRESADO') AS ADIC_100, FECHA_CREACION, (CASE WHEN ISNULL(COD_GESTION_EXTERNA,0)='' THEN 0 ELSE ISNULL(COD_GESTION_EXTERNA,0) END) AS COD_GESTION_EXTERNA, (CASE WHEN ISNULL(DES_GESTION_EXTERNA,'NO REGISTRA')='' THEN 'NO REGISTRA' ELSE ISNULL(DES_GESTION_EXTERNA,'NO REGISTRA') END) AS DES_GESTION_EXTERNA, IsNull(IMAGEN_DOC,'') as IMAGEN_DOC, ISNULL(CONVERT(VARCHAR(15),FECHA_CONSULTA_NORM,103),'NO REGISTRA') AS FECHA_CONSULTA_NORM FROM CUOTA WHERE ID_CUOTA = " & ID_CUOTA
'response.write "strSql=" & strSql
'Response.End
set rsDET=Conn.execute(strSql)
strPatentes = ""
if Not rsDET.eof Then
	strAdic1 = rsDET("ADIC_1")
	strAdic2 = rsDET("ADIC_2")
	strAdic3 = rsDET("ADIC_3")
	strNroSap = rsDET("NRO_CLIENTE_DOC")
	strAdic4 = rsDET("ADIC_4")
	strAdic5 = rsDET("ADIC_5")
	strAdic91 = rsDET("ADIC_91")
	strAdic92 = rsDET("ADIC_92")
	strAdic93 = rsDET("ADIC_93")
	strAdic94 = rsDET("ADIC_94")
	strAdic95 = rsDET("ADIC_95")
	strAdic96 = rsDET("ADIC_96")
	strAdic97 = rsDET("ADIC_97")
	strAdic98 = rsDET("ADIC_98")
	strAdic99 = rsDET("ADIC_99")
	strAdic100 = rsDET("ADIC_100")
	strFechaCreacion = rsDET("FECHA_CREACION")
	strCodGestionExt = rsDET("COD_GESTION_EXTERNA")
	strDesGestionExt = rsDET("DES_GESTION_EXTERNA")
	strImagenDoc = rsDET("IMAGEN_DOC")
	strFechaConsultaNorm= rsDET("FECHA_CONSULTA_NORM")

	strArchivoImagenes="../Archivo/BibliotecaImagenesCuota/"&cliente&"/"&trim(strImagenDoc)

End If
'response.write "<br>"&strArchivoImagenes
strSql="SELECT IsNull(ADIC_1,'ADIC_1') as ADIC_1, IsNull(NRO_CLIENTE_DOC,'NRO_CLIENTE_DOC') as NRO_CLIENTE_DOC, IsNull(ADIC_2,'ADIC_2') as ADIC_2, IsNull(ADIC_3,'ADIC_3') as ADIC_3, IsNull(ADIC_4,'ADIC_4') as ADIC_4, IsNull(ADIC_5,'ADIC_5') as ADIC_5, IsNull(ADIC_91,'ADIC_91') as ADIC_91, IsNull(ADIC_92,'ADIC_92') as ADIC_92, IsNull(ADIC_93,'ADIC_93') as ADIC_93, IsNull(ADIC_94,'ADIC_94') as ADIC_94, IsNull(ADIC_95,'ADIC_95') as ADIC_95, IsNull(ADIC_96,'ADIC_96') as ADIC_96, IsNull(ADIC_97,'ADIC_97') as ADIC_97, IsNull(ADIC_98,'ADIC_98') as ADIC_98, IsNull(ADIC_99,'ADIC_99') as ADIC_99, IsNull(ADIC_100,'ADIC_100') as ADIC_100, IsNull(COD_ULT_GES,'COD_ULT_GES') as COD_ULT_GES, IsNull(OBS_ULT_GES,'OBS_ULT_GES') as OBS_ULT_GES FROM CLIENTE WHERE COD_CLIENTE = '" & cliente & "'"
'response.write "strSql=" & strSql
'Response.End
set rsDET=Conn.execute(strSql)
if Not rsDET.eof Then
	strNombreAdic1 = rsDET("ADIC_1")
	strNombreAdic2 = rsDET("ADIC_2")
	strNombreAdic3 = rsDET("ADIC_3")
	strNombreAdic4 = rsDET("ADIC_4")
	strNombreAdic5 = rsDET("ADIC_5")
	strNombreAdic91 = rsDET("ADIC_91")
	strNombreAdic92 = rsDET("ADIC_92")
	strNombreAdic93 = rsDET("ADIC_93")
	strNombreAdic94 = rsDET("ADIC_94")
	strNombreAdic95 = rsDET("ADIC_95")
	strNombreAdic96 = rsDET("ADIC_96")
	strNombreAdic97 = rsDET("ADIC_97")
	strNombreAdic98 = rsDET("ADIC_98")
	strNombreAdic99 = rsDET("ADIC_99")
	strNombreAdic100 = rsDET("ADIC_100")
	strNombreCOD_ULT_GES= rsDET("COD_ULT_GES")
	strNombreOBS_ULT_GES= rsDET("OBS_ULT_GES")
	strNombreNroSap= rsDET("NRO_CLIENTE_DOC")
End If

%>
<title>MAS DATOS</title>
</head>
<body>
<div class="titulo_informe">OTROS DETALLE DE DEUDA </div>
<br>
<table width="380" height="167" border="1" bordercolor="#FFFFFF" class="intercalado" align="center">
	<thead>
  	<tr>
    	<td height="21" colspan="4">
			RUT DEUDOR : <%=strRutDeudor%>, DOC : <%=strNroDoc%>, CUOTA : <%=strNroCuota%>
    	</td>
  	</tr>
	</thead>
  	<tbody>
	<% If Trim(strNombreAdic1) <> "ADIC_1" Then %>
  	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic1%></td><td><%=strAdic1%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic2) <> "ADIC_2" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic2%></td><td><%=strAdic2%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic3) <> "ADIC_3" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic3%></td><td><%=strAdic3%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreNroSap) <> "NRO_CLIENTE_DOC" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreNroSap%></td><td><%=strNroSap%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic4) <> "ADIC_4" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic4%></td><td><%=strAdic4%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic5) <> "ADIC_5" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic5%></td><td><%=strAdic5%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic91) <> "ADIC_91" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic91%></td><td><%=strAdic91%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic92) <> "ADIC_92" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic92%></td><td><%=strAdic92%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic93) <> "ADIC_93" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic93%></td><td><%=strAdic93%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic94) <> "ADIC_94" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic94%></td><td><%=strAdic94%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic95) <> "ADIC_95" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic95%></td><td><%=strAdic95%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic96) <> "ADIC_96" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic96%></td><td><%=strAdic96%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic97) <> "ADIC_97" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic97%></td><td><%=strAdic97%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic98) <> "ADIC_98" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic98%></td><td><%=strAdic98%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic99) <> "ADIC_99" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic99%></td><td><%=strAdic99%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreAdic100) <> "ADIC_100" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreAdic100%></td><td><%=strAdic100%></td>
	</tr>
	<% End If%>

	<tr  height="17" bordercolor="#999999">
		<td>FECHA CREACIÓN</td><td><%=strFechaCreacion%></td>
	</tr>

	<tr  height="17" bordercolor="#999999">
		<td>FECHA CONSULTA NORM</td><td><%=strFechaConsultaNorm%></td>
	</tr>

	<% If Trim(strNombreCOD_ULT_GES) <> "COD_ULT_GES" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreCOD_ULT_GES%><td><%=strCodGestionExt%></td>
	</tr>
	<% End If%>

	<% If Trim(strNombreOBS_ULT_GES) <> "OBS_ULT_GES" Then %>
	<tr  height="17" bordercolor="#999999">
		<td><%=strNombreOBS_ULT_GES%><td><%=strDesGestionExt%></td>
	</tr>
	<% End If%>
	</tbody>
	<thead>
	<tr  height="17" bordercolor="#999999">
		<td>IMAGEN DOCUMENTO</td>
		<td align="left">
		<%If Trim(strImagenDoc) = "" Then%>
			NO DISPONIBLE
		<%Else%>
		<!--A HREF="#" onClick="ventanaFactura('../ImagenesDoc/<%=strImagenDoc%>')";-->
		<A HREF="#" onClick="ventanaFactura('<%=strArchivoImagenes%>')";>
		<img src="../imagenes/descargar_pdf.png" width=84, height=33 border="0">
		</A>
		<%End If%>
		</td>
	</tr>
	</thead>
	<FORM name="frmSend" id="frmSend" onSubmit="return enviar(this)"  method="POST" enctype="multipart/form-data" action="biblioteca_deudores.asp">

		<tr BGCOLOR="#FFFFFF">
			<td height="45" colspan=2>
				Archivo:
			<input name="File1" type="file" VALUE="<%= File1%>" size="15" maxlength="15">
			<input Name="SubmitButton" class="fondo_boton_100" Value="Cargar" Type="BUTTON" onClick="enviar();">
			</td>
		</tr>
		</table>

	</FORM>


  <% CerrarSCG() %>
</table>


</body>
</html>
<script>
function ventanaFactura (URL){
	frmSend.action = "mas_datos_adicionales.asp?ruta="+URL+"&archivo=100&ID_CUOTA=<%=ID_CUOTA%>&strRUT_DEUDOR=<%=strRutDeudor%>&cliente=<%=cliente%>";
	frmSend.submit();
}

function enviar(){
	if (frmSend.File1.value==''){
		alert("Debe seleccionar un archivo");
		frmSend.File1.focus();
		return;
	}
	frmSend.action = "mas_datos_adicionales.asp?archivo=1&ID_CUOTA=<%=ID_CUOTA%>&strRUT_DEUDOR=<%=strRutDeudor%>&cliente=<%=cliente%>";
	frmSend.submit();
}


</SCRIPT>