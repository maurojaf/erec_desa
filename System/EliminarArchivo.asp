<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../lib/asp/comunes/General/MostrarRegistro.inc"-->
<!--#include file="../lib/asp/comunes/general/JavaEfectoBotones.inc"-->
<!--#include file="../lib/asp/comunes/general/rutinasTraeCampo.inc"-->

<% ' Capa 1 ' %>
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
<!--#include file="arch_utils.asp"-->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	IntId 				=request("IntId")
    VarNombreFichero 	=request("VarNombreFichero")
    strRut 				=request("strRut")
    pagina_origen 		=request("pagina_origen")
    id_archivo 			=request("id_archivo")

	 Dim DestinationPath
	If Trim(strRut) <> "" Then
		DestinationPath=  Server.mapPath("UploadFolder") & "\" & IntId & "\" & strRut & "\" & VarNombreFichero
		DestinationPathValida=  Server.mapPath("UploadFolder") & "\" & IntId & "\" & strRut
	Else
		DestinationPath = Server.mapPath("UploadFolder") & "\" & IntId  & "\" & VarNombreFichero
	End If


	if trim(pagina_origen)="biblioteca_deudores" then
		DestinationPath = Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId & "\" & strRut & "\" & VarNombreFichero

		DestinationPathValida=  Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId & "\" & strRut

		AbrirSCG()
		strSql = "EXEC Proc_Audita_Archivo 0, 2, "&trim(session("session_idusuario"))&",'"&trim(strRut)&"', '"&triM(IntId)&"', '"&trim(archivo)&"','', "&CINT(trim(id_archivo))
		'response.end
		Conn.execute(strSql)
		'response.write strSql
		


		SQL_UPDATE_CARGA_ARCHIVO_CUOTA ="UPDATE CARGA_ARCHIVOS_CUOTA SET ACTIVO=0, FECHA_ELIMINACION='"&date()&"', USUARIO_ELIMINACION='"&trim(session("session_idusuario"))&"' WHERE ID_ARCHIVO=" &CINT(trim(id_archivo))
		Conn.execute(SQL_UPDATE_CARGA_ARCHIVO_CUOTA)

		CerrarSCG()

	end if

	if trim(pagina_origen)="biblioteca_cliente" then

		DestinationPath = Server.mapPath("../Archivo/BibliotecaClientes") & "\" & IntId & "\" & VarNombreFichero

		AbrirSCG()
		strSql = "EXEC Proc_Audita_Archivo 0, 1, "&trim(session("session_idusuario"))&",null, '"&triM(IntId)&"', '"&trim(archivo)&"','', "&CINT(trim(id_archivo))
		'response.end
		Conn.execute(strSql)
		'response.write strSql
		CerrarSCG()



	end if

	if trim(pagina_origen)="CargaArchivos" then

		DestinationPath = Server.mapPath("../Archivo/CargaArchivosAdmin") & "\" & IntId & "\" & VarNombreFichero

		AbrirSCG()
		strSql = "EXEC Proc_Audita_Archivo 0, 1, "&trim(session("session_idusuario"))&",null, '"&triM(IntId)&"', '"&trim(archivo)&"','', "&CINT(trim(id_archivo))
		'response.end
		Conn.execute(strSql)
		'response.write strSql
		CerrarSCG()		

	end if

	if trim(pagina_origen)="informe_cliente" then

		If Trim(strRut) <> "" Then
			DestinationPath=  Server.mapPath("../Archivo/InformesAnexados") & "\" & IntId & "\" & strRut & "\" & VarNombreFichero
			DestinationPathValida=  Server.mapPath("../Archivo/InformesAnexados") & "\" & IntId & "\" & strRut
		Else
			DestinationPath = Server.mapPath("../Archivo/InformesAnexados") & "\" & IntId  & "\" & VarNombreFichero
		End If


		AbrirSCG()
		strSql = "EXEC Proc_Audita_Archivo 0, 5, "&trim(session("session_idusuario"))&",null, '"&triM(IntId)&"', '"&trim(archivo)&"','', "&trim(id_archivo)
		'response.end
		Conn.execute(strSql)
		'response.write strSql
		CerrarSCG()		

	end if

	

	'response.write pagina_origen
	'Response.write "<br>DestinationPath" & DestinationPath
	'Response.write "<br>DestinationPathValida" & DestinationPathValida
	'response.End
	dim fs
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    if fs.FileExists(DestinationPath) then fs.DeleteFile(DestinationPath)

	'response.End


    If fs.FolderExists(DestinationPathValida) Then
		Set objFolder = fs.GetFolder(DestinationPathValida)
		Set colFiles = objFolder.Files
		If colFiles.Count = 0 then
			AbrirSCG()
			strSql="UPDATE DEUDOR SET FEC_SUBIDA_ULT_ARCHIVO = NULL WHERE COD_CLIENTE = '" & IntId & "' AND RUT_DEUDOR = '" & strRut& "'"
			'Response.write "strSql = " & strSql
			Conn.execute(strSql)
			CerrarSCG()
		End If

    End If



    'response.End

    Set fs = Nothing

	if trim(pagina_origen)="CargaArchivos" then

		Response.Redirect "carga_archivos_admin.asp"

	elseif trim(pagina_origen)="informe_cliente" then
	
	  	If Trim(strRut) <> "" Then
			Response.Redirect "informe_clientes.asp?strRut=" & strRut
		Else
			Response.Redirect "informe_clientes.asp"
		End If	

	else

	  	If Trim(strRut) <> "" Then
			Response.Redirect "biblioteca_deudores.asp?strRut=" & strRut
		Else
			Response.Redirect "man_ClienteForm.asp?sintNuevo=0&COD_CLIENTE=" & IntId
		End If

	end if




%>