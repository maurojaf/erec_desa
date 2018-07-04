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
	
	IntId= request("IntId")
    VarNombreFichero=request("VarNombreFichero")
    strRut=request("strRut")


	 Dim DestinationPath
	If Trim(strRut) <> "" Then
		DestinationPath=  Server.mapPath("../Archivo/InformesAnexados") & "\" & IntId & "\" & strRut & "\" & VarNombreFichero
		DestinationPathValida=  Server.mapPath("../Archivo/InformesAnexados") & "\" & IntId & "\" & strRut
	Else
		DestinationPath = Server.mapPath("../Archivo/InformesAnexados") & "\" & IntId  & "\" & VarNombreFichero
	End If

	'Response.write "<br>DestinationPath" & DestinationPath
	'Response.write "<br>DestinationPathValida" & DestinationPathValida
	'response.End
	dim fs
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    if fs.FileExists(DestinationPath) then fs.DeleteFile(DestinationPath)




    If fs.FolderExists(DestinationPathValida) Then
		Set objFolder = fs.GetFolder(DestinationPathValida)
		Set colFiles = objFolder.Files
		If colFiles.Count = 0 then
			'AbrirSCG()
			'strSql="UPDATE DEUDOR SET FEC_SUBIDA_ULT_ARCHIVO = NULL WHERE CODCLIENTE = '" & IntId & "' AND RUTDEUDOR = '" & strRut& "'"
			''Response.write "strSql = " & strSql
			'Conn.execute(strSql)
			'CerrarSCG()
		End If

    End If



    'response.End

    Set fs = Nothing

  	If Trim(strRut) <> "" Then
		Response.Redirect "informe_clientes.asp?strRut=" & strRut
	Else
		Response.Redirect "informe_clientes.asp"
	End If


%>