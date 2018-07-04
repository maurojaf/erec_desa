<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../lib/asp/comunes/General/MostrarRegistro.inc"-->
<!--#include file="../lib/asp/comunes/general/JavaEfectoBotones.inc"-->
<!--#include file="../lib/asp/comunes/general/rutinasTraeCampo.inc"-->

<% ' Capa 1 ' %>
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/odbc/ObtenerRecordSet.inc"-->
<!--#include file="../lib/asp/comunes/odbc/ObtenerRegistros.inc"-->

<% ' Capa 2 ' %>
<!--#include file="../lib/asp/comunes/recordset/Cliente.inc"-->
<!--#include file="../lib/freeaspupload.asp" -->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<%


	Response.CodePage=65001
	Response.charset ="utf-8"



   AbrirSCG()



	DestinationPath = Server.mapPath("../Archivo/BibliotecaDeudores") 

	' crear una instancia



Set Obj_FSO = createobject("scripting.filesystemobject")

Set objfolder = Obj_FSO.GetFolder (DestinationPath)



for each objsubfolder in objfolder.subfolders 'Recorre los subdirectorios del directorio actual
	
	response.write objsubfolder.name&"<br>"

	Set objfile = Obj_FSO.GetFolder (DestinationPath&"/"&objsubfolder.name)

	for each objfile in objsubfolder.subfolders 'Recorre los ficheros del directorio raiz
		response.write "&nbsp;&nbsp;&nbsp;"&objfile.name&"<br>"

		

		Set objfile2 = Obj_FSO.GetFolder (DestinationPath&"/"&objsubfolder.name&"/"&objfile.name)

		for each objfile22 in objfile2.files  'Recorre los ficheros del directorio raiz


				'response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&objfile22.name&"<br>"
			if trim(objfile22.name)<>"Thumbs.db" then

				strSql = "EXEC Proc_Audita_Archivo 1, 2, 1,'"&trim(objfile.name)&"', '"&triM(objsubfolder.name)&"', '"&trim(objfile22.name)&"', '',0 "

				'response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&strSql&"<br>"
				'Conn.execute(strSql)

				Cadena1=objfile22.name
				Cadena2="ñ"
				If InStr(Cadena1,Cadena2)>0 then
					response.write objfile22.name&"<br>"
				End if

			end if

			
		next


	next

next
		'Response.write DestinationPath



%>





