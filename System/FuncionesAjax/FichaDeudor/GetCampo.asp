<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
		
<!--#include file="../../../lib/JSON_2.0.4.asp" -->
<!--#include file="../../arch_utils.asp"-->
<!--#include file="../../../lib/comunes/rutinas/funciones.inc" -->
		
<%
	IdCampo			=	request("IdCampo")
	
	AbrirSCG()
		
		StrSql = "		SELECT  IdCampo, "
		StrSql = StrSql & "		Nombre, "
		StrSql = StrSql & "		Tipo "
		StrSql = StrSql & "FROM    dbo.Campo "
		StrSql = StrSql & "WHERE   IdCampo = " & IdCampo
		
		set rsCampo = Conn.execute(StrSql)
		
		Dim campo
		
		Set campo = jsObject()

		if not rsCampo.eof then
		
			campo("IdCampo") = rsCampo("IdCampo")
			
			campo("Nombre") = rsCampo("Nombre")
			
			campo("Tipo") = rsCampo("Tipo")
			
		else
		
			campo("IdCampo") = ""
			
			campo("Nombre") = ""
			
			campo("Tipo") = ""
			
		end if

		campo.Flush
		
	CerrarSCG()
%>