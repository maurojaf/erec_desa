<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!-- ARCHIVOS INCLUIDOS -->
		
		<!--#include file="../../arch_utils.asp"-->
		<!--#include file="../../../lib/comunes/rutinas/funciones.inc" -->
		
<%
	AbrirSCG()

		Observacion 	=	request("Observacion")
		IdCampo			=	request("IdCampo")
		IdDominio		=	request("IdDominio")
		RutDeudor		=	request("RutDeudor")
		CodigoCliente	=	request("CodigoCliente")
		CodigoUsuario	=	request("CodigoUsuario")
		Texto			=	request("Texto")
		
		if(Observacion = "") then
			Observacion = "null"
		else
			Observacion = "'"&Observacion&"'"
		end if
		
		if(Texto = "") then
			Texto = "null"
		else
			Texto = "'"&Texto&"'"
		end if
		
		StrSql="EXEC uspFichaDeudorInsert "&TRIM(Observacion)&",'"&TRIM(IdCampo)&"','"&TRIM(IdDominio)&"','"&TRIM(RutDeudor)&"','"&TRIM(CodigoCliente)&"','"&TRIM(CodigoUsuario)&"',"&TRIM(Texto)&""
		Conn.execute(StrSql)
		
	CerrarSCG()
%>