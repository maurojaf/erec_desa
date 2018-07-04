<!DOCTYPE html>
<HTML lang="es">
<HEAD>
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
   	<meta charset="utf-8">

	<!--#include file="sesion_inicio.asp"-->
	<%

	If Trim(request("HD_CLIENTE")) <> "" Then
		session("ses_codcli")=request("HD_CLIENTE")
	End If
	
	
	
	if session("Pagina") = "" then
		Pagina = "principal.asp?EmpresaId=3"
	else 
		Pagina=session("Pagina")
	end if 
	
	
	
	%>
	<TITLE>SISTEMA DE COBRANZA</TITLE>
</HEAD>
<!--
	<FRAMESET cols="*,1024,*" frameborder="NO" border="0" framespacing="0">
	<FRAME src="blank.html" SCROLLING="NO" NORESIZE MARGINWIDTH='0' MARGINHEIGHT='0' BORDER='0'> -->
	<FRAMESET rows="80,*" frameborder="NO" border="0" framespacing="0">
	    <frame NAME='topFrame' SRC='top.asp?EmpresaId=3'  SCROLLING="NO" NORESIZE MARGINWIDTH='0' MARGINHEIGHT='0' BORDER='0'>
	    <FRAME NAME='Contenido' SRC='<%=pagina%>' SCROLLING='AUTO' NORESIZE MARGINWIDTH='0' MARGINHEIGHT='0' BORDER='0'>
	</FRAMESET>
<!--	<FRAME src="blank.html" SCROLLING="NO" NORESIZE MARGINWIDTH='0' MARGINHEIGHT='0' BORDER='0'>
	</FRAMESET>-->

</HTML>
