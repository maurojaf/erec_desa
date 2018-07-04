<%

If session("session_user")="" then

	Response.write "SESION HA EXPIRADO, CIERRE EL NAVEGADOR Y VUELVA A INGRESAR"
	''Response.redirect("../index.asp")
	Response.End
End if

%>
<!--#include file="menu.asp"-->
