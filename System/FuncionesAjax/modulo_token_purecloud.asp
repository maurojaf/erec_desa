<%@  language="VBSCRIPT" codepage="65001" %>

<%
	
	session("tokenID") = Request.QueryString("tokenID")


	Response.Redirect "../principal.asp?g=si"
%>