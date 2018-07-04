<%@ Language=VBScript %>
<!--#include file="sesion.asp" -->
<!--#include file="AspSessionManager.asp" -->
<%
	Dim oAspSessionManager
	
	Set oAspSessionManager = New AspSessionManager
	
	oAspSessionManager.SerializeElements()
%>
<iframe src="http://sistemas.llacruz.cl/eRec4Core/Uf" style="border: 0; width: 100%; height: 90%;">
<%
	Set oAspSessionManager = Nothing
%>