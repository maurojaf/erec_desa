<% If InStr(Request.ServerVariables ("PATH_TRANSLATED"), "login.asp") <= 0 Then
	If Session("intCodUsuario") = "" Then %>
	<SCRIPT LANGUAGE=JavaScript>
		alert('Su sesión ha expirado')
		top.location.href='<%= Session("virtual_dir") %>/inicio.asp'
	</SCRIPT>
	<% 
	Response.End
	End If
End If
%>