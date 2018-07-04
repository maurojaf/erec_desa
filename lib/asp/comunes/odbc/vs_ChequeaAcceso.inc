<%
If Session("AplicNombreDeUsuario") <> "" Then
	Response.Write "<SCRIPT Languaje=""JavaScript"">"
	SQL = "SELECT * FROM Empleado WHERE CodUsuario='" & Session("AplicNombreDeUsuario") & "'"
	Set rsCliente = OBJdbConnection.Execute( SQL )
	Response.Write "window.defaultStatus = ""Cliente conectado: '" & rsCliente("Nombre") & " " & rsCliente("ApPaterno") & " " & rsCliente("ApMaterno") &"'"";"
	Response.Write "window.status = window.defaultStatus;"
	rsCliente.Close()
	Set rsCliente = nothing
	Response.Write "</SCRIPT>"
	
	dim avarEmpresas, totalEmpresas
	SQLQuery = "SELECT Empresa.* FROM Empresa, Empleado " & _
		" WHERE Empresa.EmpresaId = Empleado.EmpresaId " & _
		" AND Empleado.CodUsuario = '" & Session("AplicNombreDeUsuario") & "'"
	ObtenerRegistros OBJdbConnection, avarEmpresas, totalEmpresas, SQLQuery

Else
	Response.Redirect "Registro.asp"
End If
%>