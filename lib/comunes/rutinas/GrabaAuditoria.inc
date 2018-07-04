<%
Function GrabaAuditoria(strAccion, strAplicacion, strArchivo, strTabla)
	
	strSql = "INSERT INTO AUDITORIA (ID_USUARIO, LOGIN, FECHA, IP, ACCION, APLICACION, ARCHIVO, TABLA)"
	strSql = strSql & " Values (" & session("session_idusuario") & ",'" & session("session_login") & "',getdate(),'" & Mid(request.servervariables("REMOTE_ADDR"),1,19) & "','" & strAccion & "','" & strAplicacion & "','" & strArchivo & "','" & strTabla & "')"
	set rsInserta=Conn.execute(strSql)
		
End Function
%>
