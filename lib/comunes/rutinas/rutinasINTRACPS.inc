<%
Function TraeDatosUsuarioINTRACPS(strConex,intUsuario,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM USUARIO WHERE COD_USUARIO = " &  intUsuario
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeDatosUsuarioINTRACPS = rs(strCampo)
		Else
			TraeDatosUsuarioINTRACPS = ""
		End If
	Else
		TraeDatosUsuarioINTRACPS = ""
	End If

End Function
%>