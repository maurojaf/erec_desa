<%
Function TraeCampoId(strConex, strCampo, intId, strTabla, descId)
	
	If Trim(intId) = "" Then
		TraeCampoId=""
	Else
		strSql="select " & strCampo & " from " & strTabla & " where " & descId & " = " & intId 
		'Response.write strSql
		'Response.End
		set rsCliente = strConex.execute(strSql)
		if not rsCliente.eof then
			TraeCampoId = rsCliente(strCampo)
		else
			TraeCampoId = ""
		End if
		rsCliente.close
		set rsCliente=nothing
	End if
End Function

%>