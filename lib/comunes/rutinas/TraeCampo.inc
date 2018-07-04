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

Function TraeCampoId2(strConex, strCampo, intId, strTabla, descId)
	
	If Trim(intId) = "" Then
		TraeCampoId2=""
	Else
		strSql="select " & strCampo & " from " & strTabla & " where " & descId & " = '" & intId & "'"
		'Response.write strSql
		'Response.End
		set rsCliente = strConex.execute(strSql)
		if not rsCliente.eof then
			TraeCampoId2 = rsCliente(strCampo)
		else
			TraeCampoId2 = ""
		End if
		rsCliente.close
		set rsCliente=nothing
	End if
End Function

Function TraeCampoIdWhere(strConex, strCampo, strTabla, strQry)
	
	If Trim(strTabla) = "" Then
		TraeCampoIdWhere=""
	Else
		strSql="select " & strCampo & " from " & strTabla & " where " & strQry
		'Response.write strSql
		'Response.End
		set rsCliente = strConex.execute(strSql)
		if not rsCliente.eof then
			TraeCampoIdWhere = rsCliente(strCampo)
		else
			TraeCampoIdWhere = ""
		End if
		rsCliente.close
		set rsCliente=nothing
	End if
End Function

Function TraeCampoSeleccionado(strConex, strCampo, strQry)
	If Trim(strQry) = "" Then
		TraeCampoSeleccionado=""
	Else
		strSql= strQry
		'Response.write strSql
		'Response.End
		set rsCliente = strConex.execute(strSql)
		if not rsCliente.eof then
			TraeCampoSeleccionado = rsCliente(strCampo)
		else
			TraeCampoSeleccionado = ""
		End if
		rsCliente.close
		set rsCliente=nothing
	End if
End Function
%>
