<%
Function TraeDatosDeudorSCP(strConex,intCodCliente,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(intCodCliente) <> "" Then
		strSql = "SELECT " & strCampo & " FROM T_JUICIO WHERE RUT_DEUDOR = '" &  strRut & "' AND CODCLIENTE = '" & intCodCliente & "'"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeDatosDeudorSCP = rs(strCampo)
		Else
			TraeDatosDeudorSCP = ""
		End If
	Else
		TraeDatosDeudorSCP = ""
	End If

End Function

Function TraeDatosDeudorSinClienteSCP(strConex,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM T_JUICIO WHERE RUT = '" &  strRut & "'"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeDatosDeudorSinClienteSCP = rs(strCampo)
		Else
			TraeDatosDeudorSinClienteSCP = ""
		End If
	Else
		TraeDatosDeudorSinClienteSCP = ""
	End If

End Function


%>