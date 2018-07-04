<%
Function TraeNombreClienteSCG(strConex,intCodCliente)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(intCodCliente) <> "" Then
		rs.Open "SELECT * FROM CLIENTES WHERE COD_CLI = '" & intCodCliente & "'", strConex
		If Not rs.EOF Then
			TraeNombreClienteSCG = Trim(rs("DESC_CLI"))
		Else
			TraeNombreClienteSCG = "Cliente " & intCodCliente
		End If
	Else
		TraeNombreClienteSCG = "Cliente " & intCodCliente
	End If

End Function

Function TraeFechaActual(strConex)
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT convert(varchar(10),getdate(),103) as strFechaActual ", strConex
	If Not rs.EOF Then
		TraeFechaActual = Trim(rs("strFechaActual"))
	End If

End Function

Function TraeFechaActual2(strConex,strTipo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	If Trim(strTipo) = "D" Then
		strSql="SELECT Day(getdate()) as FEC "
	End if
	If Trim(strTipo) = "M" Then
		strSql="SELECT Month(getdate()) as FEC "
	End if
	If Trim(strTipo) = "A" Then
		strSql="SELECT Year(getdate()) as FEC "
	End if
	
	'Response.write "strSql="&strSql
	'Response.End
	
	rs.Open strSql, strConex
	If Not rs.EOF Then
		TraeFechaActual2 = Trim(rs("FEC"))
	Else
		TraeFechaActual2 = Trim(rs("FEC"))
	End If

End Function

Function TraeDatosDeudorSCG(strConex,intCodCliente,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(intCodCliente) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR WHERE RUT_DEUDOR = '" &  strRut & "' AND CODCLIENTE = '" & intCodCliente & "'"
		'Response.write "strSql="&strSql
		'Response.write strSql
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeDatosDeudorSCG = Trim(rs(strCampo))
		Else
			TraeDatosDeudorSCG = ""
		End If
	Else
		TraeDatosDeudorSCG = ""
	End If

End Function

Function TraeDatosDeudorSinClienteSCG(strConex,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR WHERE RUT_DEUDOR = '" &  strRut & "'"
		
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeDatosDeudorSinClienteSCG = rs(strCampo)
		Else
			TraeDatosDeudorSinClienteSCG = ""
		End If
	Else
		TraeDatosDeudorSinClienteSCG = ""
	End If

End Function

Function TraeUltimaDirDeudorSCG(strConex,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 1"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeUltimaDirDeudorSCG = rs(strCampo)
		Else
			strSql = "SELECT " & strCampo & " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 0"
			rs.Close
			rs.Open strSql , strConex
			If Not rs.EOF Then
				TraeUltimaDirDeudorSCG = rs(strCampo)
			Else
				TraeUltimaDirDeudorSCG = ""
			End If
		End If
	Else
		TraeUltimaDirDeudorSCG = ""
	End If

End Function


Function TraeUltimoFonoDeudorSCG(strConex,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 1 AND COD_AREA <> 9 AND COD_AREA <> 8"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeUltimoFonoDeudorSCG = rs(strCampo)
		Else
			strSql = "SELECT " & strCampo & " FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 0 AND COD_AREA <> 9 AND COD_AREA <> 8"
			rs.Close
			rs.Open strSql , strConex
			If Not rs.EOF Then
				TraeUltimoFonoDeudorSCG = rs(strCampo)
			Else
				TraeUltimoFonoDeudorSCG = ""
			End If
		End If
	Else
		TraeUltimoFonoDeudorSCG = ""
	End If

End Function

Function TraeUltimoMovilDeudorSCG(strConex,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 1 AND (COD_AREA=9 OR COD_AREA=8)"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeUltimoMovilDeudorSCG = rs(strCampo)
		Else
			strSql = "SELECT " & strCampo & " FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 0 AND (COD_AREA=9 OR COD_AREA=8)"
			rs.Close
			rs.Open strSql , strConex
			If Not rs.EOF Then
				TraeUltimoMovilDeudorSCG = rs(strCampo)
			Else
				TraeUltimoMovilDeudorSCG = ""
			End If
		End If
	Else
		TraeUltimoMovilDeudorSCG = ""
	End If

End Function


Function TraeRutPorNroDocSCG(strConex,strNroDoc)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strNroDoc) <> "" Then
		strSql = "SELECT DISTINCT RUT_DEUDOR AS RUT_DEUDOR FROM CUOTA WHERE NRODOC = '" &  strNroDoc & "' AND SALDO > 0"
		'Response.write strSql
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeRutPorNroDocSCG = rs("RUT_DEUDOR")
		Else
			TraeRutPorNroDocSCG = ""
		End If
	Else
		TraeRutPorNroDocSCG = ""
	End If

End Function

Function TraeProductoSCG(strConex,strCodProd)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strCodProd) <> "" Then
		strSql = "SELECT DESCRIPCION FROM PRODUCTO WHERE TIPO_PRODU = '" &  strCodProd & "'"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeProductoSCG = rs("DESCRIPCION")
		Else
			TraeProductoSCG = ""
		End If
	Else
		TraeProductoSCG = ""
	End If

End Function

Function TraeTipoDocSCG(strConex,strCodTipo, strCodCli)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strCodTipo) <> "" and Trim(strCodCli) <> "" Then
		strSql = "SELECT DESCRIPCION FROM TIPODOCUMENTO WHERE CODIGO = '" &  strCodTipo & "' AND CODCLIENTE = '" &  strCodCli & "'"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeTipoDocSCG = rs("DESCRIPCION")
		Else
			TraeTipoDocSCG = ""
		End If
	Else
		TraeTipoDoc = ""
	End If

End Function

Function TraeEstadoSCG(strConex,strCodEstado)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strCodEstado) <> "" Then
		strSql = "SELECT DESCRIP FROM ESTADOS WHERE CODIGO = '" &  strCodEstado & "'"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeEstadoSCG = rs("DESCRIP")
		Else
			TraeEstadoSCG = ""
		End If
	Else
		TraeTipoDoc = ""
	End If

End Function
%>