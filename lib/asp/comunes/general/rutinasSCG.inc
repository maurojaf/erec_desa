<%
Function TraeNombreClienteSCG(strConex,intCOD_CLIENTE)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(intCOD_CLIENTE) <> "" Then
		rs.Open "SELECT * FROM CLIENTES WHERE COD_CLIENTE = '" & intCOD_CLIENTE & "'", strConex
		If Not rs.EOF Then
			TraeNombreClienteSCG = Trim(rs("DESC_CLI"))
		Else
			TraeNombreClienteSCG = "Cliente " & intCOD_CLIENTE
		End If
	Else
		TraeNombreClienteSCG = "Cliente " & intCOD_CLIENTE
	End If

End Function

Function TraeNombreGestion(strConex,intCodigo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(intCodigo) <> "" Then
		rs.Open "SELECT * FROM GESTIONES_UNIFICADAS WHERE CODIGO = '" & intCodigo & "'", strConex
		If Not rs.EOF Then
			TraeNombreGestion = Trim(rs("CATEGORIA")) & " " & Trim(rs("SUBCATEGORIA")) & " " & Trim(rs("GESTION"))
		Else
			TraeNombreGestion = "Gestion " & intCodigo
		End If
	Else
		TraeNombreGestion = "Gestion " & intCodigo
	End If

End Function

Function TraeFechaActual(strConex)
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT convert(varchar(10),getdate(),103) as strFechaActual ", strConex
	If Not rs.EOF Then
		TraeFechaActual = Trim(rs("strFechaActual"))
	End If
End Function


Function TraeFechaMesActual(strConex,intCant)
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT convert(varchar(10),dateadd(month," & intCant & ",getdate()),103) as strFechaActual ", strConex
	If Not rs.EOF Then
		TraeFechaMesActual = Trim(rs("strFechaActual"))
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

	
	rs.Open strSql, strConex
	If Not rs.EOF Then
		TraeFechaActual2 = Trim(rs("FEC"))
	Else
		TraeFechaActual2 = Trim(rs("FEC"))
	End If

End Function

Function TraeFechaActual3(strConex,strTipo)
	Set rs = Server.CreateObject("ADODB.Recordset")
    
    If Trim(strTipo) = "D" Then
        strSql="SELECT Day(getdate()) as FEC "
    End if
    If Trim(strTipo) = "M" Then
        strSql="SELECT substring(convert(char(2),getdate(),101),1,2) as FEC"
    End if
    If Trim(strTipo) = "A" Then
        strSql="SELECT Year(getdate()) as FEC "
    End if
    
    	
	'Response.write "strSql="&strSql
	'Response.End
	
	rs.Open strSql, strConex
	If Not rs.EOF Then
		TraeFechaActual3 = Trim(rs("FEC"))
	Else
		TraeFechaActual3 = Trim(rs("FEC"))
	End If

End Function

Function TraeDatosDeudorSCG(strConex,intCOD_CLIENTE,strRut,strCampo)

	
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(intCOD_CLIENTE) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR WHERE RUT_DEUDOR = '" &  strRut & "' AND COD_CLIENTE = '" & intCOD_CLIENTE & "'"
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
			
			''Response.write strSql
			rs.Close
			rs.Open strSql , strConex
			If Not rs.EOF Then
				TraeUltimaDirDeudorSCG = rs(strCampo)
			Else
				TraeUltimaDirDeudorSCG = ""
			End If
			If UCASE(Trim(strCampo)) = "RESTO" Then
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

Function TraeUltimoFonoDeudor(strConex,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 1"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeUltimoFonoDeudor = rs(strCampo)
		Else
			strSql = "SELECT " & strCampo & " FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 0"
			rs.Close
			rs.Open strSql , strConex

			If Not rs.EOF Then
				TraeUltimoFonoDeudor = rs(strCampo)
			Else
				TraeUltimoFonoDeudor = ""
			End If
		End If
	Else
		TraeUltimoFonoDeudor = ""
	End If
End Function

Function TraeUltimoEMailDeudorSCG(strConex,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 1 "
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeUltimoEMailDeudorSCG = rs(strCampo)
		Else
		strSql = "SELECT " & strCampo & " FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 0 "
			rs.Close
			rs.Open strSql , strConex
			If Not rs.EOF Then
				TraeUltimoEMailDeudorSCG = rs(strCampo)
			Else
				TraeUltimoEMailDeudorSCG = ""
			End If
		End If
	Else
		TraeUltimoFonoDeudorSCG = ""
	End If
End Function





Function TraeUltimoMovilDeudorSCG(strConex,strRut,strCampo)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRut) <> "" Then
		strSql = "SELECT " & strCampo & " FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 1 AND (COD_AREA=9 OR COD_AREA=8 OR COD_AREA=0)"
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeUltimoMovilDeudorSCG = rs(strCampo)
		Else
			strSql = "SELECT " & strCampo & " FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" &  strRut & "' AND ESTADO = 0 AND (COD_AREA=9 OR COD_AREA=8 OR COD_AREA=0)"
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
		strSql = "SELECT DESCRIPCION FROM TIPODOCUMENTO WHERE CODIGO = '" &  strCodTipo & "' AND COD_CLIENTE = '" &  strCodCli & "'"
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
		TraeEstadoSCG = ""
	End If

End Function

Function TraeNombreDeudor(strConex,strRUTDEUDOR)
	Set rs = Server.CreateObject("ADODB.Recordset")
	If Trim(strRUTDEUDOR) <> "" Then
		strSql = "SELECT TOP 1 NOMBRE_DEUDOR,RUT_DEUDOR FROM DEUDOR WHERE RUT_DEUDOR='"& strRUTDEUDOR &"' AND NOMBRE_DEUDOR IS NOT NULL AND LTRIM(RTRIM(NOMBRE_DEUDOR)) <> '' "
		
		'RESPONSE.WRITE strSql
		rs.Open strSql , strConex
		If Not rs.EOF Then
			TraeNombreDeudor = rs("NOMBRE_DEUDOR")
		Else
			TraeNombreDeudor = ""
		End If
	Else
		TraeNombreDeudor = ""
	End If

End Function

Function TraeZona(strNombreSucursal)
	Select Case strNombreSucursal
		Case "ARICA","ANTOFAGASTA","IQUIQUE","COPIAPO"
			TraeZona="NORTE"	
		Case "VALPARAISO","LA SERENA","RANCAGUA","TALCA"
			TraeZona="CENTRO"	
		Case "CONCEPCION","TEMUCO","VALDIVIA","OSORNO","PUERTO MONTT","COYHAIQUE","PUNTA ARENAS"
			TraeZona="SUR"	
		Case "SANTIAGO"
			TraeZona="SANTIAGO"	
	End Select
End Function


%>