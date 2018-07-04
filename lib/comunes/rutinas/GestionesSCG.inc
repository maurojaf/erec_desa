<%
Function CuentaGestiones(strConex,intRutDeudor,intTipoGestion)
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	strSql = "SELECT COUNT(*) AS INTCANTGESTIONES FROM GESTIONES WHERE RUT_DEUDOR = '" & intRutDeudor & "'"
	
	Select Case UCASE(intTipoGestion)
		Case "C": 'Call, telefonica
			strSql = strSql & "AND CODCATEGORIA IN (1,2)"
		Case "T": 'Terreno
			strSql = strSql & "AND CODCATEGORIA IN (4,5)"
		Case "A": 'Gestion de Apoyo
			strSql = strSql & "AND CODCATEGORIA = 3"
	End Select
	
	rs.Open strSql, strConex
	If Not rs.EOF Then
		CuentaGestiones = Trim(rs("INTCANTGESTIONES"))
	Else
		CuentaGestiones = 0
	End If
	rs.Close

End Function

%>


