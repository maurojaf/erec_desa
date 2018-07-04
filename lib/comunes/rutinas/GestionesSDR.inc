<%
Function CuentaGestiones(strConex,intRutDeudor,intTipoGestion)
Set rs = Server.CreateObject("ADODB.Recordset")
strSql = ""
strSql = "SELECT COUNT(*) AS INTCANTGESTIONES FROM GESTION_TERRENO WHERE RUT_DEUDOR = '" & intRutDeudor & "'"
Select Case UCASE(intTipoGestion)
	Case "F": 'Call, telefonica
		strSql = strSql & " AND gtipo='F'"
	Case "T": 'Terreno
		strSql = strSql & " AND gtipo='T'"
	Case "A": 'Gestion de Apoyo
		strSql = strSql & " AND codgestion='0008'"
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


