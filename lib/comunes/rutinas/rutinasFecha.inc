<%

Function TraeFechaActual(strConex)
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT convert(varchar(10),getdate(),103) as strFechaActual ", strConex
	If Not rs.EOF Then
		TraeFechaActual = Trim(rs("strFechaActual"))
	End If
End Function

Function TraeFechaHoraActual(strConex)
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT getdate() as strFechaActual ", strConex
	If Not rs.EOF Then
		TraeFechaHoraActual = Trim(rs("strFechaActual"))
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



%>