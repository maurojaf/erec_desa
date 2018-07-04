<%
Function TraeFechaActual(strConex)
	strSql="select convert(varchar(10),getdate() ,103) as FechaActual" 
	set rsFecha = strConex.execute(strSql)
	TraeFechaActual= rsFecha("FechaActual")
	rsFecha.close
	set rsFecha=nothing
End Function

%>