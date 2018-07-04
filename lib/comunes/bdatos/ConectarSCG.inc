<%
Set ConexionSCG = Server.CreateObject("ADOdb.Connection")
ConexionSCG.ConnectionTimeout = 0
ConexionSCG.CommandTimeout = 0
ConexionSCG.Open "uid=scgweb2;pwd=scgweb2;driver={SQL Server};server=ICARUS;database=SCG"
%>