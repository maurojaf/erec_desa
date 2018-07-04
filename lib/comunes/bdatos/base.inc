<%

Set ConexionINTRACPS = Server.CreateObject("ADOdb.Connection")
ConexionINTRACPS.Open "uid=usr_intranet;pwd=usr_intranet;driver={SQL Server};server=ICARUS;database=INTRACPS"
'Session("intConfServidor") = "Ingles"
'Session("PrecisionDouble") = 10
%>