<%

Set ConnUPV = Server.CreateObject("ADODB.Connection")

Sub AbrirUPV()
	Set ConnUPV = Server.CreateObject("ADODB.Connection")
	ConnUPV.Open "driver=Sql server;Uid=Llacruz;Pwd=llacruz.,-;Database=matricula;App=Sistema Operativo Microsoft Windows;Server=200.54.110.124"
	Server.ScriptTimeout = 1800
	ConnUPV.CommandTimeOut = 10000
End Sub

Sub CerrarUPV()
	ConnUPV.close
	set ConnUPV = nothing
End Sub

%>