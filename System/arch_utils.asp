<%
On error GoTo 0

SERVIDOR= MID(request.servervariables("PATH_INFO"),2, (Instr(MID(request.servervariables("PATH_INFO"),2, LEN(request.servervariables("PATH_INFO"))),"/"))-1)
	'response.write SERVIDOR
if ucase(SERVIDOR)="EREC" then
	DATABASE ="EREC"

elseif ucase(SERVIDOR)="EREC_DEMO" then
	DATABASE ="EREC_DEMO"

elseif ucase(SERVIDOR)="EREC_DESA" then
	DATABASE ="EREC_DESA"

end if

'response.write DATABASE
Set Conn = Server.CreateObject("ADODB.Connection")
Set Conn1 = Server.CreateObject("ADODB.Connection")
Set Conn2 = Server.CreateObject("ADODB.Connection")

Sub AbrirSCG()
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open "driver=Sql server;Uid=sa;Pwd=LlacruzhuelenPiso3;Database=" & trim(DATABASE) & ";App=Sistema Operativo Microsoft Windows;Server=192.168.2.14"
	Server.ScriptTimeout = 1800
	Conn.CommandTimeOut = 10000
End Sub

Sub CerrarSCG()
	Conn.close
	set Conn = nothing
End Sub

Sub AbrirSCG1()
	Set Conn1 = Server.CreateObject("ADODB.Connection")
	Conn1.Open "driver=Sql server;Uid=sa;Pwd=LlacruzhuelenPiso3;Database=" & trim(DATABASE) & ";App=Sistema Operativo Microsoft Windows;Server=192.168.2.14"
	Server.ScriptTimeout = 1800
	Conn1.CommandTimeOut = 10000
End Sub

Sub CerrarSCG1()
	Conn1.close
	set Conn1 = nothing
End Sub

Sub AbrirSCG2()
	Set Conn2 = Server.CreateObject("ADODB.Connection")
	Conn2.Open "driver=Sql server;Uid=sa;Pwd=LlacruzhuelenPiso3;Database=" & trim(DATABASE) & ";App=Sistema Operativo Microsoft Windows;Server=192.168.2.14"
	Server.ScriptTimeout = 1800
	Conn2.CommandTimeOut = 10000
End Sub

Sub CerrarSCG2()
	Conn2.close
	set Conn2 = nothing
End Sub


%>