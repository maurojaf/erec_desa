
<%
On error GoTo 0

SERVIDOR= MID(request.servervariables("PATH_INFO"),2, (Instr(MID(request.servervariables("PATH_INFO"),2, LEN(request.servervariables("PATH_INFO"))),"/"))-1)

'response.WRITE SERVIDOR
if ucase(SERVIDOR)="EREC_DESA" then
	Database ="EREC_DESA"

elseif ucase(SERVIDOR)="EREC" then
	Database ="EREC"

elseif ucase(SERVIDOR)="EREC_DEMO" then
	Database ="EREC_DEMO"
	
else
	response.write "ERROR CONEXION BASE DE DATOS.. Contacte al administrador"
	response.end()
end if


Set Conn = Server.CreateObject("ADODB.Connection")

Sub AbrirSCG()
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open "driver=Sql server;Uid=login;Pwd=LoginLlacruz164;Database="&TRIM(Database)&";App=Sistema Operativo Microsoft Windows;Server=192.168.2.14"
	Server.ScriptTimeout = 1800
	Conn.CommandTimeOut = 10000
	if err then 
		response.write " Contacte a su administrador / error : "& err.description
	end if		

End Sub

Sub CerrarSCG()
	Conn.close
	set Conn = nothing
End Sub

 'response.write "driver=Sql server;Uid=login;Pwd=LoginLlacruz164;Database="&TRIM(Database)&";App=Sistema Operativo Microsoft Windows;Server=192.168.2.14<br>"

%>