<%

Dim ConnMySql
Set ConnMySql = Server.CreateObject("ADODB.Connection")
ConnMySql.Open "Driver={MySQL ODBC 5.1 Driver}; Server=192.168.2.20; Database=asteriskcdrdb; UID=serverCRM30; PWD=Llacruzhuelen164"

%>