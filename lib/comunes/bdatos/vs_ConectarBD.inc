<%
	' Se conecta a la Base de Datos
	Set OBJdbConnection = Server.CreateObject("ADODB.Connection")
	OBJdbConnection.Open "uid=usr_intranet;pwd=usr_intranet;driver={SQL Server};server=ICARUS;database=INTRACPS"
	
	' Si hay un error lo muestra y termina la ejecución del programa
	If Err.Number <> 0 Then Call vs_MostrarError("No se pudo conectar a la Base de Datos", true )
%>
