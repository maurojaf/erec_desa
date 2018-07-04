<SCRIPT LANGUAGE=VBScript RUNAT=Server>
' --------------------------------------------------------------------------------------------
' Función    : Conectar
' Descripción: Se conecta a la base de datos.
' Argumentos : No tiene
' Retorna    : El objeto Conectado

 Function odbc_Conectar()
    ' Se conecta a la Base de Datos
    Set objDbConnection = Server.CreateObject("ADODB.Connection")
    objDbConnection.Open Session("CadenaDeConexion")
    Set odbc_Conectar = objDbConnection
 End Function

</SCRIPT>
