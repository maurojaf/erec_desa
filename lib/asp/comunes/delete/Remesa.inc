<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub delete_Remesa (OBJdbConnection, ByVal Id, ByVal intIdCliente, ByRef Mensaje)

   Mensaje = ""
   odbc_BorrarRegistro OBJdbConnection, "SELECT * FROM REMESA WHERE COD_REMESA = " & Cstr(Id) & " AND COD_CLIENTE = '" & intIdCliente & "'"
   If Cstr(Err) = "3219" Then Mensaje = "No se puede borrar el registro "

End Sub
</SCRIPT>

