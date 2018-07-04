<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub delete_Usuario (OBJdbConnection, ByVal Id, ByRef Mensaje)

   Mensaje = ""
   odbc_BorrarRegistro OBJdbConnection, "SELECT * FROM USUARIO WHERE ID_USUARIO = " + Cstr(Id)
   If Cstr(Err) = "3219" Then Mensaje = "No se puede borrar el registro "

End Sub
</SCRIPT>
