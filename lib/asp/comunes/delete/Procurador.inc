<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub delete_Procurador (OBJdbConnection, ByVal Id, ByRef Mensaje)

   Mensaje = ""
   odbc_BorrarRegistro OBJdbConnection, "SELECT * FROM PROCURADOR WHERE ID_PROCURADOR = " + Cstr(Id)
   If Cstr(Err) = "3219" Then Mensaje = "No se puede borrar el registro "

End Sub
</SCRIPT>
