<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub delete_Demanda (OBJdbConnection, ByVal Id, ByRef Mensaje)

   Mensaje = ""
   odbc_BorrarRegistro OBJdbConnection, "SELECT * FROM DEMANDA WHERE ID_DEMANDA = " + Cstr(Id)
   If Cstr(Err) = "3219" Then Mensaje = "No se puede borrar el registro "

End Sub
</SCRIPT>
