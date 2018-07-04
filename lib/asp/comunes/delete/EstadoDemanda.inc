<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub delete_EstadoDemanda (OBJdbConnection, ByVal Id, ByRef Mensaje)

   Mensaje = ""
   odbc_BorrarRegistro OBJdbConnection, "SELECT * FROM ESTADO_DEMANDA WHERE ID_ESTADO_DEMANDA = " + Cstr(Id)
   If Cstr(Err) = "3219" Then Mensaje = "No se puede borrar el registro "

End Sub
</SCRIPT>
