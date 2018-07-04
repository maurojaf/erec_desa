<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_Usuario ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM USUARIO WHERE ID_USUARIO >= 100 ORDER BY ID_USUARIO DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>