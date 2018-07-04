<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_Actuario ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM ACTUARIO ORDER BY IDACTUARIO DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>