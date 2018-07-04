<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_EstadoDemanda ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM ESTADODEMANDA ORDER BY IDESTADODEMANDA DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>