<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_Demanda ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM DEMANDA ORDER BY IDDEMANDA ASC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>