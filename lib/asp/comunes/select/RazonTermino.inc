<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_RazonTermino ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM RAZONTERMINO ORDER BY IDRAZONTERMINO DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>