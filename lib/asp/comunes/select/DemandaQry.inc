<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_DemandaQry ( OBJdbConnection, strWhere, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM DEMANDA " & strWhere & " ORDER BY IDDEMANDA DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>