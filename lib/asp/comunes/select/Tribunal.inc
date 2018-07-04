<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_Tribunal ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM TRIBUNAL ORDER BY IDTRIBUNAL DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>