<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_Abogado ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM ABOGADO ORDER BY IDABOGADO DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>