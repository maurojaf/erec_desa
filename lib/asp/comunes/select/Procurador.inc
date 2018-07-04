<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_Procurador ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM PROCURADOR ORDER BY IDPROCURADOR DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>