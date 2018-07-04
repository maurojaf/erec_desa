<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_UsuarioQry ( OBJdbConnection, strWhere, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM USUARIO " & strWhere & " ORDER BY ID_USUARIO DESC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>