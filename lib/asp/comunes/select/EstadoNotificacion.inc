<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub select_EstadoNotificacion ( OBJdbConnection, arrResultado, intNumRegistros)
    strQuery= " SELECT * FROM ESTADO_NOTIFICACION ORDER BY CODIGO ASC"
    'Response.Write strQuery
    odbc_ObtenerRegistros OBJdbConnection, arrResultado, intNumRegistros, strQuery
End Sub
</SCRIPT>