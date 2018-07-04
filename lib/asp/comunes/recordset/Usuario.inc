<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub recordset_Usuario ( OBJdbConnection, rsRecordSet, IntId )
    If Trim( IntId ) = "" Then
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM USUARIO"
    Else
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM USUARIO WHERE ID_USUARIO = " & IntId
    End If
End Sub
</SCRIPT>