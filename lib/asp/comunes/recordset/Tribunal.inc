<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub recordset_Tribunal ( OBJdbConnection, rsRecordSet, IntId )
    If Trim( IntId ) = "" Then
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM TRIBUNAL"
    Else
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM TRIBUNAL WHERE ID_TRIBUNAL = " & IntId
    End If
End Sub
</SCRIPT>