<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub recordset_Demanda ( OBJdbConnection, rsRecordSet, IntId )
    If Trim( IntId ) = "" Then
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM DEMANDA"
    Else
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM DEMANDA WHERE ID_DEMANDA = " & IntId
    End If
End Sub
</SCRIPT>