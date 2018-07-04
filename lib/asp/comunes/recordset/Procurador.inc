<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub recordset_Procurador ( OBJdbConnection, rsRecordSet, IntId )
    If Trim( IntId ) = "" Then
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM PROCURADOR"
    Else
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM PROCURADOR WHERE ID_PROCURADOR = " & IntId
    End If
End Sub
</SCRIPT>