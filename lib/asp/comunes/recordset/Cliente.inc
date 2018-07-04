<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub recordset_Cliente ( OBJdbConnection, rsRecordSet, IntId )

    If Trim( IntId ) = "" Then
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM CLIENTE"
    Else
    	odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM CLIENTE WHERE COD_CLIENTE = '" & IntId & "'"
     End If
End Sub
</SCRIPT>