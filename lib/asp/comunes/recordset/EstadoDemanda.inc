<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub recordset_EstadoDemanda ( OBJdbConnection, rsRecordSet, IntId )
    If Trim( IntId ) = "" Then
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM ESTADO_DEMANDA"
    Else
        odbc_ObtenerRecordSet OBJdbConnection, rsRecordSet, " SELECT * FROM ESTADO_DEMANDA WHERE ID_ESTADO_DEMANDA = " & IntId
    End If
End Sub
</SCRIPT>