<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub odbc_ObtenerRecordSet ( ByRef Conn, ByRef rsRecordSet, ByVal strSqlQuery)
    'response.write strSqlQuery
    Set rsRecordSet = Conn.Execute (strSqlQuery)
End Sub
</SCRIPT>