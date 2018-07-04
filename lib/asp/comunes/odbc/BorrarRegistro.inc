<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub odbc_BorrarRegistro ( ByRef Conn, ByVal strSQLQuery )
  On Error Resume Next
  'Response.Write "<pre>" & strSQLQuery & "</pre>"
  
  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.ActiveConnection = Conn
  RS.CursorType = adOpenDynamic
  RS.LockType = adLockOptimistic
  RS.Open strSQLQuery, Conn
  If Not RS.EOF Then RS.Delete
  RS.Update
  RS.Close
  Set RS = nothing

End Sub
</SCRIPT>