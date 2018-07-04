<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub odbc_BorrarRegistros ( ByRef Conn, ByVal strSQLQuery )
  On Error Resume Next
  'Response.Write "<pre>" & strSQLQuery & "</pre>"
  
  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.ActiveConnection = Conn
  RS.CursorType = adOpenDynamic
  RS.LockType = adLockOptimistic
  RS.Open strSQLQuery, Conn
  Do While Not RS.EOF  
  	RS.Delete
   	 'RS.Update
 	  RS.MoveNext
 	   
  Loop
  RS.Close
  Set RS = nothing

End Sub

</SCRIPT>