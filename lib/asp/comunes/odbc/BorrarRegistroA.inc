<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub odbc_BorrarRegistro ( ByRef Conn, ByVal strSQLQuery )
  'On Error Resume Next
  Response.Write "<pre>" & strSQLQuery & "</pre>"

  Set RS = Server.CreateObject("ADODB.Recordset")
Response.Write "<pre>1</pre>"
  RS.ActiveConnection = Conn
Response.Write "<pre>2</pre>"
  RS.CursorType = adOpenDynamic
  RS.LockType = adLockOptimistic
Response.Write "<pre>1</pre>"
  RS.Open strSQLQuery, Conn
  Response.Write ("RS.EOF = " & RS.EOF)
  If Not RS.EOF Then 
	Response.Write "<pre>BORRARA</pre>"
  	'RS.Delete
	  'RS.Update
	  'RS.Close
  End If
  Set RS = nothing

End Sub

</SCRIPT>