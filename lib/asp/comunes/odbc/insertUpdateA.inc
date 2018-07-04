<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insertUpdate ( ByRef Conn, ByRef dicRegistro, ByVal strSQLQuery)
  ' On Error Resume Next
   'Response.Write "<pre>" & strSQLQuery & "</pre>"
   'Response.eND

  ' Realiza el UPDATE o INSERT segun el resultado anterior
  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.ActiveConnection = Conn
  RS.CursorType = adOpenKeyset
  RS.LockType = adLockOptimistic
  RS.Open strSQLQuery, Conn
  If RS.EOF Then RS.AddNew
  For Each fldLoop In RS.Fields
    Response.Write "<pre>" & UCASE( fldLoop.Name ) & "=" & Trim(dicRegistro.Item( UCASE( fldLoop.Name ) )) & "</pre>"
    If Trim(dicRegistro.Item( UCASE( fldLoop.Name ) )) = "" Then
        fldLoop.Value = NULL
    ElseIf IsNull( dicRegistro.Item( UCASE( fldLoop.Name ) ) ) Then
        fldLoop.Value = NULL
    Else
    	strValor = dicRegistro.Item( UCASE( fldLoop.Name ) )
    	Response.write "<BR>strValor1=="&strValor
    	Response.write "<BR>fldLoopType =="&CInt( fldLoop.Type )
    	'If CInt( fldLoop.Type ) = 4 Then
	'	strValor = Replace(strValor , ".", ",")
	'	Response.write "<BR>strValor2=="&strValor
	'End If
	'fldLoop.Value = strValor
	fldLoop.Value = strValor
    End If
    Response.Write "*"
  Next
  Response.eND
  Response.Write "<pre>UPDATE()</pre>"
  RS.Update()
  RS.Close
  Set RS = nothing
End Sub
</SCRIPT>