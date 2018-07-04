<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insertUpdate ( ByRef Conn, ByRef dicRegistro, ByVal strSQLQuery)
  'On Error Resume Next
   'Response.Write "<pre>" & strSQLQuery & "</pre>"

  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.ActiveConnection = Conn
  RS.CursorType = adOpenKeyset
  RS.LockType = adLockOptimistic
  RS.Open strSQLQuery, Conn
  If RS.EOF Then RS.AddNew
  For Each fldLoop In RS.Fields
 ' Response.Write "<pre>" & UCASE( fldLoop.Name ) & "=" & Trim(dicRegistro.Item( UCASE( fldLoop.Name ) )) & "</pre>"
    If Trim(dicRegistro.Item( UCASE( fldLoop.Name ) )) = "" Then
        fldLoop.Value = NULL
    ElseIf IsNull( dicRegistro.Item( UCASE( fldLoop.Name ) ) ) Then
        fldLoop.Value = NULL
    Else
        fldLoop.Value = dicRegistro.Item( UCASE( fldLoop.Name ) )
    End If
  Next
  RS.Update()
  RS.Close
  Set RS = nothing
End Sub
</SCRIPT>