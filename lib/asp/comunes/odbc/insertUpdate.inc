<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insertUpdate ( ByRef Conn, ByRef dicRegistro, ByVal strSQLQuery)
  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.ActiveConnection = Conn
  RS.CursorType = adOpenKeyset
  RS.LockType = adLockOptimistic
  RS.Open strSQLQuery, Conn
  If RS.EOF Then RS.AddNew
  For Each fldLoop In RS.Fields
    If Trim(dicRegistro.Item( UCASE( fldLoop.Name ) )) = "" Then
        fldLoop.Value = NULL
    ElseIf IsNull( dicRegistro.Item( UCASE( fldLoop.Name ) ) ) Then
        fldLoop.Value = NULL
    Else
    	strValor = dicRegistro.Item( UCASE( fldLoop.Name ) )
		If CInt( fldLoop.Type ) = 4 or CInt( fldLoop.Type ) = 5 or CInt( fldLoop.Type ) = 6 Then
			strValor = Replace(strValor , ".", ",")
		End If
		fldLoop.Value = strValor
    End If
  Next
  RS.Update()
  RS.Close
  Set RS = nothing
End Sub
</SCRIPT>