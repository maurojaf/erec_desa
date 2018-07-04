<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub odbc_ObtenerRegistro ( ByRef Conn, ByRef Registro, ByVal SQLQuery)
  ' On Error Resume Next
  Response.Write "<BR><pre> Obtener Registro SQL=  " & SQLQuery & "</pre>"

  Set RS = Conn.Execute (SQLQuery)
  If RS.EOF Then
    Set Registro = nothing
  Else
    Set Registro = CreateObject("Scripting.Dictionary")
    For Each fldLoop In RS.Fields
    Response.Write "<BR><pre>" & fldLoop.Name & "=" & fldLoop.Value & "</pre>"
    Registro.add UCASE(fldLoop.Name), fldLoop.Value
    Next
  End If
  Rs.Close()
End Sub
</SCRIPT>