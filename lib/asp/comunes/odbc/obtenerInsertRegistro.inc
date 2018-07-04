<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub odbc_ObtenerInsertRegistro ( ByRef Conn, ByVal Tabla, ByVal SQLQuery)
  ' On Error Resume Next
  Response.Write "<BR><pre> Obtener Insert Registro SQL=" & SQLQuery & "</pre>"

  Set RS = Conn.Execute (SQLQuery)
  If RS.EOF Then
    Set Registro = nothing
    Response.Write "<BR><pre> No existe el registro </pre>"
  Else
    Response.Write "<BR><pre> INSERT INTO " & Tabla & " VALUES("
    Set Registro = CreateObject("Scripting.Dictionary")
    For Each fldLoop In RS.Fields
    'Response.Write "'" & fldLoop.Name & "=" & fldLoop.Value & "',"
    Response.Write "'" & fldLoop.Value & "',"
    Registro.add UCASE(fldLoop.Name), fldLoop.Value
    Next
  End If
  Rs.Close()
End Sub
</SCRIPT>