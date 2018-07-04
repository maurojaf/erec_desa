<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub odbc_ObtenerRegistros ( ByRef Conn, ByRef ArrRegistros, ByRef NumRegistros, ByVal SQLQuery)
  ' On Error Resume Next
   'Response.Write "<pre>" & SQLQuery & "</pre>"
  'Response.Write "Nº Reg = " & NumRegistros
  
  Dim Registro, NumFilas
  NumFilas = -1
  Set RS = Conn.Execute (SQLQuery)
  If RS.EOF Then
        NumFilas = -1
        ArrRegistros = Null
  Else
    ReDim ArrRegistros( 0 )
    While Not RS.EOF
        NumFilas = NumFilas + 1
        If NumFilas > 0 Then ReDim Preserve ArrRegistros( NumFilas )
        Set Registro = nothing
        Set Registro = CreateObject("Scripting.Dictionary")
        For Each fldLoop In RS.Fields
            Registro.add UCASE(fldLoop.Name), fldLoop.Value
        Next
        RS.MoveNext
        Set ArrRegistros( NumFilas ) = Registro
    Wend

  End If
  NumRegistros = NumFilas
End Sub
</SCRIPT>