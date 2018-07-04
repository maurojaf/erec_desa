<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub odbc_ObtenerNdeNRegistros ( ByRef Conn, ByRef ArrRegistros, ByRef NumRegistros, ByVal SQLQuery, Desde, Cantidad)
  ' On Error Resume Next
  
'Response.Write "<pre>" & SQLQuery & "</pre>"

  Dim Registro
  NumFilas = -1

  Set RS = Conn.Execute (SQLQuery)
  ReDim ArrRegistros( 0 )
  If RS.EOF Then
    	NumFilas = -1
    	ArrRegistros = Null
  Else
  	RS.Move Desde


	Do While Not RS.EOF
	NumFilas = NumFilas + 1

	If Cdbl(NumFilas + 1) >  Cdbl(Cantidad) Then Exit Do
		If NumFilas > 0 Then ReDim Preserve ArrRegistros( NumFilas )
		Set Registro = nothing
		Set Registro = CreateObject("Scripting.Dictionary")
		For Each fldLoop In RS.Fields
			'Response.write UCASE(fldLoop.Name) & "=" & fldLoop.Value
			Registro.add UCASE(fldLoop.Name), fldLoop.Value
		Next
		RS.MoveNext
		Set ArrRegistros( NumFilas ) = Registro
	Loop

  End If
  NumRegistros = NumFilas
End Sub
</SCRIPT>