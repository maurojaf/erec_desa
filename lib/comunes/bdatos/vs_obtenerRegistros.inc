<SCRIPT LANGUAGE=VBScript RUNAT=Server>
' --------------------------------------------------------------------------------------------
' Procedimiento : ObtenerRegistros
' Descripción: Obtiene todos los registros de una tabla y los coloca en un arreglo de
'              dos dimensiones: la primera contiene las columnas de la tabla y la
'              segunda contiene las filas.
' Argumentos:
'             ByRef Conn               - Objeto de conexión, asume que está abierta.
'             ByRef ArrRegistros       - Arreglo para colocar los registros
'             ByRef NumRegistros       - Número de registros recuperados
'             ByVal SQLQuery           - String con la sentencia SQL
'
' Retorna:
'             Si la sentencia SQL está vacía, retorna -1 en el número de registros
'             para indicar el error.
'
Sub ObtenerRegistros ( ByRef Conn, ByRef ArrRegistros, ByRef NumRegistros, ByVal SQLQuery)
  ' Obtiene los Registros utilizando un cursor (RS) y almacenandolos en un arreglo
  ' On Error Resume Next
  'Response.Write "<pre>" & SQLQuery & "</pre>"
  Set RS = Conn.Execute (SQLQuery)
  If Err Then
    If Not IsEmpty(RS) Then
      If Not RS Is Nothing Then
        RS.Close
        Set RS = Nothing
      End If
    End If
    Call vs_MostrarError ("Error al obtener los datos de la Base de Datos: " & SQLQuery, true)
  End If
  If RS.EOF Then
    NumRegistros = 0
  Else
    ArrRegistros = RS.GetRows()
    RS.Close
    Set RS = Nothing
    NumRegistros = UBound( ArrRegistros, 2) + 1
  End If
End Sub
</SCRIPT>