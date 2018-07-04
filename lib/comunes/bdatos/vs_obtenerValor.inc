<SCRIPT LANGUAGE=VBScript RUNAT=Server>
' --------------------------------------------------------------------------------------------
' Funcion    : ObtenerValor
' Descripción: Obtiene un valor resultante de una sentencia SQL.
' Argumentos:
'             ByRef Conn               - Objeto de conexión, asume que está abierta.
'             ByVal SQLQuery           - String con la sentencia SQL
'             ByRef Estado             - Booleano que indica si resultó OK la operación
'                                        True o False
'                                       
' Retorna:    Si la sentencia SQL está vacía, retorna Null. De lo contrario retorna el valor. 
'             Se debe chequear el Estado.
'
Function ObtenerValor (ByRef Conn, ByVal SQLQuery, ByRef Estado)
  Dim RS, valor
  
  'Response.Write "<pre>" & SQLQuery & "</pre>"

  ' On Error Resume Next  
  If IsEmpty(SQLQuery) OR IsNull(SQLQuery) Then
    Estado = False
  ObtenerValor = Null
  Exit Function
  End If

  Set RS = Conn.Execute (SQLQuery)
  If Err Then
    If Not IsEmpty(RS) Then
      If Not RS Is Nothing Then
        RS.Close
        Set RS = Nothing
      End If
    End If
    Estado = False
    Call vs_MostrarError ("Error al obtener los datos de la Base de Datos: " & SQLQuery, true)
  End If
  
  Estado = True
  If rs.EOF Then
    ObtenerValor = Null
  Else
    valor = RS(0)
    If (valor = "") Or IsNull(valor) Then valor=1
    ObtenerValor = valor
    RS.Close
    Set RS = Nothing
  End If
End Function
</SCRIPT>