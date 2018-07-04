<SCRIPT LANGUAGE=VBScript RUNAT=Server>
' --------------------------------------------------------------------------------------------
' Procedimiento : ObtenerNumFilasPorPagina
' Descripción: Obtiene el número de filas a mostrar en cada página según la
'              columna NumFilPorPag de la tabla de parámetros generales
' Argumentos:
'             ByRef NumFilasPorPagina  - valor obtenido
' Retorna:
'             Retorna -1 en el parámetro para indicar el error.
'
Sub ObtenerNumFilasPorPagina ( ByRef Conn, ByRef NumFilasPorPagina)
'	' Obtiene los atributos
'	SQLQuery = "SELECT numreg_porpag FROM Param_Gen"
'    'Response.Write "<pre>" & SQLQuery & "</pre>"
'	On Error Resume Next	
'	Set RS = Conn.Execute (SQLQuery)
'	If Err Then
'		If Not IsEmpty(RS) Then
'			If Not RS Is Nothing Then
'				RS.Close
'				Set RS = Nothing
'			End If
'		End If
'		Call vs_MostrarError ("Error al obtener los datos de la Base de Datos", true)
'	End If
'	If rs.EOF Then
'		NumFilasPorPagina = -1
'	Else
'		NumFilasPorPagina = CInt( RS("numreg_porpag") )
'		RS.Close
'		Set RS = Nothing
'	End If
	
	NumFilasPorPagina = 10
End Sub
</SCRIPT>
