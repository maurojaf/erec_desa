<%
Function TraeSiNo(strValorCampo)
	Select Case strValorCampo
	    Case "True", "1", "-1", "Verdadero"
		TraeSiNo = "Si"
	    Case "False", "0", "Falso"
		TraeSiNo = "No"
	    Case Else
		TraeSiNo = "No"
	End Select
End Function

%>