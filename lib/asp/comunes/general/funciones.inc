<%

Function MesEspanol(VmesIngles)
	'Response.write "III="&UCASE(VmesIngles)
	Select Case UCASE(VmesIngles)
		Case "JANUARY":
			MesEspanol="Enero"
		Case "FEBRUARY":
			MesEspanol="Febrero"
		Case "MARCH":
			MesEspanol="Marzo"
		Case "APRIL":
			MesEspanol="Abril"
		Case "MAY":
			MesEspanol="Mayo"
		Case "JUNE":
			MesEspanol="Junio"
		Case "JULY":
			MesEspanol="Julio"
		Case "AUGUST":
			MesEspanol="Agosto"
		Case "SEPTEMBER":
			MesEspanol="Septiembre"
		Case "OCTOBER":
			MesEspanol="Octubre"
		Case "NOVEMBER":
			MesEspanol="Noviembre"
		Case "DECEMBER":
			MesEspanol="Diciembre"
		Case "":
			MesEspanol=""
		Case Else:
			MesEspanol=VmesIngles
	End Select
End Function


Function MuestraDiaMesAno(Vfecha)
	If Trim(Vfecha) <> "" Then
		If Day(Vfecha) > 9 Then
			dia = Day(Vfecha)
		Else
			dia = "0" & Day(Vfecha)
		End If
		If Month(Vfecha) > 9 Then
			mes = Month(Vfecha)
		Else
			mes = "0" & Month(Vfecha)
		End If
		MuestraDiaMesAno = dia & "/" & mes & "/" & Year(Vfecha)
	Else
		MuestraDiaMesAno = ""
	End if

End Function

Function Formatear(Vcadena)
	IF Vcadena = "" THEN
		Vcadena = " "
	ELSE
		Vcadena = Replace(Vcadena,"'","|")
		Vcadena = Replace(Vcadena,Chr(13),"<br>")
	END IF
	Formatear = Vcadena
End Function

Function DesFormatear(Vcadena)
	Vcadena = Replace(Vcadena,"|","'")
	Vcadena = Replace(Vcadena,"<br>",Chr(13))
	DesFormatear = Vcadena
End Function

Function ValNulo(Valor, Tipo)
    If IsNull(Valor) Or Trim(Valor) = "" Then
        Select Case Tipo
            Case "N" 'Numerico
               ValNulo = 0
            Case "C" 'Caracteres
                ValNulo = ""
            Case "F" 'Fecha
               ValNulo = Format("01-01-1900", "DD-MM-YYYY")
       End Select
    Else
        If Tipo = "N" Then
             ValNulo = Valor
        ElseIf Tipo = "F" Then
            ValNulo = Valor
        ElseIf Tipo = "C" Then
            ValNulo = Trim(Valor)
        End If
    End If
End Function


Function FN(Valor, Cant)
	FN=Trim(FormatNumber(ValNulo(Valor,"N"), Cant))
End Function

Function PoneIzq(Valor,intCaract)
	If Len(Trim(Valor)) = 1 Then
		PoneIzq = Trim(intCaract) & Trim(valor)
	Else
		PoneIzq = Trim(valor)
	End if
End Function

Function PN(Valor)
	If Trim(valor)="" or isNull(Valor) Then
		PN = "NULL"
	Else
		PN = Valor
	
	End if
End Function

%>