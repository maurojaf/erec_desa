<SCRIPT LANGUAGE=VBScript RUNAT=Server>
' --------------------------------------------------------------------------------------------
' Función    : Minimo
' Descripción: Obtiene el mínimo entre 2 valores
' Argumentos : los 2 valores
' Retorna    : El minimo de los 2 valores del argumento

function general_Minimo (val1, val2)
  If val1 <= val2 Then
    general_Minimo = val1
  Else
    general_Minimo = val2
  End If
End Function

Function FormatDateCompare(dDate)
dim year
dim month
dim day
dim syear
dim smonth
dim sday
dim sdate
	year = datepart("yyyy", dDate)
	month = datepart("m", dDate)
	day = datepart("d", dDate)
	'		response.write dDate & " : dDate    " &  year & " : year        " & month & " : month      " & day & " : day     "'
	syear = year
	if month < 10 then
		smonth = "0" & month
	else
		smonth = month
	end if
	if day < 10 then
		sday = "0" & day
	else
		sday = day
	end if
	sdate = smonth & "/" & sday & "/" & syear
	FormatDateCompare = sdate
end function
</SCRIPT>
