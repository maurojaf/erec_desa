<%



Function DownloadFile(file)

  Dim strAbsFile
  Dim strFileExtension
  Dim objFSO
  Dim objFile
  Dim objStream
  strAbsFile = Server.MapPath(file)

  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

  If objFSO.FileExists(strAbsFile) Then
	Set objFile = objFSO.GetFile(strAbsFile)

	Response.Clear
	Response.AddHeader "Content-Disposition", "attachment; filename=" & objFile.Name
	Response.ContentType = "application/octet-stream"
	Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Open
	objStream.Type = 1
	Response.CharSet = "UTF-8"
	objStream.LoadFromFile(strAbsFile)
	Response.BinaryWrite(objStream.Read)
	objStream.Close
	Set objStream = Nothing
	Set objFile = Nothing
  Else
	Response.Clear
	Response.Write("No such file exists.")
  End If
  Set objFSO = Nothing
End Function



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


Function FormatoFecha(Vdia,Vmes,Vagno)

	' en español
		'FormatoFecha = Vdia & "-" & Vmes & "-" & Vagno
	' en ingles
		FormatoFecha = Vdia & "-" & Vmes & "-" & Vagno

End Function

Function FormatoFechaServ(Vdia,Vmes,Vagno)

	' en ingles
	If Session("intConfServidor") = "Ingles" Then
		FormatoFechaServ = Vmes & "-" & Vdia & "-" & Vagno
	Else
		FormatoFechaServ = Vdia & "-" & Vmes & "-" & Vagno
	End if
End Function

Function FechaIng(strFecha)

	' Cambia formato dd/mm/aaaa por mm/dd/aaaa
	
	FechaIng = Mid(strFecha,4,2) & "/" & Mid(strFecha,1,2) & "/" & Mid(strFecha,7,4)
	
	
End Function

Function FechaMasvida(strFecha)

	' Cambia formato dd/mm/aaaa por aaaa/mm/dd
	
	FechaMasvida = Mid(strFecha,1,2) & "-" & Mid(strFecha,4,2) & "-" & Mid(strFecha,7,4)
	
End Function

Function Link()
	colores = "aLink=Red vLink=Black Link=Black"
	' nego y gris
	Link = colores
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

Function CambiaCero(intValor,intTipo)
	If Trim(intValor) = "0" or Trim(intValor) = "" Then
		If Trim(intTipo) = "N" Then
			CambiaCero="Null"
		End If
		If Trim(intTipo) = "0" Then
			CambiaCero=""
		End If
	Else
		CambiaCero=intValor
	End If

End Function

Function ChequeaImagen(intProducto,intTipo)
	connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../asp/CISA.mdb")
	connStr = Replace(connStr,"asp\asp","asp")
	Set rs = Server.CreateObject("ADODB.Recordset")
	'Response.write "SELECT * FROM FILES WHERE IDFILES = " & intProducto & " AND IDTIPO = '" & intTipo & "'"
	'Response.End
	If Trim(intProducto)="" Then intProducto="9999999"
	If Trim(intTipo)="" Then intProducto="9999999"
	rs.Open "SELECT * FROM FILES WHERE IDFILES = " & intProducto & " AND IDTIPO = '" & intTipo & "'", connStr
	If Not rs.EOF Then
		ChequeaImagen = 1
	Else
		ChequeaImagen = 0
	End If

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
            ValNulo = Cdbl(Valor)
        ElseIf Tipo = "F" Then
            ValNulo = Valor
        ElseIf Tipo = "C" Then
            ValNulo = Trim(Valor)
        End If
    End If
End Function


Function FN(Valor, Cant)
	If Trim(Valor) = "" or IsNull(Valor) Then Valor = 0
	FN=Trim(FormatNumber(Valor, Cant))
End Function

Function InteresCuota(intDiasMora,intMaximaAnual,intMontoCuota)
	If IsNumeric(intDiasMora) and IsNumeric(intMaximaAnual) and IsNumeric(intMontoCuota) Then
		InteresCuota = Round(intDiasMora *  (intMaximaAnual / 360) * intMontoCuota,0)
	Else
		InteresCuota = 0
	End if
End Function


Function PoneIzq(Valor,intCaract)
	If Len(Trim(Valor)) = 1 Then
		PoneIzq = Trim(intCaract) & Trim(valor)
	Else
		PoneIzq = Trim(valor)
	End if
End Function


function formatDate(format, intTimeStamp)
  dim unUDate, A

  ' Test to see if intTimeStamp looks valid. If not, they have passed a normal date
  if not (isnumeric(intTimeStamp)) then
    if isdate(intTimeStamp) then
      intTimeStamp = DateDiff("S", "01/01/1970 00:00:00", intTimeStamp)
    else
      response.write "Date Invalid"
      exit function
    end if
  end if
  
  if (intTimeStamp=0) then 
    unUDate = now()
  else
    unUDate = DateAdd("s", intTimeStamp, "01/01/1970 00:00:00")
  end if

  unUDate = trim(unUDate)

  dim startM : startM = InStr(1, unUDate, "/", vbTextCompare) + 1
  dim startY : startY = InStr(startM, unUDate, "/", vbTextCompare) + 1
  dim startHour : startHour = InStr(startY, unUDate, " ", vbTextCompare) + 1
  dim startMin : startMin = InStr(startHour, unUDate, ":", vbTextCompare) + 1

  dim dateDay : dateDay = mid(unUDate, 1, 2)
  dim dateMonth : dateMonth = mid(unUDate, startM, 2)
  dim dateYear : dateYear = mid(unUDate, startY, 4)
  dim dateHour : dateHour = mid(unUDate, startHour, 2)
  dim dateMinute : dateMinute = mid(unUDate, startMin, 2)
  dim dateSecond : dateSecond = mid(unUDate, InStr(startMin, unUDate, ":", vbTextCompare) + 1, 2)

  format = replace(format, "%Y", right(dateYear, 4))
  format = replace(format, "%y", right(dateYear, 2))
  format = replace(format, "%m", dateMonth)
  format = replace(format, "%n", cint(dateMonth))
  format = replace(format, "%F", monthname(cint(dateMonth)))
  format = replace(format, "%M", left(monthname(cint(dateMonth)), 3))
  format = replace(format, "%d", dateDay)
  format = replace(format, "%j", cint(dateDay))
  format = replace(format, "%h", mid(unUDate, startHour, 2))
  format = replace(format, "%g", cint(mid(unUDate, startHour, 2)))

  if (cint(dateHour) > 12) then
    A = "PM"
  else
    A = "AM"
  end if
  format = replace(format, "%A", A)
  format = replace(format, "%a", lcase(A))

  if (A = "PM") then format = replace(format, "%H", left("0" & dateHour - 12, 2))
  format = replace(format, "%H", dateHour)
  if (A = "PM") then format = replace(format, "%G", left("0" & cint(dateHour) - 12, 2))
  format = replace(format, "%G", cint(dateHour))

  format = replace(format, "%i", dateMinute)
  format = replace(format, "%I", cint(dateMinute))
  format = replace(format, "%s", dateSecond)
  format = replace(format, "%S", cint(dateSecond))
  format = replace(format, "%L", WeekDay(unUDate))
  format = replace(format, "%D", left(WeekDayName(WeekDay(unUDate)), 3))
  format = replace(format, "%l", WeekDayName(WeekDay(unUDate)))
  format = replace(format, "%U", intTimeStamp)
  format = replace(format, "11%O", "11th")
  format = replace(format, "1%O", "1st")
  format = replace(format, "12%O", "12th")
  format = replace(format, "2%O", "2nd")
  format = replace(format, "13%O", "13th")
  format = replace(format, "3%O", "3rd")
  format = replace(format, "%O", "th")

  formatDate = format

end function

function calcula_base_exponente(base, exponente)
   'response.write "Base: " & base
   'response.write "<br>Exponente: " & exponente
   
   if (exponente = 0) then
      if (base = 0) then
         'cero elevado a cero = indeterminación
         calcula_base_exponente = null
      else
         'cualquier número elevado a 0 es 1
         calcula_base_exponente = 1
      end if
   else
      'para exponentes o potencias que no son el cero
      dim resultado
      resultado = 1
      'desde 1 hasta el exponente
      for i=1 to exponente
         'tenemos que multiplicar por la base
         resultado = resultado * base
      next
      calcula_base_exponente = resultado
   end if
end function 

Function PonerPrimeraLetraEnMayuscula(texto)

	if len(texto) > 0 then
		textoMayuscula=UCase(left(texto,1)) & Lcase(mid(texto, 2, len(texto)))
	end if
	PonerPrimeraLetraEnMayuscula = textoMayuscula

End Function


%>