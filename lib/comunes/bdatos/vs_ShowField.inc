<SCRIPT RUNAT=Server LANGUAGE="VBScript">
Dim strFormMode

'---- FieldAttributeEnum Values ----
Const adFldUpdatable = &H00000004
Const adFldUnknownUpdatable = &H00000008
Const adFldIsNullable = &H00000020

'---- DataTypeEnum Values ----
Const adUnsignedTinyInt = 17
Const adBoolean = 11
Const adLongVarChar = 201
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205
Const adVarChar = 200
Const adWVarChar = 202
Const adBSTR = 8
Const adChar = 129
Const adWChar = 130
Const adDateTime = 135
'---- Other Values ----
Const dfMaxSize = 100

'-------------------------------------------------------------------------------
' Propósito:  Generar el código de inicio de una celda en HTML
' Asume: 	  Que posteriormente se cerrará la celda y el tag FONT
' Parámetros: colSpan	- el número de celdas que ocupa a su ancho el campo
'             width  	- ancho de la celda
' Returns:    (nada)
' Hecho por:  Carlos Navarro (29/11/1998)
'-------------------------------------------------------------------------------

Function InicioCampo( colSpan, width, title )

	' Escribe la celda donde va el título
	Response.Write "<TR BGCOLOR=" & strBgColDet & "><TD WIDTH=""10%"" ALIGN=LEFT NOWRAP VALIGN=TOP>"
	Response.Write "<FONT SIZE=-2>&nbsp;&nbsp;" & title & "&nbsp;&nbsp;</FONT></TD>"

	' Escribe el inicio la celda donde va el campo
	Response.Write "<TD"
	If colSpan > 1 Then
		Response.Write " COLSPAN=" & colSpan
	End If
	Response.Write " WIDTH=""90%"" >"
	Response.Write "<FONT SIZE=-1>"
End Function

'-------------------------------------------------------------------------------
' Propósito:  Generar el código de fin de una celda en HTML
' Returns:    (nada)
' Hecho por:  Carlos Navarro (29/11/1998)
'-------------------------------------------------------------------------------

Function FinCampo(  )
	Response.Write "</FONT></TD></TR>" & CHR(13)
End Function

'-------------------------------------------------------------------------------
' Purpose:  Substitutes Empty for Null and trims leading/trailing spaces
' Inputs:   varTemp	- the target value
' Returns:	The processed value
'-------------------------------------------------------------------------------

Function ConvertNull(varTemp)
	If IsNull(varTemp) Then
		ConvertNull = ""
	Else
		ConvertNull = Trim(varTemp)
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Embeds bracketing quotes around the string
' Inputs:   varTemp	- the target value
' Returns:	The processed value
'-------------------------------------------------------------------------------

Function QuotedString(varTemp)
	If IsNull(varTemp) Then
		QuotedString = Chr(34) & Chr(34)
	Else
		QuotedString = Chr(34) & CStr(varTemp) & Chr(34)
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Tests string to see if it is a URL by looking for protocol
' Inputs:   varTemp	- the target value
' Returns:	True - if is URL, False if not
'-------------------------------------------------------------------------------

Function IsURL(varTemp)
	IsURL = True
	If UCase(Left(Trim(varTemp), 6)) = "HTTP:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 6)) = "FILE:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 8)) = "MAILTO:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 5)) = "FTP:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 8)) = "GOPHER:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 6)) = "NEWS:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 7)) = "HTTPS:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 8)) = "TELNET:/" Then Exit Function
	If UCase(Left(Trim(varTemp), 6)) = "NNTP:/" Then Exit Function
	IsURL = False
End Function

'-------------------------------------------------------------------------------
' Purpose:  Tests whether the field in the recordset is updatable
' Assumes: 	That the recordset containing the field is open
' Inputs:   strFieldName	- the name of the field in the recordset
' Returns:	True if updatable, False if not
'-------------------------------------------------------------------------------

Function CanUpdateField(strFieldName, resultSet)
	Dim intUpdatable
	intUpdatable = (adFldUpdatable Or adFldUnknownUpdatable)
	CanUpdateField = True
	If (resultSet(strFieldName).Attributes And intUpdatable) = False Then
		CanUpdateField = False
	End If
End Function

'-------------------------------------------------------------------------------
' Purpose:  Handles the display of a field from a recordset depending
'			on its data type, attributes, and the current mode.
' Assumes: 	That the recordset containing the field is open
'			That strFormMode is initialized
' Inputs:   strFieldName 	- the name of the field in the recordset
'			strLabel		- the label to display
'			blnIdentity		- identity field flag
'			avarLookup		- array of lookup values
'           resultSet		- Result Set (por Carlos). Permite generar list box para otras tablas y campos en general., ya que por defecto toma la tabla inicial del "wizard"
' Observaciones:
'	1. Se eliminaron los tags TD y TR para generar la tabla fuera de la función.
'	2. Se agregó el tipo 200 (varchar en Oracle, para que aparezcan como TEXTAREA en vez de INPUT
'-------------------------------------------------------------------------------

Sub ShowField(strFieldName, strLabel, blnIdentity, avarLookup, resultSet)
	Dim blnFieldRequired
	Dim intMaxSize
	Dim intInputSize
	Dim strOption1State
	Dim strOption2State
	Dim strFieldValue
	Dim nPos

	strFieldValue = ""
	nPos=Instr(strFieldName,".")
	Do While nPos > 0
		strFieldName= Mid (strFieldName, nPos+1)
		nPos=Instr(strFieldName,".")
	Loop
	' If not in Edit form mode then set value to empty so doesn't display
	strFieldValue = ""
	If strFormMode = "Edit" Then strFieldValue = RTrim(resultSet(strFieldName))

	' See if the field is required by checking the attributes
	blnFieldRequired = False
	If (resultSet(strFieldName).Attributes And adFldIsNullable) = 0 Then
		blnFieldRequired = True
	End If

	' Set values for the MaxLength and Size attributes
	intMaxSize = dfMaxSize

	intInputSize = resultSet(strFieldName).DefinedSize + 2

	If strFormMode <> "Filtrar" Then intMaxSize = intInputSize - 2

	' If the field is not updatable, then handle
	' it like an Identity column and exit
	If Not CanUpdateField(strFieldName, resultSet) Then

		' Special handling if Binary
		Select Case resultSet(strFieldName).Type
			Case adBinary, adVarBinary, adLongVarBinary		'Binary
				Response.Write "[Binary]"
			Case Else
				Select Case strFormMode
					Case "Edit"
						Response.Write ConvertNull(strFieldValue)
						Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strFieldName)
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
					Case "Nuevo"
						Response.Write "[Automático]"
						Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strFieldName)
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
					Case "Filtrar"
						Response.Write "< TYPE=Text NAME=" & QuotedString(strFieldName)
						Response.Write " SIZE=" & intInputSize
						Response.Write " MAXLENGTH=" & intMaxSize
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
				End Select
		End Select
		Exit Sub
	End If

	' Handle lookups using a select and options
	If Not IsNull(avarLookup) Then
		Response.Write "<SELECT NAME=" & QuotedString(strFieldName) & ">"
		' Add blank entry if not required or in filter mode
		If Not blnFieldRequired Or strFormMode = "Filtrar" Then
			If (strFormMode = "Filtrar" Or strFormMode = "Nuevo") Then
				Response.Write "<OPTION SELECTED>"
			Else
				Response.Write "<OPTION>"
			End If
		End If

		' Loop thru the rows in the array
		For intRow = 0 to UBound(avarLookup, 2)
			Response.Write "<OPTION VALUE=" & QuotedString(avarLookup(0, intRow))
            If strFormMode = "Edit" Then
				If ConvertNull(avarLookup(0, intRow)) = ConvertNull(strFieldValue) Then
               		Response.Write " SELECTED"
				End If
            End If
           	Response.Write ">"
			Response.Write ConvertNull(avarLookup(1, intRow))
		Next
		Response.Write "</SELECT>"
		If blnFieldRequired And strFormMode = "Nuevo" Then
			Response.Write "  Obligatorio"
		End If
		Exit Sub
	End If


'Acá se puede ver el tipo de datos y el campo HTML que genera dependiendo de cada caso
'Response.Write "<PRE>" & strFieldName & ".Type=" & Cstr(resultSet(strFieldName).Type)
	' Evaluate data type and handle appropriately
	Select Case resultSet(strFieldName).Type

		Case adBoolean, adUnsignedTinyInt 'Boolean
			If strFormMode = "Filtrar" Then

				strOption1State = " >Sí"
				strOption2State = " >No"
			Else
				Select Case strFieldValue
					Case "True", "1", "-1"
						strOption1State = " CHECKED>Sí"
						strOption2State = " >No"
					Case "False", "0"
						strOption1State = " >Sí"
						strOption2State = " CHECKED>No"
					Case Else
						strOption1State = " >Sí"
						strOption2State = " CHECKED>No"
				End Select
			End If
			Response.Write "<INPUT TYPE=Radio VALUE=1 NAME=" & QuotedString(strFieldName) & strOption1State
			Response.Write "<INPUT TYPE=Radio VALUE=0 NAME=" & QuotedString(strFieldName) & strOption2State
			If strFormMode = "Filtrar" Then
				Response.Write "<INPUT TYPE=Radio NAME=" & QuotedString(strFieldName) & " CHECKED>Neither"
			End If

		Case adBinary, adVarBinary, adLongVarBinary		'Binary
			Response.Write "[Binary]"

		Case adLongVarChar, adLongVarWChar, adVarChar, 203, 202		'Memo, Varchar
			If strFieldName = "Estado" And strFieldValue = "" Then	strFieldValue = "Activo"

			If strFieldName = "Contenido" Then
				Response.Write "<TEXTAREA NAME=" & QuotedString(strFieldName) & " ROWS=10 COLS=60>"
				Response.Write Server.HTMLEncode(ConvertNull(strFieldValue))
				Response.Write "</TEXTAREA>"
			ElseIf strFieldName = "Bajada" Then
				Response.Write "<TEXTAREA NAME=" & QuotedString(strFieldName) & " ROWS=3 COLS=60>"
				Response.Write Server.HTMLEncode(ConvertNull(strFieldValue))
				Response.Write "</TEXTAREA>"
			ElseIf strFieldName = "DesLarga" Then
				Response.Write "<TEXTAREA NAME=" & QuotedString(strFieldName) & " ROWS=3 COLS=60>"
				Response.Write Server.HTMLEncode(ConvertNull(strFieldValue))
				Response.Write "</TEXTAREA>"
			ElseIf strFieldName = "Titulo" Then
				Response.Write "<INPUT TYPE=TEXT NAME=" & QuotedString(strFieldName) & _
				" SIZE=" & Cstr(intSize) & " MAXLENGTH=" & Cstr(intMaxSize) & " VALUE='" & Server.HTMLEncode(ConvertNull(strFieldValue)) & "'>"
			Else
				If intMaxSize > 80 Then
					intSize=80
				Else
					intSize=intMaxSize
				End If
				Response.Write "<INPUT TYPE=TEXT NAME=" & QuotedString(strFieldName) & _
						" SIZE=" & Cstr(intSize) & " MAXLENGTH=" & Cstr(intMaxSize) & " VALUE='" & Server.HTMLEncode(ConvertNull(strFieldValue)) & "'>"
			End If

		Case adDateTime, 7									'DateTime (Format 'DD/MM/YYYY')
		  strDate = ""
		  If strFormMode = "Edit" Then
			  dayDate = Cstr( Day(resultSet(strFieldName)) )
			  monthDate = Cstr( Month(resultSet(strFieldName)) )
			  yearDate = Cstr( Year(resultSet(strFieldName)) )
			  If Len( dayDate ) = 1 Then dayDate = "0" & dayDate
			  If Len( monthDate ) = 1 Then monthDate = "0" & monthDate
			  strDate = dayDate & "/" & monthDate & "/" & yearDate
		  End If
		  Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strFieldName)
		  Response.Write " SIZE=12"
		  Response.Write " MAXLENGTH=10"
		  Response.Write " ONCHANGE='chkFecha(this);'"
		  Response.Write " VALUE=""" & strDate & """ > (DD/MM/YYYY)"

		Case Else
			Dim nType
			nType=resultSet(strFieldName).Type

			If (nType <> adVarChar) and (nType <> adWVarChar) and (nType <> adBSTR) and (nType <> adChar) and (nType <> adWChar)  Then
				intInputSize = (intInputSize-2)*3+2
				If strFormMode <> "Filtrar" Then intMaxSize = intInputSize - 2
			End If
			If (nType = 135)  Then
			' Fecha - Hora
				intInputSize = 20
				intMaxSize = 20
			End If

			If blnIdentity Then
				Select Case strFormMode
					Case "Edit"
						Response.Write ConvertNull(strFieldValue)
						Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strFieldName)
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
					Case "Nuevo"
						Response.Write "[Automático]"
						Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strFieldName)
						Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
					Case "Filtrar"
						Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strFieldName) & " SIZE=" & tInputSize
						Response.Write " MAXLENGTH=" & tMaxSize & " VALUE=" & QuotedString(strFieldValue) & " >"
				End Select
			Else
				If intInputSize > 80 Then
				  Response.Write "<TEXTAREA NAME=" & QuotedString(strFieldName) & " ROWS=2 COLS=50>"
				  Response.Write Server.HTMLEncode(ConvertNull(strFieldValue))
				  Response.Write "</TEXTAREA>"
				ElseIf strFieldName = "RUT" Then
					Response.Write "<INPUT TYPE=TEXT NAME=" & QuotedString(strFieldName) & _
							" SIZE=9 MAXLENGTH=9 VALUE='" & Server.HTMLEncode(ConvertNull(strFieldValue)) & "'>"
					Response.Write " <B>-</B> <INPUT TYPE=TEXT NAME='DV' SIZE=1 MAXLENGTH=1 VALUE=''>"
				ElseIf InStr( strFieldName , "Color" ) > 0 Then
				Response.Write "		<TABLE BORDER=1 CELLSPACING=0 CELLPADDING=0>"
				Response.Write "<TR>"
				Response.Write "<TD><INPUT TYPE=TEXT SIZE=6 MAXLENGTH=6 NAME=""" & strFieldName & """ VALUE=""" & Server.HTMLEncode(ConvertNull(strFieldValue)) & """ STYLE=""font-family:Arial; font-size:8pt;""></TD>"
				Response.Write "<TD bgcolor=""#00FFFF""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='00FFFF';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#000000""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='000000';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#0000FF""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='0000FF';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#FF00FF""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='FF00FF';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#808080""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='808080';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#008000""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='008000';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#00FF00""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='00FF00';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#800000""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='800000';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#000080""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='000080';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#808000""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='808000';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#800080""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='800080';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#FF0000""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='FF0000';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#C0C0C0""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='C0C0C0';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#008080""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='008080';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#FFFFFF""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='FFFFFF';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD bgcolor=""#FFFF00""><a href=""#"" onClick=""document.forms[0]." & strFieldName & ".value='FFFF00';""><IMG SRC=""TRANS.GIF"" BORDER=0></a></td>"
				Response.Write "<TD><A HREF='#' onClick=""Selector_Item('colores.asp', 'SelectorProductos', '" & strFieldName & "'); return( false );"">Otro</a></TD>"
				Response.Write "</TR>"
				Response.Write "</TABLE>"
				Else
				  Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strFieldName)
				  Response.Write " SIZE=" & intInputSize
				  Response.Write " MAXLENGTH=" & intMaxSize
				  Response.Write " VALUE=" & QuotedString(strFieldValue) & " >"
				  ' Check for special field types
				  Select Case UCase(Left(resultSet(strFieldName).Name, 4))
				 	Case "IMG_"
						If strFieldValue <> "" Then
							Response.Write "<BR><BR><IMG SRC=" & QuotedString(strFieldValue) & "><BR>&nbsp;<BR>"
						End If
					Case "URL_"
						If strFieldValue <> "" Then
							Response.Write "&nbsp;&nbsp;<A HREF=" & QuotedString(strFieldValue) & ">"
							Response.Write "Go"
							Response.Write "</A>"
						End If
					Case Else
						If IsURL(strFieldValue) Then
							Response.Write "&nbsp;&nbsp;<A HREF=" & QuotedString(strFieldValue) & ">"
							Response.Write "Go"
							Response.Write "</A>"
						End If
				  End Select
				End If
			End If
	End Select
   	If blnFieldRequired And strFormMode = "Nuevo" Then
		Response.Write "  Obligatorio"
	End If
End Sub
</SCRIPT>