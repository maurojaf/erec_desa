<!--#include file="JavaScriptBuscadores.inc"-->
<!--#include file="JavaScriptChequea.inc"-->
<!--#include file="Rutinas.inc"-->

<SCRIPT LANGUAGE="JavaScript">

function Terminar( PaginaTerminar ) {

    self.location.href= PaginaTerminar ;
    return false
}

</SCRIPT>

<SCRIPT RUNAT=Server LANGUAGE="VBScript">


Dim strFormMode

'-------------------------------------------------------------------------------
' Propósito:    Generar el código de inicio de una celda en HTML
' Asume:        Que posteriormente se cerrará la celda y el tag FONT
' Parámetros:   colSpan - el número de celdas que ocupa a su ancho el campo
'               width   - ancho de la celda
'-------------------------------------------------------------------------------

Function InicioCampo( colSpan, width, title )

    ' Escribe la celda donde va el título
    Response.Write "<TR><TD WIDTH=""10%"" ALIGN=RIGHT NOWRAP VALIGN=TOP>"
    Response.Write "<FONT SIZE=-1><B>&nbsp;&nbsp;" & title & "&nbsp;&nbsp;</B></FONT></TD>"

    ' Escribe el inicio la celda donde va el campo
    Response.Write "<TD NOWRAP"
    If colSpan > 1 Then
        Response.Write " COLSPAN=" & colSpan
    End If
    Response.Write " WIDTH=""90%"" >"
    Response.Write "<FONT SIZE=-1>"
End Function

'-------------------------------------------------------------------------------
' Propósito:  Generar el código de fin de una celda en HTML
' Returns:    (nada)
'-------------------------------------------------------------------------------

Function FinCampo(  )
    Response.Write "</FONT></TD></TR>" & CHR(13)
End Function

Function FormatoNumero( Numero, Largo )
    strNumero = Cstr( Numero )
    If Len( strNumero ) < Largo Then strNumero = Replace(Space(Largo-Len( strNumero )), " ", "0") & strNumero
    FormatoNumero = strNumero
End Function


'-------------------------------------------------------------------------------
' Propósito:    Sustituye el valor Nulo por vacío y elimina los espacios
'               en blanco antes y despues del string
' Entradas:     strVarTemp - el valor destino
' Returns:      El valor procesado
'-------------------------------------------------------------------------------

Function ConvertNull(strVarTemp)
    If IsNull(strVarTemp) Then
        ConvertNull = ""
    Else
        ConvertNull = Trim(strVarTemp)
    End If
End Function

'-------------------------------------------------------------------------------
' Propósito:    Pone apostrofes (") alrededor del string
' Entradas:     strVarTemp - el valor destino
' Retorna:      El valor procesado
'-------------------------------------------------------------------------------

Function QuotedString(strVarTemp)
    If IsNull(strVarTemp) Then
        QuotedString = Chr(34) & Chr(34)
    Else
        QuotedString = Chr(34) & CStr(strVarTemp) & Chr(34)
    End If
End Function

'-------------------------------------------------------------------------------
' Propósito:    Testea el string para ver si es una URL viendo el protocolo
' Entradas:     strVarTemp - el valor destino
' Retorna:      True - si es URL, False no lo es
'-------------------------------------------------------------------------------

Function IsURL(strVarTemp)
    IsURL = True
    If UCase(Left(Trim(strVarTemp), 6)) = "HTTP:/" Then Exit Function
    If UCase(Left(Trim(strVarTemp), 6)) = "FILE:/" Then Exit Function
    If UCase(Left(Trim(strVarTemp), 8)) = "MAILTO:/" Then Exit Function
    If UCase(Left(Trim(strVarTemp), 5)) = "FTP:/" Then Exit Function
    If UCase(Left(Trim(strVarTemp), 8)) = "GOPHER:/" Then Exit Function
    If UCase(Left(Trim(strVarTemp), 6)) = "NEWS:/" Then Exit Function
    If UCase(Left(Trim(strVarTemp), 7)) = "HTTPS:/" Then Exit Function
    If UCase(Left(Trim(strVarTemp), 8)) = "TELNET:/" Then Exit Function
    If UCase(Left(Trim(strVarTemp), 6)) = "NNTP:/" Then Exit Function
    IsURL = False
End Function

'-------------------------------------------------------------------------------
' Propósito:    Chequea si el campo en el recordset es modificable
' Asume:        Que el recordset que contiene el campo está abierto
' Entradas:     strNombreCampo  - el nombre del campo en el recordset
' Retorna:      True si es modificable, False si no lo es
'-------------------------------------------------------------------------------

Function CanUpdateField(strNombreCampo, rsRegistro)
    Dim intUpdatable
    intUpdatable = (adFldUpdatable Or adFldUnknownUpdatable)
    CanUpdateField = True
    If (rsRegistro(strNombreCampo).Attributes And intUpdatable) = False Then
        CanUpdateField = False
    End If
End Function

'-------------------------------------------------------------------------------
' Propósito:    Maneja el despliegue de un campo desde un recordset dependiendo
'               de su tipo de dato, atributos, y el modo actual.
' Asume:        Que el recordset que contiene el campo está abierto
'               Que strFormMode está inicializado
' Entrada:      strNombreCampo  - El nombre del campo en rsRegistro
'               blnEsIdentity   - Indicador si es identity
'               arrForaneo  - Arreglo de registro para una llave foránea
'               rsRegistro  - Result Set del registro actual
'-------------------------------------------------------------------------------

Sub general_MostrarCampo(strNombreCampo, blnEsIdentity, strNombreCampoAux, arrForaneo, rsRegistro)
    Dim blnEsRequerido
    Dim intMaxTamano
    Dim intTamano
    Dim strOption1Estado
    Dim strOption2Estado
    Dim strValorCampo
    Dim nPos

    strValorCampo = ""
    nPos=Instr(strNombreCampo,".")
    Do While nPos > 0
        strNombreCampo= Mid (strNombreCampo, nPos+1)
        nPos=Instr(strNombreCampo,".")
    Loop
    ' If not in Edit form mode then set value to empty so doesn't display
    strValorCampo = ""
    If strFormMode = "Edit" Then strValorCampo = RTrim(rsRegistro(strNombreCampo))

    ' Ve si el campo es requeridochequeando sus atributos
    blnEsRequerido = False

    If (rsRegistro(strNombreCampo).Attributes And adFldIsNullable) = 0 Then
        blnEsRequerido = True
    End If

    ' Setea valores MaxLength y Size dependiendo de los atributos
    intMaxTamano = dfMaxSize

    intTamano = rsRegistro(strNombreCampo).DefinedSize + 2

    If strFormMode <> "Filtrar" Then intMaxTamano = intTamano - 2

    ' Si el campo no es modificable, then handle
    ' it like an Identity column and exit
    If Not CanUpdateField(strNombreCampo, rsRegistro) Then

        ' Special handling if Binary
        Select Case rsRegistro(strNombreCampo).Type
            Case adBinary, adVarBinary, adLongVarBinary 'Binary
                Response.Write "[Binary]"
            Case Else
                Select Case strFormMode
                    Case "Edit"
                        Response.Write ConvertNull(strValorCampo)
                        Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strNombreCampo)
                        Response.Write " VALUE=" & QuotedString(strValorCampo) & " >"
                    Case "Nuevo"
                        Response.Write "[Automático]"
                        Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strNombreCampo)
                        Response.Write " VALUE=" & QuotedString(strValorCampo) & " >"
                    Case "Filtrar"
                        Response.Write "< TYPE=Text NAME=" & QuotedString(strNombreCampo)
                        Response.Write " SIZE=" & intTamano
                        Response.Write " MAXLENGTH=" & intMaxTamano
                        Response.Write " VALUE=" & QuotedString(strValorCampo) & " >"
                End Select
        End Select
        Exit Sub
    End If

    ' Handle lookups using a select and options
    If Not IsNull(arrForaneo) Then
    	strMovRet = UCASE( Mid( rsRegistro.Source, InStr( rsRegistro.Source, "FROM" ) + 5 , 26) )

		Response.Write "<SELECT NAME=" & QuotedString(strNombreCampo) & ">"
		' Add blank entry if not required or in filter mode
        'If Not blnEsRequerido Or strFormMode = "Filtrar" Then
            If (strFormMode = "Filtrar" Or strFormMode = "Nuevo") Then
                Response.Write "<OPTION SELECTED>"
            Else
                Response.Write "<OPTION>"
            End If
        'End If

        ' Loop thru the rows in the array
        For intRow = 0 to UBound(arrForaneo)
            'a = arrForaneo.keys
            'Response.Write "<OPTION VALUE=" & QuotedString(arrForaneo(intRow).Item(0))

            Response.Write "<OPTION VALUE=" & QuotedString(arrForaneo(intRow).Item(strNombreCampo))
            If strFormMode = "Edit" Then
                If ConvertNull(arrForaneo(intRow).Item(strNombreCampo)) = ConvertNull(strValorCampo) Then
                    Response.Write " SELECTED"
                End If
            End If
            Response.Write ">"
            Response.Write ConvertNull(arrForaneo(intRow).Item(strNombreCampoAux))
        Next
        Response.Write "</SELECT>"
        If blnEsRequerido And strFormMode = "Nuevo" Then
            Response.Write "  <IMG BORDER=0 src='images/asterisco.gif' WIDTH=10>"
        End If
        Exit Sub
    End If


    'Acá se puede ver el tipo de datos y el campo HTML que genera dependiendo de cada caso
    'Response.Write "<PRE>" & strNombreCampo & ".Type=" & Cstr(rsRegistro(strNombreCampo).Type)
    'Response.Write "<PRE> Valor=" & Cstr(strValorCampo)
    'Evaluate data type and handle appropriately


    Select Case rsRegistro(strNombreCampo).Type

        Case adBoolean   'Boolean
            If strFormMode = "Filtrar" Then

                strOption1Estado = " >Sí"
                strOption2Estado = " >No"
            Else
                Select Case strValorCampo
                    Case "True", "1", "-1", "Verdadero"
                        strOption1Estado = " CHECKED>Sí"
                        strOption2Estado = " >No"

                    Case "False", "0", "Falso"
                        strOption1Estado = " >Sí"
                        strOption2Estado = " CHECKED>No"
                    Case Else
                        strOption1Estado = " >Sí"
                        strOption2Estado = " CHECKED>No"
                End Select
            End If
		Response.Write "<INPUT TYPE=Radio VALUE=1 NAME=" & QuotedString(strNombreCampo) & strOption1Estado
            	Response.Write "<INPUT TYPE=Radio VALUE=0 NAME=" & QuotedString(strNombreCampo) & strOption2Estado

            If strFormMode = "Filtrar" Then
                Response.Write "<INPUT TYPE=Radio NAME=" & QuotedString(strNombreCampo) & " CHECKED>Neither"
            End If

        Case adBinary, adVarBinary, adLongVarBinary     'Binary
            Response.Write "[Binary]"

        Case adLongVarChar, adLongVarWChar, adVarChar       'Memo, Varchar
            If strNombreCampo = "Estado" And strValorCampo = "" Then    strValorCampo = "Activo"

            If strNombreCampo = "Contenido" Then
                Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=10 COLS=40>"
                Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
                Response.Write "</TEXTAREA>"
            ElseIf strNombreCampo = "Resumen" Then
                Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=3 COLS=40>"
                Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
                Response.Write "</TEXTAREA>"
            ElseIf strNombreCampo = "CARACTERISTICAS" Then
	    	Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=4 COLS=40>"
	    	Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
		Response.Write "</TEXTAREA>"
	    ElseIf strNombreCampo = "DIRECCION" or strNombreCampo = "NOMBRE" Then
	    	Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=45 MAXLENGTH =" & intMaxTamano & " VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "'>"
	    ElseIf strNombreCampo = "PASSWORD" Then
		Response.Write "<INPUT TYPE=PASSWORD NAME=" & QuotedString(strNombreCampo) & " SIZE=" & intTamano & " MAXLENGTH=" & intMaxTamano & " VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "'>"
            ElseIf strNombreCampo = "DesLarga" or strNombreCampo = "COMENTARIOS" Then
                Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=3 COLS=40>"
                Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
                Response.Write "</TEXTAREA>"
	    ElseIf strNombreCampo = "DESCPROPIEDAD" or strNombreCampo = "DESCCLIENTE" Then
	    	Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=4 COLS=100>"
                Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
                Response.Write "</TEXTAREA>"                
            ElseIf strNombreCampo = "REEMPLAZA_A" Then
		If strValorCampo<>"" then
			Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=12 VALUE='" & LlenaRut(Server.HTMLEncode(ConvertNull(strValorCampo))) & "' onChange='return checkRutField(this.value)'>"
		Else
		strModulo = UCASE( Mid( rsRegistro.Source, InStr( rsRegistro.Source, "FROM" ) + 5 , 3) )
			If strModulo = "RH_" Or strModulo = "RQ_"  Then
				Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=" & intMaxTamano & " VALUE=" & QuotedString(strValorCampo) & "' ondblclick=""Selector_Item('../../remun/Administracion/tc_bus_Trabajador.asp', 'SelectorCodigos', 'REEMPLAZA_A', 700, 450); return( false );"">"
			End If
		End if
            Else
                If intMaxTamano > 80 Then
                    intSize=45
                Else
                    intSize=intMaxTamano
                End If







                if trim(strNombreCampo)="RUT_USUARIO" then

                    Response.Write "<INPUT TYPE=TEXT onblur=ValidaRut(this) NAME=" & QuotedString(strNombreCampo) & _
                        " SIZE=" & Cstr(intSize) & " MAXLENGTH=" & Cstr(intMaxTamano) & " VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "'>"

                elseif trim(strNombreCampo)="CLAVE" then

                    Response.Write "<INPUT TYPE=PASSWORD disabled NAME=" & QuotedString(strNombreCampo) & _
                        " SIZE=" & Cstr(intSize) & " MAXLENGTH=" & Cstr(intMaxTamano) & " VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "'>"

                else

                    Response.Write "<INPUT TYPE=TEXT NAME=" & QuotedString(strNombreCampo) & _
                        " SIZE=" & Cstr(intSize) & " MAXLENGTH=" & Cstr(intMaxTamano) & " VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "'>"

                end if













            End If

        Case adDateTime                                 'DateTime (Format 'DD/MM/YYYY')
          strDate = ""
          If strFormMode = "Edit" Then
           if not isnull(rsRegistro(strNombreCampo)) then
              dayDate = Cstr( Day(rsRegistro(strNombreCampo)) )
              monthDate = Cstr( Month(rsRegistro(strNombreCampo)) )
              yearDate = Cstr( Year(rsRegistro(strNombreCampo)) )
              If Len( dayDate   ) = 1 Then dayDate = "0" & dayDate
              If Len( monthDate ) = 1 Then monthDate = "0" & monthDate
              strDate = dayDate & "/" & monthDate & "/" & yearDate
             else
             	strDate=Empty
            end if
          End If

		  Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo)
		  Response.Write " SIZE=12"
		  Response.Write " MAXLENGTH=10"

          If strNombreCampo = "FECHA_REQUISICION" And strFormMode = "Nuevo" Then
			  Response.Write " VALUE='" & FormatoNumero( Day(Now()), 2 ) & "/" & FormatoNumero( Month(Now()) , 2 ) & "/" & FormatoNumero( Year(Now()) , 4 ) & "'"
		  Else
			  Response.Write " VALUE='" & strDate & "'"
		  End If
		  Response.Write " ONCHANGE='chkFecha(this);'"
		  Response.Write " VALUE=""" & strDate & """ > (DD/MM/YYYY)"

        Case Else
            Dim nType
            nType=rsRegistro(strNombreCampo).Type

            If (nType <> adVarChar) and (nType <> adWVarChar) and (nType <> adBSTR) and (nType <> adChar) and (nType <> adWChar)  Then
                intTamano = (intTamano-2)*1.5+2
                If strFormMode <> "Filtrar" Then intMaxTamano = intTamano - 2
            End If
            If (nType = 135)  Then
            ' Fecha - Hora
                intTamano = 20
                intMaxTamano = 20
            End If

            If blnEsIdentity Then
                Select Case strFormMode
                    Case "Edit"
                        Response.Write ConvertNull(strValorCampo)
                        Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strNombreCampo)
                        Response.Write " VALUE=" & QuotedString(strValorCampo) & " >"
                    Case "Nuevo"
                        Response.Write "[Automático]"
                        Response.Write "<INPUT TYPE=Hidden NAME=" & QuotedString(strNombreCampo)
                        Response.Write " VALUE=" & QuotedString(strValorCampo) & " >"
                    Case "Filtrar"
                        Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) & " SIZE=" & tEntradasize
                        Response.Write " MAXLENGTH=" & tMaxSize & " VALUE=" & QuotedString(strValorCampo) & " >"
                End Select
            Else
                If intTamano > 80 Then
                  Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=2 COLS=40>"
                  Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
                  Response.Write "</TEXTAREA>"
                ElseIf strNombreCampo = "RUT" Then
                    If strValorCampo<>"" then
                        Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=12 VALUE='" & LlenaRut(Server.HTMLEncode(ConvertNull(strValorCampo))) & "' onChange='return checkRutField(this.value)'>"
                    Else
                    	strModulo = UCASE( Mid( rsRegistro.Source, InStr( rsRegistro.Source, "FROM" ) + 5 , 3) )
                    	If strModulo = "RH_" Or strModulo = "RQ_"  Then
	                    	Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=12 VALUE=" & QuotedString(strValorCampo) & "' ondblclick=""Selector_Item('../../remun/Administracion/tc_bus_Trabajador.asp', 'SelectorCodigos', 'RUT', 700, 450); return( false ); "" onChange='return checkRutField(this.value)'>"
	                	Else
	                	    Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=12 VALUE=" & QuotedString(strValorCampo) & "' ondblclick=""Selector_Item('../../remun/Administracion/tc_bus_MaestroRut.asp', 'SelectorCodigos', 'RUT', 700, 450); return( false );"" onChange='return checkRutField(this.value)'>"
		                End If
                    End if
		ElseIf strNombreCampo = "PASSWORD" Then
			Response.Write "<INPUT TYPE=PASSWORD NAME=" & QuotedString(strNombreCampo) & " SIZE=" & intTamano & " MAXLENGTH=" & intMaxTamano & " VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "'>"
                Else
                  Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo)
                  Response.Write " SIZE=" & intTamano
                  Response.Write " MAXLENGTH=" & intMaxTamano
                  Response.Write " VALUE=" & QuotedString(strValorCampo) & " >"
                  ' Check for special field types
                  Select Case UCase(Left(rsRegistro(strNombreCampo).Name, 4))
                    Case "IMG_"
                        If strValorCampo <> "" Then
                            Response.Write "<BR><BR><IMG SRC=" & QuotedString(strValorCampo) & "><BR>&nbsp;<BR>"
                        End If
                    Case "URL_"
                        If strValorCampo <> "" Then
                            Response.Write "&nbsp;&nbsp;<A HREF=" & QuotedString(strValorCampo) & ">"
                            Response.Write "Go"
                            Response.Write "</A>"
                        End If
                    Case Else
                        If IsURL(strValorCampo) Then
                            Response.Write "&nbsp;&nbsp;<A HREF=" & QuotedString(strValorCampo) & ">"
                            Response.Write "Go"
                            Response.Write "</A>"
                        End If
                  End Select
                End If
            End If
    End Select
    If blnEsRequerido And strFormMode = "Nuevo" Then
        Response.Write "  <IMG BORDER=0 src='images/asterisco.gif' WIDTH=10>"
    End If
End Sub
</SCRIPT>
