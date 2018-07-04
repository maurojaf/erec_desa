<SCRIPT RUNAT=Server LANGUAGE="VBScript">

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
        If Trim(strNombreCampo) = "COD_TIPTRATO" Then
 		Response.Write "<SELECT NAME=" & QuotedString(strNombreCampo) & " onChange=""ValidaTrato(this.value);return false;"">"

        Else
             	If Trim(strNombreCampo) = "TIPO_ASOCF" Then
 			Response.Write "<SELECT NAME=" & QuotedString(strNombreCampo) & " onChange=""MuestraSeccion(this.value);return false;"">"
        	else

        		If Trim(strNombreCampo) = "TIPO_GRUPO" Then
		 		Response.Write "<SELECT NAME=" & QuotedString(strNombreCampo) & " onChange=""VerGrupo(this.value);return false;"">"
        		else
        			Response.Write "<SELECT NAME=" & QuotedString(strNombreCampo) & ">"
        		end if
        	end if
        End If
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
            Response.Write "  <IMG BORDER=0 src='../../imagenes/asterisco.gif' WIDTH=10>"
        End If
        Exit Sub
    End If


    'Acá se puede ver el tipo de datos y el campo HTML que genera dependiendo de cada caso
    'Response.Write "<PRE>" & strNombreCampo & ".Type=" & Cstr(rsRegistro(strNombreCampo).Type)
    'Evaluate data type and handle appropriately


    Select Case rsRegistro(strNombreCampo).Type

        Case adBoolean              'Boolean
            If strFormMode = "Filtrar" Then

                strOption1Estado = " >Sí"
                strOption2Estado = " >No"
            Else
                Select Case strValorCampo
                    Case "True", "1", "-1"
                        strOption1Estado = " CHECKED>Sí"
                        strOption2Estado = " >No"

                    Case "False", "0"
                        strOption1Estado = " >Sí"
                        strOption2Estado = " CHECKED>No"
                    Case Else
                        strOption1Estado = " >Sí"
                        strOption2Estado = " CHECKED>No"
                End Select
            End If
		If strNombreCampo = "TRASPASA_DH" Then
			strOption1Estado = " >Debe"
                	strOption2Estado = " >Haber"
                	Select Case strValorCampo
			   Case "True", "1", "-1"
			        strOption1Estado = " CHECKED>Debe"
			        strOption2Estado = " >Haber"

			   Case "False", "0"
			        strOption1Estado = " >Debe"
			        strOption2Estado = " CHECKED>Haber"
			   Case Else
			        strOption1Estado = " >Debe"
			        strOption2Estado = " >Haber"
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
                Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=10 COLS=60>"
                Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
                Response.Write "</TEXTAREA>"
            ElseIf strNombreCampo = "Resumen" Then
                Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=3 COLS=60>"
                Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
                Response.Write "</TEXTAREA>"
			ElseIf strNombreCampo = "PASSWORD" Then
					Response.Write "<INPUT TYPE=PASSWORD NAME=" & QuotedString(strNombreCampo) & " SIZE=" & intTamano & " MAXLENGTH=" & intMaxTamano & " VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "'>"
            ElseIf strNombreCampo = "DesLarga" Then
                Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=3 COLS=60>"
                Response.Write Server.HTMLEncode(ConvertNull(strValorCampo))
                Response.Write "</TEXTAREA>"
            ElseIf strNombreCampo = "FUNRES" or strNombreCampo = "EXPCON" Then
                Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=3 COLS=60>"
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
                    intSize=80
                Else
                    intSize=intMaxTamano
                End If
                Response.Write "<INPUT TYPE=TEXT NAME=" & QuotedString(strNombreCampo) & _
                        " SIZE=" & Cstr(intSize) & " MAXLENGTH=" & Cstr(intMaxTamano) & " VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "'>"
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
		  Response.Write " SIZE=10"
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
                intTamano = (intTamano-2)*3+2
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
                  Response.Write "<TEXTAREA NAME=" & QuotedString(strNombreCampo) & " ROWS=2 COLS=50>"
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
                ElseIf InStr( strNombreCampo, "ROL_" ) > 0  Then
				    Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=" & intMaxTamano & " VALUE=" & QuotedString(strValorCampo) & "' ondblclick=""Selector_Item_Descripcion('../Administracion/rh_bus_Maestro.asp', 'SelectorCodigos', '" & strNombreCampo & "', '" & strNombreCampoAux & "', 700, 380); return( false );"">"
                ElseIf InStr(strNombreCampo, "COD_CTACON")>0 Then
				    Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=" & intMaxTamano & " VALUE=" & QuotedString(strValorCampo) & "' ondblclick=""Selector_Item('../../remun/Administracion/tc_bus_cuentascontables.asp', 'SelectorCodigos', '" & strNombreCampo & "', 700, 380); return( false );"">"
                ElseIf InStr(strNombreCampo, "CTA_")>0 Then
				    Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=" & intMaxTamano & " VALUE=" & QuotedString(strValorCampo) & "' ondblclick=""Selector_Item('../../remun/Administracion/tc_bus_cuentascontables.asp', 'SelectorCodigos', '" & strNombreCampo & "', 700, 380); return( false );"">"
				ElseIf strNombreCampo = "BRUTO" Then
					Response.Write "<INPUT TYPE=Text NAME=" & QuotedString(strNombreCampo) &" SIZE=12 MAXLENGTH=12 VALUE='" & Server.HTMLEncode(ConvertNull(strValorCampo)) & "' onChange='return cambiaLiquido(this.value)'>"
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
        Response.Write "  <IMG BORDER=0 src='../../imagenes/asterisco.gif' WIDTH=10>"
    End If
End Sub
</SCRIPT>
