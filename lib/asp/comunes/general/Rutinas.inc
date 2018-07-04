<!--#include file="LlenaRut.inc"-->
<!--#include file="LimpiaRut.inc"-->
<!--#include file="LlenaMonto.inc"-->


<SCRIPT RUNAT=Server LANGUAGE="VBScript">

'-------------------------------------------------------------------------------
' Propósito:    
' Entradas:     
' Returns:      
'-------------------------------------------------------------------------------

Function CHNULL(strVar, strTipo)

    If strTipo = "N" then
    	If isNull(strVar) or Trim(strVar) = "" then
    		strVar = 0
    	End If
    End If
    If strTipo = "C" then
		If isNull(strVar) or Trim(strVar) = "" then
			strVar = ""
		End If
    End If
    If strTipo = "D" then
		If isNull(strVar) or Trim(strVar) = "" then
			strVar = "01/01/1900"
		End If
    End If
    
    CHNULL = strVar

End Function

'-------------------------------------------------------------------------------
' Propósito: Formatear el campo período del modo AAAA-MM o AAAA-M
' Entradas : strVarPer, el cual contiene el valor de un período dado
' Returns  : Retorna el Período en uno de los formatos antes mencionados
'-------------------------------------------------------------------------------

Function LlenaPeriodo(strVarPer)
    Dim strPerTemp
    ' Obtiene los parametros
    ' Pone caracter - al campo Periodo
    
    
    If strVarPer = Null then
    	strPerTemp=""
    Else
    	Largo = Len(Trim(CStr(strVarPer)))
    	strVarPer = Trim(CStr(strVarPer))
   
    	If Largo=0 then
    		strPerTemp=""
    	End if
   	If Largo = 5 then
	        strPerTemp = Mid(strVarPer,1,4)&"-"&Mid(strVarPer,5,1)
    	End if
   	
   	
    	If Largo = 6 then
        	strPerTemp = Mid(strVarPer,1,4)&"-"&Mid(strVarPer,5,2)
    	End if
    	    	    	
    End if
    LlenaPeriodo = strPerTemp
    
End Function
'-------------------------------------------------------------------------------
' Propósito: Formatear el campo período del modo AAAA-MM o AAAA-M y separa el mes
' Entradas : strPer, el cual contiene el valor de un período dado
' Returns  : Retorna el mes en palabras
'-------------------------------------------------------------------------------

Function NombrePeriodo(strPer)
Dim strPerTMP
Dim valMes
Dim valAno

' Obtiene los parametros
' Pone caracter - al campo Periodo
    
If strPer = Null then
    	strPerTMP=""
Else
    	Largo = Len(Trim(CStr(strPer)))
    	strPer = Trim(CStr(strPer))
   
    	If Largo=0 then
    		strPerTMP=""
    	End if
   	If Largo = 5 then
	        valAno = cint(Mid(strPer,1,4))
	        valMes = cint(Mid(strPer,5,1))
	        
    	End if
   	If Largo = 6 then
        	valAno = cint(Mid(strPer,1,4))
        	valMes = cint(Mid(strPer,5,2))
    	End if

    	if Cint(valmes)<>0 then
    		select case valMes
    		case "1"
    			strPerTMP="Enero"
    		case "2"
    			strPerTMP="Febrero"
    		case "3"
    			strPerTMP="Marzo"
    		case "4"
    			strPerTMP="Abril"
    		case "5"
    			strPerTMP="Mayo"
    		case "6"
    			strPerTMP="Junio"
    		case "7"
    			strPerTMP="Julio"
    		case "8"
    			strPerTMP="Agosto"
    		case "9"
    			strPerTMP="Septiembre"
    		case "10"
    			strPerTMP="Octubre"
    		case "11"
    			strPerTMP="Noviembre"
    		case "12"
    			strPerTMP="Diciembre"
    			
    		end select
    	end if
    	    	    	
End if
if strPerTMP="" then
	NombrePeriodo=""
else
	NombrePeriodo = strPerTMP &" - "& valAno
end if

End Function


'-------------------------------------------------------------------------------
' Propósito: Validar campo período del de acuerdo al formato de entrada
' Entradas : strPeriodo, strFormato
' Returns  : Retorna True o False
'-------------------------------------------------------------------------------

Function ValidaPeriodo(strPeriodo,strFormato)
	    
	ValidaPeriodo = "True"
	If Trim(strFormato) = "aaaamm" then

		If Len(Trim(strPeriodo)) = 5 Then
			strPeriodo = Mid(strPeriodo,1,4) & "0" & Mid(strPeriodo,5,1)
		End If

		If Len(Trim(strPeriodo)) < 5 OR Len(Trim(strPeriodo)) > 6 Then
			ValidaPeriodo = "False"
		Else
			If Cint(Mid(strPeriodo,1,4)) < 2000 then 
				ValidaPeriodo = "False"
			End If
			If (Cint(Mid(strPeriodo,5,2)) > 12) OR (Cint(Mid(strPeriodo,5,2)) < 1) then 
				ValidaPeriodo = "False"
			End If
		End If

	End If

End Function


'-------------------------------------------------------------------------------
' Propósito: Completa string con el caracter indicado
' Entradas : cadena,largo,caracter
' Returns  : la cadena con el caracter indicado
'-------------------------------------------------------------------------------

Function completa_caracter(cadena,largo,caracter,sentido)
	If IsNull(cadena) then
		cadena = String(largo, caracter)
	Else
		If Len(cadena) < largo then
			llena=largo-len(cadena)
			If sentido = "D" then
				cadena = cadena & String(llena, caracter)
			Else
				cadena = String(llena, caracter) & cadena
			End If
		Else
			cadena = Mid(cadena,1,largo)
		End If
	End If
	completa_caracter=cadena
End Function



'-------------------------------------------------------------------------------
' Propósito: Obtiene el nombre del mes
' Entradas : número de mes
' Returns  : la cadena con el nombre de mes
'-------------------------------------------------------------------------------

Function NombreMes(nro)
	cadena = "Mes no defenido"
	select case nro
	case 0 cadena = "Capital Inicial"	
	case 1 cadena = "Enero"
	case 2 cadena = "Febrero"
	case 3 cadena = "Marzo"
	case 4 cadena = "Abril"
	case 5 cadena = "Mayo"
	case 6 cadena = "Junio"
	case 7 cadena = "Julio"
	case 8 cadena = "Agosto"
	case 9 cadena = "Septiembre"
	case 10 cadena = "Octubre"
	case 11 cadena = "Noviembre"
	case 12 cadena = "Diciembre"
	End Select
	NombreMes=cadena
End Function
	
	
	
'-------------------------------------------------------------------------------
' Propósito: Busca descripcion
' Entradas : codigo a buscar, largo de la busqueda y una coleccion
' Returns  : la cadena con la descripcion
'-------------------------------------------------------------------------------

Function BuscarAsignacion(codigo,Largo,arrAsigna)
Dim cont

  For cont=1 To Largo
	IF TRIM(arrAsigna(cont).item("DONDEASIGNA")) = Trim(Codigo) Then
			response.write arrAsigna(cont).item("DES_PARAMETRO")
	End if
  Next

End Function

'-------------------------------------------------------------------------------
' Propósito : Formatear fecha como dd/mm/yyyy
' Entradas  : fecha no formateada
' Returns   : cadena con fecha en formato dd/mm/yyyy
'-------------------------------------------------------------------------------

Function SetearFecha(fecha)
Dim dia, mes, agno, estring

dia = Day(fecha)
mes = Month(fecha)
agno= Year(fecha)
estring = ""

if (dia < 10) then
	estring = "0"+Trim(dia)+"/"
else
	estring = Trim(dia)+"/"
end if

if (mes < 10) Then
	estring = estring+"0"+Trim(mes)+"/"+Trim(agno)
else
	estring = estring+Trim(mes)+"/"+Trim(agno)
end if


SetearFecha = estring
End Function


'-------------------------------------------------------------------------------
' Propósito : formatear el campo estado.
' Entradas  : valor booleano
' Returns   : cadena indicando si es abierto o cerrado
'
Function ObtenerValor(valor)
Dim temp

if IsNull(valor) Then
	temp = ""
else
	if (valor) Then
		temp="Abierta"
	else
		if not(valor) then
			temp="cerrada"
		end if
	end if
end if
ObtenerValor=temp
End Function
	
</SCRIPT>