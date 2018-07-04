<SCRIPT RUNAT=Server LANGUAGE="VBScript">
'-------------------------------------------------------------------------------
' Propósito: Dar formato de salida a campos de tipo money
' Entradas : strVarMonoto
' Returns  : Retorna el valor monetario con puntos
'-------------------------------------------------------------------------------

Function LlenaMonto(strVarMonto)
    Dim strMontoTemp
    ' Obtiene los parametros
    ' Pone caracteres . y - al campo rut
    
    if strVarMonto="" then
    	exit function
    end if

    
    If strVarMonto = Null or Trim(strVarMonto)="" then
    	strMontoTemp="0"
    Else
        if IsNumeric(strVarMonto)  then
    		strMontoTemp = FormatNumber(strVarMonto,0)
    	else
    		strVarMonto =0 
    	end if
    End if
    LlenaMonto = strMontoTemp
End Function
	
</SCRIPT>