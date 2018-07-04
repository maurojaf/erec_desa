<SCRIPT RUNAT=Server LANGUAGE="VBScript">
Function LlenaRut(strVarRut)
    Dim strRutTemp
    ' Obtiene los parametros
    ' Pone caracteres . y - al campo rut
    
    
    If (strVarRut = Null) or IsNull(strVarRut) then
    	strRutTemp=""
    Else
    	Largo = Len(Trim(CStr(strVarRut)))
    	strVarRut = Trim(CStr(strVarRut))
   
    	If Largo=0 then
    		strRutTemp=""
    	End if
   
       	If Largo=0 then
       		strRutTemp=""
       	End if
      
       	If Largo = 1 then
           	strRutTemp = strVarRut
       	End if
       	If Largo = 2 then
           	strRutTemp = Left(strVarRut,1) & "-" & Right(strVarRut,1)
       	End if
       	If Largo = 3 then
           	strRutTemp = Left(strVarRut,2) & "-" & Right(strVarRut,1)
       	End if
       	If Largo = 4 then
           	strRutTemp = Mid(Right(strVarRut,4),1,3) & "-" & Right(strVarRut,1)
       	End if
       	If Largo = 5 then
           	strRutTemp = Left(strVarRut,1) & "." & Mid(Right(strVarRut,4),1,3) & "-" & Right(strVarRut,1)
       	End if
       	If Largo = 6 then
           	strRutTemp = Left(strVarRut,2) & "." & Mid(Right(strVarRut,4),1,3) & "-" & Right(strVarRut,1)
       	End if
    	If Largo = 7 then
        	strRutTemp = Mid(Right(strVarRut,7),1,3) & "." & Mid(Right(strVarRut,4),1,3) & "-" & Right(strVarRut,1)
    	End if
    	If Largo = 8 then
    	    strRutTemp = Left(strVarRut,1) & "." & Mid(Right(strVarRut,7),1,3) & "." & Mid(Right(strVarRut,4),1,3) & "-" & Right(strVarRut,1)
    	End if
    	If Largo = 9 then
    	    strRutTemp = Left(strVarRut,2) & "." & Mid(Right(strVarRut,7),1,3) & "." & Mid(Right(strVarRut,4),1,3) & "-" & Right(strVarRut,1)
    	End if
    End if
    LlenaRut = strRutTemp
End Function
</SCRIPT>