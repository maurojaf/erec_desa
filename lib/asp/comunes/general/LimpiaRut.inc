<SCRIPT RUNAT=Server LANGUAGE="VBScript">
function LimpiaRut(strVarRut)
    Dim strRutTemp
    ' Obtiene los parametros'
    
    ' Quita caracteres extraños al campo rut, si los tuviera
    For i=1 TO  LEN(strVarRut) 
        strRutTemp=Mid(strVarRut,i,1)
        If ( strRutTemp <> " " AND strRutTemp <> "." AND strRutTemp <> "-" ) then
            strTemp = strTemp & strRutTemp
        End if
    Next
    If LEN(strTemp) = 7 then
        strTemp = "00" & strTemp
    End If
    If LEN(strTemp) = 8 then
        strTemp = "0" & strTemp
    End If
    LimpiaRut = strTemp
End Function
</SCRIPT>
