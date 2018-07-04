<SCRIPT RUNAT=Server LANGUAGE="VBScript">
function SoloNumeros(strVarRut)
    Dim strRutTemp
    ' Obtiene los parametros'
    
    ' Quita caracteres extraños al campo rut, si los tuviera
    For i=1 TO  LEN(strVarRut) 
        strRutTemp=Mid(strVarRut,i,1)
        If (IsNumeric(strRutTemp)) then
            strTemp = strTemp & strRutTemp
        End if
    Next
    SoloNumeros = strTemp
End Function
</SCRIPT>