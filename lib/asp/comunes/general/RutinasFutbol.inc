
<SCRIPT RUNAT=Server LANGUAGE="VBScript">

'-------------------------------------------------------------------------------
' Propósito:
' Entradas:
' Returns:
'-------------------------------------------------------------------------------

Function PTS(intGF, intGC)

    If intGF > intGC then
    	PTS = 3
    End If
    If intGF < intGC then
		PTS = 0
    End If
    If intGF = intGC then
		PTS = 1
    End If

End Function

Function PG(intGF, intGC)

    If intGF > intGC then
    	PG = 1
    Else
    	PG = 0
    End If

End Function

Function PE(intGF, intGC)

    If intGF = intGC then
    	PE = 1
    Else
    	PE = 0
    End If

End Function

Function PP(intGF, intGC)

    If intGF < intGC then
    	PP = 1
    Else
    	PP = 0
    End If

End Function

</SCRIPT>