<SCRIPT RUNAT=Server LANGUAGE="VBScript">

Function FormatoNumero( Numero, Largo )
    strNumero = Cstr( Numero )
    If Len( strNumero ) < Largo Then strNumero = Replace(Space(Largo-Len( strNumero )), " ", "0") & strNumero
    FormatoNumero = strNumero
End Function

</SCRIPT>
