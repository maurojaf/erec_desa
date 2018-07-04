<%

strTituloTabla="DATOS DEL DEUDOR"

If Trim(hdd_cod_cliente) <> "" Then
    strNombreCliente = TraeNombreClienteSCG(conexionSCG,hdd_cod_cliente)
End if

If Trim(hdd_rut_deudor) <> "" and Trim(hdd_cod_cliente) <> "" Then

    strNombreDeudor = TraeDatosDeudorSCG(conexionSCG,hdd_cod_cliente,hdd_rut_deudor,"NOMBREDEUDOR")
    strDirCompletaDeudor = TraeUltimaDirDeudorSCG(conexionSCG,hdd_rut_deudor,"CALLE") & " " & TraeUltimaDirDeudorSCG(conexionSCG,hdd_rut_deudor,"NUMERO") & " " & TraeUltimaDirDeudorSCG(conexionSCG,hdd_rut_deudor,"RESTO")
    strComunaDeudor = TraeUltimaDirDeudorSCG(conexionSCG,hdd_rut_deudor,"COMUNA")
    strTelefonoDeudor = TraeUltimoFonoDeudorSCG(conexionSCG,hdd_rut_deudor,"CODAREA") & "-" & TraeUltimoFonoDeudorSCG(conexionSCG,hdd_rut_deudor,"TELEFONO")
    strMovilDeudor = "0" & TraeUltimoMovilDeudorSCG(conexionSCG,hdd_rut_deudor,"CODAREA") & "-" & TraeUltimoMovilDeudorSCG(conexionSCG,hdd_rut_deudor,"TELEFONO")
    If len(strTelefonoDeudor) = 1  Then strTelefonoDeudor = Replace(strTelefonoDeudor,"-","")
    If len(strMovilDeudor) = 2 Then strMovilDeudor = Replace(strMovilDeudor,"0-","")
    If Trim(strMovilDeudor) = "" Then strMovilDeudor = "S/F"
    If Trim(strTelefonoDeudor) = "" Then strTelefonoDeudor = "S/F"
    strEmail=""
    strCodPostal=""
    If Trim(strEmail) = "" Then strEmail = "NO REGISTRADO"
    If Trim(strCodPostal) = "" Then strCodPostal = "NO REGISTRADO"
End If

%>
<br>
<table width="95%" border="0" cellspacing="0" cellpadding="0" class="SycFondoTableAdm">
    <tr>
        <td align="center" valign="top">
            <table border="0" width="100%" cellspacing="1" cellpadding="1" class="SycFondoTableAdm">
                <tr class="SycFondoTitTabAdm" height=25>
                    <td align="center" colspan=8><font class="TituloDatos"><%=UCASE(strTituloTabla)%></font></td>
                </tr>
                <tr>
                    <td class="DatosBlanco" width="15%"><font class="LabelDatos">&nbsp;R.U.T.</font></td>
                    <td class="DatosDeudorTexto" width="35%"><font class="TextoDatos">&nbsp;<%=hdd_rut_deudor%></font></td>
                    <td class="DatosBlanco" width="15%"><font class="LabelDatos">&nbsp;DIRECCION:</font></td>
                    <td class="DatosDeudorTexto" width="35%"><font class="TextoDatos">&nbsp;<%=strDirCompletaDeudor%></font></td>
                </tr>
                <tr>
                    <td class="DatosBlanco"><font class="LabelDatos">&nbsp;NOMBRE:</font></td>
                    <td class="DatosDeudorTexto"><font class="TextoDatos">&nbsp;<%=strNombreDeudor%></font></td>
                    <td class="DatosBlanco"><font class="LabelDatos">&nbsp;COMUNA:</font></td>
                    <td class="DatosDeudorTexto"><font class="TextoDatos">&nbsp;<%=strComunaDeudor%></font></td>
                </tr>
                <tr>
                    <td class="DatosBlanco" width="16%"><font class="LabelDatos">&nbsp;EMAIL:</font></td>
                    <td class="DatosDeudorTexto" width="16%"><font class="TextoDatos">&nbsp;<%=strEmail%></font></td>
                    <td class="DatosBlanco"><font class="LabelDatos">&nbsp;COD.POSTAL:</font></td>
                    <td class="DatosDeudorTexto"><font class="TextoDatos">&nbsp;<%=strCodPostal%></font></td>
                </tr>
                <tr>
                    <td class="DatosBlanco" width="16%"><font class="LabelDatos">&nbsp;FONO RED FIJA:</font></td>
                    <td class="DatosDeudorTexto" width="16%"><font class="TextoDatos">&nbsp;<%=strTelefonoDeudor%></font></td>
                    <td class="DatosBlanco" width="16%"><font class="LabelDatos">&nbsp;M�VIL:</font></td>
                    <td class="DatosDeudorTexto" width="16%"><font class="TextoDatos">&nbsp;<%=strMovilDeudor%></font></td>
                </tr>
            </table>
        </td>
    </tr>
</table>

