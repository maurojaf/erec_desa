<%

Function TraeFechaActual(strConex)
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT convert(varchar(10),getdate(),103) as strFechaActual ", strConex
    If Not rs.EOF Then
        TraeFechaActual = Trim(rs("strFechaActual"))
    End If
End Function

Function TraeFechaActual2(strConex,strTipo)
    Set rs = Server.CreateObject("ADODB.Recordset")
    
    If Trim(strTipo) = "D" Then
        strSql="SELECT Day(getdate()) as FEC "
    End if
    If Trim(strTipo) = "M" Then
        strSql="SELECT substring(convert(char(2),getdate(),101),1,2) as FEC"
    End if
    If Trim(strTipo) = "A" Then
        strSql="SELECT Year(getdate()) as FEC "
    End if
    
    'Response.write "strSql="&strSql
    'Response.End
    
    rs.Open strSql, strConex
    If Not rs.EOF Then
        TraeFechaActual2 = Trim(rs("FEC"))
    Else
        TraeFechaActual2 = Trim(rs("FEC"))
End If

End Function
Function TraeRutSDR(strConex,NroDoc)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If not isnull(NroDoc) Then
        strSql = ""
        strSql = "SELECT RUT_DEUDOR FROM Cuota WHERE NroDoc = '" & NroDoc & "'"
        rs.Open strSql, strConex
        If Not rs.EOF Then
            TraeRutSDR = Trim(rs("RUT_DEUDOR"))
        Else
            TraeRutSDR = "Folio " & NroDoc
        End If
    Else
            TraeRutSDR = "Folio " & NroDoc
    End If
End Function

Function TraeNombreClienteSDR(strConex,intCodCliente)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Trim(intCodCliente) <> "" Then
        rs.Open "SELECT * FROM CLIENTES WHERE COD_CLI = '" & intCodCliente & "'", strConex
        If Not rs.EOF Then
            TraeNombreClienteSDR = Trim(rs("DESC_CLI"))
        Else
            TraeNombreClienteSDR = "Cliente " & intCodCliente
        End If
    Else
            TraeNombreClienteSDR = "Cliente " & intCodCliente
    End If
End Function

Function TraeFechaActual(strConex)
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT convert(varchar(10),getdate(),103) as strFechaActual ", strConex
    If Not rs.EOF Then
        TraeFechaActual = Trim(rs("strFechaActual"))
    End If
End Function

Function TraeDatosDeudorSDR(strConex,intCodCliente,strRut,strCampo)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Trim(intCodCliente) <> "" Then
        strSql=""
        strSql = "SELECT " & strCampo & " FROM DEUDOR WHERE RUT_DEUDOR = '" &  strRut & "' AND CODCLIENTE = '" & intCodCliente & "'"
        'Response.write "strSql="&strSql
        'Response.write strSql
        rs.Open strSql , strConex
        If Not rs.EOF Then
            TraeDatosDeudorSDR = Trim(rs(strCampo))
        Else
            TraeDatosDeudorSDR = ""
        End If
    Else
            TraeDatosDeudorSDR = ""
    End If
End Function

Function TraeDirrSDR(strConex,strRut,strCampo)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Trim(strRut) <> "" Then
        strSql=""
        strSql = "SELECT " & strCampo & " FROM DEUDOR WHERE RUT_DEUDOR = '" &  strRut & "'"
        rs.Open strSql , strConex
        If Not rs.EOF Then
            TraeDirrSDR = rs(strCampo)
        Else
            strSql = "SELECT " & strCampo & " FROM DEUDOR WHERE RUT_DEUDOR = '" &  strRut & "'"
            rs.Close
            rs.Open strSql , strConex
            If Not rs.EOF Then
                TraeDirrSDR = rs(strCampo)
            Else
                TraeDirrSDR = ""
            End If
        End If
    Else
        TraeDirrSDR = ""
    End If
End Function

Function TraeRutPorNroDocSDR(strConex,strNroDoc)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Trim(strNroDoc) <> "" Then
        strSql=""
        strSql = "SELECT DISTINCT RUT_DEUDOR AS RUT_DEUDOR FROM CUOTA WHERE NRODOC = '" &  strNroDoc & "' AND SALDO >= 0"
        rs.Open strSql , strConex
        If Not rs.EOF Then
            TraeRutPorNroDocSDR = rs("RUT_DEUDOR")
        Else
            TraeRutPorNroDocSDR = ""
        End If
    Else
            TraeRutPorNroDocSDR = ""
    End If
End Function

Function TraeProductoSDR(strConex,strCodProd)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Trim(strCodProd) <> "" Then
        strSql=""
        strSql = "SELECT DESCRIPCION FROM PRODUCTO WHERE TIPO_PRODU = '" &  strCodProd & "'"
        rs.Open strSql , strConex
        If Not rs.EOF Then
            TraeProductoSDR = ucase(rs("DESCRIPCION"))
        Else
            TraeProductoSDR = ""
        End If
    Else
        TraeProductoSDR = ""
    End If
End Function

Function TraeTipoDocSDR(strConex,strCodTipo, strCodCli)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Trim(strCodTipo) <> "" and Trim(strCodCli) <> "" Then
        strSql = ""
        strSql = "SELECT DESCRIPCION FROM TIPODOCUMENTO WHERE CODIGO = '" &  strCodTipo & "' AND CODCLIENTE = '" &  strCodCli & "'"
        rs.Open strSql , strConex
        If Not rs.EOF Then
            TraeTipoDocSDR = rs("DESCRIPCION")
        Else
            TraeTipoDocSDR = ""
        End If
    Else
            TraeTipoDocSDR = ""
    End If
End Function

Function TraeEstadoSDR(strConex,strCodEstado)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Trim(strCodEstado) <> "" Then
        strSql = "SELECT DESCRIP FROM ESTADOS WHERE CODIGO = '" &  strCodEstado & "'"
        rs.Open strSql , strConex
        If Not rs.EOF Then
            TraeEstadoSDR = rs("DESCRIP")
        Else
            TraeEstadoSDR = ""
        End If
    Else
            TraeEstadoSDR = ""
    End If
End Function

Function TraeTelefono(strConex,strRut)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Trim(strRut) <> "" Then
        strSql = "exec parsea_telefono '" &  strRut & "'"
        rs.Open strSql , strConex
        If Not rs.EOF Then
            TraeTelefono = rs("COD_AREA")&"-"&rs("telefono")
        Else
            TraeTelefono = ""
        End If
    Else
            TraeTelefono = ""
    End If
End Function


%>

