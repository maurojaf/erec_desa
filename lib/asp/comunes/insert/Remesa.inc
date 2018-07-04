<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_Remesa( Conn, dicRemesa)
    If Request("strFormMode") = "Nuevo" Then
        strSQLQuery = "INSERT INTO REMESA (COD_REMESA, COD_CLIENTE, NOMBRE, DESCRIPCION, FECHA_LLEGADA, FECHA_CARGA, ACTIVO)"
        strSQLQuery = strSQLQuery & " VALUES (" & dicRemesa("COD_REMESA") & ",'" & dicRemesa("COD_CLIENTE") & "','" & dicRemesa("NOMBRE") & "','" & dicRemesa("DESCRIPCION") & "','" & dicRemesa("FECHA_LLEGADA") & "','" & dicRemesa("FECHA_CARGA") & "'," & dicRemesa("ACTIVO")  & ")"
        'Response.write "strSQLQuery=" & strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT NOMBRE, DESCRIPCION, FECHA_LLEGADA, FECHA_CARGA, ACTIVO FROM REMESA WHERE COD_REMESA = " + Cstr(dicRemesa.Item("COD_REMESA")) & " AND COD_CLIENTE = '" + Cstr(dicRemesa.Item("COD_CLIENTE")) & "'"
    	InsertUpdate Conn, dicRemesa, strSql
    End If
End Sub
</SCRIPT>

