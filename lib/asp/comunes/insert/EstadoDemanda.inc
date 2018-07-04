<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_EstadoDemanda( Conn, dicEstadoDemanda)
    If Request("strFormMode") = "Nuevo" Then
        strSQLQuery = "INSERT INTO ESTADO_DEMANDA (NOM_ESTADO_DEMANDA, ACTIVO)"
        strSQLQuery = strSQLQuery & " VALUES ('" & dicEstadoDemanda("NOM_ESTADO_DEMANDA") & "','" & dicEstadoDemanda("ACTIVO") & "')"
        'Response.write strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT NOM_ESTADO_DEMANDA, ACTIVO FROM ESTADO_DEMANDA WHERE ID_ESTADO_DEMANDA = " + Cstr(dicEstadoDemanda.Item("ID_ESTADO_DEMANDA"))
        InsertUpdate Conn, dicEstadoDemanda, strSql
    End If
End Sub
</SCRIPT>



		 
		 
