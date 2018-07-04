<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_Procurador( Conn, dicProcurador)
    If Request("strFormMode") = "Nuevo" Then
        strSQLQuery = "INSERT INTO PROCURADOR (NOM_PROCURADOR, ACTIVO)"
        strSQLQuery = strSQLQuery & " VALUES ('" & dicProcurador("NOMPROCURADOR") & "','" & dicProcurador("ACTIVO") & "')"
        'Response.write strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT NOM_PROCURADOR, ACTIVO FROM PROCURADOR WHERE ID_PROCURADOR = " + Cstr(dicProcurador.Item("ID_PROCURADOR"))
        InsertUpdate Conn, dicProcurador, strSql
    End If
End Sub
</SCRIPT>



		 
		 
