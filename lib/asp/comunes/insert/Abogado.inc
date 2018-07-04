<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_Abogado( Conn, dicAbogado)
    If Request("strFormMode") = "Nuevo" Then
        strSQLQuery = "INSERT INTO ABOGADO (NOM_ABOGADO, ACTIVO)"
        strSQLQuery = strSQLQuery & " VALUES ('" & dicAbogado("NOMABOGADO") & "','" & dicAbogado("ACTIVO") & "')"
        'Response.write strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT NOM_ABOGADO, ACTIVO FROM ABOGADO WHERE ID_ABOGADO = " + Cstr(dicAbogado.Item("ID_ABOGADO"))
        InsertUpdate Conn, dicAbogado, strSql
    End If
End Sub
</SCRIPT>



		 
		 
