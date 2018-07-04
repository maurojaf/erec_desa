<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_Actuario( Conn, dicActuario)
    If Request("strFormMode") = "Nuevo" Then
        strSQLQuery = "INSERT INTO ACTUARIO (NOM_ACTUARIO, COD_INTERNO, ACTIVO)"
        strSQLQuery = strSQLQuery & " VALUES ('" & dicActuario("NOMACTUARIO") & "','" &  dicActuario("COD_INTERNO") & "'," & dicActuario("ACTIVO") & ")"
        'Response.write strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT NOM_ACTUARIO, COD_INTERNO, ACTIVO FROM ACTUARIO WHERE ID_ACTUARIO = " + Cstr(dicActuario.Item("ID_sACTUARIO"))
        InsertUpdate Conn, dicActuario, strSql
    End If
End Sub
</SCRIPT>
	


		 
		 
