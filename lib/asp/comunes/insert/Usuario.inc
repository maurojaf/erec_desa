<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_Usuario( Conn, dicUsuario)
    If Request("strFormMode") = "Nuevo" Then
        strSQLQuery = "INSERT INTO USUARIO (RUT_USUARIO, ANEXO, NOMBRES_USUARIO, APELLIDOS_USUARIO, PERFIL, LOGIN, CLAVE, PERFIL_SUP, PERFIL_ADM, PERFIL_COB, PERFIL_CAJA, PERFIL_PROC, PERFIL_FULL, PERFIL_EMP, ACTIVO)"
        strSQLQuery = strSQLQuery & " VALUES ('" & dicUsuario("RUT_USUARIO") & "','" & dicUsuario("ANEXO") & "','" &  dicUsuario("NOMBRES_USUARIO") & "','" & dicUsuario("APELLIDOS_USUARIO") & "','COB_GER','" & dicUsuario("LOGIN") & "','" & dicUsuario("CLAVE") & "'," & dicUsuario("PERFIL_SUP") & "," & dicUsuario("PERFIL_ADM") & "," & dicUsuario("PERFIL_COB") & "," & dicUsuario("PERFIL_CAJA") & "," & dicUsuario("PERFIL_PROC") & "," & dicUsuario("PERFIL_FULL") & "," & dicUsuario("PERFIL_EMP") & "," & dicUsuario("ACTIVO") & ")"
        'Response.write strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT RUT_USUARIO, ANEXO, NOMBRES_USUARIO, APELLIDOS_USUARIO, LOGIN, CLAVE, PERFIL_SUP, PERFIL_ADM, PERFIL_COB, PERFIL_CAJA, PERFIL_PROC, PERFIL_FULL, PERFIL_EMP, ACTIVO FROM USUARIO WHERE ID_USUARIO = " + Cstr(dicUsuario.Item("ID_USUARIO"))
        InsertUpdate Conn, dicUsuario, strSql
    End If
End Sub
</SCRIPT>



		 
		 
