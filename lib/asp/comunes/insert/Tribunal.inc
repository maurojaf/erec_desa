<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_Tribunal( Conn, dicTribunal)
    If Request("strFormMode") = "Nuevo" Then
        strSQLQuery = "INSERT INTO TRIBUNAL (NOM_TRIBUNAL, RECEPTOR, DIRECCION, COMUNA, FONO_1, FONO_2,ACTIVO)"
        strSQLQuery = strSQLQuery & " VALUES ('" & dicTribunal("NOM_TRIBUNAL") & "','" & dicTribunal("RECEPTOR") & "','" & dicTribunal("DIRECCION") & "','" & dicTribunal("COMUNA") & "','" & dicTribunal("FONO1") & "','" & dicTribunal("FONO2") & "'," & dicTribunal("ACTIVO") & ")"
        'Response.write strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT NOM_TRIBUNAL, RECEPTOR, DIRECCION, COMUNA, FONO_1, FONO_2, ACTIVO FROM TRIBUNAL WHERE ID_TRIBUNAL = " + Cstr(dicTribunal.Item("ID_TRIBUNAL"))
        InsertUpdate Conn, dicTribunal, strSql
    End If
End Sub
</SCRIPT>



		 
		 
