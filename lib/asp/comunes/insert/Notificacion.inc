<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_notificacion( Conn, dicNotificacion)
    If Request("strFormMode") = "Nuevo" OR strNuevo = "1" Then
        strSQLQuery = "INSERT INTO DEMANDA_NOTIF (ID_DEMANDA, ID_NOTIFICACION, FECHA, VALOR, ID_ESTADO_NOTIF, BOLETA, PATENTE, ID_USUARIO, OBSERVACIONES)"
        strSQLQuery = strSQLQuery & " VALUES (" & dicNotificacion("ID_DEMANDA") & "," & dicNotificacion("ID_NOTIFICACION") & ",'" & dicNotificacion("FECHA") & "'," & dicNotificacion("VALOR") & "," & dicNotificacion("ID_ESTADONOTIF") & ",'" & dicNotificacion("BOLETA") & "','" & dicNotificacion("PATENTE") & "'," & dicNotificacion("ID_USUARIO") & ",'" & dicNotificacion("OBSERVACIONES") & "')"
        'Response.write strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT FECHA, VALOR, ID_ESTADO_NOTIF, BOLETA, PATENTE, ID_USUARIO, OBSERVACIONES FROM DEMANDA_NOTIF WHERE ID_DEMANDA = " + Cstr(dicNotificacion.Item("ID_DEMANDA")) & " AND ID_NOTIFICACION = " + Cstr(dicNotificacion.Item("ID_NOTIFICACION"))
        InsertUpdate Conn, dicNotificacion, strSql
    End If
End Sub
</SCRIPT>



		 
		 
