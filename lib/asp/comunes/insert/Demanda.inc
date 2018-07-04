<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_Demanda( Conn, dicDemanda)
    If Request("strFormMode") = "Nuevo" Then
        strSQLQuery = "INSERT INTO DEMANDA (RUT_DEUDOR,COD_CLIENTE, ID_TRIBUNAL, ROLANO, FECHA_INGRESO, FECHA_CADUCIDAD, ID_PROCURADOR, ID_ABOGADO, ID_ACTUARIO, MONTO, FECHA_COMPARENDO, HORA_COMPARENDO, RAZON_TERMINO, GASTOS_JUDICIALES, HONORARIOS, INTERESES, INDEM_COMPENSATORIA, TOTAL_APAGAR, ID_ESTADO, FECHA_PRES_ESCRITO, TIPO_DEMANDA)"
        strSQLQuery = strSQLQuery & " VALUES ('" & dicDemanda("RUT_DEUDOR") & "','" & dicDemanda("COD_CLIENTE") & "'," & dicDemanda("ID_TRIBUNAL") & " , '" &dicDemanda("ROLANO") & "','" & dicDemanda("FECHA_INGRESO") & "','" & dicDemanda("FECHA_CADUCIDAD") & "'," & dicDemanda("ID_PROCURADOR") & "," & dicDemanda("ID_ABOGADO") & "," & dicDemanda("ID_ACTUARIO") & "," & dicDemanda("MONTO") & ",'" & dicDemanda("FECHA_COMPARENDO") & "','" & dicDemanda("HORA_COMPARENDO") & "'," & dicDemanda("RAZON_TERMINO") & "," & dicDemanda("GASTOS_JUDICIALES") & "," & dicDemanda("HONORARIOS") & "," & dicDemanda("INTERESES") & "," & dicDemanda("INDEM_COMPENSATORIA") & "," & dicDemanda("TOTAL_APAGAR") & "," & dicDemanda("ID_ESTADO") & ",'" & dicDemanda("FECHA_PRES_ESCRITO") & "','" & dicDemanda("TIPO_DEMANDA") & "')"
        'Response.write strSQLQuery
        'Response.End
        set RS = Conn.execute(strSQLQuery)
    Else
    	strSql="SELECT RUT_DEUDOR, COD_CLIENTE, ID_TRIBUNAL, ROLANO, FECHA_INGRESO, FECHA_CADUCIDAD, ID_PROCURADOR, ID_ABOGADO, ID_ACTUARIO, MONTO,FECHA_COMPARENDO, HORA_COMPARENDO, RAZON_TERMINO, GASTOS_JUDICIALES, HONORARIOS, INTERESES, INDEM_COMPENSATORIA, TOTAL_APAGAR, ID_ESTADO, FECHA_PRES_ESCRITO, TIPO_DEMANDA FROM DEMANDA WHERE ID_DEMANDA = " + Cstr(dicDemanda.Item("ID_DEMANDA"))
        'Response.write strSql
        'Response.End 
        InsertUpdate Conn, dicDemanda, strSql
    End If
End Sub
</SCRIPT>



		 
		 
