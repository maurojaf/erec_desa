<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub insert_Cliente( Conn, dicCliente)
    If Request("strFormMode") = "Nuevo" Then
        strSql = "INSERT INTO CLIENTE (COD_CLIENTE, DESCRIPCION, NOMBRE_FANTASIA, RAZON_SOCIAL, RUT, DIRECCION, COMUNA, FONO_1, FONO_2, NOM_CONTACTO, EMAIL_CONTACTO, TASA_MAX_CONV, IC_PORC_CAPITAL, HON_PORC_CAPITAL, PIE_PORC_CAPITAL, TIPO_INTERES, ACTIVO, GASTOS_OPERACIONALES, GASTOS_ADMINISTRATIVOS, GASTOS_OPERACIONALES_CD, GASTOS_ADMINISTRATIVOS_CD,"
        strSql = strSql & " ADIC_1, ADIC_2, ADIC_3, ADIC_4, ADIC_5, ADIC_91, ADIC_92, ADIC_93, ADIC_94, ADIC_95, USA_CUSTODIO, COLOR_CUSTODIO, INTERES_MORA, TIPO_CLIENTE, EXPIRACION_CONVENIO, EXPIRACION_ANULACION, COD_MONEDA, COD_TIPODOCUMENTO_HON, MESES_TD_HON,ADIC1_DEUDOR,ADIC2_DEUDOR,ADIC3_DEUDOR,NOMBRE_CONV_PAGARE, RETIRO_SABADO, USA_HONORARIOS, FORMULA_HONORARIOS, USA_INTERESES, FORMULA_INTERESES, FORMULA_HONORARIOS_FACT)"
        strSql = strSql & " VALUES ('" & dicCliente("COD_CLIENTE") & "','" & dicCliente("DESCRIPCION") & "','" & dicCliente("NOMBRE_FANTASIA") & "','" & dicCliente("RAZON_SOCIAL") & "','" & dicCliente("RUT") & "','" & dicCliente("DIRECCION") & "','" & dicCliente("COMUNA") & "','" & dicCliente("FONO_1") & "','" & dicCliente("FONO_2") & "','" & dicCliente("NOM_CONTACTO") & "','" & dicCliente("EMAIL_CONTACTO") & "'," & dicCliente("IC_PORC_CAPITAL") & "," & dicCliente("IC_PORC_CAPITAL") & "," & dicCliente("HON_PORC_CAPITAL") & "," & dicCliente("PIE_PORC_CAPITAL") & ",'" & dicCliente("TIPO_INTERES") & "'," & dicCliente("ACTIVO") & "," & dicCliente("GASTOS_OPERACIONALES") & "," & dicCliente("GASTOS_ADMINISTRATIVOS") & "," & dicCliente("GASTOS_OPERACIONALES_CD") & "," & dicCliente("GASTOS_ADMINISTRATIVOS_CD") 
        strSql = strSql & ",'" & dicCliente("ADIC_1") & "','" & dicCliente("ADIC_2") & "','" & dicCliente("ADIC_3") & "','" & dicCliente("ADIC_4") & "','" & dicCliente("ADIC_5") & "','" & dicCliente("ADIC_91") & "','" & dicCliente("ADIC_92") & "','" & dicCliente("ADIC_93") & "','" & dicCliente("ADIC_94")  & "','" & dicCliente("ADIC_95") & "','" & dicCliente("USA_CUSTODIO") & "','" & dicCliente("COLOR_CUSTODIO") & "'," & dicCliente("INTERES_MORA") & ",'" & dicCliente("TIPO_CLIENTE") & "'," & dicCliente("EXPIRACION_CONVENIO") & "," & dicCliente("EXPIRACION_ANULACION") & ",'" & dicCliente("COD_MONEDA") & "','" & dicCliente("COD_TIPODOCUMENTO_HON") & "'," & dicCliente("MESES_TD_HON") & ",'" & dicCliente("ADIC1_DEUDOR") & "','" & dicCliente("ADIC2_DEUDOR") & "','" & dicCliente("ADIC3_DEUDOR") & "','" & dicCliente("NOMBRE_CONV_PAGARE") & "'," & dicCliente("RETIRO_SABADO") & "," & dicCliente("USA_HONORARIOS") & ",'" & dicCliente("FORMULA_HONORARIOS") & "'," & dicCliente("USA_INTERESES") & ",'" & dicCliente("FORMULA_INTERESES") & "','" & dicCliente("FORMULA_HONORARIOS_FACT") & "')" 
        'Response.write "strSql=" & strSql
        'Response.End
        set RS = Conn.execute(strSql)
    Else
    	strSql="SELECT DESCRIPCION, NOMBRE_FANTASIA, RAZON_SOCIAL, RUT, DIRECCION, COMUNA, FONO_1, FONO_2, NOM_CONTACTO, EMAIL_CONTACTO, TASA_MAX_CONV, IC_PORC_CAPITAL, HON_PORC_CAPITAL, PIE_PORC_CAPITAL, TIPO_INTERES, ACTIVO, GASTOS_OPERACIONALES, GASTOS_ADMINISTRATIVOS , GASTOS_OPERACIONALES_CD, GASTOS_ADMINISTRATIVOS_CD, ADIC_1, ADIC_2, ADIC_3, ADIC_4, ADIC_5, ADIC_91, ADIC_92, ADIC_93, ADIC_94, ADIC_95, USA_CUSTODIO, COLOR_CUSTODIO, INTERES_MORA, TIPO_CLIENTE, EXPIRACION_CONVENIO, EXPIRACION_ANULACION,COD_MONEDA, COD_TIPODOCUMENTO_HON, MESES_TD_HON, ADIC1_DEUDOR,ADIC2_DEUDOR,ADIC3_DEUDOR,NOMBRE_CONV_PAGARE, RETIRO_SABADO, USA_HONORARIOS, FORMULA_HONORARIOS, USA_INTERESES, FORMULA_INTERESES, FORMULA_HONORARIOS_FACT FROM CLIENTE WHERE COD_CLIENTE = '" & Cstr(dicCliente.Item("COD_CLIENTE")) & "'"
    	InsertUpdate Conn, dicCliente, strSql
    End If
End Sub
</SCRIPT>



		 
		 
