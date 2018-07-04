<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->

<%

Response.CodePage=65001
Response.charset ="utf-8"

AbrirSCG()

accion_ajax		=request.querystring("accion_ajax")



if trim(accion_ajax)="guarda_cliente" Then

   		COD_CLIENTE 				=request.querystring("COD_CLIENTE")
		DESCRIPCION 				=request.querystring("DESCRIPCION")
		RAZON_SOCIAL 				=request.querystring("RAZON_SOCIAL")		
		NOMBRE_FANTASIA 			=request.querystring("NOMBRE_FANTASIA")		
		RUT 						=request.querystring("RUT")		
		REP_LEGAL 					=request.querystring("REP_LEGAL")		
		EMAIL_CONTACTO 				=request.querystring("EMAIL_CONTACTO")		
		ACTIVO 						=request.querystring("ACTIVO")		
		TASA_MAX_CONV 				=request.querystring("TASA_MAX_CONV")		
		IC_PORC_CAPITAL 			=request.querystring("IC_PORC_CAPITAL")		
		HON_PORC_CAPITAL 			=request.querystring("HON_PORC_CAPITAL")		
		PIE_PORC_CAPITAL 			=request.querystring("PIE_PORC_CAPITAL")		
		TIPO_INTERES 				=request.querystring("TIPO_INTERES")		
		TIPO_CLIENTE 				=request.querystring("TIPO_CLIENTE")		
		GASTOS_OPERACIONALES 		=request.querystring("GASTOS_OPERACIONALES")		
		GASTOS_ADMINISTRATIVOS 		=request.querystring("GASTOS_ADMINISTRATIVOS")		
		GASTOS_OPERACIONALES_CD 	=request.querystring("GASTOS_OPERACIONALES_CD")		
		GASTOS_ADMINISTRATIVOS_CD 	=request.querystring("GASTOS_ADMINISTRATIVOS_CD")		
		ADIC_1 						=request.querystring("ADIC_1")		
		ADIC_2 						=request.querystring("ADIC_2")		
		ADIC_3 						=request.querystring("ADIC_3")		
		ADIC_4 						=request.querystring("ADIC_4")		
		ADIC_5 						=request.querystring("ADIC_5")		
		ADIC_91 					=request.querystring("ADIC_91")		
		ADIC_92 					=request.querystring("ADIC_92")		
		ADIC_93 					=request.querystring("ADIC_93")		
		ADIC_94 					=request.querystring("ADIC_94")		
		ADIC_95 					=request.querystring("ADIC_95")		
		USA_CUSTODIO 				=request.querystring("USA_CUSTODIO")		
		COLOR_CUSTODIO 				=request.querystring("COLOR_CUSTODIO")		
		ADIC_96 					=request.querystring("ADIC_96")		
		ADIC_97 					=request.querystring("ADIC_97")		
		ADIC_98 					=request.querystring("ADIC_98")		
		ADIC_99 					=request.querystring("ADIC_99")		
		ADIC_100 					=request.querystring("ADIC_100")				
		NOMBRE_CONV_PAGARE 			=request.querystring("NOMBRE_CONV_PAGARE")		
		INTERES_MORA 				=request.querystring("INTERES_MORA")		
		EXPIRACION_CONVENIO 		=request.querystring("EXPIRACION_CONVENIO")		
		EXPIRACION_ANULACION 		=request.querystring("EXPIRACION_ANULACION")		
		COD_MONEDA 					=request.querystring("COD_MONEDA")		
		COD_TIPODOCUMENTO_HON 		=request.querystring("COD_TIPODOCUMENTO_HON")		
		MESES_TD_HON 				=request.querystring("MESES_TD_HON")		
		ADIC1_DEUDOR 				=request.querystring("ADIC1_DEUDOR")		
		ADIC2_DEUDOR 				=request.querystring("ADIC2_DEUDOR")		
		ADIC3_DEUDOR 				=request.querystring("ADIC3_DEUDOR")		
		RETIRO_SABADO 				=request.querystring("RETIRO_SABADO")		
		COD_ULT_GES 				=request.querystring("COD_ULT_GES")		
		OBS_ULT_GES 				=request.querystring("OBS_ULT_GES")		
		FORMULA_HONORARIOS 			=request.querystring("FORMULA_HONORARIOS")		
		FORMULA_HONORARIOS_FACT 	=request.querystring("FORMULA_HONORARIOS_FACT")		
		FORMULA_INTERESES 			=request.querystring("FORMULA_INTERESES")		
		USA_HONORARIOS 				=request.querystring("USA_HONORARIOS")		
		USA_INTERESES 				=request.querystring("USA_INTERESES")		
		USA_SUBCLIENTE 				=request.querystring("USA_SUBCLIENTE")		
		USA_REPLEGAL 				=request.querystring("USA_REPLEGAL")		
		USA_PROTESTOS 				=request.querystring("USA_PROTESTOS")		
		NRO_CLIENTE_DOC 			=request.querystring("NRO_CLIENTE_DOC")		
		NRO_CLIENTE_DEUDOR 			=request.querystring("NRO_CLIENTE_DEUDOR")			
		DIRECCION					=request.querystring("DIRECCION")		 
		USA_COB_INTERNA 			=request.querystring("USA_COB_INTERNA")
		
       strSql = "INSERT INTO CLIENTE (COD_CLIENTE, DESCRIPCION, NOMBRE_FANTASIA, RAZON_SOCIAL, RUT, DIRECCION, EMAIL_CONTACTO, TASA_MAX_CONV, IC_PORC_CAPITAL, HON_PORC_CAPITAL, PIE_PORC_CAPITAL, TIPO_INTERES, ACTIVO, GASTOS_OPERACIONALES, GASTOS_ADMINISTRATIVOS, GASTOS_OPERACIONALES_CD, GASTOS_ADMINISTRATIVOS_CD,"
        strSql = strSql & " ADIC_1, ADIC_2, ADIC_3, ADIC_4, ADIC_5, ADIC_91, ADIC_92, ADIC_93, ADIC_94, ADIC_95, USA_CUSTODIO, COLOR_CUSTODIO, INTERES_MORA, TIPO_CLIENTE, EXPIRACION_CONVENIO, EXPIRACION_ANULACION, COD_MONEDA, COD_TIPODOCUMENTO_HON, MESES_TD_HON,ADIC1_DEUDOR,ADIC2_DEUDOR,ADIC3_DEUDOR,NOMBRE_CONV_PAGARE, RETIRO_SABADO, USA_HONORARIOS, FORMULA_HONORARIOS, USA_INTERESES, FORMULA_INTERESES, FORMULA_HONORARIOS_FACT)"
        strSql = strSql & " VALUES ('" & trim(COD_CLIENTE) & "','" & trim(DESCRIPCION) & "','" & trim(NOMBRE_FANTASIA) & "','" & trim(RAZON_SOCIAL) & "','" & trim(RUT) & "','" & trim(DIRECCION) & "','" & trim(EMAIL_CONTACTO) & "'," & trim(IC_PORC_CAPITAL) & "," & trim(IC_PORC_CAPITAL) & "," & trim(HON_PORC_CAPITAL) & "," & trim(PIE_PORC_CAPITAL) & ",'" & trim(TIPO_INTERES) & "'," & trim(ACTIVO) & "," & trim(GASTOS_OPERACIONALES) & "," & trim(GASTOS_ADMINISTRATIVOS) & "," & trim(GASTOS_OPERACIONALES_CD) & "," & trim(GASTOS_ADMINISTRATIVOS_CD) 
        strSql = strSql & ",'" & trim(ADIC_1) & "','" & trim(ADIC_2) & "','" & trim(ADIC_3) & "','" & trim(ADIC_4) & "','" & trim(ADIC_5) & "','" & trim(ADIC_91) & "','" & trim(ADIC_92) & "','" & trim(ADIC_93) & "','" & trim(ADIC_94)  & "','" & trim(ADIC_95) & "','" & trim(USA_CUSTODIO) & "','" & trim(COLOR_CUSTODIO) & "'," & trim(INTERES_MORA) & ",'" & trim(TIPO_CLIENTE) & "'," & trim(EXPIRACION_CONVENIO) & "," & trim(EXPIRACION_ANULACION) & ",'" & trim(COD_MONEDA) & "','" & trim(COD_TIPODOCUMENTO_HON) & "'," & trim(MESES_TD_HON) & ",'" & trim(ADIC1_DEUDOR) & "','" & trim(ADIC2_DEUDOR) & "','" & trim(ADIC3_DEUDOR) & "','" & trim(NOMBRE_CONV_PAGARE) & "'," & trim(RETIRO_SABADO) & "," & trim(USA_HONORARIOS) & ",'" & trim(FORMULA_HONORARIOS) & "'," & trim(USA_INTERESES) & ",'" & trim(FORMULA_INTERESES) & "','" & trim(FORMULA_HONORARIOS_FACT) & "')" 

        Conn.execute(strSql)
        if err then
        	Response.write err.description
        end if
        'Response.write strSql

elseif trim(accion_ajax)="verifica_cliente" Then
	COD_CLIENTE =request.querystring("COD_CLIENTE")

	sql_sel ="select * from CLIENTE where COD_CLIENTE=" & COD_CLIENTE
	set rs_sel =conn.execute(sql_sel)

	if not rs_sel.eof Then
		%>
			<input type="hidden" name="valida_cliente" id="valida_cliente" value="S">
		<%
	else
		%>
			<input type="hidden" name="valida_cliente" id="valida_cliente" value="N">
		<%
	end if
	'Response.write sql_sel


elseif trim(accion_ajax)="update_cliente" then

   		COD_CLIENTE 				=request.querystring("COD_CLIENTE")
		DESCRIPCION 				=request.querystring("DESCRIPCION")
		RAZON_SOCIAL 				=request.querystring("RAZON_SOCIAL")		
		NOMBRE_FANTASIA 			=request.querystring("NOMBRE_FANTASIA")		
		RUT 						=request.querystring("RUT")		
		REP_LEGAL 					=request.querystring("REP_LEGAL")		
		EMAIL_CONTACTO 				=request.querystring("EMAIL_CONTACTO")		
		ACTIVO 						=request.querystring("ACTIVO")		
		TASA_MAX_CONV 				=request.querystring("TASA_MAX_CONV")		
		IC_PORC_CAPITAL 			=request.querystring("IC_PORC_CAPITAL")		
		HON_PORC_CAPITAL 			=request.querystring("HON_PORC_CAPITAL")		
		PIE_PORC_CAPITAL 			=request.querystring("PIE_PORC_CAPITAL")		
		TIPO_INTERES 				=request.querystring("TIPO_INTERES")		
		TIPO_CLIENTE 				=request.querystring("TIPO_CLIENTE")		
		GASTOS_OPERACIONALES 		=request.querystring("GASTOS_OPERACIONALES")		
		GASTOS_ADMINISTRATIVOS 		=request.querystring("GASTOS_ADMINISTRATIVOS")		
		GASTOS_OPERACIONALES_CD 	=request.querystring("GASTOS_OPERACIONALES_CD")		
		GASTOS_ADMINISTRATIVOS_CD 	=request.querystring("GASTOS_ADMINISTRATIVOS_CD")		
		ADIC_1 						=request.querystring("ADIC_1")		
		ADIC_2 						=request.querystring("ADIC_2")		
		ADIC_3 						=request.querystring("ADIC_3")		
		ADIC_4 						=request.querystring("ADIC_4")		
		ADIC_5 						=request.querystring("ADIC_5")		
		ADIC_91 					=request.querystring("ADIC_91")		
		ADIC_92 					=request.querystring("ADIC_92")		
		ADIC_93 					=request.querystring("ADIC_93")		
		ADIC_94 					=request.querystring("ADIC_94")		
		ADIC_95 					=request.querystring("ADIC_95")		
		USA_CUSTODIO 				=request.querystring("USA_CUSTODIO")		
		COLOR_CUSTODIO 				=request.querystring("COLOR_CUSTODIO")		
		ADIC_96 					=request.querystring("ADIC_96")		
		ADIC_97 					=request.querystring("ADIC_97")		
		ADIC_98 					=request.querystring("ADIC_98")		
		ADIC_99 					=request.querystring("ADIC_99")		
		ADIC_100 					=request.querystring("ADIC_100")				
		NOMBRE_CONV_PAGARE 			=request.querystring("NOMBRE_CONV_PAGARE")		
		INTERES_MORA 				=request.querystring("INTERES_MORA")		
		EXPIRACION_CONVENIO 		=request.querystring("EXPIRACION_CONVENIO")		
		EXPIRACION_ANULACION 		=request.querystring("EXPIRACION_ANULACION")		
		COD_MONEDA 					=request.querystring("COD_MONEDA")		
		COD_TIPODOCUMENTO_HON 		=request.querystring("COD_TIPODOCUMENTO_HON")		
		MESES_TD_HON 				=request.querystring("MESES_TD_HON")		
		ADIC1_DEUDOR 				=request.querystring("ADIC1_DEUDOR")		
		ADIC2_DEUDOR 				=request.querystring("ADIC2_DEUDOR")		
		ADIC3_DEUDOR 				=request.querystring("ADIC3_DEUDOR")		
		RETIRO_SABADO 				=request.querystring("RETIRO_SABADO")		
		COD_ULT_GES 				=request.querystring("COD_ULT_GES")		
		OBS_ULT_GES 				=request.querystring("OBS_ULT_GES")		
		FORMULA_HONORARIOS 			=request.querystring("FORMULA_HONORARIOS")		
		FORMULA_HONORARIOS_FACT 	=request.querystring("FORMULA_HONORARIOS_FACT")		
		FORMULA_INTERESES 			=request.querystring("FORMULA_INTERESES")		
		USA_HONORARIOS 				=request.querystring("USA_HONORARIOS")		
		USA_INTERESES 				=request.querystring("USA_INTERESES")		
		USA_SUBCLIENTE 				=request.querystring("USA_SUBCLIENTE")		
		USA_REPLEGAL 				=request.querystring("USA_REPLEGAL")		
		USA_PROTESTOS 				=request.querystring("USA_PROTESTOS")		
		NRO_CLIENTE_DOC 			=request.querystring("NRO_CLIENTE_DOC")		
		NRO_CLIENTE_DEUDOR 			=request.querystring("NRO_CLIENTE_DEUDOR")			
		DIRECCION					=request.querystring("DIRECCION")		 
		USA_COB_INTERNA 			=request.querystring("USA_COB_INTERNA")

	sql_update ="UPDATE  CLIENTE  "
sql_update = sql_update & " SET  DESCRIPCION='"&trim(DESCRIPCION)&"', RAZON_SOCIAL='"&trim(RAZON_SOCIAL)&"', NOMBRE_FANTASIA='"&trim(NOMBRE_FANTASIA)&"', "
sql_update = sql_update & " RUT='"&trim(RUT)&"', REP_LEGAL='"&trim(REP_LEGAL)&"', EMAIL_CONTACTO='"&trim(EMAIL_CONTACTO)&"', ACTIVO='"&trim(ACTIVO)&"', TASA_MAX_CONV='"&trim(TASA_MAX_CONV)&"', "
sql_update = sql_update & " IC_PORC_CAPITAL='"&trim(IC_PORC_CAPITAL)&"', HON_PORC_CAPITAL='"&trim(HON_PORC_CAPITAL)&"', PIE_PORC_CAPITAL='"&trim(PIE_PORC_CAPITAL)&"', TIPO_INTERES='"&trim(TIPO_INTERES)&"', "
sql_update = sql_update & " TIPO_CLIENTE='"&trim(TIPO_CLIENTE)&"', GASTOS_OPERACIONALES='"&trim(GASTOS_OPERACIONALES)&"', GASTOS_ADMINISTRATIVOS='"&trim(GASTOS_ADMINISTRATIVOS)&"',  "
sql_update = sql_update & " GASTOS_OPERACIONALES_CD='"&trim(GASTOS_OPERACIONALES_CD)&"', GASTOS_ADMINISTRATIVOS_CD='"&trim(GASTOS_ADMINISTRATIVOS_CD)&"', ADIC_1='"&trim(ADIC_1)&"', "
sql_update = sql_update & " ADIC_2='"&trim(ADIC_2)&"', ADIC_3='"&trim(ADIC_3)&"', ADIC_4='"&trim(ADIC_4)&"', ADIC_5='"&trim(ADIC_5)&"', ADIC_91='"&trim(ADIC_91)&"', ADIC_92='"&trim(ADIC_92)&"', "
sql_update = sql_update & " ADIC_93='"&trim(ADIC_93)&"', ADIC_94='"&trim(ADIC_94)&"', ADIC_95='"&trim(ADIC_95)&"', USA_CUSTODIO='"&trim(USA_CUSTODIO)&"', COLOR_CUSTODIO='"&trim(COLOR_CUSTODIO)&"', "
sql_update = sql_update & " ADIC_96='"&trim(ADIC_96)&"', ADIC_97='"&trim(ADIC_97)&"', ADIC_98='"&trim(ADIC_98)&"', ADIC_99='"&trim(ADIC_99)&"', ADIC_100='"&trim(ADIC_100)&"', NOMBRE_CONV_PAGARE='"&trim(NOMBRE_CONV_PAGARE)&"', "
sql_update = sql_update & " INTERES_MORA='"&trim(INTERES_MORA)&"', EXPIRACION_CONVENIO='"&trim(EXPIRACION_CONVENIO)&"', EXPIRACION_ANULACION='"&trim(EXPIRACION_ANULACION)&"', "
sql_update = sql_update & " COD_MONEDA='"&trim(COD_MONEDA)&"', COD_TIPODOCUMENTO_HON='"&trim(COD_TIPODOCUMENTO_HON)&"', MESES_TD_HON='"&trim(MESES_TD_HON)&"', ADIC1_DEUDOR='"&trim(ADIC1_DEUDOR)&"', ADIC2_DEUDOR='"&trim(ADIC2_DEUDOR)&"', "
sql_update = sql_update & " ADIC3_DEUDOR='"&trim(ADIC3_DEUDOR)&"', RETIRO_SABADO='"&trim(RETIRO_SABADO)&"', COD_ULT_GES='"&trim(COD_ULT_GES)&"', OBS_ULT_GES='"&trim(OBS_ULT_GES)&"', FORMULA_HONORARIOS='"&trim(FORMULA_HONORARIOS)&"', FORMULA_HONORARIOS_FACT='"&trim(FORMULA_HONORARIOS_FACT)&"', FORMULA_INTERESES='"&trim(FORMULA_INTERESES)&"', "
sql_update = sql_update & " USA_HONORARIOS='"&trim(USA_HONORARIOS)&"', USA_INTERESES='"&trim(USA_INTERESES)&"', USA_SUBCLIENTE='"&trim(USA_SUBCLIENTE)&"', USA_REPLEGAL='"&trim(USA_REPLEGAL)&"', USA_PROTESTOS='"&trim(USA_PROTESTOS)&"', NRO_CLIENTE_DOC='"&trim(NRO_CLIENTE_DOC)&"', NRO_CLIENTE_DEUDOR='"&trim(NRO_CLIENTE_DEUDOR)&"', "
sql_update = sql_update & " DIRECCION='"&trim(DIRECCION)&"', USA_COB_INTERNA='"&trim(USA_COB_INTERNA)&"' "
sql_update = sql_update & " WHERE  COD_CLIENTE  =" & trim(COD_CLIENTE)

	Conn.execute(sql_update)
    if err then
    	Response.write err.description
    end if

	'Response.write sql_update
end if

%>
