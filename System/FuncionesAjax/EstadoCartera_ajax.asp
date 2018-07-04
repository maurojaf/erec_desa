<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->

<%

Response.CodePage=65001
Response.charset ="utf-8"

AbrirSCG()

accion_ajax		=request.querystring("accion_ajax")



if trim(accion_ajax)="filtra_estado_cartera" Then

	COD_CLIENTE 			=request.querystring("COD_CLIENTE")
	ID_USUARIO 				=request.querystring("ID_USUARIO")
	fecha_asignacion_desde 	=request.querystring("fecha_asignacion_desde")
	fecha_asignacion_hasta 	=request.querystring("fecha_asignacion_hasta")
	COD_ESTADO_COBRANZA 	=request.querystring("COD_ESTADO_COBRANZA")
	fecha_gestion_desde 	=request.querystring("fecha_gestion_desde")
	fecha_gestion_hasta 	=request.querystring("fecha_gestion_hasta")
	TIPO_COBRANZA 			=request.querystring("TIPO_COBRANZA")
	CB_CAMPANA 				=request.querystring("CB_CAMPANA")
	CB_RUBRO 				=request.querystring("CB_RUBRO")
	CB_TIPODOC 				=request.querystring("CB_TIPODOC")

	'Response.write COD_CLIENTE &"<br>"&ID_USUARIO&"<br>"&fecha_asignacion_desde&"<br>"&fecha_asignacion_hasta&"<br>"&COD_ESTADO_COBRANZA&"<br>"&fecha_gestion_desde&"<br>"&fecha_gestion_hasta&"<br>"&TIPO_COBRANZA&"<br>"&CB_CAMPANA&"<br>"&CB_RUBRO&"<br>"&CB_TIPODOC


	sql_det =" SELECT ISNULL(COUNT(PP2.RUT_DEUDOR),0) AS TOTAL_RUT, ISNULL(SUM(SALDO_RUT),0) AS SALDO_RUT, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR "
	sql_det = sql_det & " PP2.DIR_SA=1	) THEN 1 ELSE 0 END)),0) AS RUT_GESTIONABLES, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 "
	sql_det = sql_det & " OR PP2.DIR_SA=1) THEN SALDO_RUT ELSE 0 END)),0) AS MONTO_GESTIONABLE, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=1 OR PP2.TEL_SA=1) AND (PP2.EMAIL_VA=0 AND PP2.EMAIL_SA=0) "
	sql_det = sql_det & " AND (PP2.DIR_VA=0 AND PP2.DIR_SA=0)) THEN 1 ELSE 0 END)),0) AS RUT_GES_TEL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=0 AND PP2.TEL_SA=0) AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1) "
	sql_det = sql_det & " AND (PP2.DIR_VA=0 AND PP2.DIR_SA=0)) THEN 1 ELSE 0 END)),0) AS RUT_GES_MAIL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=0 AND PP2.TEL_SA=0) AND (PP2.EMAIL_VA=0 AND PP2.EMAIL_SA=0) "
	sql_det = sql_det & " AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS RUT_GES_DIR, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=1 OR PP2.TEL_SA=1) AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1) "
	sql_det = sql_det & " AND (PP2.DIR_VA=0 AND PP2.DIR_SA=0)) THEN 1 ELSE 0 END)),0) AS RUT_GES_TEL_MAIL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=1 OR PP2.TEL_SA=1) AND (PP2.EMAIL_VA=0 AND PP2.EMAIL_SA=0) "
	sql_det = sql_det & " AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS RUT_GES_TEL_DIR, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=0 AND PP2.TEL_SA=0) AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1) "
	sql_det = sql_det & " AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS RUT_GES_MAIL_DIR, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=1 OR PP2.TEL_SA=1) AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1) "
	sql_det = sql_det & " AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS RUT_GES_TEL_MAIL_DIR, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS CASOS_GESTIONADOS, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN SALDO_RUT ELSE 0 END)),0) AS MONTO_GESTIONADOS, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL=0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_PENDIENTES_CON_FONO, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL=0 AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_PENDIENTES_CON_MAIL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL=0 AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_PENDIENTES_CON_DIRECCION, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_POSITIVA, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN SALDO_RUT ELSE 0 END)),0) AS MONTO_GESTION_POSITIVA, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE=0 AND PP2.TT_GEST_GENERAL>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1)) "
	sql_det = sql_det & " THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_NEGATIVA_CON_FONO, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE=0 AND PP2.TT_GEST_GENERAL>0 AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1)) "
	sql_det = sql_det & " THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_NEGATIVA_CON_MAIL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE=0 AND PP2.TT_GEST_GENERAL>0 AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) "
	sql_det = sql_det & " THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_NEGATIVA_CON_DIRECCION, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GTIT>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_POSITIVA_TITULAR, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GTIT>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN SALDO_RUT ELSE 0 END)),0) AS MONTO_GESTION_POSITIVA_TITULAR, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND PP2.TT_GTIT=0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_GESTION_POSITIVA_TERCERO_CON_FONO, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND PP2.TT_GTIT=0 AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_GESTION_POSITIVA_TERCERO_CON_MAIL, "
	sql_det = sql_det & "  ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND PP2.TT_GTIT=0 AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_GESTION_POSITIVA_TERCERO_CON_DIRECCION "

	sql_det = sql_det & " FROM "

	sql_det = sql_det & " (SELECT PP.RUT_DEUDOR, "
	sql_det = sql_det & " (SELECT SUM(SALDO) FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO WHERE C.RUT_DEUDOR=PP. RUT_DEUDOR AND C.COD_CLIENTE=PP.COD_CLIENTE AND ED.ACTIVO=1 "

		if trim(TIPO_COBRANZA)="INTERNA" THEN
			sql_det = sql_det & " AND C.CUSTODIO IS not NULL "

		elseif trim(TIPO_COBRANZA)="EXTERNA" THEN
			sql_det = sql_det & " AND C.CUSTODIO IS NULL "
		end if
	sql_det = sql_det & " /* DEJAR PARAMETRICO) */) AS SALDO_RUT, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_TEL_VA >0 THEN 1 ELSE 0 END) AS TEL_VA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_TEL_SA >0 THEN 1 ELSE 0 END) AS TEL_SA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_TEL_NV >0 THEN 1 ELSE 0 END) AS TEL_NV, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_EMAIL_VA >0 THEN 1 ELSE 0 END) AS EMAIL_VA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_EMAIL_SA >0 THEN 1 ELSE 0 END) AS EMAIL_SA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_EMAIL_NV >0 THEN 1 ELSE 0 END) AS EMAIL_NV, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_DIR_VA >0 THEN 1 ELSE 0 END) AS DIR_VA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_DIR_SA >0 THEN 1 ELSE 0 END) AS DIR_SA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_DIR_NV >0 THEN 1 ELSE 0 END) AS DIR_NV, "
	sql_det = sql_det & " SUM(PP.GEST_GENERAL) AS TT_GEST_GENERAL, "
	sql_det = sql_det & " SUM(PP.GEST_TEL) AS TT_GEST_TEL, "
	sql_det = sql_det & " SUM(PP.GEST_MAIL) AS TT_GEST_MAIL, "
	sql_det = sql_det & " SUM(PP.GEST_DIR) AS TT_GDIR, "
	sql_det = sql_det & " SUM(PP.GEST_EFE) AS TT_GEFE, "
	sql_det = sql_det & " SUM(PP.GEST_TIT) AS TT_GTIT "

	sql_det = sql_det & " FROM  "
	sql_det = sql_det & " (SELECT D.RUT_DEUDOR,D.COD_CLIENTE, "
	sql_det = sql_det & " (CASE WHEN G.ID_GESTION IS NOT NULL THEN 1 ELSE 0 END) AS GEST_GENERAL, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GTEL),0) AS GEST_TEL, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GMAIL),0) AS GEST_MAIL, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GDIR),0) AS GEST_DIR, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GEFE),0) AS GEST_EFE, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GTIT),0) AS GEST_TIT, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_TEL_VA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_TEL_SA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_TEL_NV, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_EMAIL_VA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_EMAIL_SA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_EMAIL_NV, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_DIR_VA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_DIR_SA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_DIR_NV "
	sql_det = sql_det & " FROM CUOTA C INNER JOIN DEUDOR D ON C.RUT_DEUDOR = D.RUT_DEUDOR AND C.COD_CLIENTE = D.COD_CLIENTE "
	sql_det = sql_det & " 			 INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO "
	sql_det = sql_det & " 			 LEFT JOIN GESTIONES_CUOTA GC ON C.ID_CUOTA = GC.ID_CUOTA "
	sql_det = sql_det & " 			 LEFT JOIN GESTIONES G ON GC.ID_GESTION = G.ID_GESTION "
	sql_det = sql_det & " 			 LEFT JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA AND "
	sql_det = sql_det & " 					  G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA AND "
	sql_det = sql_det & " 					  G.COD_GESTION = GTG.COD_GESTION AND "
	sql_det = sql_det & " 					  G.COD_CLIENTE = GTG.COD_CLIENTE "
						  
	sql_det = sql_det & " WHERE ED.ACTIVO=1  "
	sql_det = sql_det & " AND D.COD_CLIENTE IN (" & trim(COD_CLIENTE) &")"


	IF TRIM(fecha_gestion_desde)<>"" AND TRIM(fecha_gestion_hasta)="" THEN
		sql_det = sql_det & " AND G.FECHA_INGRESO > CONVERT(DATETIME, '"&TRIM(fecha_gestion_desde)&"') "

	ELSEIF TRIM(fecha_gestion_desde)="" AND TRIM(fecha_gestion_hasta)<>"" THEN
		sql_det = sql_det & " AND G.FECHA_INGRESO < CONVERT(DATETIME, '"&TRIM(fecha_gestion_hasta)&"') "

	ELSEIF TRIM(fecha_gestion_desde)<>"" AND TRIM(fecha_gestion_hasta)<>"" THEN
		sql_det = sql_det & " AND G.FECHA_INGRESO BETWEEN CONVERT(DATETIME,'"&TRIM(fecha_gestion_desde)&"') AND CONVERT(DATETIME,'"&TRIM(fecha_gestion_hasta)&"') "	

	END IF


	if trim(CB_TIPODOC)<>"" then
		sql_det = sql_det & " AND C.TIPO_DOCUMENTO ='"&trim(CB_TIPODOC)&"' " 
	end if

	if trim(CB_RUBRO)<>"" then
		sql_det = sql_det & " AND ISNULL(D.ADIC_2,'OTRO') = '"&trim(CB_RUBRO)&"' " 
	end if


	IF TRIM(fecha_asignacion_desde)<>"" AND TRIM(fecha_asignacion_hasta)="" THEN
		sql_det = sql_det & " AND ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION) > CONVERT(DATETIME, '"&TRIM(fecha_asignacion_desde)&"') "

	ELSEIF TRIM(fecha_asignacion_desde)="" AND TRIM(fecha_asignacion_hasta)<>"" THEN
		sql_det = sql_det & " AND ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION) < CONVERT(DATETIME, '"&TRIM(fecha_asignacion_hasta)&"') "

	ELSEIF TRIM(fecha_asignacion_desde)<>"" AND TRIM(fecha_asignacion_hasta)<>"" THEN
		sql_det = sql_det & " AND ISNULL(C.FECHA_ESTADO_CUSTODIO,C.FECHA_CREACION) BETWEEN CONVERT(DATETIME,'"&TRIM(fecha_asignacion_desde)&"') AND CONVERT(DATETIME,'"&TRIM(fecha_asignacion_hasta)&"') "	

	END IF


	if trim(TIPO_COBRANZA)="INTERNA" THEN
		sql_det = sql_det & " AND D.CUSTODIO IS not NULL "

	elseif trim(TIPO_COBRANZA)="EXTERNA" THEN
		sql_det = sql_det & " AND D.CUSTODIO IS NULL "
	end if

	IF TRIM(CB_CAMPANA)<>"" THEN
		sql_det = sql_det & " AND G.ID_CAMPANA = '"&TRIM(CB_CAMPANA)&"'"

	END IF

	IF TRIM(ID_USUARIO)<>"" THEN
		sql_det = sql_det & " AND ISNULL(C.USUARIO_ASIG,0) = '"&TRIM(ID_USUARIO)&"'"

	END IF

	IF TRIM(COD_ESTADO_COBRANZA)<>"" THEN
		sql_det = sql_det & " AND D.ETAPA_COBRANZA = '"&TRIM(COD_ESTADO_COBRANZA)&"'"

	END IF
	
	sql_det = sql_det & " GROUP BY D.RUT_DEUDOR,D.COD_CLIENTE,G.ID_GESTION,GTG.PRIORIDAD_GTEL,GTG.PRIORIDAD_GMAIL,GTG.PRIORIDAD_GDIR,GTG.PRIORIDAD_GEFE,GTG.PRIORIDAD_GTIT "
	sql_det = sql_det & " ) AS PP "	 	 
	sql_det = sql_det & " GROUP BY PP.COD_CLIENTE,PP.RUT_DEUDOR,PP. TOTAL_TEL_VA,TOTAL_TEL_SA,TOTAL_TEL_NV,TOTAL_EMAIL_VA, "
	sql_det = sql_det & " TOTAL_EMAIL_SA,TOTAL_EMAIL_NV,TOTAL_DIR_VA,TOTAL_DIR_SA,TOTAL_DIR_NV "
	sql_det = sql_det & " ) AS PP2	"
				 
	set rs_det = conn.execute(sql_det)			 

	if not rs_det.eof then 

		TOTAL_RUT      									=rs_det("TOTAL_RUT")
		SALDO_RUT										=rs_det("SALDO_RUT")

		RUT_GESTIONABLES								=rs_det("RUT_GESTIONABLES")
		MONTO_GESTIONABLE								=rs_det("MONTO_GESTIONABLE")

		RUT_GES_TEL 									=rs_det("RUT_GES_TEL")
		RUT_GES_MAIL 									=rs_det("RUT_GES_MAIL")
		RUT_GES_DIR 									=rs_det("RUT_GES_DIR")
		RUT_GES_TEL_MAIL 								=rs_det("RUT_GES_TEL_MAIL")
		RUT_GES_TEL_DIR 								=rs_det("RUT_GES_TEL_DIR")
		RUT_GES_MAIL_DIR 								=rs_det("RUT_GES_MAIL_DIR")
		RUT_GES_TEL_MAIL_DIR 							=rs_det("RUT_GES_TEL_MAIL_DIR")
		
		CASOS_GESTIONADOS 								=rs_det("CASOS_GESTIONADOS")
		MONTO_GESTIONADOS 								=rs_det("MONTO_GESTIONADOS")


		CASOS_PENDIENTES_CON_FONO 						=rs_det("CASOS_PENDIENTES_CON_FONO")
		CASOS_PENDIENTES_CON_MAIL 						=rs_det("CASOS_PENDIENTES_CON_MAIL")
		CASOS_PENDIENTES_CON_DIRECCION 					=rs_det("CASOS_PENDIENTES_CON_DIRECCION")
		
		CASOS_GESTION_POSITIVA 							=rs_det("CASOS_GESTION_POSITIVA")
		MONTO_GESTION_POSITIVA 							=rs_det("MONTO_GESTION_POSITIVA")

		CASOS_GESTION_NEGATIVA_CON_FONO 				=rs_det("CASOS_GESTION_NEGATIVA_CON_FONO")
		CASOS_GESTION_NEGATIVA_CON_MAIL 				=rs_det("CASOS_GESTION_NEGATIVA_CON_MAIL")
		CASOS_GESTION_NEGATIVA_CON_DIRECCION 			=rs_det("CASOS_GESTION_NEGATIVA_CON_DIRECCION")
		
		CASOS_GESTION_POSITIVA_TITULAR 					=rs_det("CASOS_GESTION_POSITIVA_TITULAR")
		MONTO_GESTION_POSITIVA_TITULAR 					=rs_det("MONTO_GESTION_POSITIVA_TITULAR")

		CASOS_GESTION_POSITIVA_TERCERO_CON_FONO 		=rs_det("CASOS_GESTION_POSITIVA_TERCERO_CON_FONO")
		CASOS_GESTION_POSITIVA_TERCERO_CON_MAIL 		=rs_det("CASOS_GESTION_POSITIVA_TERCERO_CON_MAIL")
		CASOS_GESTION_POSITIVA_TERCERO_CON_DIRECCION 	=rs_det("CASOS_GESTION_POSITIVA_TERCERO_CON_DIRECCION")

		CASOS_NO_GESTIONABLES							=CINT(TOTAL_RUT)-CINT(RUT_GESTIONABLES)
		MONTO_NO_GESTIONABLES 							=(SALDO_RUT)-(MONTO_GESTIONABLE)

		CASOS_NO_GESTIONADOS 							=CINT(RUT_GESTIONABLES)-CINT(CASOS_GESTIONADOS)
		MONTO_NO_GESTIONADOS 							=MONTO_GESTIONABLE-MONTO_GESTIONADOS

		CASOS_GESTION_NEGATIVA 							=CINT(CASOS_GESTIONADOS)-CINT(CASOS_GESTION_POSITIVA)
		MONTO_GESTION_NEGATIVA 							=MONTO_GESTIONADOS-MONTO_GESTION_POSITIVA

		CASOS_GESTION_POSITIVA_TERCERO 					=CINT(CASOS_GESTION_POSITIVA)-CINT(CASOS_GESTION_POSITIVA_TITULAR)
		MONTO_GESTION_POSITIVA_TERCERO 					=MONTO_GESTION_POSITIVA-MONTO_GESTION_POSITIVA_TITULAR

	Else

		TOTAL_RUT      									=""
		SALDO_RUT										=""
		RUT_GESTIONABLES								=""
		MONTO_GESTIONABLE								=""
		RUT_GES_TEL 									=""
		RUT_GES_MAIL 									=""
		RUT_GES_DIR 									=""
		RUT_GES_TEL_MAIL 								=""
		RUT_GES_TEL_DIR 								=""
		RUT_GES_MAIL_DIR 								=""
		RUT_GES_TEL_MAIL_DIR 							=""
		CASOS_GESTIONADOS 								=""
		MONTO_GESTIONADOS 								=""
		CASOS_PENDIENTES_CON_FONO 						=""
		CASOS_PENDIENTES_CON_MAIL 						=""
		CASOS_PENDIENTES_CON_DIRECCION 					=""
		CASOS_GESTION_POSITIVA 							=""
		MONTO_GESTION_POSITIVA 							=""
		CASOS_GESTION_NEGATIVA_CON_FONO 				=""
		CASOS_GESTION_NEGATIVA_CON_MAIL 				=""
		CASOS_GESTION_NEGATIVA_CON_DIRECCION 			=""
		CASOS_GESTION_POSITIVA_TITULAR 					=""
		MONTO_GESTION_POSITIVA_TITULAR 					=""
		CASOS_GESTION_POSITIVA_TERCERO_CON_FONO 		=""
		CASOS_GESTION_POSITIVA_TERCERO_CON_MAIL 		=""
		CASOS_GESTION_POSITIVA_TERCERO_CON_DIRECCION 	=""
		CASOS_NO_GESTIONADOS 							=""
		MONTO_NO_GESTIONADOS 							=""

	end if


%>

	

		<div class="cargados">
			<div class="titulo_carga">CASOS CARGADOS</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=0&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=cint(TOTAL_RUT)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(SALDO_RUT,0)%></div>
		</div>

		<div class="gestionables">
			<div class="titulo_carga">
				GESTIONABLES 
				<img class="iconos_detalle_gestiones" id="id_gestionables" src="../Imagenes/48px-Crystal_Clear_app_kexi.png" alt="Detalle carga">
			</div> 			
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=1&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(RUT_GESTIONABLES)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTIONABLE,0)%></div>
		</div>

		<div class="contenido_detalle_ges">
			<div class="titulo_carga_detalle">INFORMACIÓN ADICIONAL GESTIONABLES</div><br>
			<span class="titulo_carga_det">Con Télefono: </span><%=trim(RUT_GES_TEL)%><br>
			<span class="titulo_carga_det">Con Email: </span><%=trim(RUT_GES_MAIL)%><br>
			<span class="titulo_carga_det">Con Dirección: </span><%=trim(RUT_GES_DIR)%><br><br>
			<span class="titulo_carga_det">Con Télefono-Email: </span><%=trim(RUT_GES_TEL_MAIL)%><br><br>
			<span class="titulo_carga_det">Con Télefono-Dirección: </span><%=trim(RUT_GES_TEL_DIR)%><br>
			<span class="titulo_carga_det">Con Email-Dirección: </span><%=trim(RUT_GES_MAIL_DIR)%><br>
			<span class="titulo_carga_det">Con Télefono-Email-Dirección: </span><%=trim(RUT_GES_TEL_MAIL_DIR)%><br>	
			<input type="text" name="sacar_foco" readonly id="sacar_foco">		
		</div>

		<div class="no_gestionables">
			<div class="titulo_carga_negativa">INUBICABLES</div>
			<div class="cuerpo_carga_negativa">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=2&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_NO_GESTIONABLES)%></a>
			</div>
			<div class="cuerpo_monto_negativa">Monto: $<%=FormatNumber(MONTO_NO_GESTIONABLES,0)%></div>
		</div>	
		<div class="titulo_carga_linea">&nbsp;</div>
		<div class="gestionados">
			
			<div class="titulo_carga">GESTIONADOS</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=3&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_GESTIONADOS)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTIONADOS,0)%></div>	
		</div>	

		<div class="pendientes">
			<div class="titulo_carga">
				PENDIENTES
				<img class="iconos_detalle_gestiones" id="id_pendientes" src="../Imagenes/48px-Crystal_Clear_app_kexi.png" alt="Detalle carga">
			</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=4&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_NO_GESTIONADOS)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_NO_GESTIONADOS,0)%></div>
		</div>
		<div class="titulo_carga_linea_pos">&nbsp;</div>
		<div class="gestion_positiva">
			<div class="titulo_carga">GESTIÓN POSITIVA</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=5&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_GESTION_POSITIVA)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTION_POSITIVA,0)%></div>	
		</div>

		<div class="gestion_negativa">
			<div class="titulo_carga">
				GESTIÓN NEGATIVA 
				<img class="iconos_detalle_gestiones" id="id_negativa" src="../Imagenes/48px-Crystal_Clear_app_kexi.png" alt="Detalle carga">
			</div>			
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=6&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_GESTION_NEGATIVA)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTION_NEGATIVA,0)%></div>
		</div>

		<div class="contenido_detalle_negativa">
			<div class="titulo_carga_detalle">INFORMACIÓN ADICIONAL GESTIONABLES</div><br>
			<span class="titulo_carga_det">Con Télefono: </span><%=trim(CASOS_GESTION_NEGATIVA_CON_FONO)%><br>
			<span class="titulo_carga_det">Con Email: </span><%=trim(CASOS_GESTION_NEGATIVA_CON_MAIL)%><br>
			<span class="titulo_carga_det">Con Dirección: </span><%=trim(CASOS_GESTION_NEGATIVA_CON_DIRECCION)%><br>			
			<input type="text" name="sacar_foco_neg" readonly id="sacar_foco_neg">		
		</div>
		<div class="titulo_carga_linea_titular">&nbsp;</div>
		<div class="titular">
			<div class="titulo_carga">TITULAR</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=7&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_GESTION_POSITIVA_TITULAR)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTION_POSITIVA_TITULAR,0)%></div>
		</div>

		<div class="tercero">
			<div class="titulo_carga">
				TERCERO
				<img class="iconos_detalle_gestiones" id="id_tercero" src="../Imagenes/48px-Crystal_Clear_app_kexi.png" alt="Detalle carga">
			</div>
			
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=8"><%=CINT(CASOS_GESTION_POSITIVA_TERCERO)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTION_POSITIVA_TERCERO,0)%></div>	
		</div>
		
	

<%
	'Response.write sql_det

elseif trim(accion_ajax)="refresa_cobranza" then
	COD_CLIENTE 	=request.querystring("COD_CLIENTE")
	intVerCobExt 	="1"

	'###### CLIETNE INTERNO EXTERNO
	strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
	strSql = strSql & " FROM CLIENTE CL"
	strSql = strSql & " WHERE CL.COD_CLIENTE = '" & COD_CLIENTE & "'"

	set RsCli=conn.execute(strSql)
	If not RsCli.eof then
		intUsaCobInterna = RsCli("USA_COB_INTERNA")
	End if
	'Response.write strSql
%>

	<select style="width:243px;"  name="TIPO_COBRANZA" id="TIPO_COBRANZA" >
		<%If Trim(intVerCobExt) = "1" and Trim(intUsaCobInterna) = "1" Then%>
			<option value="0" <%If Trim(strCobranza) ="" Then Response.write "SELECTED"%>>TODOS</option>
		<%End If%>
		
		<%If Trim(intUsaCobInterna) = "1" Then%>
			<option value="INTERNA" <%If Trim(strCobranza) ="INTERNA" Then Response.write "SELECTED"%>>INTERNA</option>
		<%End If%>
		
		<%If Trim(intVerCobExt) = "1" Then%>
			<option value="EXTERNA" <%If Trim(strCobranza) ="EXTERNA" Then Response.write "SELECTED"%>>EXTERNA</option>
		<%End If%>
	</select>

<%

elseif trim(accion_ajax)="refresa_ejecutivo" then
	COD_CLIENTE =request.querystring("COD_CLIENTE")
	'###### EJECUTIVO
	sql_usuario= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
	sql_usuario= sql_usuario & " FROM USUARIO U "
	sql_usuario= sql_usuario & " INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & trim(COD_CLIENTE)
	sql_usuario= sql_usuario & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
	set rsUsuario=Conn.execute(sql_usuario)
%>
	<select style="width:240px;"  name="ID_USUARIO" id="ID_USUARIO" >
		<option value="">TODOS</option>
		<%if not rsUsuario.eof then%>
			<%do while not rsUsuario.eof%>
			<option value="<%=trim(rsUsuario("ID_USUARIO"))%>"><%=trim(rsUsuario("LOGIN"))%></option>
			<%rsUsuario.movenext
			loop%>
		<%end if%>
	</select>
<%

elseif trim(accion_ajax)="refresa_rubro" then
	COD_CLIENTE =request.querystring("COD_CLIENTE")

%>
	<select style="width:243px;" name="CB_RUBRO" ID="CB_RUBRO">
		<option value="" <%if Trim(strRubro)="" then response.Write("Selected") end if%>>SELECCIONE</option>
		<%

		ssql="SELECT DISTINCT ISNULL(ADIC_2,'OTRO') AS ADIC_2 FROM DEUDOR  WHERE COD_CLIENTE = '" & COD_CLIENTE & "' ORDER BY ADIC_2"
		set rsTemp= Conn.execute(ssql)
		if not rsTemp.eof then
			do until rsTemp.eof%>
			<option value="<%=rsTemp("ADIC_2")%>"<%if strRubro=rsTemp("ADIC_2") then response.Write("Selected") End If%>><%=rsTemp("ADIC_2")%></option>
			<%
			rsTemp.movenext
			loop
		end if

		%>
	</select>

<%

elseif trim(accion_ajax)="refresa_tipo_doc" then
	COD_CLIENTE =request.querystring("COD_CLIENTE")

%>
	<select style="width:240px;" name="CB_TIPODOC" id="CB_TIPODOC">
		<option value="">TODOS</option>
		<%
		strSql="SELECT DISTINCT COD_TIPO_DOCUMENTO, NOM_TIPO_DOCUMENTO"
		strSql=strSql & " FROM CUOTA LEFT JOIN TIPO_DOCUMENTO ON TIPO_DOCUMENTO = COD_TIPO_DOCUMENTO"
		strSql=strSql & " WHERE CUOTA.COD_CLIENTE = '" & COD_CLIENTE & "' AND COD_TIPO_DOCUMENTO is not null "
		strSql=strSql & " ORDER BY NOM_TIPO_DOCUMENTO ASC"

		set rsTemp= Conn.execute(strSql)
		if not rsTemp.eof then
			do until rsTemp.eof%>
			<option value="<%=rsTemp("COD_TIPO_DOCUMENTO")%>"<%if Trim(intTipoDoc)=Trim(rsTemp("COD_TIPO_DOCUMENTO")) then response.Write("Selected") End If%>><%=rsTemp("NOM_TIPO_DOCUMENTO")%></option>
			<%
			rsTemp.movenext
			loop
		end if
		%>
	</select>

<%	
end if

%>
