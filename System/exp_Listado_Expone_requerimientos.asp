<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.buffer = true
	response.contentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Informe.xls"	
%>
<html lang="es">
<HEAD>
    <meta charset="utf-8">
	<!--#include file="arch_utils.asp"-->
<%

	Response.CodePage = 65001
	Response.charset="utf-8"
	
	AbrirSCG()


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())


	strCobranza 		=request.querystring("CB_COBRANZA")
	strTipoGestion 		=request.querystring("CMB_TIPO_GESTION")
	strEstadoProceso 	=request.querystring("CMB_ESTADO_PROCESO")
	dtmInicio			=request.querystring("inicio")
	dtmTermino 			=request.querystring("termino")
	strEjecutivo 		=request.querystring("CB_EJECUTIVO")
	strCodCliente		=request.querystring("COD_CLIENTE")
	intInicioContador	=request.querystring("inicia_contador")
	strRutDeudor 		=request.querystring("RUT_DEUDOR")
	HORA_CONSULTA  		=request.querystring("HORA_CONSULTA")
	FECHA_CONSULTA  	=request.querystring("FECHA_CONSULTA")
	CH_CP_ADJUNTO 		=request.querystring("CH_CP_ADJUNTO")

	'response.write FECHA_CONSULTA &"<br>"&HORA_CONSULTA
	sql_sel_casos = " "
	sql_sel_casos = sql_sel_casos & " SELECT "
	sql_sel_casos = sql_sel_casos & " VV.ID_GESTION, "   
	sql_sel_casos = sql_sel_casos & " VV.ID_CUOTA, "
	sql_sel_casos = sql_sel_casos & " VV.COD_CLIENTE, "   
	sql_sel_casos = sql_sel_casos & " VV.CUSTODIO, "   
	sql_sel_casos = sql_sel_casos & " CONVERT(VARCHAR, VV.FECHA_INGRESO_GESTION,103) FECHA_INGRESO_GESTION, "  
	sql_sel_casos = sql_sel_casos & " SUBSTRING(CONVERT(VARCHAR, VV.FECHA_INGRESO_GESTION,108),1,5) FECHA_INGRESO_GESTION_HORA, "  
	sql_sel_casos = sql_sel_casos & " VV.RUT_DEUDOR, "   
	sql_sel_casos = sql_sel_casos & " VV.NOMBRE_DEUDOR, "  
	sql_sel_casos = sql_sel_casos & " VV.SALDO_CUOTA , "    
	sql_sel_casos = sql_sel_casos & " VV.FECHA_VENC , "    
	sql_sel_casos = sql_sel_casos & " VV.FECHA_GESTION, "   
	sql_sel_casos = sql_sel_casos & " VV.HORA_INGRESO, "   
	sql_sel_casos = sql_sel_casos & " VV.MONTO_GESTION, "   
	sql_sel_casos = sql_sel_casos & " FORMA_NORMALIZACION = ISNULL(CFP.DESC_FORMA_PAGO,'NO ESPEC.'), "   
	sql_sel_casos = sql_sel_casos & " LUGAR_GESTION = ISNULL(ISNULL(UPPER(FR.NOMBRE+' '+FR.UBICACION), " 
	sql_sel_casos = sql_sel_casos & " upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.COMUNA)),'NO ESPEC.'),  "    
	sql_sel_casos = sql_sel_casos & " VV.NRO_DOC_PAGO, "   
	sql_sel_casos = sql_sel_casos & " VV.OBSERVACIONES_CAMPO, "    
	sql_sel_casos = sql_sel_casos & " VV.TIPO_MODULO, "   
	sql_sel_casos = sql_sel_casos & " VV.ACUMULADO, "   
	sql_sel_casos = sql_sel_casos & " ISNULL(CONVERT(VARCHAR, VV.FECHA_CONSULTA,103),'NO CONSULT') MIN_FECHA_CONSULTA, "  
	sql_sel_casos = sql_sel_casos & " SUBSTRING(CONVERT(VARCHAR, VV.FECHA_CONSULTA,108),1,5) MIN_FECHA_CONSULTA_HORA, "  
	sql_sel_casos = sql_sel_casos & " VV.FORMA_PAGO, "   
	sql_sel_casos = sql_sel_casos & " VV.ID_DIRECCION_COBRO_DEUDOR, "   
	sql_sel_casos = sql_sel_casos & " VV.ID_FORMA_RECAUDACION, "
	sql_sel_casos = sql_sel_casos & " VV.ID_USUARIO_ASIG, "   
	sql_sel_casos = sql_sel_casos & " U.LOGIN, "    
	sql_sel_casos = sql_sel_casos & " VV.PROCESO , "  
	sql_sel_casos = sql_sel_casos & " VV.ID_PROCESO, "   
	'sql_sel_casos = sql_sel_casos & " [dbo].[concatena_cuotas_estados] (VV.ID_GESTION,'ACTIVAS') AS CUOTAS_ACTIVAS, " 
	'sql_sel_casos = sql_sel_casos & " [dbo].[concatena_cuotas_estados] (VV.ID_GESTION,'CANCELADAS') AS CUOTAS_CANCELADAS, "  
	'sql_sel_casos = sql_sel_casos & " [dbo].[concatena_cuotas_estados] (VV.ID_GESTION,'RETIRADAS') AS CUOTAS_RETIRADAS, "  
	'sql_sel_casos = sql_sel_casos & " [dbo].[concatena_cuotas_estados] (VV.ID_GESTION,'NO ASIGNABLE') AS CUOTAS_NO_ASIGNABLES, "  
	sql_sel_casos = sql_sel_casos & " CASE WHEN VV.ID_ARCHIVO IS NOT NULL THEN 1 ELSE 0 END CANTIDAD_DOCUMENTOS, " 
	sql_sel_casos = sql_sel_casos & " GESTIONSOLA = CASE WHEN (VV.TIPO_MODULO = 2 AND VV.FECHA_CONSULTA_NORM IS NULL) "
	sql_sel_casos = sql_sel_casos & " THEN 'INDICA QUE PAGO' "
	sql_sel_casos = sql_sel_casos & " WHEN (VV.TIPO_MODULO = 1 OR VV.TIPO_MODULO = 11) AND ( CFP.ID_FORMA_PAGO = 'TR' OR CFP.ID_FORMA_PAGO = 'DP') "
	sql_sel_casos = sql_sel_casos & " THEN 'COMPROMISO D & T' "
	sql_sel_casos = sql_sel_casos & " WHEN (VV.TIPO_MODULO = 2 AND (VV.FECHA_INGRESO_GESTION + VV.HORA_INGRESO) < CAST(ISNULL(VV.FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(VV.FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) > GETDATE ()) "
	sql_sel_casos = sql_sel_casos & " THEN 'INDICA PAGO EN CONSULTA' "
	sql_sel_casos = sql_sel_casos & " WHEN (VV.TIPO_MODULO = 2 AND (VV.FECHA_INGRESO_GESTION + VV.HORA_INGRESO) < CAST(ISNULL(VV.FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) AND (CAST(ISNULL(VV.FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME) + 7) < GETDATE ()) "
	sql_sel_casos = sql_sel_casos & " THEN 'INDICA PAGO NO RESP.' "
	sql_sel_casos = sql_sel_casos & " WHEN (VV.TIPO_MODULO = 2 AND (VV.FECHA_INGRESO_GESTION + VV.HORA_INGRESO) > CAST(ISNULL(VV.FECHA_CONSULTA_NORM,'01-01-1900') AS DATETIME)) "
	sql_sel_casos = sql_sel_casos & " THEN 'REITERA INDICA PAGO' "
	sql_sel_casos = sql_sel_casos & " ELSE 'PAGO NO APLICADO' "
	sql_sel_casos = sql_sel_casos & " END , "
	sql_sel_casos = sql_sel_casos & " VV.RUT_SUBCLIENTE, "	 
	sql_sel_casos = sql_sel_casos & " VV.NOMBRE_SUBCLIENTE, "
	sql_sel_casos = sql_sel_casos & " VV.NRO_DOC, "
	sql_sel_casos = sql_sel_casos & " VV.NRO_CUOTA, "
	sql_sel_casos = sql_sel_casos & " TD.NOM_TIPO_DOCUMENTO, "
	sql_sel_casos = sql_sel_casos & " VV.VALOR_CUOTA, "
	sql_sel_casos = sql_sel_casos & " VV.SALDO, "
	sql_sel_casos = sql_sel_casos & " VV.INTERLOCUTOR, "
	sql_sel_casos = sql_sel_casos & " VV.SUCURSAL, "
	sql_sel_casos = sql_sel_casos & " CASE WHEN (VV.TIPO_MODULO IN (1,11) ) "
	sql_sel_casos = sql_sel_casos & " THEN ISNULL(CONVERT(VARCHAR(10),VV.FECHA_COMPROMISO,103),'') "
	sql_sel_casos = sql_sel_casos & " WHEN ( VV.TIPO_MODULO IN (2) AND VV.FECHA_PAGO IS NOT NULL) "
	sql_sel_casos = sql_sel_casos & " THEN CONVERT(VARCHAR(10),VV.FECHA_PAGO,103) "
	sql_sel_casos = sql_sel_casos & " WHEN ( VV.TIPO_MODULO IN (6) AND VV.FECHA_PAGO IS NOT NULL) "
	sql_sel_casos = sql_sel_casos & " THEN CONVERT(VARCHAR(10),VV.FECHA_INGRESO_GESTION ,103) "
	sql_sel_casos = sql_sel_casos & " ELSE 'NO ESPEC' "
	sql_sel_casos = sql_sel_casos & " END AS FECHA_NORMALIZACION "	

	sql_sel_casos = sql_sel_casos & " FROM VIEW_CASOS_GESTION_APOYO VV "	   
	sql_sel_casos = sql_sel_casos & " LEFT JOIN CAJA_FORMA_PAGO CFP ON VV.FORMA_PAGO = CFP.ID_FORMA_PAGO "	   
	sql_sel_casos = sql_sel_casos & " LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION=VV.ID_DIRECCION_COBRO_DEUDOR "	   
	sql_sel_casos = sql_sel_casos & " LEFT JOIN FORMA_RECAUDACION FR ON FR.ID_FORMA_RECAUDACION=VV.ID_FORMA_RECAUDACION "	   
	sql_sel_casos = sql_sel_casos & " INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON VV.COD_CATEGORIA = GTC.COD_CATEGORIA "	   
	sql_sel_casos = sql_sel_casos & " INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTS ON VV.COD_CATEGORIA = GTS.COD_CATEGORIA AND VV.COD_SUB_CATEGORIA = GTS.COD_SUB_CATEGORIA "	   
	sql_sel_casos = sql_sel_casos & " LEFT JOIN USUARIO U ON VV.ID_USUARIO_ASIG = U.ID_USUARIO "	    
	sql_sel_casos = sql_sel_casos & " LEFT JOIN TIPO_DOCUMENTO TD ON VV.TIPO_DOCUMENTO = TD.COD_TIPO_DOCUMENTO "	
 

  	sql_sel_casos = sql_sel_casos & " WHERE VV.COD_CLIENTE ='"&TRIM(strCodCliente)&"' AND VV.TIPO_MODULO = 3 "
  	
  	if trim(strRutDeudor)<>"" then
  		sql_sel_casos = sql_sel_casos & " AND VV.RUT_DEUDOR ='"&trim(strRutDeudor)&"' "
  	end if


  	if trim(strCobranza)="INTERNA" then
  		sql_sel_casos = sql_sel_casos & " AND VV.CUSTODIO IS NOT NULL "
  	ElseIf trim(strCobranza)="EXTERNA" then
  		sql_sel_casos = sql_sel_casos & " AND VV.CUSTODIO IS  NULL "
  	end if

  	if trim(dtmInicio)<>"" and trim(dtmTermino)<>"" then
  		sql_sel_casos = sql_sel_casos & " AND  convert(datetime, VV.FECHA_INGRESO_GESTION) BETWEEN  convert(datetime, '"&trim(dtmInicio)&"') AND  convert(datetime, '"&trim(dtmTermino)&"')"
  	end  if

  	if trim(dtmInicio)="" and trim(dtmTermino)<>"" then
  		sql_sel_casos = sql_sel_casos & " AND  convert(datetime, VV.FECHA_INGRESO_GESTION) <= convert(datetime, '"&trim(dtmTermino)&"')"
  	end  if

  	if trim(dtmInicio)<>"" and trim(dtmTermino)="" then
  		sql_sel_casos = sql_sel_casos & " AND  convert(datetime, VV.FECHA_INGRESO_GESTION) >=  convert(datetime, '"&trim(dtmInicio)&"')"
  	end  if  	  	

  	if trim(strEjecutivo)<>"" then
  		sql_sel_casos = sql_sel_casos & " AND  U.ID_USUARIO='"&strEjecutivo&"'"
  	end if

  	if  trim(strEstadoProceso)<>""  then
		sql_sel_casos = sql_sel_casos & " AND VV.ID_PROCESO in ("&trim(strEstadoProceso)&") "

  	end if 

  	if trim(FECHA_CONSULTA)<>"" then
  		sql_sel_casos = sql_sel_casos & " AND convert(varchar,VV.FECHA_CONSULTA, 103) = '"&trim(FECHA_CONSULTA)&"' "

  	end if

  	if trim(HORA_CONSULTA)<>"" then
  		sql_sel_casos = sql_sel_casos & " AND SUBSTRING(convert(varchar,VV.FECHA_CONSULTA, 108), 1, 5) = '"&trim(HORA_CONSULTA)&"' "

  	end if


	sql_sel_casos = sql_sel_casos & " ORDER BY VV.ID_GESTION,  VV.COD_CLIENTE,   VV.CUSTODIO, "	   
	sql_sel_casos = sql_sel_casos & " VV.FECHA_INGRESO_GESTION,    VV.RUT_DEUDOR,   VV.NOMBRE_DEUDOR,    VV.FECHA_GESTION,  "
	sql_sel_casos = sql_sel_casos & " VV.HORA_INGRESO,   VV.MONTO_GESTION,      CFP.DESC_FORMA_PAGO, FR.NOMBRE, "	  
	sql_sel_casos = sql_sel_casos & " FR.UBICACION  ,DD.CALLE  ,DD.NUMERO  ,DD.RESTO  ,DD.COMUNA,  VV.NRO_DOC_PAGO, "
	sql_sel_casos = sql_sel_casos & " VV.OBSERVACIONES_CAMPO,     VV.TIPO_MODULO,  "  
	sql_sel_casos = sql_sel_casos & " VV.ACUMULADO,   VV.FORMA_PAGO,   VV.ID_DIRECCION_COBRO_DEUDOR,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_FORMA_RECAUDACION,   VV.ID_USUARIO_ASIG,   U.LOGIN, VV.PROCESO ,  VV.ID_PROCESO "


	
	'response.write sql_sel_casos&".."
''	response.end()

  	set rs_casos_gestion = conn.execute(sql_sel_casos)
	if err then
		Response.write "ERROR : " & err.description
		Response.end()
	end if
%>
</head>
<body>
	<table border="1" bordercolor="#848484" cellSpacing="0" cellPadding="0">
	<%
	if not rs_casos_gestion.eof then
	%>
		<tr style="background-color:#ccc; height:24px;">
			<td>ESTADO REAL</td>
			<td>N° CP</td>
			<td>FECHA PAGO</td>
			<td>OBSERVACION</td>
			<td>FECHA INGRESO</td>
			<td>FECHA CONSULTA</td>
			<td>ID CUOTA</td>
			<td>TIPO GESTION</td>
			<td>RUT CLIENTE</td>
			<td>NOMBRE CLIENTE</td>
			<td>RUT DEUDOR</td>
			<td>NOMBRE DEUDOR</td>
			<td>NRO DOC</td>
			<td>NRO CUOTA</td>
			<td>TIPO DOC</td>
			<td>FECHA VENC</td>
			<td>MONTO CAPITAL</td>
			<td>SALDO ACTIVO</td>
			<td>INTERLOCUTOR</td>
			<td>SEDE</td>
			<td>MONTO PAGADO</td>
			<td>FORMA PAGO</td>
			<td>FECHA PAGO</td>
			<td>LUGAR PAGO</td>
			<td>EJECUTIVO</td>	
			<td width="1000">OBSERVACIÓN</td>
		</tr>
	<%

		do while not rs_casos_gestion.eof
		IF  (i mod 2)=1 then
			bgcolor="#F2F2F2"
		Else
			bgcolor="#FFFFFF"
		end if
		i = i + 1
		intIdGestion			=rs_casos_gestion("ID_GESTION")
		intIdcuota 				=rs_casos_gestion("ID_CUOTA")
		intCodCliente			=rs_casos_gestion("COD_CLIENTE")
		strCustodio				=rs_casos_gestion("CUSTODIO")
		dtmFechaINgresoGestion 	=rs_casos_gestion("FECHA_INGRESO_GESTION")
		dtmFechaINgresoGestionHora	=rs_casos_gestion("FECHA_INGRESO_GESTION_HORA")			
		strRutDeudor			=rs_casos_gestion("RUT_DEUDOR")
		strNombreDeudor			=rs_casos_gestion("NOMBRE_DEUDOR")
		intSaldoDeudor			=rs_casos_gestion("SALDO_CUOTA")
		dtmFechaVenc			=rs_casos_gestion("FECHA_VENC")
		dtmFechaGestion			=rs_casos_gestion("FECHA_GESTION")
		strHOraIngreso 			=rs_casos_gestion("HORA_INGRESO")
		intMontoGestion			=rs_casos_gestion("MONTO_GESTION")
		strFormaNormalizacion 	=rs_casos_gestion("FORMA_NORMALIZACION")
		strLugarGestion			=rs_casos_gestion("LUGAR_GESTION")
		intNroDocPago 			=rs_casos_gestion("NRO_DOC_PAGO")
		strObservacionesCampo 	=rs_casos_gestion("OBSERVACIONES_CAMPO")
		strTipoModulo 			=rs_casos_gestion("TIPO_MODULO")
		strAcumulado 			=rs_casos_gestion("ACUMULADO")
		dtmFechaConsult_norm 	=rs_casos_gestion("MIN_FECHA_CONSULTA")
		dtmFechaConsultNormHora =rs_casos_gestion("MIN_FECHA_CONSULTA_HORA")			
		strEjecutivo 			=rs_casos_gestion("LOGIN")
		strProceso 				=rs_casos_gestion("PROCESO")

		'intCuotasActivas 		=rs_casos_gestion("CUOTAS_ACTIVAS")
		'intCuotasCanceladas		=rs_casos_gestion("CUOTAS_CANCELADAS")
		'intCuotasRetiradas		=rs_casos_gestion("CUOTAS_RETIRADAS")
		'intCuotasNoAsignables	=rs_casos_gestion("CUOTAS_NO_ASIGNABLES")
		'intCantidadDocumentos 	=rs_casos_gestion("CANTIDAD_DOCUMENTOS")

		strFormaPago 		 	=rs_casos_gestion("FORMA_PAGO")				
		intIdDireccionCobro 	=rs_casos_gestion("ID_DIRECCION_COBRO_DEUDOR")				
		intIdFormarecaudacion 	=rs_casos_gestion("ID_FORMA_RECAUDACION")				
		intUsuarioAsig 			=rs_casos_gestion("ID_USUARIO_ASIG")				
		intIdProceso		 	=rs_casos_gestion("ID_PROCESO")				
		strTipoGestion 		 	=rs_casos_gestion("GESTIONSOLA")				
		strRutSubCliente 		=rs_casos_gestion("RUT_SUBCLIENTE")				
		strNombreSubcliente 	=rs_casos_gestion("NOMBRE_SUBCLIENTE")				
		intNroDoc 			 	=rs_casos_gestion("NRO_DOC")				
		intNroCuota 		 	=rs_casos_gestion("NRO_CUOTA")				
		strNomTipoDOcumento 	=rs_casos_gestion("NOM_TIPO_DOCUMENTO")				
		intValorCuota 		 	=rs_casos_gestion("VALOR_CUOTA")				
		intSaldo 				=rs_casos_gestion("SALDO")	
		strInterlocutor			=rs_casos_gestion("INTERLOCUTOR")	
		strSucursal				=rs_casos_gestion("SUCURSAL")	
		dtmFechaNormalizacion	=rs_casos_gestion("FECHA_NORMALIZACION")	


		'If Trim(intCuotasActivas) <> "" Then
		'	strTextoDocAct 		= "Doc.Asociados : " & intCuotasActivas & "<BR>"
		'End If

		'If Trim(intCuotasCanceladas) <> "" Then
		'	strTextoDocPag 		= "Doc.Cancelados : " & intCuotasCanceladas & "<BR>"
		'End If

		'If Trim(intCuotasRetiradas) <> "" Then
		'	strTextoDocRet 		= "Doc.Desasignados : " & intCuotasRetiradas & "<BR>"
		'End If

		'If Trim(intCuotasNoAsignables) <> "" Then
		'	strTextoDocNoAsig 	= "Doc.No Asignable : " & intCuotasNoAsignables & "<BR>"
		'End If
		
		'strTextoDoc ="Nombre deudor :"&strNombreDeudor &"<br>" & strTextoDocAct & strTextoDocPag & strTextoDocRet & strTextoDocNoAsig



		%>
		<tr class="td_hover" bgcolor="<%=bgcolor%>">
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td valign="top"><%=trim(dtmFechaINgresoGestion)%></td>
			<td valign="top"><%=trim(dtmFechaConsult_norm)%></td>
			<td valign="top"><%=trim(intIdcuota)%></td>
			<td valign="top"><%=trim(strTipoGestion)%></td>
			<td valign="top"><%=trim(strRutSubCliente)%></td>
			<td valign="top"><%=trim(strNombreSubcliente)%></td>
			<td valign="top"><%=trim(strRutDeudor)%></td>
			<td valign="top"><%=trim(strNombreDeudor)%></td>
			<td valign="top"><%=trim(intNroDoc)%></td>
			<td valign="top"><%=trim(intNroCuota)%></td>
			<td valign="top"><%=trim(strNomTipoDOcumento)%></td>
			<td valign="top"><%=trim(dtmFechaVenc)%></td>
			<td valign="top"><%=trim(intValorCuota)%></td>
			<td valign="top"><%=trim(intSaldo)%></td>
			<td valign="top"><%=trim(strInterlocutor)%></td>
			<td valign="top"><%=trim(strSucursal)%></td>
			<td valign="top"><%=trim(intMontoGestion)%></td>
			<td valign="top"><%=trim(strFormaPago)%></td>
			<td valign="top"><%=trim(dtmFechaNormalizacion)%></td>
			<td valign="top"><%=trim(strLugarGestion)%></td>
			<td valign="top"><%=trim(strEjecutivo)%></td>
			<td valign="top"><%=trim(strObservacionesCampo)%></td>
		</tr>

			<%
		Response.flush()
		rs_casos_gestion.movenext
		loop

	Else
		Response.write "<BR>SIN REGISTROS SEGÚN PARAMETROS DE BÚSQUEDA"
	end if	

%>

</body>
</html>