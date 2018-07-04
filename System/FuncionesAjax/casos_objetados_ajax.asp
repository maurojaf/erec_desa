<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../../lib/asp/comunes/general/rutinasBooleano.inc"-->
<%

Response.CodePage = 65001
Response.charset="utf-8"



accion_ajax =request("accion_ajax")

abrirscg()

if trim(accion_ajax)="filtra_usuario" then

	CB_COBRANZA =request.querystring("CB_COBRANZA")

	if trim(CB_COBRANZA)="INTERNA" then
		PERFIL_EMP =1
	end if
	
	if trim(CB_COBRANZA)="EXTERNA" then
		PERFIL_EMP =0
	end if	

	sql_usuario ="SELECT DISTINCT U.ID_USUARIO, LOGIN "
	sql_usuario = sql_usuario & " FROM USUARIO U "
	sql_usuario = sql_usuario & " INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO "
	sql_usuario = sql_usuario & " AND UC.COD_CLIENTE = '"&trim(session("ses_codcli"))&"' "
	sql_usuario = sql_usuario & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1 "

	if trim(CB_COBRANZA)<>"TODOS" then
		sql_usuario = sql_usuario & " AND U.PERFIL_EMP= " &  trim(PERFIL_EMP)
	end if

	SET rs_sel = conn.execute(sql_usuario)
	if err then
		Response.write "ERROR : " & err.description
		Response.end()
	end if

	%>
    <select name="CB_EJECUTIVO"  id="CB_EJECUTIVO">
        <option value="">TODOS</option>  
        <%DO WHILE NOT rs_sel.eof%>      
        	<option value="<%=trim(rs_sel("ID_USUARIO"))%>"><%=trim(rs_sel("LOGIN"))%></option>    
        <%rs_sel.movenext
        loop%>
    </select> 
<%

elseif trim(accion_ajax)="refresa_objetados" then
	
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


	session("Ftro_EstadoProcesoCasosObj") = strEstadoProceso
	session("Ftro_EjecAsigCasosObj") = strEjecutivo
	
	''response.write strEstadoProceso	
	''response.end
	
	''strEstadoProceso = session("Ftro_EstadoProcesoCasosObj")
	
	sql_sel_casos = " "
	sql_sel_casos = sql_sel_casos & " SELECT VV.ID_GESTION,  "
	sql_sel_casos = sql_sel_casos & " VV.COD_CLIENTE,  "
	sql_sel_casos = sql_sel_casos & " VV.CUSTODIO,  "
	sql_sel_casos = sql_sel_casos & " CONVERT(VARCHAR, MIN(VV.FECHA_INGRESO_GESTION),103) FECHA_INGRESO_GESTION,  " 
	sql_sel_casos = sql_sel_casos & " SUBSTRING(CONVERT(VARCHAR, MIN(VV.FECHA_INGRESO_GESTION),108),1,5) FECHA_INGRESO_GESTION_HORA,  " 	
	sql_sel_casos = sql_sel_casos & " VV.RUT_DEUDOR,  "
	sql_sel_casos = sql_sel_casos & " VV.NOMBRE_DEUDOR, "  
	sql_sel_casos = sql_sel_casos & " SUM(VV.SALDO_CUOTA) SUM_SALDO_CUOTA,   "
	sql_sel_casos = sql_sel_casos & " max(VV.FECHA_VENC) MIN_FECHA_VENC,   "
	sql_sel_casos = sql_sel_casos & " DATEDIFF(D,max(VV.FECHA_VENC),GETDATE()) MIN_DIA_MORA, "
	sql_sel_casos = sql_sel_casos & " VV.FECHA_GESTION,  "
	sql_sel_casos = sql_sel_casos & " VV.HORA_INGRESO,  "
	sql_sel_casos = sql_sel_casos & " VV.MONTO_GESTION,  "   
	sql_sel_casos = sql_sel_casos & " FORMA_NORMALIZACION = ISNULL(CFP.DESC_FORMA_PAGO,'NO ESPEC.'),  "
	sql_sel_casos = sql_sel_casos & " LUGAR_GESTION = ISNULL(ISNULL(UPPER(FR.NOMBRE+' '+FR.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.COMUNA)),'NO ESPEC.'),   "
	sql_sel_casos = sql_sel_casos & " VV.NRO_DOC_PAGO,  "
	sql_sel_casos = sql_sel_casos & " VV.OBSERVACIONES_CAMPO,   "
	sql_sel_casos = sql_sel_casos & " VV.TIPO_MODULO,  "
	sql_sel_casos = sql_sel_casos & " VV.ACUMULADO,  "
	sql_sel_casos = sql_sel_casos & " CONVERT(VARCHAR, MIN(VV.FECHA_CONSULTA),103) MIN_FECHA_CONSULTA, " 
	sql_sel_casos = sql_sel_casos & " SUBSTRING(CONVERT(VARCHAR, MIN(VV.FECHA_CONSULTA),108),1,5) MIN_FECHA_CONSULTA_HORA, " 
	sql_sel_casos = sql_sel_casos & " VV.FORMA_PAGO,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_DIRECCION_COBRO_DEUDOR,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_FORMA_RECAUDACION,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_USUARIO_ASIG,  "
	sql_sel_casos = sql_sel_casos & " U.LOGIN,   "
	sql_sel_casos = sql_sel_casos & " VV.PROCESO , "
	sql_sel_casos = sql_sel_casos & " VV.ID_PROCESO,  "
	sql_sel_casos = sql_sel_casos & " [dbo].[concatena_cuotas_estados] (VV.ID_GESTION,'ACTIVAS') AS CUOTAS_ACTIVAS, "          
	sql_sel_casos = sql_sel_casos & " [dbo].[concatena_cuotas_estados] (VV.ID_GESTION,'CANCELADAS') AS CUOTAS_CANCELADAS, "
	sql_sel_casos = sql_sel_casos & " [dbo].[concatena_cuotas_estados] (VV.ID_GESTION,'RETIRADAS') AS CUOTAS_RETIRADAS, "
	sql_sel_casos = sql_sel_casos & " [dbo].[concatena_cuotas_estados] (VV.ID_GESTION,'NO ASIGNABLE') AS CUOTAS_NO_ASIGNABLES, "
	sql_sel_casos = sql_sel_casos & " SUM(CASE WHEN VV.ID_ARCHIVO IS NOT NULL THEN 1 ELSE 0 END) CANTIDAD_DOCUMENTOS, "
	sql_sel_casos = sql_sel_casos & " VV.OBSERVACION_CONSULTA, "
	sql_sel_casos = sql_sel_casos & " VV.FECHA_AGENDAMIENTO "
	sql_sel_casos = sql_sel_casos & " FROM VIEW_CASOS_GESTION_APOYO VV  "
	sql_sel_casos = sql_sel_casos & " LEFT JOIN CAJA_FORMA_PAGO CFP ON VV.FORMA_PAGO = CFP.ID_FORMA_PAGO  "
	sql_sel_casos = sql_sel_casos & " LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION=VV.ID_DIRECCION_COBRO_DEUDOR  "
	sql_sel_casos = sql_sel_casos & " LEFT JOIN FORMA_RECAUDACION FR ON FR.ID_FORMA_RECAUDACION=VV.ID_FORMA_RECAUDACION  "
	sql_sel_casos = sql_sel_casos & " INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON VV.COD_CATEGORIA = GTC.COD_CATEGORIA  "
	sql_sel_casos = sql_sel_casos & " INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTS ON VV.COD_CATEGORIA = GTS.COD_CATEGORIA AND VV.COD_SUB_CATEGORIA = GTS.COD_SUB_CATEGORIA  "
	sql_sel_casos = sql_sel_casos & " LEFT JOIN USUARIO U ON VV.ID_USUARIO_ASIG = U.ID_USUARIO   "


  	sql_sel_casos = sql_sel_casos & " WHERE VV.COD_CLIENTE ='"&TRIM(strCodCliente)&"' AND VV.TIPO_MODULO = 3"
  	
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

	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
		sql_sel_casos = sql_sel_casos & " 	AND  U.ID_USUARIO = '" & session("session_idusuario") & "'"
	Else
	  	if trim(strEjecutivo)<>"" then
	  		sql_sel_casos = sql_sel_casos & " AND  U.ID_USUARIO='"&strEjecutivo&"'"
	  	end if
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

	sql_sel_casos = sql_sel_casos & " GROUP BY VV.ID_GESTION, " 
	sql_sel_casos = sql_sel_casos & " VV.COD_CLIENTE,  "
	sql_sel_casos = sql_sel_casos & " VV.CUSTODIO,  "
	sql_sel_casos = sql_sel_casos & " VV.FECHA_INGRESO_GESTION,   "
	sql_sel_casos = sql_sel_casos & " VV.RUT_DEUDOR,  "
	sql_sel_casos = sql_sel_casos & " VV.NOMBRE_DEUDOR,   "
	sql_sel_casos = sql_sel_casos & " VV.FECHA_GESTION,  "
	sql_sel_casos = sql_sel_casos & " VV.HORA_INGRESO,  "
	sql_sel_casos = sql_sel_casos & " VV.MONTO_GESTION,     "
	sql_sel_casos = sql_sel_casos & " CFP.DESC_FORMA_PAGO,"
	sql_sel_casos = sql_sel_casos & " FR.NOMBRE "
	sql_sel_casos = sql_sel_casos & " ,FR.UBICACION "
	sql_sel_casos = sql_sel_casos & " ,DD.CALLE "
	sql_sel_casos = sql_sel_casos & " ,DD.NUMERO "
	sql_sel_casos = sql_sel_casos & " ,DD.RESTO "
	sql_sel_casos = sql_sel_casos & " ,DD.COMUNA, "
	sql_sel_casos = sql_sel_casos & " VV.NRO_DOC_PAGO,   "
	sql_sel_casos = sql_sel_casos & " VV.OBSERVACIONES_CAMPO,    "
	sql_sel_casos = sql_sel_casos & " VV.TIPO_MODULO,   "
	sql_sel_casos = sql_sel_casos & " VV.ACUMULADO,  "
	sql_sel_casos = sql_sel_casos & " VV.FORMA_PAGO,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_DIRECCION_COBRO_DEUDOR,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_FORMA_RECAUDACION,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_USUARIO_ASIG,  "
	sql_sel_casos = sql_sel_casos & " U.LOGIN,   " 
	sql_sel_casos = sql_sel_casos & " VV.PROCESO , "
	sql_sel_casos = sql_sel_casos & " VV.ID_PROCESO, VV.OBSERVACION_CONSULTA, VV.FECHA_AGENDAMIENTO  "

	if trim(CH_CP_ADJUNTO)="1" then
		sql_sel_casos = sql_sel_casos & " HAVING SUM(CASE WHEN VV.ID_ARCHIVO IS NOT NULL THEN 1 ELSE 0 END) > 0   "
	end if

	if trim(CH_CP_ADJUNTO)="2" then
		sql_sel_casos = sql_sel_casos & " HAVING SUM(CASE WHEN VV.ID_ARCHIVO IS NOT NULL THEN 1 ELSE 0 END) = 0   "
	end if

	sql_sel_casos = sql_sel_casos & " ORDER BY VV.ID_GESTION, " 
	sql_sel_casos = sql_sel_casos & " VV.COD_CLIENTE,  "
	sql_sel_casos = sql_sel_casos & " VV.CUSTODIO,  "
	sql_sel_casos = sql_sel_casos & " VV.FECHA_INGRESO_GESTION,   "
	sql_sel_casos = sql_sel_casos & " VV.RUT_DEUDOR,  "
	sql_sel_casos = sql_sel_casos & " VV.NOMBRE_DEUDOR,   "
	sql_sel_casos = sql_sel_casos & " VV.FECHA_GESTION,  "
	sql_sel_casos = sql_sel_casos & " VV.HORA_INGRESO,  "
	sql_sel_casos = sql_sel_casos & " VV.MONTO_GESTION,     "
	sql_sel_casos = sql_sel_casos & " CFP.DESC_FORMA_PAGO,"
	sql_sel_casos = sql_sel_casos & " FR.NOMBRE "
	sql_sel_casos = sql_sel_casos & " ,FR.UBICACION "
	sql_sel_casos = sql_sel_casos & " ,DD.CALLE "
	sql_sel_casos = sql_sel_casos & " ,DD.NUMERO "
	sql_sel_casos = sql_sel_casos & " ,DD.RESTO "
	sql_sel_casos = sql_sel_casos & " ,DD.COMUNA, "
	sql_sel_casos = sql_sel_casos & " VV.NRO_DOC_PAGO,   "
	sql_sel_casos = sql_sel_casos & " VV.OBSERVACIONES_CAMPO,    "
	sql_sel_casos = sql_sel_casos & " VV.TIPO_MODULO,   "
	sql_sel_casos = sql_sel_casos & " VV.ACUMULADO,  "
	sql_sel_casos = sql_sel_casos & " VV.FORMA_PAGO,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_DIRECCION_COBRO_DEUDOR,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_FORMA_RECAUDACION,  "
	sql_sel_casos = sql_sel_casos & " VV.ID_USUARIO_ASIG,  "
	sql_sel_casos = sql_sel_casos & " U.LOGIN,   " 
	sql_sel_casos = sql_sel_casos & " VV.PROCESO , "
	sql_sel_casos = sql_sel_casos & " VV.ID_PROCESO, VV.OBSERVACION_CONSULTA, VV.FECHA_AGENDAMIENTO  "	

	
	'response.write sql_sel_casos
  	set rs_casos_gestion = conn.execute(sql_sel_casos)
	if err then
		Response.write "ERROR : " & err.description
		Response.end()
	end if
%>
	<table class="" style="width:100%;" border="0" cellSpacing="0" cellPadding="0">
	<%
	if not rs_casos_gestion.eof then
	intContador =0
	do while not rs_casos_gestion.eof
		intContador = intContador + 1

		if cint(intContador) <= cint(intInicioContador) then

			intIdGestion			=rs_casos_gestion("ID_GESTION")
			intCodCliente			=rs_casos_gestion("COD_CLIENTE")
			strCustodio				=rs_casos_gestion("CUSTODIO")
			dtmFechaINgresoGestion 	=rs_casos_gestion("FECHA_INGRESO_GESTION")
			dtmFechaINgresoGestionHora	=rs_casos_gestion("FECHA_INGRESO_GESTION_HORA")			
			strRutDeudor			=rs_casos_gestion("RUT_DEUDOR")
			strNombreDeudor			=rs_casos_gestion("NOMBRE_DEUDOR")
			intSaldoDeudor			=rs_casos_gestion("SUM_SALDO_CUOTA")
			dtmFechaVenc			=rs_casos_gestion("MIN_FECHA_VENC")
			intDiaMora				=rs_casos_gestion("MIN_DIA_MORA")
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

			intCuotasActivas 		=rs_casos_gestion("CUOTAS_ACTIVAS")
			intCuotasCanceladas		=rs_casos_gestion("CUOTAS_CANCELADAS")
			intCuotasRetiradas		=rs_casos_gestion("CUOTAS_RETIRADAS")
			intCuotasNoAsignables	=rs_casos_gestion("CUOTAS_NO_ASIGNABLES")
			intCantidadDocumentos 	=rs_casos_gestion("CANTIDAD_DOCUMENTOS")
			strObservacionConsulta 	=rs_casos_gestion("OBSERVACION_CONSULTA")
			dtmFechaAgendamiento 	=rs_casos_gestion("FECHA_AGENDAMIENTO")

			If Trim(intCuotasActivas) <> "" Then
				strTextoDocAct 		= "<b>Doc.Asociados :</b> " & intCuotasActivas & "<BR><br>"
			End If

			If Trim(intCuotasCanceladas) <> "" Then
				strTextoDocPag 		= "<b>Doc.Cancelados :</b> " & intCuotasCanceladas & "<BR><br>"
			End If

			If Trim(intCuotasRetiradas) <> "" Then
				strTextoDocRet 		= "<b>Doc.Desasignados :</b> " & intCuotasRetiradas & "<BR><br>"
			End If

			If Trim(intCuotasNoAsignables) <> "" Then
				strTextoDocNoAsig 	= "<b>Doc.No Asignable :</b> " & intCuotasNoAsignables & "<BR><br>"
			End If

			
			strTextoDoc ="<b>Nombre deudor :</b> "&strNombreDeudor &"<br><br>" & strTextoDocAct & strTextoDocPag & strTextoDocRet & strTextoDocNoAsig



			%>
			<tr class="td_hover">
				<td width="20">
				<%if trim(strProceso)<>"EN CONSULTA" and trim(strProceso)<>"NO RESPONDIDO" then%>
					<input type="checkbox" name="CH_CASOS_APOYO" ID="CH_CASOS_APOYO"  value="<%=intIdGestion%>">
				<%end if%>
				</td>
				<td width="20"><%=trim(intContador)%></td>
				<td width="100">CASOS OBJETADOS</td>
				<td width="100"><%=trim(strProceso)%></td>
				<td width="80" 
					<%if not isnull(dtmFechaConsult_norm) then
						if trim(strObservacionConsulta)="" or isnull(strObservacionConsulta) then
							strObservacionConsulta = "SIN OBSERVACIÓN"
						end if	

						dtmFechaAgendamiento =mid(dtmFechaAgendamiento,1, len(dtmFechaAgendamiento)-3)


					%>
						title="<%="<table><tr><td width='140'><b>HORA CONSULTA</b></td><td>:"&dtmFechaConsultNormHora&"</td></tr><tr><td><b>OBSERVACIÓN</b></td><td>:"&ucase(strObservacionConsulta)&"</td></tr><tr><td><b>FECHA AGENDAMIENTO</b></td><td>:"&dtmFechaAgendamiento&"</td></tr></table>"%>"
					<%else%>
						title="NO CONSULTADO"
					<%end if%>>
					<%
					if isnull(dtmFechaConsult_norm) then
						Response.write "NO CONSULT"
					else
						response.write trim(dtmFechaConsult_norm)
					end if%>

				</td>
				<td width="40"><%=trim(strAcumulado)%></td>

				<td width="70" title=""><%=trim(dtmFechaINgresoGestion)%></td>
				<td width="80" onclick="bt_trae_cuotas_vista('<%=intIdGestion%>','<%=strRutDeudor%>','<%=intCodCliente%>')"><a href="#"><%=trim(strRutDeudor)%></a></td>
				<td width="80"><%=FN(intSaldoDeudor,0)%></td>
				<td width="70"><%=trim(intDiaMora)%></td>
				<td width="70"><%=trim(dtmFechaGestion)%></td>
				<td width="70" title="<%=trim(intMontoGestion)%>">
					<%
 					if trim(intMontoGestion)="NO INGRESADO" then
						response.write mid(intMontoGestion,1,10)
					else
						response.write FN(intMontoGestion,0)
					end if%>
				</td>

				<td width="100" title="<%=trim(strFormaNormalizacion)%>"><%=mid(strFormaNormalizacion,1,15)%></td>
				<td width="100" title="<%=trim(strLugarGestion)%>">&nbsp;<%=mid(strLugarGestion,1,12)%></td>
				<td width="70"><%=trim(intNroDocPago)%></td>
				<td width="100"><%=trim(strEjecutivo)%></td>
				<td width="30" align="center" title="<%=strObservacionesCampo%>">
					<img src="../imagenes/priorizar_normal.png" border="0">
				</td>
				<td width="30" align="center" title="<%=strTextoDoc%>">
					<img src="../imagenes/bt_editar.png" width="20" height="20" border="0">
				</td>				
				<td width="50" align="center">
					<%IF trim(intCantidadDocumentos)>0 then%>
						<img src="../imagenes/48px-Crystal_Clear_app_kappfinder.png" id="imagen_muestra_cuotas_<%=intIdGestion%>" width="20" height="20" style="cursor:pointer;" border="0" onclick="busca_cuotas('<%=intIdGestion%>','<%=strRutDeudor%>')">
						<img src="../imagenes/48px-Crystal_Clear_app_kappfinder.png" id="imagen_oculta_cuotas_<%=intIdGestion%>" width="20" style="display:none;" height="20" style="cursor:pointer;" border="0" onclick="oculta_cuotas('<%=intIdGestion%>','<%=strRutDeudor%>')">
					<%else%>
						<img src="../imagenes/48px-Crystal_Clear_app_kappfinder_rojo.png" id="imagen_muestra_cuotas_<%=intIdGestion%>" width="20" height="20" style="cursor:pointer;" border="0" onclick="busca_cuotas('<%=intIdGestion%>','<%=strRutDeudor%>')">
						<img src="../imagenes/48px-Crystal_Clear_app_kappfinder_rojo.png" id="imagen_oculta_cuotas_<%=intIdGestion%>" width="20" style="display:none;" height="20" style="cursor:pointer;" border="0" onclick="oculta_cuotas('<%=intIdGestion%>','<%=strRutDeudor%>')">

					<%end if%>

				</td>
			</tr>
			<tr>
				<td colspan="19" id="refresca_busca_cuotas_<%=intIdGestion%>"></td>
			</tr>
		<%end if

	Response.flush()
	rs_casos_gestion.movenext
	loop

	%>
<%
	Else
		Response.write "<BR>SIN REGISTROS SEGÚN PARAMETROS DE BÚSQUEDA"
	end if	

	fin = intInicioContador +20
	%>				
	</table>
	<br>
	<br>
	<%'if cint(intContador) > cint(fin) then%>
	<div class="fondo_boton_100 mas_registros" onclick="bt_mostrar_mas_registros('<%=fin%>')">
		Mas registros <%=fin%>	
	</div>
	<%'end if%>
		
<%	

elseif trim(accion_ajax)="refresca_resumen" then
	
	strCobranza 		=request.querystring("CB_COBRANZA")
	strTipoGestion 		=request.querystring("CMB_TIPO_GESTION")
	strEstadoProceso 	=request.querystring("CMB_ESTADO_PROCESO")
	dtmInicio			=request.querystring("inicio")
	dtmTermino 			=request.querystring("termino")
	strEjecutivo 		=request.querystring("CB_EJECUTIVO")
	strCodCliente		=request.querystring("COD_CLIENTE")

	strRutDeudor 		=request.querystring("RUT_DEUDOR")
	HORA_CONSULTA  		=request.querystring("HORA_CONSULTA")
	FECHA_CONSULTA  	=request.querystring("FECHA_CONSULTA")
	CH_CP_ADJUNTO 		=request.querystring("CH_CP_ADJUNTO")
	sql_resumen = ""
	sql_resumen = sql_resumen & " SELECT "
	sql_resumen = sql_resumen & " 'CASOS OBJETADOS' TIPO_GESTION, "
	sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'NO RESPONDIDO' THEN 1 ELSE 0 END),0) AS TOTAL_DOC_NO_RESPONDIDO, "
	sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'NO RESPONDIDO' THEN SALDO_CUOTA ELSE 0 END),0) AS TOTAL_SALDO_NO_RESPONDIDO, "
	sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'NO RESPONDIDO' THEN RUT_DEUDOR END)) AS TOTAL_CASOS_NO_RESPONDIDO, "
	sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'NO RESPONDIDO' THEN ID_GESTION END)) AS TOTAL_GESTIONES_NO_RESPONDIDO,"
	sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'NO PROCESADO' THEN 1 ELSE 0 END),0) AS TOTAL_DOC_NO_PROCESADO, "
	sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'NO PROCESADO' THEN SALDO_CUOTA ELSE 0 END),0) AS TOTAL_SALDO_NO_PROCESADO, "
	sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'NO PROCESADO' THEN RUT_DEUDOR END)) AS TOTAL_CASOS_NO_PROCESADO, "
	sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'NO PROCESADO' THEN ID_GESTION END)) AS TOTAL_GESTIONES_NO_PROCESADO,"


	sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'EN CONSULTA' THEN 1 ELSE 0 END),0) AS TOTAL_DOC_NO_CONSULTA, "
	sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'EN CONSULTA' THEN SALDO_CUOTA ELSE 0 END),0) AS TOTAL_SALDO_NO_CONSULTA, "
	sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'EN CONSULTA' THEN RUT_DEUDOR END)) AS TOTAL_CASOS_NO_CONSULTA, "
	sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'EN CONSULTA' THEN ID_GESTION END)) AS TOTAL_GESTIONES_NO_CONSULTA "
	sql_resumen = sql_resumen & " FROM VIEW_CASOS_GESTION_APOYO "
	sql_resumen = sql_resumen & " WHERE COD_CLIENTE ='"&TRIM(strCodCliente)&"' AND TIPO_MODULO = 3"
 	

  	if trim(strCobranza)="INTERNA" then ' ELLOS NO LLACRUZ
  		sql_resumen = sql_resumen & " AND CUSTODIO IS NOT NULL "

  	ElseIf trim(strCobranza)="EXTERNA" then ' COBRANZA LLACRUZ'
  		sql_resumen = sql_resumen & " AND CUSTODIO IS  NULL "

  	end if

  	if trim(dtmInicio)<>"" and trim(dtmTermino)<>"" then
  		sql_resumen = sql_resumen & " AND  convert(datetime, FECHA_INGRESO_GESTION) BETWEEN  convert(datetime, '"&trim(dtmInicio)&"') AND  convert(datetime, '"&trim(dtmTermino)&"')"

  	end  if
  	if trim(dtmInicio)="" and trim(dtmTermino)<>"" then
  		sql_resumen = sql_resumen & " AND  convert(datetime, FECHA_INGRESO_GESTION) <= convert(datetime, '"&trim(dtmTermino)&"')"

  	end  if
  	if trim(dtmInicio)<>"" and trim(dtmTermino)="" then
  		sql_resumen = sql_resumen & " AND  convert(datetime, FECHA_INGRESO_GESTION) >=  convert(datetime, '"&trim(dtmInicio)&"')"

  	end  if  	  	

	If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
		sql_resumen = sql_resumen & " 	AND  ID_USUARIO_ASIG = '" & session("session_idusuario") & "'"
	Else
	  	if trim(strEjecutivo)<>"" then
	  		sql_resumen = sql_resumen & " AND  ID_USUARIO_ASIG='"&strEjecutivo&"'"
	  	end if
	end if

  	if trim(FECHA_CONSULTA)<>"" then
  		sql_resumen = sql_resumen & " AND convert(varchar,FECHA_CONSULTA, 103) = '"&trim(FECHA_CONSULTA)&"' "

  	end if

  	if trim(HORA_CONSULTA)<>"" then
  		sql_resumen = sql_resumen & " AND SUBSTRING(convert(varchar,FECHA_CONSULTA, 108), 1, 5) = '"&trim(HORA_CONSULTA)&"' "

  	end if

  	if trim(strRutDeudor)<>"" then
  		sql_resumen = sql_resumen & " AND RUT_DEUDOR ='"&trim(strRutDeudor)&"' "
  	end if

	if trim(CH_CP_ADJUNTO)="1" then
		sql_resumen = sql_resumen &  " AND ID_ARCHIVO IS NOT NULL  "
	end if

	if trim(CH_CP_ADJUNTO)="2" then
		sql_resumen = sql_resumen & "  AND ID_ARCHIVO IS NULL "
	end if

  	'response.write sql_resumen
	
	set rs_resumen =conn.execute(sql_resumen)
	if err then
		Response.write "ERROR : " & err.description
		Response.end()
	end if	
	
	if not rs_resumen.eof then

		strtipoGestion 					=rs_resumen("TIPO_GESTION")
		IntTotalDocNoRespondido 		=rs_resumen("TOTAL_DOC_NO_RESPONDIDO")
		IntTotalSaldoNoRespondido 		=rs_resumen("TOTAL_SALDO_NO_RESPONDIDO")
		IntTotalCasosNoRespondido	 	=rs_resumen("TOTAL_CASOS_NO_RESPONDIDO")
		intTotalGestionesNoRespondido	=rs_resumen("TOTAL_GESTIONES_NO_RESPONDIDO")
		intTotalDocNoProcesado 			=rs_resumen("TOTAL_DOC_NO_PROCESADO")
		intTotalSaldoNoProcesado 		=rs_resumen("TOTAL_SALDO_NO_PROCESADO")
		intTotalCasosNoProcesado	 	=rs_resumen("TOTAL_CASOS_NO_PROCESADO")
		intTotalGestionesNoProcesado	=rs_resumen("TOTAL_GESTIONES_NO_PROCESADO")
		intTotalDocNoConsulta 			=rs_resumen("TOTAL_DOC_NO_CONSULTA")
		intTotalSaldoNoConsulta 		=rs_resumen("TOTAL_SALDO_NO_CONSULTA")
		intTotalCasosNoConsulta	 		=rs_resumen("TOTAL_CASOS_NO_CONSULTA")
		intTotalGestionesNoConsulta	 	=rs_resumen("TOTAL_GESTIONES_NO_CONSULTA") 

	end if

%>
	<table class="" style="width:78%; margin-left:5%;" align="left" cellSpacing="0" cellPadding="0" border="0">
		<tr>
			<td style="width:100px;" class=""></td>
			<td style="width:100px;" class="estilo_columna_individual td_bordes" align="center">NO RESPONDIDOS</td>
			<td style="width:100px;" class="estilo_columna_individual td_bordes" align="center">NO PROCESADOS</td>
			<td style="width:100px;" class="estilo_columna_individual td_bordes" align="center">EN CONSULTA</td>
			<td style="width:100px;background-color:#424242;" class="estilo_columna_individual td_bordes" align="center">TOTAL ACUM. GESTION</td>
		</tr>
		<tr>
			<td style="width:100px;" class="estilo_columna_individual" align="center">TIPO GESTIÓN</td>
			<td class="" align="center">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" height="25" align="center"><b>Total Doc</b></td>
					<td width="33%" height="25" align="center"><b>Total casos</b></td>
					<td width="33%" height="25" align="center"><b>Monto Doc</b></td>
				</tr>					
				</table>
			</td>
			<td class="" align="center">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" height="25" align="center"><b>Total Doc</b></td>
					<td width="33%" height="25" align="center"><b>Total casos</b></td>
					<td width="33%" height="25" align="center"><b>Monto Doc</b></td>
				</tr>					
				</table>
			</td>
			<td class="" align="center">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" height="25" align="center"><b>Total Doc</b></td>
					<td width="33%" height="25" align="center"><b>Total casos</b></td>
					<td width="33%" height="25" align="center"><b>Monto Doc</b></td>
				</tr>					
				</table>
			</td>
			<td class="" align="center">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" height="25" align="center"><b>Totales Doc</b></td>
					<td width="33%" height="25" align="center"><b>Totales casos</b></td>
					<td width="33%" height="25" align="center"><b>Monto tot Doc</b></td>
				</tr>					
				</table>
			</td>			
		</tr>
		<tr>
			 	
			<td style="width:100px;" class="estilo_columna_individual td_bordes" align="center" >CASOS OBJETADOS</td>
			<td class="td_bordes">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalDocNoRespondido,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalCasosNoRespondido,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalSaldoNoRespondido,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalDocNoProcesado,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalCasosNoProcesado,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalSaldoNoProcesado,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalDocNoConsulta,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalCasosNoConsulta,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalSaldoNoConsulta,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalDocNoRespondido)+(intTotalDocNoProcesado)+(intTotalDocNoConsulta),0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalCasosNoRespondido)+(intTotalCasosNoProcesado)+(intTotalCasosNoConsulta),0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalSaldoNoRespondido)+(intTotalSaldoNoProcesado)+(intTotalSaldoNoConsulta),0)%></td>
				</tr>
				</table>
			</td>			
		</tr>

		<tr>
			<td style="width:100px; background-color:#424242;" class="estilo_columna_individual" align="center">TOTAL ACUM. PROCESOS</td>
			<td class="td_bordes" bgcolor="#D8D8D8">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalDocNoRespondido,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalCasosNoRespondido,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalSaldoNoRespondido,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes" bgcolor="#D8D8D8">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalDocNoProcesado,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalCasosNoProcesado,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalSaldoNoProcesado,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes" bgcolor="#D8D8D8">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalDocNoConsulta,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalCasosNoConsulta,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalSaldoNoConsulta,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes" bgcolor="#D8D8D8">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalDocNoRespondido)+(intTotalDocNoProcesado)+(intTotalDocNoConsulta),0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalCasosNoRespondido)+(intTotalCasosNoProcesado)+(intTotalCasosNoConsulta),0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalSaldoNoRespondido)+(intTotalSaldoNoProcesado)+(intTotalSaldoNoConsulta),0)%></td>
				</tr>
				</table>
			</td>			
		</tr>

	</table>


	<table style="float:right; margin-right:5%;" border="0">
	<tr>
		<td>
			<input 	type="button"  	class="fondo_boton_100" STYLE="MARGIN-TOP:2PX;" value="Ver" id="boton_ver" onClick="consulta_resumen();consulta_detalle();"><BR>
			<input  type="button" id="boton_exportar" 	class="fondo_boton_100" STYLE="MARGIN-TOP:2PX;" Value="Exportar" onClick="exportar();"><BR>

			<%If TraeSiNo(session("perfil_adm")) = "Si" OR TraeSiNo(session("perfil_sup")) = "Si" Then%>
				<input  type="button" 	class="fondo_boton_100" STYLE="MARGIN-TOP:2PX;" Value="Procesar" onClick="ventana_procesa();">	
			<%End if%>
			
		</td>
	</tr>

	</table>
	<br>
	<br>		
<%


elseif trim(accion_ajax)="proceso_casos" then

	concat_CH_ID_GESTION 	=request.querystring("concat_CH_ID_GESTION")
	OBSERVACION_CONSULTA 	=request.querystring("OBSERVACION_CONSULTA")
	strObservacionConsulta  =request.querystring("TX_OBSERVACION_CONSULTA")
	dtmFechaAgendamiento 	=request.querystring("TX_FECHA_AGENDAMIENTO")
	strHoraAgendamiento 	=request.querystring("TX_HORA_AGENDAMIENTO")

	valor_CH_ID_GESTION 	=split(concat_CH_ID_GESTION,"*")
	total_CH_ID_GESTION 	=ubound(valor_CH_ID_GESTION)
	fecha 					=now()
	for indice = 1 to total_CH_ID_GESTION

		SQL_SEL ="SELECT ID_CUOTA "
		SQL_SEL = SQL_SEL  & " FROM VIEW_CASOS_GESTION_APOYO "
		SQL_SEL = SQL_SEL  & " WHERE ID_GESTION=" & valor_CH_ID_GESTION(indice)
		SET RS_SEL = CONN.execute(SQL_SEL)
		if not RS_SEL.EOF then
			do while not RS_SEL.EOF
			sql_update ="UPDATE GESTIONES_CUOTA "
			sql_update = sql_update & " SET FECHA_CONSULTA = '"&trim(fecha)&"', USUARIO_CONSULTA='"&trim(session("session_idusuario"))&"', OBSERVACION_CONSULTA='"&trim(strObservacionConsulta)&"', FECHA_AGENDAMIENTO='"&trim(dtmFechaAgendamiento)&" "&trim(strHoraAgendamiento)&"'"
			sql_update = sql_update & " WHERE Id_Cuota = '"&RS_SEL("ID_CUOTA")&"'"
			sql_update = sql_update & " AND Id_Gestion = '"&valor_CH_ID_GESTION(indice)&"'"
			conn.execute(sql_update)
			RS_SEL.movenext
			loop

		END IF
	
	next 

elseif trim(accion_ajax)="refresca_busca_cuotas" then
	IntIdGestion 	= request.querystring("intIdGestion")
	strRutDeudor 	= request.querystring("RUT_DEUDOR")
	strCodCliente 	= session("ses_codcli")
	concat_id_cuota = ""

	SQL_SEL ="SELECT ID_CUOTA "
	SQL_SEL = SQL_SEL  & " FROM VIEW_CASOS_GESTION_APOYO "
	SQL_SEL = SQL_SEL  & " WHERE ID_GESTION=" & IntIdGestion
	SET RS_SEL = CONN.execute(SQL_SEL)
	IF NOT RS_SEL.EOF THEN
		do while not rs_sel.eof 
			concat_id_cuota = concat_id_cuota & "," & RS_SEL("ID_CUOTA")		

		rs_sel.movenext		
		loop

		strIDCuotas 	= mid(concat_id_cuota,2,len(concat_id_cuota))

		%>
			<input type="hidden" name="strIDCuotas" id="strIDCuotas" value="<%=strIDCuotas%>">
		<%

	END IF

end if

cerrarscg()

%>


