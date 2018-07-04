<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/asp/comunes/general/RutinasVarias.inc" -->
<!--#include file="../../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../../lib/asp/comunes/general/rutinasSCG.inc" -->

<style type="text/css">
<!--
.Estilo1 {font-size: 14px;font-weight: bold;}
.Estilo13 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; }
.Estilo2 {font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif;}
.Estilo5 {font-size: 11px; font-weight: bold;}
.Estilo8 {font-size: 11px}
.Estilo9 {font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; }
.Estilo12 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.Estilo14 {font-size: 10px}
.Estilo15 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
.Estilo26 {font-size: 12px}
.Estilo28 {font-family: "Courier New", Courier, monospace}
.Estilo32 {font-size: 13px}
.Estilo22 {font-size: 13px; font-family: "Courier New", Courier, monospace;}
.Estilo36 {font-family: Verdana, Arial, Helvetica, sans-serif;font-weight: bold;}
.Estilo37 {font-family: Verdana, Arial, Helvetica, sans-serif; }
.Estilo38 {font-size: 13px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.Estilo40 {	font-family: Verdana, Arial, Helvetica, sans-serif;	color: #FF0000;	font-size: 11px;font-weight: bold;}
.Estilo41 {color: #FFFFFF}
.Estilo33 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 14px}
.Estilo34 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px}

H1.SaltoDePagina {PAGE-BREAK-AFTER: always}
	.transpa {
	background-color: transparent;
	border: 1px solid #FFFFFF;
	text-align:center
	}
-->

</style>
<meta charset="utf-8">
<%

Response.CodePage = 65001
Response.charset="utf-8"


accion_ajax 		=request.querystring("accion_ajax")
abrirscg()

			

if trim(accion_ajax)="envio_plan_pago" then
	

	CB_TIPO 				=request.queryString("CB_TIPO")
	CB_FPAGO 				=request.queryString("CB_FPAGO")	
	rut 					=request.queryString("rut")
	strCodCliente 			=request.queryString("strCodCliente")
	cuotas_rut				=request.queryString("cuotas_rut")
	NOMBRE_DEUDOR 			=request.queryString("NOMBRE_DEUDOR")
	txt_observacion_email 	=request.queryString("txt_observacion_email")
	CORREO_SALIENTE 		=request.queryString("CORREO_SALIENTE")
	adj_pdf					=request.querystring("adj_pdf")	
	adj_excel				=request.querystring("adj_excel")
	concat_con_copia 		=request.querystring("concat_con_copia")			
	COD_CORREO 				=request.querystring("COD_CORREO")
	fecha_generar_documentos 	=request.querystring("fecha_generar_documentos")

	fecha_generar_documentos    =replace(fecha_generar_documentos,"/","-")

	SERVIDOR= MID(request.servervariables("PATH_INFO"),2, (Instr(MID(request.servervariables("PATH_INFO"),2, LEN(request.servervariables("PATH_INFO"))),"/"))-1)

	if ucase(SERVIDOR)="EREC" then
		email 			=request.querystring("email")

	elseif ucase(SERVIDOR)="EREC_DEMO" then
		email 			="soporte@llacruz.cl;currutia@llacruz.cl"

	elseif ucase(SERVIDOR)="EREC_DESA" then
		email 			="soporte@llacruz.cl;currutia@llacruz.cl"

	end if	


	sql_sel ="SELECT correo_electronico, nombres_usuario, apellido_paterno, anexo "
	sql_sel = sql_sel & " FROM USUARIO "
	sql_sel = sql_sel & " WHERE ID_USUARIO=" & Trim(session("session_idusuario")) 
	SET rs_email = Conn.execute(sql_sel)
	if not rs_email.eof then
		correo_electronico 	=rs_email("correo_electronico")
		nombres_usuario 	=rs_email("nombres_usuario")
		apellido_paterno    =rs_email("apellido_paterno")
		anexo 				=rs_email("anexo")
	else
		correo_electronico 	=""
		nombres_usuario 	=""
		apellido_paterno    =""
		anexo 				=""
	end if		

	if trim(anexo)<>"" then
		des_anexo ="al anexo " & anexo & " u "
	end if	

	strSql=" SELECT USA_SUBCLIENTE, USA_INTERESES, USA_HONORARIOS, USA_PROTESTOS, FORMULA_HONORARIOS, " 
	strSql= strSql & " FORMULA_INTERESES,TASA_MAX_CONV, DESCRIPCION, TIPO_INTERES "
	strSql= strSql & " FROM CLIENTE WHERE COD_CLIENTE ='" & strCodCliente & "'"


	set rsTasa=Conn.execute(strSql)
	if not rsTasa.eof then
		intTasaMax 			= rsTasa("TASA_MAX_CONV")
		strDescripcion 		= rsTasa("DESCRIPCION")
		strTipoInteres 		= rsTasa("TIPO_INTERES")
		strNomFormHon 		= ValNulo(rsTasa("FORMULA_HONORARIOS"),"C")
		strNomFormInt 		= ValNulo(rsTasa("FORMULA_INTERESES"),"C")
		strUsaSubCliente	= rsTasa("USA_SUBCLIENTE")
		strUsaInteres 		= rsTasa("USA_INTERESES")
		strUsaHonorarios 	= rsTasa("USA_HONORARIOS")
		strUsaProtestos 	= rsTasa("USA_PROTESTOS")

	Else
		intTasaMax = 1
		strDescripcion = ""
		strTipoInteres = ""
	end if




		strSql = "SELECT RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, convert(int,VALOR_CUOTA) VALOR_CUOTA , convert(int,isnull(dbo." & trim(strNomFormInt) & "(ID_CUOTA),0)) as INTERESES, convert(int,isnull(dbo." & trim(strNomFormHon) & "(ID_CUOTA),0)) as HONORARIOS, ID_CUOTA, NRO_DOC, NRO_CUOTA, NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO, convert(int, isnull(GASTOS_PROTESTOS,0)) GASTOS_PROTESTOS, CUENTA, convert(char, FECHA_VENC, 103) FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, CUSTODIO, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, convert(int,SALDO) SALDO, ( "
		strSql = strSql & " 		SELECT MIN(FECHA_VENC) "
		strSql = strSql & " 		FROM CUOTA, TIPO_DOCUMENTO  "
		strSql = strSql & " 		WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & strCodCliente & "' AND SALDO > 0  "
		strSql = strSql & " 		AND ESTADO_DEUDA IN (	SELECT ESTADO_DEUDA  "
		strSql = strSql & " 								FROM ESTADO_DEUDA "
		strSql = strSql & " 								WHERE ACTIVO = 1) "
		strSql = strSql & " 		AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO "
		strSql = strSql & " 		AND CUOTA.ID_CUOTA IN ("&cuotas_rut&") "
		strSql = strSql & " 		GROUP BY CUOTA.RUT_DEUDOR ) MAX_FECHA_VENC "

		strSql = strSql & " FROM CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & strCodCliente & "' AND SALDO > 0 AND ESTADO_DEUDA IN (SELECT ESTADO_DEUDA FROM ESTADO_DEUDA WHERE ACTIVO = 1) AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO AND CUOTA.ID_CUOTA IN ("&cuotas_rut&")"
		strSql = strSql & " ORDER BY CUOTA.FECHA_VENC ASC "

		set rsTemp= Conn.execute(strSql)
		'Response.write "strSql=" & strSql

		intTasaMensual = 2/100
		intTasaDiaria = intTasaMensual/30
		intCorrelativo = 0
		strArrID_CUOTA=""
		intTotSelSaldo= 0
		intTotSelIntereses= 0
		intTotSelProtestos= 0
		intTotSelHonorarios= 0
		Do while not rsTemp.eof
		
			intSaldo 			= rsTemp("SALDO")
			strNroDoc 			= rsTemp("NRO_DOC")
			strFechaVenc 		= rsTemp("FECHA_VENC")
			strTipoDoc 			= rsTemp("TIPO_DOCUMENTO")
			strNroCuota 		= rsTemp("NRO_CUOTA")
			intAntiguedad 		= ValNulo(rsTemp("ANTIGUEDAD"),"N")
			intIntereses 		= rsTemp("INTERESES")
			intHonorarios 		= rsTemp("HONORARIOS")
			intValorCapital 	= rsTemp("VALOR_CUOTA")
			intAbono 			= (intValorCapital) - (intSaldo)
			intProtestos 		= ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")
			strArrID_CUOTA 		= strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")
			MAX_FECHA_VENC 		= rsTemp("MAX_FECHA_VENC")
			intTotDoc 			= (intSaldo)+(intIntereses)+(intProtestos)+(intHonorarios)
			intTotSelSaldo 		= (intTotSelSaldo)+(intSaldo)
			intTotSelIntereses 	= (intTotSelIntereses)+(intIntereses)
			intTotSelProtestos 	= (intTotSelProtestos)+(intProtestos)
			intTotSelHonorarios = (intTotSelHonorarios)+(intHonorarios)
			intTotSelValorAbono = (intTotSelValorAbono)+(intAbono)
			intTotSelDoc 		= (intTotSelDoc)+(intTotDoc)
			intTotValorCapital  = (intTotValorCapital) + (intValorCapital)
			intCorrelativo = intCorrelativo + 1
		rsTemp.movenext
		loop
	

	SQL_SEL =" SELECT NOM_CORREO, ASUNTO_CORREO, CORREO_SALIENTE, CUERPO_CORREO, FIRMA, COD_CORREO "
	SQL_SEL = SQL_SEL & " FROM ENVIO_CORREO "
	SQL_SEL = SQL_SEL & " WHERE COD_CLIENTE =  " & strCodCliente & " AND COD_CORREO='"&trim(COD_CORREO)&"'"
	SQL_SEL = SQL_SEL & " ORDER BY ORDEN DESC "
	SET RS_SEL = Conn.execute(SQL_SEL)
	if not RS_SEL.eof then
		NOM_CORREO			=rs_sel("NOM_CORREO")
		ASUNTO_CORREO		=rs_sel("ASUNTO_CORREO")
		CORREO_SALIENTE		=rs_sel("CORREO_SALIENTE")
		CUERPO_CORREO		=rs_sel("CUERPO_CORREO")
		FIRMA				=rs_sel("FIRMA")
		COD_CORREO			=rs_sel("COD_CORREO")
	end if
'Response.write SQL_SEL&"<br>"
	
com = Chr(34)   
str_concat = "<div style="&com&"text-align: RIGHT;"&com&"><img src="&com&session("ses_ruta_web")&"/Imagenes/Logos/"&strCodCliente&"/Logo.jpg"&com&"></div>"
str_concat = str_concat &  "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"3"&com&"><b>Estimado(a)</b></font></div>"	

str_concat = str_concat &  "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"3"&com&"><b>Sr (a). "&trim(NOMBRE_DEUDOR)&"</b></font></div>"	

str_concat = str_concat &  "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"3"&com&"><b>RUT: "&trim(ucase(rut))&"</b></font></div><br>"	


str_concat = str_concat & "<TABLE ALIGN="&com&"CENTER"&com&" WIDTH="&com&"100%"&com&" BORDER="&com&"0"&com&" BORDERCOLOR="&com&"#000000"&com&" CELLSPACING="&com&"0"&com&" CELLPADDING="&com&"1"&com&">"
str_concat = str_concat & "	<tr>"
str_concat = str_concat & "		<td><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">"&trim(CUERPO_CORREO)&"</font></td>"
str_concat = str_concat & "	</tr>"
str_concat = str_concat & "</table>"
str_concat = str_concat & "<br>"

if trim(txt_observacion_email)<>"" then

	str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"3"&com&"><b>Observacion:</b></font></div>"

	str_concat = str_concat &"<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">"&trim(txt_observacion_email)&"</font></div><br>"

end if

str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"3"&com&"><b>El resumen de cobranza es:</b></font></div><br>"

str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">Total documentos:"&intCorrelativo&"</font></div>"
str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">Vencimiento mayor: "&MAX_FECHA_VENC&"</font></div>"
str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">Capital: "&formatNumber(intTotValorCapital,0)&"</font></div>"

If Trim(strUsaProtestos)="1" Then
	str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">Protesto: "&formatNumber(intTotSelProtestos,0)&"</font></div>"
end if

if trim(strUsaInteres)="1" then
	str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">Interés: "&formatNumber(intTotSelIntereses,0)&"</font></div>"
end if

If Trim(strUsaHonorarios)="1" Then
	str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">Gastos de cobranza: "&formatNumber(intTotSelHonorarios,0)&"</font></div>"
end if

str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">Abonado: "&formatNumber(intTotSelValorAbono,0)&"</font></div>"


str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"3"&com&"><b>---------------------------------------</b></font></div>"
str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"2"&com&">Total a pagar: "&formatNumber((intTotSelSaldo +(intTotSelIntereses)+(intTotSelHonorarios) + (intTotSelProtestos)), 0)&"</font></div><br><br>"

str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#000"&com&" size="&com&"2"&com&"><b>Si usted canceló algunas de los documentos en cobranza, favor enviar vía correo electrónico el (los) comprobantes de pago asociados a éstos y a la vez fecha de pago de los que aún están pendientes.</b></font></div><br><br>"

str_concat = str_concat & "<div style="&com&"text-align: left;"&com&"><font color="&com&"#585858"&com&"  face="&com&"Verdana"&com&" size="&com&"3"&com&"><b>Comuníquese con su ejecutivo (a) "&PonerPrimeraLetraEnMayuscula(nombres_usuario)&" "&PonerPrimeraLetraEnMayuscula(apellido_paterno)&" al teléfono fijo 28978900 ("&trim(des_anexo)&" opción cobranzas) o correo electrónico "&trim(lcase(correo_electronico))&"</b></font></div>"


str_concat = str_concat & "<br><br><br>"

str_concat = str_concat &"<TABLE ALIGN="&com&"CENTER"&com&" WIDTH="&com&"100%"&com&" BORDER="&com&"0"&com&" BORDERCOLOR="&com&"#000000"&com&" CELLSPACING="&com&"0"&com&" CELLPADDING="&com&"1"&com&">"
str_concat = str_concat & "	<tr>"
str_concat = str_concat & "		<td ALIGN="&com&"left"&com&"><img src="&session("ses_ruta_web")&FIRMA&"></td>"
str_concat = str_concat & "	</tr>"	
str_concat = str_concat & "</table>"
 

			cuotas_rut 			= replace(cuotas_rut,"'","")
			concat_con_copia 	= replace(concat_con_copia,"***",";")

			sql_sel ="Exec Proc_Envio_mail_deudor 'PLAN_PAGO','"&lcase(email)&"','"&trim(str_concat)&"','"&trim(session("session_idusuario"))&"' ,'"&trim(rut)&"','("&trim(cuotas_rut)&")','"&strCodCliente&"','"&trim(CORREO_SALIENTE)&"','"&trim(adj_pdf)&"','"&trim(adj_excel)&"','"&trim(concat_con_copia)&"','"&trim(ASUNTO_CORREO)& " Rut: "&trim(rut)&" Nombre: "&trim(ucase(NOMBRE_DEUDOR))&"','"&trim(fecha_generar_documentos)&"','"&TRIM(replace(session("ses_ruta_sitio_Fisica"),"/","\"))&"'"	

			conn.execute(sql_sel)
			
					
			Response.WRITE "<FONT STYLE='font-family:"&com&"VERDANA"&com&"; font-size:14px; background-color:#CEF6D8;'>ENVIADO CORRECTAMENTE, FAVOR REVISAR TU BANDEJA DE ENTRADA</FONT>"


elseif trim(accion_ajax)="crea_correo_electronico" then

	email 			=request.querystring("email")
	NOMBRE_DEUDOR 	=request.querystring("NOMBRE_DEUDOR")
	cuotas_rut 		=request.querystring("cuotas_rut")
	rut 			=request.querystring("rut")
	strCodCliente 	=request.queryString("strCodCliente")
	IntId 			=strCodCliente
	fecha_generar_documentos =request.querystring("fecha_generar_documentos")
	fecha_generar_documentos    =replace(fecha_generar_documentos,"/","-")
	
	

	DestinationPath = Server.mapPath("../../Archivo/BibliotecaNotificacionDeudores") & "\" & IntId  & "\" & rut

	' crear una instancia
	set Obj_FSO = createobject("scripting.filesystemobject")

	If not Obj_FSO.FolderExists(Server.mapPath("../../Archivo/BibliotecaNotificacionDeudores") & "\" & IntId) = True Then ' verifica la existencia del archivo
		Obj_FSO.CreateFolder(Server.mapPath("../../Archivo/BibliotecaNotificacionDeudores") & "\" & IntId) 

		If not Obj_FSO.FolderExists(Server.mapPath("../../Archivo/BibliotecaNotificacionDeudores") & "\" & IntId  & "\" & rut) = True Then 
			Obj_FSO.CreateFolder(Server.mapPath("../../Archivo/BibliotecaNotificacionDeudores") & "\" & IntId  & "\" & rut) 
		End if	
	else
		If not Obj_FSO.FolderExists(Server.mapPath("../../Archivo/BibliotecaNotificacionDeudores") & "\" & IntId  & "\" & rut) = True Then 
			Obj_FSO.CreateFolder(Server.mapPath("../../Archivo/BibliotecaNotificacionDeudores") & "\" & IntId  & "\" & rut) 
		End if	
	End if

	AbrirScg1()	

	sql_sel ="SELECT correo_electronico, ID_USUARIO "
	sql_sel = sql_sel & " FROM USUARIO "
	sql_sel = sql_sel & " WHERE ID_USUARIO in (372, 268) order by ID_USUARIO desc "
	set rs_email_sup = conn1.execute(sql_sel)

	sql_sel_email	="SELECT ID_EMAIL, lower(Email) correo_electronico, NOMBRE_DEUDOR "
	sql_sel_email = sql_sel_email & " FROM DEUDOR_EMAIL de "
	sql_sel_email = sql_sel_email & " INNER JOIN DEUDOR d ON d.RUT_DEUDOR=de.RUT_DEUDOR AND D.COD_CLIENTE = '" & strCodCliente & "'"
	sql_sel_email = sql_sel_email & " where de.RUT_DEUDOR ='"&rut&"' and lower(email) != '"&trim(lcase(email))&"' AND de.ESTADO IN (1,0)"
	SET rs_email_deudor = conn1.execute(sql_sel_email)
	'response.write sql_sel_email

	sql_sel ="SELECT correo_electronico, nombres_usuario, apellido_paterno "
	sql_sel = sql_sel & " FROM USUARIO "
	sql_sel = sql_sel & " WHERE ID_USUARIO=" & Trim(session("session_idusuario")) 
	SET rs_email = conn1.execute(sql_sel)
	if not rs_email.eof then
		correo_electronico 	=rs_email("correo_electronico")
		nombres_usuario 	=rs_email("nombres_usuario")
		apellido_paterno    =rs_email("apellido_paterno")
		

		if trim(correo_electronico)<>""	and not isnull(correo_electronico) then


			SQL_SEL =" SELECT NOM_CORREO, ASUNTO_CORREO, CORREO_SALIENTE, CUERPO_CORREO, FIRMA, COD_CORREO "
			SQL_SEL = SQL_SEL & " FROM ENVIO_CORREO "
			SQL_SEL = SQL_SEL & " WHERE COD_CLIENTE =  " & strCodCliente
			SQL_SEL = SQL_SEL & " ORDER BY ORDEN ASC "
			SET RS_SEL = conn1.execute(SQL_SEL)
			if not RS_SEL.eof then
			
			'Response.write "strCodCliente=" & strCodCliente

			%>
				
				<div>
					<div class="bt_enviar_correo" onclick="bt_tipo_correo('<%=trim(lcase(email))%>')">
						<span class="texto_enviar">Enviar</span>
					</div>
					<div class="contenido_correo">
						<table class="tabla_contenido">
							<tr>
								<td class="titulo_envio" valign="top" >De:</td>
								<td class="contenido_envio" valign="top" id="email_de"></td>
								
							</tr>
							<tr>
								<td  class="titulo_envio" valign="top">Para:</td>
								<td  class="contenido_envio" valign="top" id="email_para"><%=trim(lcase(email))%></td>
							</tr>
							<tr>
								<td class="titulo_envio" valign="top">CC:</td>
								<td class="contenido_envio" valign="top">
									<select name="con_copia" id="con_copia" multiple>
										<option value="cobranza@@llacruz.cl">cobranza@llacruz.cl</option>


										<%if not rs_email_deudor.eof then
											do while not rs_email_deudor.eof
											%>
												<option value="<%=trim(rs_email_deudor("correo_electronico"))%>"><%=trim(lcase(rs_email_deudor("correo_electronico")))%></option>
											<%
											rs_email_deudor.movenext
											loop

										end if%>


									</select>

								</td>
							</tr>
							<tr>

								<td class="titulo_envio" valign="top">Asunto:</td>
								<td class="contenido_envio" id="email_asunto"></td>
							</tr>
							<tr>
								<td class="titulo_envio" valign="top">Adjunto:</td>
								<td class="contenido_envio" valign="middle" >
									<input type="checkbox" name="adj_pdf" checked id="adj_pdf" value="1"><a href="#" onclick="bt_descargar('../Archivo/BibliotecaNotificacionDeudores/<%=strCodCliente%>/<%=rut%>/<%=fecha_generar_documentos%>_detalle_deuda.pdf')">&nbsp;<img border="0" src="../Imagenes/icono_pdf.jpg" width="25" height="25" alt="" style="cursor:pointer;"> Detalle_deuda.pdf</a>
									&nbsp;&nbsp;
									<input type="checkbox" checked name="adj_excel" id="adj_excel" value="1"><a href="../Archivo/BibliotecaNotificacionDeudores/<%=strCodCliente%>/<%=rut%>/<%=fecha_generar_documentos%>_detalle_deuda.csv">&nbsp;<img src="../Imagenes/icono_excel.jpg" border="0" width="25" height="25" alt=""> Detalle_deuda.csv</a>									
								</td>
							</tr>																						
							<tr>
								<td class="titulo_envio" valign="top">Observacion:</td>
								<td class="contenido_envio" valign="top"><textarea class="textarea_email" name="txt_observacion_email" id="txt_observacion_email"></textarea>
	</td>
							</tr>													
						</table>
					</div>					
				</div>

						
			<%
				do while not RS_SEL.eof
				%>
					<div class="opcion_envio_correo">
						<div><img src="../Imagenes/bt_email.png" onclick="bt_visualiza_correo('<%=RS_SEL("COD_CORREO")%>','<%=rut%>','<%=NOMBRE_DEUDOR%>','<%=trim(nombres_usuario)%>','<%=trim(apellido_paterno)%>','<%=trim(correo_electronico)%>','<%=trim(cuotas_rut)%>','')" alt="<%=trim(RS_SEL("NOM_CORREO"))%>"></div>
						<div><%=RS_SEL("NOM_CORREO")%></div>
					</div>
				<%

				RS_SEL.movenext
				loop
			else
				response.write "<font style='color:#B40404;font-family=Verdana'>No posee habilitadas cuentas de correo, favor comuníquese con el administrador.</font>"
			end if

%>
			<div id="envio_correo_plan_pago"></div>

<%
		else

			response.write "<font style='color:#B40404;font-family=Verdana'>No posee ingresada su cuenta de correo, favor comuníquese con el administrador.</font>"	

		end if

	else
		response.write "<font style='color:#B40404;font-family=Verdana'>Usuario no existe en sistema, favor comuníquese con el administrador.</font>"	
	end if

	CerrarScg1()


%>
	<div id="visualiza_correo"></div>
<%
elseif trim(accion_ajax)="visualiza_correo" then
	COD_CORREO 				=request.querystring("COD_CORREO")
	rut						=request.querystring("rut")
	strCodCliente 			=request.queryString("strCodCliente")
	NOMBRE_DEUDOR			=request.querystring("NOMBRE_DEUDOR")
	txt_observacion_email 	=request.querystring("txt_observacion_email")
	nombres_usuario			=request.querystring("nombres_usuario")
	apellido_paterno		=request.querystring("apellido_paterno")
	correo_electronico 		=request.querystring("correo_electronico")
	cuotas_rut 				=request.querystring("cuotas_rut")

	strSql=" SELECT USA_SUBCLIENTE, USA_INTERESES, USA_HONORARIOS, USA_PROTESTOS, FORMULA_HONORARIOS, " 
	strSql= strSql & " FORMULA_INTERESES,TASA_MAX_CONV, DESCRIPCION, TIPO_INTERES "
	strSql= strSql & " FROM CLIENTE WHERE COD_CLIENTE ='" & strCodCliente & "'"


	set rsTasa=Conn.execute(strSql)
	if not rsTasa.eof then
		intTasaMax 			= rsTasa("TASA_MAX_CONV")
		strDescripcion 		= rsTasa("DESCRIPCION")
		strTipoInteres 		= rsTasa("TIPO_INTERES")
		strNomFormHon 		= ValNulo(rsTasa("FORMULA_HONORARIOS"),"C")
		strNomFormInt 		= ValNulo(rsTasa("FORMULA_INTERESES"),"C")
		strUsaSubCliente	= rsTasa("USA_SUBCLIENTE")
		strUsaInteres 		= rsTasa("USA_INTERESES")
		strUsaHonorarios 	= rsTasa("USA_HONORARIOS")
		strUsaProtestos 	= rsTasa("USA_PROTESTOS")

	Else
		intTasaMax = 1
		strDescripcion = ""
		strTipoInteres = ""
	end if

	sql_sel ="SELECT correo_electronico, nombres_usuario, apellido_paterno, anexo "
	sql_sel = sql_sel & " FROM USUARIO "
	sql_sel = sql_sel & " WHERE ID_USUARIO=" & Trim(session("session_idusuario")) 
	SET rs_email = Conn.execute(sql_sel)
	if not rs_email.eof then
		correo_electronico 	=rs_email("correo_electronico")
		nombres_usuario 	=rs_email("nombres_usuario")
		apellido_paterno    =rs_email("apellido_paterno")
		anexo 				=rs_email("anexo")
	else
		correo_electronico 	=""
		nombres_usuario 	=""
		apellido_paterno    =""
		anexo 				=""
	end if		

	if trim(anexo)<>"" then
		des_anexo ="al anexo " & anexo & " u "
	end if


	SQL_SEL =" SELECT NOM_CORREO, ASUNTO_CORREO, CORREO_SALIENTE, CUERPO_CORREO, FIRMA, COD_CORREO "
	SQL_SEL = SQL_SEL & " FROM ENVIO_CORREO "
	SQL_SEL = SQL_SEL & " WHERE COD_CORREO =  " & TRIM(COD_CORREO)
	SQL_SEL = SQL_SEL & " ORDER BY ORDEN ASC "
	SET RS_SEL = conn.execute(SQL_SEL)
	if not RS_SEL.eof then	
		NOM_CORREO 			=RS_SEL("NOM_CORREO")
		ASUNTO_CORREO 		=RS_SEL("ASUNTO_CORREO")
		CORREO_SALIENTE 	=RS_SEL("CORREO_SALIENTE")
		CUERPO_CORREO 		=RS_SEL("CUERPO_CORREO")
		FIRMA 				=RS_SEL("FIRMA")
		COD_CORREO 			=RS_SEL("COD_CORREO")


		strSql = "SELECT RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, convert(int,VALOR_CUOTA) VALOR_CUOTA , convert(int,isnull(dbo." & trim(strNomFormInt) & "(ID_CUOTA),0)) as INTERESES, convert(int,isnull(dbo." & trim(strNomFormHon) & "(ID_CUOTA),0)) as HONORARIOS, ID_CUOTA, NRO_DOC, NRO_CUOTA, NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO, convert(int, isnull(GASTOS_PROTESTOS,0)) GASTOS_PROTESTOS, CUENTA, convert(char, FECHA_VENC, 103) FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, CUSTODIO, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, convert(int,SALDO) SALDO, ( "
		strSql = strSql & " 		SELECT MIN(FECHA_VENC) "
		strSql = strSql & " 		FROM CUOTA, TIPO_DOCUMENTO  "
		strSql = strSql & " 		WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & strCodCliente & "' AND SALDO > 0  "
		strSql = strSql & " 		AND ESTADO_DEUDA IN (	SELECT ESTADO_DEUDA  "
		strSql = strSql & " 								FROM ESTADO_DEUDA "
		strSql = strSql & " 								WHERE ACTIVO = 1) "
		strSql = strSql & " 		AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO "
		strSql = strSql & " 		AND CUOTA.ID_CUOTA IN ("&cuotas_rut&") "
		strSql = strSql & " 		GROUP BY CUOTA.RUT_DEUDOR ) MAX_FECHA_VENC "

		strSql = strSql & " FROM CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & strCodCliente & "' AND SALDO > 0 AND ESTADO_DEUDA IN (SELECT ESTADO_DEUDA FROM ESTADO_DEUDA WHERE ACTIVO = 1) AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO AND CUOTA.ID_CUOTA IN ("&cuotas_rut&")"
		strSql = strSql & " ORDER BY CUOTA.FECHA_VENC ASC "
		set rsTemp= Conn.execute(strSql)
		'Response.write "strSql=" & strSql

		intTasaMensual = 2/100
		intTasaDiaria = intTasaMensual/30
		intCorrelativo = 0
		strArrID_CUOTA=""
		intTotSelSaldo= 0
		intTotSelIntereses= 0
		intTotSelProtestos= 0
		intTotSelHonorarios= 0
		Do while not rsTemp.eof
		
			intSaldo 			= rsTemp("SALDO")
			strNroDoc 			= rsTemp("NRO_DOC")
			strFechaVenc 		= rsTemp("FECHA_VENC")
			strTipoDoc 			= rsTemp("TIPO_DOCUMENTO")
			strNroCuota 		= rsTemp("NRO_CUOTA")
			intAntiguedad 		= ValNulo(rsTemp("ANTIGUEDAD"),"N")
			intIntereses 		= rsTemp("INTERESES")
			intHonorarios 		= rsTemp("HONORARIOS")
			intValorCapital 	= rsTemp("VALOR_CUOTA")
			intAbono 			= (intValorCapital) - (intSaldo)
			intProtestos 		= ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")
			strArrID_CUOTA 		= strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")
			MAX_FECHA_VENC 		= rsTemp("MAX_FECHA_VENC")
			intTotDoc 			= (intSaldo)+(intIntereses)+(intProtestos)+(intHonorarios)
			intTotSelSaldo 		= (intTotSelSaldo)+(intSaldo)
			intTotSelIntereses 	= (intTotSelIntereses)+(intIntereses)
			intTotSelProtestos 	= (intTotSelProtestos)+(intProtestos)
			intTotSelHonorarios = (intTotSelHonorarios)+(intHonorarios)
			intTotSelValorAbono = (intTotSelValorAbono)+(intAbono)
			intTotSelDoc 		= (intTotSelDoc)+(intTotDoc)
			intTotValorCapital  = (intTotValorCapital) + (intValorCapital)

			intCorrelativo = intCorrelativo + 1
		rsTemp.movenext		
		loop

		%>
		<input type="hidden" name="NOM_CORREO" 		id="NOM_CORREO" 		value="<%=trim(NOM_CORREO)%>">			
		<input type="hidden" name="ASUNTO_CORREO" 	id="ASUNTO_CORREO" 		value="<%=trim(ASUNTO_CORREO)%>">
		<input type="hidden" name="CORREO_SALIENTE" id="CORREO_SALIENTE" 	value="<%=trim(CORREO_SALIENTE)%>">
		<input type="hidden" name="CUERPO_CORREO" 	id="CUERPO_CORREO" 		value="<%=trim(CUERPO_CORREO)%>">
		<input type="hidden" name="FIRMA" 			id="FIRMA" 				value="<%=trim(FIRMA)%>">
		<input type="hidden" name="COD_CORREO" 		id="COD_CORREO" 		value="<%=trim(COD_CORREO)%>">
		<input type="hidden" name="NOMBRE_DEUDOR" 	id="NOMBRE_DEUDOR" 		value="<%=trim(NOMBRE_DEUDOR)%>">
		
		<br>
		<br>
		<br>
		<br>
		<br>
		<br>
		<div>
			<div class="imagens_email"><img src="../Imagenes/Logos/<%=strCodCliente%>/Logo.jpg" alt=""></div>
			<div class="encabezado_email">Estimado(a)</div>	
			<div class="encabezado_email">Sr (a). <%=trim(NOMBRE_DEUDOR)%></div>
			<div class="encabezado_email">RUT: <%=trim(ucase(rut))%></div>

			<div class="cuerpo_email"><%=trim(CUERPO_CORREO)%></div>
			<%if trim(txt_observacion_email)<>"" then%>
				<div class="encabezado_email" id="visual_txt_observacion_email_titulo">Observacion:</div>
				<div class="cuerpo_email" id="visual_txt_observacion_email"><%=trim(txt_observacion_email)%></div>
			<%else%>
				<div class="encabezado_email" id="visual_txt_observacion_email_titulo"></div>
				<div class="cuerpo_email" id="visual_txt_observacion_email"></div>						
			<%end if%>
			<div class="encabezado_email">El resumen de cobranza es:</div>
			<div class="resumen_email">Total documentos: <%=intCorrelativo%></div>
			<div class="resumen_email">Vencimiento mayor: <%=MAX_FECHA_VENC%></div>
			<div class="resumen_email">Capital: <%=formatNumber(intTotValorCapital,0)%></div>
			<%If Trim(strUsaProtestos)="1" Then%>
				<div class="resumen_email">Protesto: <%=formatNumber(intTotSelProtestos,0)%></div>
			<%end if%>
			<%if trim(strUsaInteres)="1" then%>
				<div class="resumen_email">Interés: <%=formatNumber(intTotSelIntereses,0)%></div>
			<%end if%>
			<%If Trim(strUsaHonorarios)="1" Then%>
				<div class="resumen_email">Gastos de cobranza: <%=formatNumber(intTotSelHonorarios,0)%></div>
			<%end if%>
			<div class="resumen_email">Abonado: <%=formatNumber(intTotSelValorAbono,0)%></div>
			<div class="comunicacion_email">---------------------------------------</div>
			<div class="resumen_email">Total a pagar <%=formatNumber((intTotSelSaldo +(intTotSelIntereses)+(intTotSelHonorarios) + (intTotSelProtestos)), 0)%></div>
			<div class="importante_email"><p>Si usted canceló algunas de los documentos en cobranza, favor enviar vía correo electrónico el (los) comprobantes de pago asociados a éstos y a la vez fecha de pago de los que aún están pendientes.</p></div>
			<div class="comunicacion_email"><p>Comuníquese con nosotros al teléfono fijo +56 2 28978989 o al correo electrónico cobranza@llacruz.cl</p></div>
			<div class="firma_email"><img src="<%=session("ses_ruta_web")&FIRMA%>" alt=""></div>
		
		</div>		

		<%	



	end if

elseif trim(accion_ajax)="actualiza_fecha_hora" then
%>
	<input type="hidden" name="fecha_generar_documentos" id="fecha_generar_documentos" value="<%=trim(replace(replace(now(),":","-")," ","_"))%>">
<%
end if
%>

 