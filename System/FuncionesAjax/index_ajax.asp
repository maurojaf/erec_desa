<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../../lib/asp/comunes/general/rutinasBooleano.inc"-->
<%

Response.CodePage = 65001
Response.charset  ="utf-8"

accion_ajax =request.querystring("accion_ajax")
AbrirSCG()

if trim(accion_ajax)="crear_cookie" then

	usuario_nombre 	=request.querystring("usuario_nombre")
	contrasena 		=request.querystring("contrasena")

	response.cookies("usuario_nombre")=usuario_nombre
	response.cookies("contrasena")=contrasena

	response.cookies("usuario_nombre").Expires=Date+20
	response.cookies("contrasena").Expires=Date+20

elseif trim(accion_ajax)="eliminar_cookie" then
	
	Response.Cookies("contrasena").Expires= Date() -1
	Response.Cookies("usuario_nombre").Expires= Date() -1


elseif trim(accion_ajax)="verifica_login" then

	login 	=request.querystring("usuario_nombre")
	clave 	=request.querystring("contrasena")

	sql_sel="SELECT ID_USUARIO, rut_usuario, nombres_usuario, apellido_paterno, " 
	sql_sel= sql_sel & " apellido_materno, fecha_nacimiento, correo_electronico, telefono_contacto, "
	sql_sel= sql_sel & " perfil, LOGIN, DECRYPTBYPASSPHRASE('C1traseña',CLAVE) CLAVE, PERFIL_ADM, perfil_cob, ACTIVO, perfil_proc, perfil_sup, "
	sql_sel= sql_sel & " PERFIL_CAJA, perfil_emp, PERFIL_FULL, perfil_back, gestionador_preventivo, "
	sql_sel= sql_sel & " anexo, observaciones_usuario, COD_AREA, CONVERT(VARCHAR(8),GETDATE(),108) AS HH , CONVERT(VARCHAR(10),GETDATE(),103) AS FH, TIPO_SOFTPHONE, substring(nombres_usuario,1,1)+substring(apellido_paterno,1,1)+SUBSTRING(rut_usuario, (len(substring(rut_usuario, 1, LEN(rut_usuario)-2))-3),4) CLAVE_AUTOMATICA "
	sql_sel= sql_sel & " FROM USUARIO "
	sql_sel= sql_sel & " WHERE LOGIN='" & trim(login) & "' AND DECRYPTBYPASSPHRASE('C1traseña',CLAVE) = '" & trim(clave) & "' "

	set rsUSU=Conn.execute(sql_sel)
	if not rsUSU.eof then
		ACTIVO								=TraeSiNo(Trim(rsUSU("ACTIVO")))
		'Response.write ACTIVO
		if trim(ACTIVO)="Si" then

			session("COLTABBG") 				= TraeCampoId(Conn, "COLOR_TABLA_BG", 1, "PARAMETRO_SISTEMA", "ID")
			session("COLTABBG2") 				= TraeCampoId(Conn, "COLOR_TABLA_BG_2", 1, "PARAMETRO_SISTEMA", "ID")

			SERVIDOR= MID(request.servervariables("PATH_INFO"),2, (Instr(MID(request.servervariables("PATH_INFO"),2, LEN(request.servervariables("PATH_INFO"))),"/"))-1)
			
			if ucase(SERVIDOR)="EREC" then
				session("ses_ruta_sitio") 	= "D:\app\EREC"
				session("ses_ruta_web") 	= "http://sistemas.llacruz.cl/erec"
				session("ses_ruta_sitio_Fisica") 	=  "\\sistemas.llacruz.cl\erec" 
			elseif ucase(SERVIDOR)="EREC_DEMO" then
				session("ses_ruta_sitio") 	= "D:\app\EREC_DEMO"
				session("ses_ruta_web") 	= "http://sistemas.llacruz.cl/erec_demo"
				session("ses_ruta_sitio_Fisica") 	=  "\\sistemas.llacruz.cl\EREC_DEMO" 
			elseif ucase(SERVIDOR)="EREC_DESA" then
				session("ses_ruta_sitio") 	= "D:\app\EREC_DESA"
				session("ses_ruta_web") 	= "http://sistemas.llacruz.cl/erec_desa"
				session("ses_ruta_sitio_Fisica") 	=  "\\sistemas.llacruz.cl\EREC_DESA" 
			end if
			
			strSql="SELECT VALOR FROM UNIDAD_FOMENTO WHERE CONVERT(VARCHAR(10),FECHA,103)=CONVERT(VARCHAR(10),GETDATE(),103)"
			set rsUF=Conn.execute(strSql)
			if not rsUF.eof then
				session("valor_uf") 	= rsUF("VALOR")
				session("valor_moneda") = rsUF("VALOR")
			Else
				session("valor_uf") = 22000
				session("valor_moneda") = 22000
			End If


			strSql="SELECT IsNull(PERMITE_NO_VALIDAR_FONOS,'N') as PERMITE_NO_VALIDAR_FONOS FROM PARAMETROS"
			set rsParam=Conn.execute(strSql)
			if not rsParam.eof then
				session("permite_no_validar_fonos") = rsParam("PERMITE_NO_VALIDAR_FONOS")
			Else
				session("permite_no_validar_fonos") = "S"
			End If
	'Variables de session desde BBDD
			session("ses_clave") 		 =clave
			session("ses_codcli") 		 =request("CB_CLIENTE")
			session("session_idusuario") =rsUSU("ID_USUARIO")
			session("session_user") 	 =rsUSU("RUT_USUARIO")
			session("session_login") 	 =rsUSU("LOGIN")
			session("session_tipo") 	 =rsUSU("PERFIL")
			session("perfil_adm") 		 =rsUSU("PERFIL_ADM")
			session("perfil_caja") 		 =rsUSU("PERFIL_CAJA")
			session("perfil_emp") 		 =rsUSU("PERFIL_EMP")
			session("perfil_sup") 		 =rsUSU("PERFIL_SUP")
			session("perfil_full") 		 =rsUSU("PERFIL_FULL")
			session("perfil_cob") 		 =rsUSU("PERFIL_COB")
			session("nombre_user") 		 =TRIM(rsUSU("NOMBRES_USUARIO")) & " " & TRIM(rsUSU("APELLIDO_PATERNO"))& " " & TRIM(rsUSU("APELLIDO_MATERNO"))
			session("iniciosesion") 	 =rsUSU("FH") & " - " & rsUSU("HH")
			session("tipo_softphone")	 =rsUSU("TIPO_SOFTPHONE")

			strSql = "SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario")
			set rsClientes=Conn.execute(strSql)
			strClientes=""
			Do While Not rsClientes.eof

				If strValCli="" Then
					session("ses_codcli") = rsClientes("COD_CLIENTE")
					strValCli="1"
				End If

				strClientes = strClientes & rsClientes("COD_CLIENTE") & ","
				rsClientes.movenext
			Loop

			strClientes = Mid(strClientes,1,len(strClientes)-1)
			session("strCliUsuarios") = strClientes

			strSql="SELECT NOMBRE_CONV_PAGARE,COD_MONEDA FROM CLIENTE WHERE COD_CLIENTE = '" & request("CB_CLIENTE") & "'"
			set rsCliente=Conn.execute(strSql)
			if not rsCliente.eof then
				'Response.write "NOMBRE_CONV_PAGARE=" & strSql
				session("NOMBRE_CONV_PAGARE") = rsCliente("NOMBRE_CONV_PAGARE")
				session("COD_MONEDA") = rsCliente("COD_MONEDA")

			Else
				session("NOMBRE_CONV_PAGARE") = "Convenio"
				session("COD_MONEDA") = 1

			End If

			strSql = "INSERT INTO LOG_CRMCOBROS (ID_USUARIO, LOGIN, FECHA, IP, IP_HOST, IP_LOCAL, IP_CLIENTE, TIPO_INGRESO, BLOQUEO)"
			strSql = strSql & " Values ("&trim(session("session_idusuario"))&",'"&trim(LOGIN)&"',getdate(),'" & Mid(request.servervariables("REMOTE_ADDR"),1,19) & "','" & Mid(request.servervariables("REMOTE_HOST"),1,19) & "','" & Mid(request.servervariables("LOCAL_ADDR"),1,19) & "','" & Mid(request.servervariables("HTTP_CLIENT_IP"),1,19) & "','1','0')"
			Conn.execute(strSql)




			sql_up ="update LOG_CRMCOBROS   "
			sql_up = sql_up & " set BLOQUEO=0 "
			sql_up = sql_up & " WHERE TIPO_INGRESO = 0 AND IP = '"& Mid(request.servervariables("REMOTE_ADDR"),1,19) &"' AND DATEDIFF(MI,FECHA,GETDATE()) < 5"
			Conn.execute(sql_up)

			'response.write sql_sel
		%>
			<input type="hidden" name="usuario_validado" id="usuario_validado" value="S">
		<%	

			'response.write trim(clave)&"<br>"&trim(rsUSU("CLAVE_AUTOMATICA"))
			if ucase(trim(clave))=ucase(trim(rsUSU("CLAVE_AUTOMATICA"))) then
			%>
				<input type="hidden" name="primer_ingreso" id="primer_ingreso" value="S">
			<%
			else
			%>
				<input type="hidden" name="primer_ingreso" id="primer_ingreso" value="N">
			<%
			end if



		else

			'response.write "no activo"

			strSql = "INSERT INTO LOG_CRMCOBROS (ID_USUARIO, LOGIN, FECHA, IP, IP_HOST, IP_LOCAL, IP_CLIENTE, TIPO_INGRESO, BLOQUEO)"
			strSql = strSql & " Values (NULL,'"&trim(login)&"',getdate(),'" & Mid(request.servervariables("REMOTE_ADDR"),1,19) & "','" & Mid(request.servervariables("REMOTE_HOST"),1,19) & "','" & Mid(request.servervariables("LOCAL_ADDR"),1,19) & "','" & Mid(request.servervariables("HTTP_CLIENT_IP"),1,19) & "','0','1')"
			Conn.execute(strSql)	

			%>
				<input type="hidden" name="usuario_validado" id="usuario_validado" value="N">
				<input type="hidden" name="mensaje_error" id="mensaje_error" value="Usuario no activo, contactese con su administrador">
			<%
		end if
		

	else

		'response.write "no existe"
			strSql = "INSERT INTO LOG_CRMCOBROS (ID_USUARIO, LOGIN, FECHA, IP, IP_HOST, IP_LOCAL, IP_CLIENTE, TIPO_INGRESO, BLOQUEO)"
			strSql = strSql & " Values (NULL,'"&trim(login)&"',getdate(),'" & Mid(request.servervariables("REMOTE_ADDR"),1,19) & "','" & Mid(request.servervariables("REMOTE_HOST"),1,19) & "','" & Mid(request.servervariables("LOCAL_ADDR"),1,19) & "','" & Mid(request.servervariables("HTTP_CLIENT_IP"),1,19) & "','0','1')"

		Conn.execute(strSql)

			sql_sel ="SELECT COUNT(*) contador "
			sql_sel = sql_sel &	" FROM LOG_CRMCOBROS "
			sql_sel = sql_sel &	" WHERE TIPO_INGRESO = 0 "
			sql_sel = sql_sel &	" AND IP = '"&Mid(request.servervariables("REMOTE_ADDR"),1,19)&"' "
			sql_sel = sql_sel &	" AND DATEDIFF(MI,FECHA,GETDATE()) < 5 and BLOQUEO=1  "
			set rs_sel = conn.execute(sql_sel)
			if err then 
				response.write sql_sel &" / error : "& err.description
			end if		

			if not rs_sel.eof then
				contador_fallidos =rs_sel("contador")
			else
				contador_fallidos =0
			end if

			intentos =3-cint(contador_fallidos)
			
				mensaje ="Quedan "& intentos &" intentos de logeo validos"

	%>
		<input type="hidden" name="usuario_validado" id="usuario_validado" value="N">
		<input type="hidden" name="mensaje_error" id="mensaje_error" value="Usuario y/o contraseña no valida <%=trim(mensaje)%>">
	<%
	end if

elseif trim(accion_ajax)="envia_correo_olvidado" then
	login_usuario 	=request.querystring("login_usuario")

	sql_sel ="exec proc_envia_contrassena_olvidada '"&trim(login_usuario)&"','ENVIA' "
	set rs_sel = Conn.execute(sql_sel)
	if err then
		response.write sql_sel &" / error : "& err.description
		response.end()
	end if

	if not rs_sel.eof then
		response.write rs_sel("mensaje")
	end if
	%>
		<input type="hidden" id="mensaje" name="mensaje" value="<%=trim(rs_sel("mensaje"))%>">
		<input type="hidden" id="accion" name="accion" value="<%=trim(rs_sel("accion"))%>">
	<%




elseif trim(accion_ajax)="verifica_usuario" then
	login_usuario 	=request.querystring("login_usuario")

	sql_sel ="exec proc_envia_contrassena_olvidada '"&trim(login_usuario)&"','VERIFICA' "
	set rs_sel = Conn.execute(sql_sel)
	if err then
		response.write sql_sel &" / error : "& err.description
		response.end
	end if

	if not rs_sel.eof then
		Response.WRITE rs_sel("mensaje")
	end if

	%>
		<input type="hidden" id="mensaje" name="mensaje" value="<%=trim(rs_sel("mensaje"))%>">
		<input type="hidden" id="accion" name="accion" value="<%=trim(rs_sel("accion"))%>">
	<%


elseif trim(accion_ajax)="verifica_log" then
	IP		=Mid(request.servervariables("REMOTE_ADDR"),1,19)

	sql_sel ="SELECT COUNT(*) contador "
	sql_sel = sql_sel &	" FROM LOG_CRMCOBROS "
	sql_sel = sql_sel &	" WHERE TIPO_INGRESO = 0 "
	sql_sel = sql_sel &	" AND IP = '"&trim(IP)&"' "
	sql_sel = sql_sel &	" AND DATEDIFF(MI,FECHA,GETDATE()) < 5 and BLOQUEO=1  "
	set rs_sel = conn.execute(sql_sel)
	if err then 
		response.write sql_sel &" / error : "& err.description
	end if		

	if not rs_sel.eof then
		contador_fallidos =rs_sel("contador")
	else
		contador_fallidos =0
	end if
	%>
		<input type="hidden" name="contador_fallidos"  id="contador_fallidos" value="<%=trim(contador_fallidos)%>">
	<%
end if
 

%>
