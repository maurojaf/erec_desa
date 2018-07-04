<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../arch_utils.asp"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"

accion_ajax 		=request("accion_ajax")
IDusuarioCarga  	= session("session_idusuario")
abrirscg()

if session("session_idusuario") ="" then
	response.write("<b><center>SU Sesion Ha Expirado,VUELVA A INGRESAR</center></b>")
	response.end
end if 
if trim(accion_ajax)="filtra_usuario" then

	CB_COBRANZA =request.querystring("CB_COBRANZA")

	if trim(CB_COBRANZA)="INTERNA" then
		PERFIL_EMP =1
	end if
	if trim(CB_COBRANZA)="EXTERNA" then
		PERFIL_EMP =0
	end if



	ssql="EXEC PROC_USUARIOS_POR_CLIENTE '" & TRIM(session("ses_codcli")) & "'," & trim(PERFIL_EMP)
	'response.write(ssql)
	
	SET rs_sel = conn.execute(ssql)
	
	if err then
		Response.write "ERROR : " & err.description
		Response.end()
	end if
	
	%>
    <select name="id_usuario" id="id_usuario" style="width:100%;">
        <option value="">SELECIONA USUARIO</option>  
        <%DO WHILE NOT rs_sel.eof%>      
        	<option value="<%=trim(rs_sel("ID_USUARIO"))%>"><%=trim(rs_sel("NOMBRE_USUARIO"))%></option>    
        <%rs_sel.movenext
        loop%>
    </select> 
	<%


elseif trim(accion_ajax)="select_tabla_temporal" then

	CB_COBRANZA 		=request.querystring("CB_COBRANZA")
	ESTADO 				=""	
	ck_dv_rut 			=request.querystring("ck_dv_rut")
	ck_activa_id_email 	=request.querystring("ck_activa_id_email")
	tipo 				=request.querystring("tipo")
	   
	 'Response.write ck_dv_rut &"<br>" 	  
	if trim(CB_COBRANZA)="INTERNA" then
		PERFIL_EMP =1
	end if
	if trim(CB_COBRANZA)="EXTERNA" then
		PERFIL_EMP =0
	end if
	 
   ssql="EXEC PROC_GESTIONES_MASIVAS_VALIDA '" & TRIM(session("ses_codcli")) & "','"&TRIM(CB_COBRANZA)&"'," & trim(PERFIL_EMP)  & "," &  trim(session("session_idusuario")) 


	set rs_valida = conn.execute(ssql)
  	
	
    if not rs_valida.eof then
    %>
		<div id="Layer1" style="width:850px; height:285px; overflow: scroll;">
		<table class="table_masiva_select estilo_columnas" cellspacing="0" border="0" cellspacing="0">
            <thead>
                <tr>
                    <td class="class_contador_select">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    <td class="class_campos_select">
                        Rut deudor                        
                    </td>
                    <td class="class_campos_select">Fecha ingreso</td>
                    <td class="class_campos_select">Hora ingreso</td>
                    <td class="class_campos_select">Observación</td>
                    <td class="class_campos_select">ID usuario</td>

                    <%if trim(tipo) ="2" then%>

	                    <%if trim(ck_activa_id_email)="1" then%>
	                    	<td class="class_campos_select" id="id_medio">ID medio</td>
	                    <%else%>	
	                    	<td class="class_campos_select">E-mail asociado</td>
	                    <%end if%>

	                <%elseif trim(tipo) <>"0" then%>
	                	<td class="class_campos_select" id="id_medio">ID medio</td>
	                <%end if%>
					
                </tr>
            </thead>
            <tbody>			
    <%
	i =0
    	do while not rs_valida.eof
		
			bgcolor_rut 	="" 
			bgcolor_fecha	=""
			bgcolor_hora 	=""
			bgcolor_obser 	=""
			bgcolor_usuario	=""
			bgcolor_medio 	=""
			bgcolor_mail 	=""
			var_error 		="" 

			If ( i Mod 2 )= 1 Then
				bgcolor = "#F2F2F2"
			Else
				bgcolor = "#FFFFFF"
			End If
			i = i + 1  
' response.write(rs_valida("estado_rut"))
' response.end

			if trim(rs_valida("estado_rut"))="RUT INVALIDO" or trim(rs_valida("RUT_DEUDOR"))="" then
				bgcolor 			="#F8E0E0"
				bgcolor_rut 		="#F78181"
				ESTADO_RUT			="RUT SIN ASIGNACION VIGENTE"	
				
			else
				bgcolor 			=""
				bgcolor_rut 		=""
				ESTADO_RUT			=""								
				
			end if

			if trim(rs_valida("estado_fecha"))="FECHA INVALIDA" or trim(rs_valida("FECHA_INGRESO"))="" then
				bgcolor 			="#F8E0E0"
				bgcolor_fecha 		="#F78181"
				ESTADO_FECHA		="FECHA INVALIDA"	
								
			else
				bgcolor 			=""
				bgcolor_fecha 		=""
				ESTADO_FECHA		=""				
				
			end if

			if trim(rs_valida("estado_hora"))="HORA INVALIDA" or trim(rs_valida("HORA_INGRESO"))="" then
				bgcolor 			="#F8E0E0"
				bgcolor_hora 		="#F78181"
				ESTADO_HORA			="HORA INVALIDA"		
												
			else
				bgcolor 			=""
				bgcolor_hora 		=""
				ESTADO_HORA			=""	
				
			end if

			if trim(rs_valida("estado_usuario"))="USUARIO INVALIDO" or trim(rs_valida("ID_USUARIO"))="" then
				bgcolor 			="#F8E0E0"
				bgcolor_usuario		="#F78181"
				ESTADO_USUARIO 		="USUARIO INVALIDO"	
							
			else
				bgcolor 			=""
				bgcolor_usuario		=""
				ESTADO_USUARIO 		=""										
				
			end if

			if trim(tipo)<>"0" then

				if trim(rs_valida("estato_id_medio"))="ID MEDIO INVALIDO" or trim(rs_valida("ID_MEDIO"))="" then
					bgcolor 			="#F8E0E0"
					bgcolor_medio		="#F78181"
					ESTADO_MEDIO 		="ID MEDIO INVALID0"		
									
				else
					bgcolor 			=""
					bgcolor_medio		=""
					ESTADO_MEDIO 		=""						
					
				end if

			end if

			if trim(rs_valida("estado_observacion"))="LARGO SUPERADO" then
				bgcolor 			="#F8E0E0"
				bgcolor_obser		="#F78181"
				ESTADO_OBSERVACION  ="LARGO OBSERVACION SUPERADO"
						
			else
				bgcolor 			=""
				bgcolor_obser		=""
				ESTADO_OBSERVACION  =""	
				
			end if

			if trim(rs_valida("estado_email"))="EMAIL INVALIDO"  then
				bgcolor 			="#F8E0E0"
				bgcolor_mail		="#F78181"
				ESTADO_EMAIL		="EMAIL INVALIDO"
					 
			else
				bgcolor 			=""
				bgcolor_mail		=""
				ESTADO_EMAIL		=""
			
			end if	

			If IsEmpty(rs_valida("OBSERVACIONES")) Then
				bgcolor 			="#F8E0E0"
				bgcolor_obser		="#F78181"
				ESTADO_OBSERVACION	="OBSERVACION INVALIDA"	
				
			else
				bgcolor 			=""
				bgcolor_obser		=""
				ESTADO_OBSERVACION	=""	
				
			end if	
			'response.write(ck_activa_id_email)
%>
			
			<tr id="tr_<%=trim(rs_valida("ID"))%>" >
			    <td bgcolor="<%=bgcolor%>" class="class_registros" >
					<img src="../Imagenes/bt_eliminar.png" width="10" height="10" onclick="bt_elimina('<%=trim(rs_valida("ID"))%>')">
					<%=I%>
			    </td>
			    <td bgcolor="<%=bgcolor%>"  class="class_registros" id="accion_rut_<%=trim(rs_valida("ID"))%>" TITLE="<%=trim(ESTADO_RUT)%>" >
			    	<span style="background-color:<%=trim(bgcolor_RUT)%>" onclick="bt_edita('rut', '<%=trim(rs_valida("ID"))%>', '<%=trim(rs_valida("RUT_DEUDOR"))%>')"><%=trim(rs_valida("RUT_DEUDOR"))%>&nbsp;&nbsp;</span>
					<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreso_gestiones" value="<%=trim(ESTADO_RUT)%>">
			    </td>
			    <td bgcolor="<%=bgcolor%>" class="class_registros" id="accion_fecha_<%=trim(rs_valida("ID"))%>" TITLE="<%=trim(ESTADO_FECHA)%>" >

			    	<span style="background-color:<%=trim(bgcolor_fecha)%>" onclick="bt_edita('fecha', '<%=trim(rs_valida("ID"))%>', '<%=trim(rs_valida("FECHA_INGRESO"))%>')"><%=trim(rs_valida("FECHA_INGRESO"))%>&nbsp;&nbsp;</span>

			    	<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreso_gestiones" value="<%=trim(ESTADO_FECHA)%>">

			    </td>

			    <td bgcolor="<%=bgcolor%>" class="class_registros" id="accion_hora_<%=trim(rs_valida("ID"))%>" TITLE="<%=trim(ESTADO_HORA)%>">
			    	
			    	<span style="background-color:<%=trim(bgcolor_hora)%>"  onclick="bt_edita('hora', '<%=trim(rs_valida("ID"))%>', '<%=trim(rs_valida("HORA_INGRESO"))%>')"><%=trim(rs_valida("HORA_INGRESO"))%>&nbsp;&nbsp;</span>

			    	<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreso_gestiones" value="<%=trim(ESTADO_HORA)%>">

			    </td>
			    <td bgcolor="<%=bgcolor%>" class="class_registros" id="accion_observaciones_<%=trim(rs_valida("ID"))%>" TITLE="<%=trim(ESTADO_OBSERVACION)%>">
	
			    	<span style="background-color:<%=trim(bgcolor_obser)%>" onclick="bt_edita('observaciones', '<%=trim(rs_valida("ID"))%>', '<%=trim(rs_valida("OBSERVACIONES"))%>')"><%=trim(rs_valida("OBSERVACIONES"))%>&nbsp;&nbsp;</span>

			    	<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreos_gestiones" value="<%=trim(ESTADO_OBSERVACION)%>">

			    </td>
			    <td bgcolor="<%=bgcolor%>" class="class_registros" id="accion_usuario_<%=trim(rs_valida("ID"))%>" TITLE="<%=trim(ESTADO_USUARIO)%>">
					    
			    	<span style="background-color:<%=trim(bgcolor_usuario)%>" onclick="bt_edita('usuario', '<%=trim(rs_valida("ID"))%>', '<%=trim(rs_valida("ID_USUARIO"))%>')"><%=trim(rs_valida("ID_USUARIO"))%>&nbsp;&nbsp;</span>

			    	<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreso_gestiones" value="<%=trim(ESTADO_USUARIO)%>">


			    </td>


			    <%if trim(tipo) ="2" then%>

			    	<%if trim(ck_activa_id_email)="1" then%>

					    <td bgcolor="<%=bgcolor%>" class="class_registros" id="accion_medio_<%=trim(rs_valida("ID"))%>" TITLE="<%=trim(ESTADO_MEDIO)%>">
						    
					    	<span style="background-color:<%=trim(bgcolor_medio)%>" onclick="bt_edita('medio', '<%=trim(rs_valida("ID"))%>', '<%=trim(rs_valida("ID_MEDIO"))%>')"><%=trim(rs_valida("ID_MEDIO"))%>&nbsp;&nbsp;</span>
							<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreso_gestiones" value="<%=trim(ESTADO_MEDIO)%>">
					    </td>

					<%else%>

					    <td bgcolor="<%=bgcolor%>" class="class_registros" id="accion_mail_<%=trim(rs_valida("ID"))%>" TITLE="<%=trim(ESTADO_EMAIL)%>">
						    
					    	<span style="background-color:<%=trim(bgcolor_mail)%>" onclick="bt_edita('mail', '<%=trim(rs_valida("ID"))%>', '<%=trim(rs_valida("DESCRIPCION_EMAIL"))%>')"><%=trim(rs_valida("DESCRIPCION_EMAIL"))%>&nbsp;&nbsp;</span>

					    	<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreso_gestiones" value="<%=trim(ESTADO_EMAIL)%>">
					    </td>

					<%end if%>				   

			    <%elseif trim(tipo) <>"0" then%>

				    <td bgcolor="<%=bgcolor%>" class="class_registros" id="accion_medio_<%=trim(rs_valida("ID"))%>" TITLE="<%=trim(ESTADO_MEDIO)%>">
					    
				    	<span style="background-color:<%=trim(bgcolor_medio)%>" onclick="bt_edita('medio', '<%=trim(rs_valida("ID"))%>','<%=trim(rs_valida("ID_MEDIO"))%>')"><%=trim(rs_valida("ID_MEDIO"))%>&nbsp;&nbsp;</span>
				    	<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreso_gestiones" value="<%=trim(ESTADO_MEDIO)%>">
				    </td>				   

				<%end if%> 
			    
			</tr>         
			
<%	
		
		Response.flush()
    	rs_valida.movenext
		'i = i + 1
    	loop
 %>
 		</tbody>
		</table>
		</div>
	
		
		
 <%
    end if


elseif trim(accion_ajax)="elimina_tabla_temp"	 then

			sql_LimpiaTabla ="EXEC PROC_GESTIONES_MASIVAS_INGRESA 2,'" &  IDusuarioCarga & "'"
			'response.end
			conn.execute(sql_LimpiaTabla)
	         'conn.execute(sql_delete)

elseif trim(accion_ajax)="elimina_Registros_Errores" then
	
	error_largo_areglo =""
	
	CB_COBRANZA 		=request.form("CB_COBRANZA")
	ESTADO 				=""	
	ck_dv_rut 			=request.querystring("ck_dv_rut")
	ck_activa_id_email 	=request.querystring("ck_activa_id_email")
	tipo 				=request.querystring("tipo")
	   
	if trim(CB_COBRANZA)="INTERNA" then
		PERFIL_EMP =1
	end if
	if trim(CB_COBRANZA)="EXTERNA" then
		PERFIL_EMP =0
	end if
	
		 
	ssql="EXEC PROC_GESTIONES_MASIVAS_VALIDA '" & TRIM(session("ses_codcli")) & "','"&TRIM(CB_COBRANZA)&"'," & trim(PERFIL_EMP)  & "," &  trim(session("session_idusuario")) & ",1"
	'response.write("<br/>pase"& ssql )
	set rs_ValidaRegistros = conn.execute(ssql)
	
	if  NOT rs_ValidaRegistros.eof then 
	
		if  rs_ValidaRegistros("RegistrosValidos")  > 0 then 
			error_largo_areglo	 =""'' si tiene registros se envia en blanco el string
		else
			error_largo_areglo	 ="Sin Registros"
		end if 
	else 
		error_largo_areglo	 ="Sin Registros"
	end if 
	
%>				
<input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
						 
<%elseif trim(accion_ajax)="modifica_accion" then

	id 	 		=request.querystring("id")
	campo 		=request.querystring("campo")
	valor   	=request.querystring("valor")
	CB_COBRANZA =request.querystring("CB_COBRANZA")
	ck_dv_rut 	=request.querystring("ck_dv_rut")
%>

	<script type="text/javascript">
		$(document).ready(function(){

			//$('.fecha_validada_').datepicker( {changeMonth: true,changeYear: true})

			$('input[id*="input_hora_"]').TimepickerInputMask({
			    seconds: false
			}); 



		})
	</script>  
<% 
	if trim(campo)="fecha" then
%>
		<input type="text" class="fecha_validada_"  readonly="readonly" style="width:100px" name="input_fecha_<%=trim(id)%>" id="input_fecha_<%=trim(id)%>" value="<%=trim(valor)%>">
<%
	
	elseif trim(campo)="hora" then

%>
		<input type="text" class="input2" style="width:100px" name="input_hora_<%=trim(id)%>" id="input_hora_<%=trim(id)%>" value="<%=trim(valor)%>">

<%
	elseif trim(campo)="usuario" then

		CB_COBRANZA =request.querystring("CB_COBRANZA")

		if trim(CB_COBRANZA)="INTERNA" then
			PERFIL_EMP =1
		end if
		if trim(CB_COBRANZA)="EXTERNA" then
			PERFIL_EMP =0
		end if
		
	
		sql="EXEC PROC_USUARIOS_POR_CLIENTE '" & TRIM(session("ses_codcli")) & "'," & trim(PERFIL_EMP)
		SET rs_sel = conn.execute(sql)
		
		if err then
			Response.write "ERROR : " & err.description
			Response.end()
		end if
		%>
	    <select name="id_usuario_<%=id%>" style="width:100px"  id="id_usuario_<%=id%>" onchange="bt_actualiza_usuario('<%=trim(id)%>',this.value)">
	        <option value="">SELECIONA USUARIO</option>  
	        <%DO WHILE NOT rs_sel.eof%>      
	        	<option value="<%=trim(rs_sel("ID_USUARIO"))%>"><%=trim(rs_sel("NOMBRE_USUARIO"))%></option>    
	        <%rs_sel.movenext
	        loop%>
	    </select> 
		<%


	elseif trim(campo)	="rut" then

%>
		<input type="text" onblur="ValidaRut_con_digito(this, '<%=trim(id)%>')" style="width:100px" name="input_rut_<%=trim(id)%>" id="input_rut_<%=trim(id)%>" value="<%=trim(valor)%>">
<%
	else
%>
		<input type="text" style="width:100px" name="input_<%=trim(campo)%>_<%=trim(id)%>" id="input_<%=trim(campo)%>_<%=trim(id)%>" value="<%=trim(valor)%>">
<%
	end if

elseif trim(accion_ajax)="actualiza_campo" then

	id_campo		=request.querystring("id_campo")
	valor_input		=request.querystring("valor_input")
	campo 			=request.querystring("campo")
	CB_COBRANZA 	=request.querystring("CB_COBRANZA")
	tipo 			=request.querystring("tipo")
	campo_update 	=""
	com 			=Chr(34)
	invalido 		="N"
	style 			=""
	ESTADO 			=""

	'Response.write tipo &"s<br>"
	if trim(campo) ="rut" then
		campo_update ="RUT_DEUDOR"
	elseif trim(campo) ="fecha" then
		campo_update ="FECHA_INGRESO"
	elseif trim(campo) ="hora" then
		campo_update ="HORA_INGRESO"
	elseif trim(campo) ="observaciones" then
		campo_update ="OBSERVACIONES"
	elseif trim(campo) ="usuario" then
		campo_update ="ID_USUARIO"
	elseif trim(campo) ="medio" then
		campo_update ="ID_MEDIO"
	elseif trim(campo) ="mail" then
		campo_update ="DESCRIPCION_EMAIL"
	end if

	
	if trim(campo_update)<>"" THEN
		sql_update = "PROC_GESTIONES_MASIVAS_MODIFICA_DATOS " &id_campo & ",'" & IDusuarioCarga & "',1,'" & campo_update & "','" &  TRIM(valor_input)&"'"
		conn.execute(sql_update)
		
		if trim(campo) ="rut" then
			sql_valida_rut2 = "PROC_GESTIONES_MASIVAS_MODIFICA_DATOS " & ID_CAMPO & ",'" & IDUSUARIOCARGA & "',5,'" & CB_COBRANZA & "','" &  valor_input &"','" &  trim(session("ses_codcli")) &"' "
		set rs_valida_rut = conn.execute(sql_valida_rut2)

	 		if rs_valida_rut.eof then
	 			style 	="style="&com&"background-color:#F78181"&com
	 			ESTADO 	="RUT SIN ASIGNACION VIGENTE"
	 		end if
	 	end if


		if trim(campo) ="usuario" then
			CB_COBRANZA =request.querystring("CB_COBRANZA")

			if trim(CB_COBRANZA)="INTERNA" then
				PERFIL_EMP =1
			end if
			if trim(CB_COBRANZA)="EXTERNA" then
				PERFIL_EMP =0
			end if

			
			sql_valida_usuario="EXEC PROC_USUARIOS_POR_CLIENTE '" & TRIM(session("ses_codcli")) & "'," & trim(PERFIL_EMP) &",'" &  trim(valor_input) & "'"
			set rs_valida_usuario = conn.execute(sql_valida_usuario)			
			
			if rs_valida_usuario.eof then
				style 	="style="&com&"background-color:#F78181"&com
				ESTADO 	="USUARIO INVALIDO"
			end if
		end if

		
		if trim(campo) ="medio" then
				if trim(tipo)="1" then  '''' TELEFONO
						SQL_VALIDA_TEL = "PROC_GESTIONES_MASIVAS_MODIFICA_DATOS " &ID_CAMPO & ",'" & IDUSUARIOCARGA & "',2,'" & "" & "','" &  "" &"'"
						SET RS_VALIDA_USUARIO = CONN.EXECUTE(SQL_VALIDA_TEL)
						
						if rs_valida_usuario.eof then
							style 	="style="&com&"background-color:#F78181"&com
							ESTADO 	="ID MEDIO INVALID0"
						end if	
				elseif trim(tipo)="2" then  ''' EMAIL
				
						SQL_VALIDA_TEL = "PROC_GESTIONES_MASIVAS_MODIFICA_DATOS " & ID_CAMPO & ",'" & IDUSUARIOCARGA & "',3,'" & "" & "','" &  "" &"'"
						SET RS_VALIDA_USUARIO = CONN.EXECUTE(SQL_VALIDA_TEL)
						'response.write(SQL_VALIDA_TEL)
						'response.end
						
						if rs_valida_usuario.eof then
							style 	="style="&com&"background-color:#F78181"&com
							ESTADO 	="EMAIL INVALIDO"
						end if	
				elseif trim(tipo)="3" then '''' DIRECCION
						SQL_VALIDA_TEL = "PROC_GESTIONES_MASIVAS_MODIFICA_DATOS " & ID_CAMPO & ",'" & IDUSUARIOCARGA & "',4,'" & "" & "','" &  "" &"'"
						
						SET RS_VALIDA_USUARIO = CONN.EXECUTE(SQL_VALIDA_TEL)
						if rs_valida_usuario.eof then
							style 	="style="&com&"background-color:#F78181"&com
							ESTADO 	="ID MEDIO INVALID0"
						end if	
				else
						style 	="style="&com&"background-color:#F78181"&com
						ESTADO 	="ID MEDIO INVALID0"
				end if
		end if

		
		if trim(campo)="mail" then
			Cadena1=valor_input
			Cadena2="@"
			If InStr(Cadena1,Cadena2)>0 then
				style 	=""
				ESTADO 	=""
			Else
				style 	="style="&com&"background-color:#F78181"&com
				ESTADO 	="EMAIL INVALIDO"
			End if
		end if

		%>

		<%if trim(valor_input)="" then%>

			<span title="<%=trim(ESTADO)%>" <%=trim(style)%> onclick="bt_edita('<%=campo%>', '<%=trim(id_campo)%>', '<%=trim(valor_input)%>')"><%=trim(valor_input)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>

		<%else%>
		
			<span title="<%=trim(ESTADO)%>" <%=trim(style)%> onclick="bt_edita('<%=campo%>', '<%=trim(id_campo)%>', '<%=trim(valor_input)%>')"><%=trim(valor_input)%>&nbsp;&nbsp;</span>

		<%end if%>
		<input type="hidden" readonly="readonly" name="estado_ingreso_gestiones" id="estado_ingreso_gestiones" value="<%=trim(ESTADO)%>">
		
		
<%
end if
%>


		
<%	
elseif trim(accion_ajax)="elimina_registro" then
		id_campo 	=request.querystring("id_campo")

		sql_elimina ="EXEC PROC_GESTIONES_MASIVAS_INGRESA 3,'" &  IDusuarioCarga & "'," & id_campo
		'response.write(sql_elimina)
		'response.end
		conn.execute(sql_elimina)
	
elseif trim(accion_ajax)="guardar_valores" then

	var_rut_deudor		=request("var_rut_deudor")
	var_fecha_ingreso	=request("var_fecha_ingreso")	
	var_hora_ingreso	=request("var_hora_ingreso")	
	var_observacion		=request("var_observacion")
	var_id_usuario		=request("var_id_usuario")
	var_id_medio		=request("var_id_medio")
	var_email 			=request("var_email")
	ck_id_usuario		=request("ck_id_usuario")		
	ck_hora_ingreso		=request("ck_hora_ingreso")
	ck_observaciones	=request("ck_observaciones")
	ck_fecha_ingreso	=request("ck_fecha_ingreso")
	ck_activa_id_email 	=request("ck_activa_id_email")
	cod_tipo_gestion 	=request("cod_tipo_gestion")
	tipo 				=request("tipo")
	ck_dv_rut 			=request("ck_dv_rut")
	contador 			=0
	com 				=Chr(34)
	
	'Response.write cod_tipo_gestion


	ValorNuevo_rut 		=split(var_rut_deudor,CHR(10))
	ValorNuevo_fecha	=split(var_fecha_ingreso,CHR(10))
	ValorNuevo_hora		=split(var_hora_ingreso,CHR(10))
	ValorNuevo_obser	=split(var_observacion,CHR(10))
	ValorNuevo_usuario	=split(var_id_usuario,CHR(10))
	ValorNuevo_medio 	=split(var_id_medio,CHR(10))
	ValorNuevo_email 	=split(var_email,CHR(10))
	
	total_rut 			=ubound(ValorNuevo_rut)
	total_fecha			=ubound(ValorNuevo_fecha)
	total_hora 			=ubound(ValorNuevo_hora)
	total_obser			=ubound(ValorNuevo_obser)
	total_usuario		=ubound(ValorNuevo_usuario)
	total_medio			=ubound(ValorNuevo_medio)
	total_email			=ubound(ValorNuevo_email)
	error_largo_areglo	=""
	


	if cint(total_rut)>cint(total_fecha) and trim(ck_fecha_ingreso)="1"  then
		Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'>Total registros fecha no corresponden a rut</div>"
		error_largo_areglo ="fecha"
		%>
		 <input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
		<%		
		On Error Resume Next
		if err then
			Response.write err.description
			Response.end()
		end if

	end if

	if cint(total_rut)>cint(total_hora) and trim(ck_hora_ingreso)="1"  then
		Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'>Total registros hora no corresponden a rut</div>"
		error_largo_areglo ="hora"
		%>
		 <input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
		<%			
		On Error Resume Next
		if err then
			Response.write err.description
			Response.end()
		end if

	end if

	if cint(total_rut)>cint(total_obser) and trim(ck_observaciones)="1" then
		Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'>Total registros observacion no corresponden a rut</div>"
		error_largo_areglo ="Observación"
		%>
		 <input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
		<%			
		On Error Resume Next
		if err then
			Response.write err.description
			Response.end()
		end if
	end if

	if cint(total_rut)>cint(total_usuario) and trim(ck_id_usuario)="1" then
		Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'>Total registros usuario no corresponden a rut</div>"
		error_largo_areglo ="usuario"
		%>
		 <input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
		<%			
		On Error Resume Next
		if err then
			Response.write err.description
			Response.end()
		end if
	end if


	if cint(total_rut)>cint(total_email) and trim(ck_activa_id_email)<>"1" and trim(tipo)="2"  then
		Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'>Total registros descripcion mail no corresponden a rut</div>"
		error_largo_areglo ="email"
		%>
		 <input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
		<%		
		On Error Resume Next
		if err then
			Response.write err.description
			Response.end()
		end if

	end if

	if trim(tipo)<>"0" then

		if cint(total_rut)>cint(total_medio) and trim(ck_activa_id_email)="1"  and trim(tipo)="2" then

			Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'>Total registros medio mail no corresponden a rut</div>"
			error_largo_areglo ="medio"
			%>
			 <input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
			<%			
			On Error Resume Next
			if err then
				Response.write err.description
				Response.end()
			end if

		end if

		if cint(total_rut)>cint(total_medio) and trim(tipo)<>"2" then

			Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'>Total registros medio no corresponden a rut</div>"
			error_largo_areglo ="medio"
			%>
			 <input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
			<%			
			On Error Resume Next
			if err then
				Response.write err.description
				Response.end()
			end if

		end if

	end if

			''''''' LIMPIAMOS la tabla 
			sql_delete ="EXEC PROC_GESTIONES_MASIVAS_INGRESA 2,'" &  IDusuarioCarga & "'"
			conn.execute(sql_delete)
	
	For indice = 0 to total_rut 

		'response.write ValorNuevo_rut(indice)&"<br>."

		contador 		=cint(contador) + 1


		if trim(ck_id_usuario)="1" then
			valor_usuario =ValorNuevo_usuario(indice)
		else
			valor_usuario =var_id_usuario
		end if

		if trim(ck_hora_ingreso)="1" then
			valor_hora  =ValorNuevo_hora(indice)
		else
			valor_hora 	=var_hora_ingreso
		end if

		if trim(ck_observaciones)="1" then
			valor_obser =ValorNuevo_obser(indice)
		else
			valor_obser =var_observacion
		end if

		if trim(ck_fecha_ingreso)="1" then
			valor_fecha =ValorNuevo_fecha(indice)
		else
			valor_fecha =var_fecha_ingreso
		end if	


		if tipo<>"2"  then
			valor_medio =ValorNuevo_medio(indice) 

		else

			if trim(ck_activa_id_email)<>"1" then
				valor_email =ValorNuevo_email(indice)
			end if	

			if trim(ck_activa_id_email)="1" then
				valor_medio =ValorNuevo_medio(indice) 
			end if	

		end if	
		
		if isnull(valor_usuario) or trim(valor_usuario)="" then
			valor_usuario 	=null
		end if	

		if isnull(valor_hora) or trim(valor_hora)="" then
			valor_hora 		=null
		end if	

		if isnull(valor_obser) or trim(valor_obser)="" then
			valor_obser 	=null
		end if

		if isnull(valor_fecha) or trim(valor_fecha)="" then
			valor_fecha 	=null
		end if				

		if isnull(valor_medio) or trim(valor_medio)="" then
			valor_medio 	=null
		end if

		if trim(tipo)="0" then
			valor_medio=""
		end if

		
	if TRIM(ValorNuevo_rut(indice))<>CHR(10) and TRIM(ValorNuevo_rut(indice))<>CHR(13) and TRIM(ValorNuevo_rut(indice))<>"" then
				''' REVISA EL RUT
				if trim(ck_dv_rut)="1" then
				
						tur 		=""
						mult 		=0
						suma 		=0
						valor 		=0
						codigo_veri =""
						tur=strreverse(trim(ValorNuevo_rut(indice))) 
						mult = 2 
						for i = 1 to len(tur) 
							if mult > 7 then 
								mult = 2 
							end if 
							suma = mult * mid(tur,i,1) + suma 
							mult = mult +1 
						next 

						valor = 11 - (suma mod 11) 
						if valor = 11 then 
							codigo_veri = "0" 
						elseif valor = 10 then 
							codigo_veri = "k" 
						else 
							codigo_veri = trim(valor) 
						end if 
						rut =trim(ValorNuevo_rut(indice))&"-"&trim(codigo_veri)

					elseif trim(ck_dv_rut)="2" then
						rut = mid(TRIM(ValorNuevo_rut(indice)), 1 ,len(TRIM(ValorNuevo_rut(indice)))-1) &"-"& mid(TRIM(ValorNuevo_rut(indice)), len(TRIM(ValorNuevo_rut(indice))) , 1)
					else
						rut = TRIM(ValorNuevo_rut(indice))
				end if


			
			ID				= TRIM(contador)
			RutDeudor 		= TRIM(rut) 
			fechaIngreso	= TRIM(valor_fecha)
			HoraIngreso		= TRIM(valor_hora)
			Observacion		= TRIM(valor_obser)
			idUsuario		= TRIM(valor_usuario)
			idMedio			= TRIM(valor_medio)
			DescripcionMail = TRIM(valor_email)
			CodGestion		= TRIM(cod_tipo_gestion)
			CodCliente      = TRIM(session("ses_codcli"))
			IDusuarioCarga  = session("session_idusuario")
			
			if trim(ck_activa_id_email)="1"  then
					DescripcionMail = "."
			end if
			
			if trim(ck_activa_id_email)="1"  then
						DescripcionMail = "."
						if not isNumeric(idmedio) and trim(tipo) = "2"  then 
						Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'><center>Ingreso de Id Deben ser Numericos</center></div>"
						Response.end
				end if 
			else 
				if not isNumeric(idmedio) and trim(tipo)<> "2"  then 
					Response.write "<div style='background-color:#F8E0E0; color:#DF0101; font-family: 'Verdana'; font-size:10px; width:100%;'><center>Ingreso de Id Deben ser Numericos</center></div>"
					Response.end
				end if 		
			end if
			
			 if DescripcionMail <> "." then 
				DescripcionMail = replace(DescripcionMail,"'","")
			end if 
			
					sql_insert ="EXEC PROC_GESTIONES_MASIVAS_INGRESA 1,'" &  IdUsuarioCarga & "'," & ID & ",'" &  rutDeudor  & "','" & fechaIngreso  & "','" & HoraIngreso  & "','" & Observacion & "','" &  idUsuario  & "','" & idMedio   & "','" & DescripcionMail  & "','" & CodGestion  & "','" & CodCliente  & "'" 
			
		'	response.write(sql_insert)
			conn.execute(sql_insert)
		
	end if
		
		
		

next

		%>
		 <input type="hidden" name="error_largo_areglo" id="error_largo_areglo" value="<%=trim(error_largo_areglo)%>"> 
		<%


elseif trim(accion_ajax)="procesa_gestiones" then

	cod_tipo_gestion 	=request.querystring("cod_tipo_gestion")
	tipo 				=request.querystring("tipo")	
	ck_activa_id_email 	=request.querystring("ck_activa_id_email")	
	CB_COBRANZA =request.querystring("CB_COBRANZA")

	sql_procesa ="EXEC proc_Ingreso_Gestion_Masiva '"&TRIM(cod_tipo_gestion)&"','"&TRIM(session("session_idusuario"))&"','"&trim(tipo)&"','"&trim(ck_activa_id_email)&"','" & CB_COBRANZA &"'"
	'Response.write sql_procesa
	'Response.END
	conn.execute(sql_procesa)

	'Response.write sql_procesa&"<br>"&cod_tipo_gestion&"<br>"&tipo&"<br>"&ck_activa_id_email


end if

cerrarscg()

%>


