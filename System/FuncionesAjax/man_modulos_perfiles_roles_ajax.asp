<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->


<%

Response.CodePage = 65001
Response.charset="utf-8"

AbrirSCG()
accion_ajax =request.querystring("accion_ajax")

'response.write accion_ajax
if trim(accion_ajax) ="modulo" then
%>
	<span class="titulo_secundario nombre_modulo">Nombre módulo <input class="input" type="text" id="nombre_modulo" name="nombre_modulo" value=""></span><input class="boton_mantenedor" type="button" onclick="bt_guardar_modulo()" value="Guardar módulo" >
	
	<%sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
	sql_sel = sql_sel &"FROM MODULO "
	set rs_sel =Conn.execute(sql_sel)
	if err then 
		Response.write sql_sel & " / ERROR : " & err.description
		Response.end()
	end if	
	%>

	<div id="contenedor_tabla">
	<table class="tabla_mantenedor" border="0">
		<thead>
			<tr>
				<th class="table_cod">Código</th>
				<th>Nombre</th>
				<th>Fecha registro</th>
			</tr>
		</thead>		
		<tbody>
			<%if not rs_sel.eof then%>
			<%do while not rs_sel.eof
			   If ( i Mod 2 )= 1 Then
					bgcolor = "#E0F2F7"
			   Else
					bgcolor = "#FFFFFF"
			   End If
			   i = i + 1

			%>
			<tr bgcolor="<%=bgcolor%>">
				<td class="table_cod">&nbsp;<%=trim(rs_sel("mod_codigo"))%></td>
				<td>&nbsp;<%=trim(rs_sel("mod_nombre"))%></td>
				<td>&nbsp;<%=trim(rs_sel("mod_fecha_registro"))%></td>
			</tr>
			<%rs_sel.movenext
			loop%>
			<%else%>
			<tr>
				<td colspan="3">Sin registros</td>
			</tr>
			<%end if%>
		</tbody>
		
	</table>
	</div>
<%
elseif trim(accion_ajax) ="perfil" then
%>

	<span class="titulo_secundario nombre_modulo">Nombre Perfil <input class="input" type="text" id="nombre_perfil" name="nombre_perfil" value=""></span><input class="boton_mantenedor" type="button" onclick="bt_guardar_perfil()" value="Guardar perfil" >

	<div class="div_datos_modulo">
		
		<%sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
		sql_sel = sql_sel &"FROM MODULO "
		set rs_sel =Conn.execute(sql_sel)
		if err then 
			Response.write sql_sel & " / ERROR : " & err.description
			Response.end()
		end if	
		%>	
		<table class="tabla_mantenedor" border="0">
			<caption>Selecciona de módulo</caption>
			<thead>
				<tr>
					<th class=""></th>
					<th class="table_cod">Código</th>
					<th>Nombre</th>
					<th>Fecha registro</th>
				</tr>
			</thead>		
			<tbody>
				<%if not rs_sel.eof then%>
				<%do while not rs_sel.eof
				   If ( i Mod 2 )= 1 Then
						bgcolor = "#E0F2F7"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1

				%>
				<tr bgcolor="<%=bgcolor%>">
					<td class="table_cod"><input type="checkbox" id="ck_mod_codigo" name="ck_mod_codigo" value="<%=trim(rs_sel("mod_codigo"))%>"></td>
					<td class="table_cod">&nbsp;<%=trim(rs_sel("mod_codigo"))%></td>
					<td>&nbsp;<%=trim(rs_sel("mod_nombre"))%></td>
					<td>&nbsp;<%=trim(rs_sel("mod_fecha_registro"))%></td>
				</tr>
				<%rs_sel.movenext
				loop%>
				<%else%>
				<tr>
					<td colspan="3">Sin registros</td>
				</tr>
				<%end if%>
			</tbody>
			
		</table>

	</div>

	<div id="buscar_perfil" CLASS="titulo_principal Estilo13">Buscar perfiles creados</div>

	<div id="muestra_perfil">
		<%sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
		sql_sel = sql_sel &"FROM MODULO "
		set rs_sel_mod =Conn.execute(sql_sel)
		if err then 
			Response.write sql_sel & " / ERROR : " & err.description
			Response.end()
		end if	
		%>
		<div class="div_datos_accion" id="div_datos_accion">
			<%sql_sel =" SELECT per_codigo, per_nombre, mod_nombre, convert(varchar,  per_fecha_registro, 103) per_fecha_registro "
			sql_sel = sql_sel & " FROM MODULO_PERFIL moda "
			sql_sel = sql_sel & " INNER JOIN MODULO mod ON mod.mod_codigo=moda.mod_codigo "
			sql_sel = sql_sel & " order by mod_nombre, per_nombre desc "
			set rs_sel =Conn.execute(sql_sel)
			if err then 
				Response.write sql_sel & " / ERROR : " & err.description
				Response.end()
			end if	
			%>	
			<table class="tabla_mantenedor_accion" border="0">
				<caption> 
					<%if not rs_sel_mod.eof then%>
					<span>
						<select class="select_modulo" id="select_modulo_acccion" name="select_modulo_acccion" onchange="bt_select_modulo_perfil()">
							<option value="">Selecciona módulo</option>
							<option value="TODO">Todo</option>
							<%do while not rs_sel_mod.eof%>
								<option value="<%=trim(rs_sel_mod("mod_codigo"))%>"><%=trim(rs_sel_mod("mod_nombre"))%></option>
							<%rs_sel_mod.movenext
							loop%>
					 	</select>					 	
					</span>
					<%end if%>
				</caption>
				<thead>
					<tr>
						<th class="table_cod">Código</th>
						<th>Nombre perfil</th>
						<th>Nombre módulo</th>
						<th>Fecha registro</th>
					</tr>
				</thead>		
				<tbody>
					<%if not rs_sel.eof then%>
						<%do while not rs_sel.eof
						   If ( i Mod 2 )= 1 Then
								bgcolor = "#E0F2F7"
						   Else
								bgcolor = "#FFFFFF"
						   End If
						   i = i + 1
						%>
							<tr bgcolor="<%=bgcolor%>">
								<td class="table_cod"><%=trim(rs_sel("per_codigo"))%></td>
								<td><%=trim(rs_sel("per_nombre"))%></td>
								<td><%=trim(rs_sel("mod_nombre"))%></td>
								<td><%=trim(rs_sel("per_fecha_registro"))%></td>
							</tr>

						<%rs_sel.movenext
						loop%>
					<%else%>

					<%end if%>
				</tbody>
				
			</table>


		</div>
	</div>




<%
elseif trim(accion_ajax) ="cagar_info_contenedor" then 
	mod_codigo 					=request.querystring("mod_codigo")
	seleccion_relacion_perfil 	=request.querystring("seleccion_relacion_perfil")
	'response.write seleccion_relacion_perfil
	if trim(mod_codigo)<>"" then

	sql_sel =" SELECT acc_codigo, acc_nombre, mod_nombre, convert(varchar,  acc_fecha_registro, 103) acc_fecha_registro "

	if seleccion_relacion_perfil<>0 then
		sql_sel = sql_sel & ", ( "
		sql_sel = sql_sel & " SELECT COUNT(*) "
		sql_sel = sql_sel & " FROM MODULO_PERFIL_ACCION pa "
		sql_sel = sql_sel & " where per_codigo="&trim(seleccion_relacion_perfil)&" AND pa.acc_codigo=moda.acc_codigo "
		sql_sel = sql_sel & " ) cantidad "
	end if

	sql_sel = sql_sel & " FROM MODULO_ACCION moda "
	sql_sel = sql_sel & " INNER JOIN MODULO mod ON mod.mod_codigo=moda.mod_codigo "
	sql_sel = sql_sel & " WHERE moda.mod_codigo="& trim(mod_codigo)
	sql_sel = sql_sel & " order by mod_nombre, acc_nombre desc "
	set rs_sel =Conn.execute(sql_sel)
	if err then 
		Response.write sql_sel & " / ERROR : " & err.description
		Response.end()
	end if	
	'Response.write sql_sel
	%>	
	<table class="tabla_mantenedor_per_accion" border="0">
		<thead>
			<tr>
				<th class="table_cod">Código</th>
				<th>Nombre acción</th>
				<th>Nombre módulo</th>
			</tr>
		</thead>		
		<tbody>
			<%if not rs_sel.eof then%>
				<%do while not rs_sel.eof
				   If ( i Mod 2 )= 1 Then
						bgcolor = "#E0F2F7"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1

				   if seleccion_relacion_perfil>0 then
				   		if rs_sel("cantidad") >0 then			   		
				   			bgcolor="#F6D8CE"
				   		end if

				   end if
				%>
					<tr bgcolor="<%=bgcolor%>">
						<td class="table_cod">
							<input type="checkbox" name="seleccion_relacion_accion" id="seleccion_relacion_accion" value="<%=trim(rs_sel("acc_codigo"))%>"> 
							&nbsp;
							<%=trim(rs_sel("acc_codigo"))%>
						</td>
						<td><%=trim(rs_sel("acc_nombre"))%></td>
						<td><%=trim(rs_sel("mod_nombre"))%></td>
					</tr>

				<%rs_sel.movenext
				loop%>
			<%else%>

			<%end if%>
		</tbody>
		
	</table>
<%
	end if

elseif trim(accion_ajax) ="accion" then
%>
	<span class="titulo_secundario nombre_modulo">Nombre Acción <input class="input" type="text" id="nombre_accion" name="nombre_accion" value=""></span><input class="boton_mantenedor" type="button" onclick="bt_guardar_accion()" value="Guardar acción" >


	<div class="div_datos_modulo">
		
		<%sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
		sql_sel = sql_sel &"FROM MODULO "
		set rs_sel =Conn.execute(sql_sel)
		if err then 
			Response.write sql_sel & " / ERROR : " & err.description
			Response.end()
		end if	
		%>	
		<table class="tabla_mantenedor" border="0">
			<caption>Selecciona de módulo</caption>
			<thead>
				<tr>
					<th class=""></th>
					<th class="table_cod">Código</th>
					<th>Nombre</th>
					<th>Fecha registro</th>
				</tr>
			</thead>		
			<tbody>
				<%if not rs_sel.eof then%>
				<%do while not rs_sel.eof
				   If ( i Mod 2 )= 1 Then
						bgcolor = "#E0F2F7"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1

				%>
				<tr bgcolor="<%=bgcolor%>">
					<td class="table_cod"><input type="checkbox" id="ck_mod_codigo" name="ck_mod_codigo" value="<%=trim(rs_sel("mod_codigo"))%>"></td>
					<td class="table_cod">&nbsp;<%=trim(rs_sel("mod_codigo"))%></td>
					<td>&nbsp;<%=trim(rs_sel("mod_nombre"))%></td>
					<td>&nbsp;<%=trim(rs_sel("mod_fecha_registro"))%></td>
				</tr>
				<%rs_sel.movenext
				loop%>
				<%else%>
				<tr>
					<td colspan="3">Sin registros</td>
				</tr>
				<%end if%>
			</tbody>
			
		</table>

	</div>

	<div id="buscar_perfil" CLASS="titulo_principal Estilo13">Buscar acciones creadas</div>

	<div id="muestra_perfil">
		<%sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
		sql_sel = sql_sel &"FROM MODULO "
		set rs_sel_mod =Conn.execute(sql_sel)
		if err then 
			Response.write sql_sel & " / ERROR : " & err.description
			Response.end()
		end if	
		%>
		<div class="div_datos_accion" id="div_datos_accion">
			<%sql_sel =" SELECT acc_codigo, acc_nombre, mod_nombre, convert(varchar,  acc_fecha_registro, 103) acc_fecha_registro "
			sql_sel = sql_sel & " FROM MODULO_ACCION moda "
			sql_sel = sql_sel & " INNER JOIN MODULO mod ON mod.mod_codigo=moda.mod_codigo "
			sql_sel = sql_sel & " order by mod_nombre, acc_nombre desc "
			set rs_sel =Conn.execute(sql_sel)
			if err then 
				Response.write sql_sel & " / ERROR : " & err.description
				Response.end()
			end if	
			%>	
			<table class="tabla_mantenedor_accion" border="0">
				<caption> 
					<%if not rs_sel_mod.eof then%>
					<span>
						<select class="select_modulo" id="select_modulo_acccion" name="select_modulo_acccion" onchange="bt_select_modulo_acccion()">
							<option value="">Selecciona módulo</option>
							<option value="TODO">Todo</option>
							<%do while not rs_sel_mod.eof%>
								<option value="<%=trim(rs_sel_mod("mod_codigo"))%>"><%=trim(rs_sel_mod("mod_nombre"))%></option>
							<%rs_sel_mod.movenext
							loop%>
					 	</select>					 	
					</span>
					<%end if%>
				</caption>
				<thead>
					<tr>
						<th class="table_cod">Código</th>
						<th>Nombre acción</th>
						<th>Nombre módulo</th>
						<th>Fecha registro</th>
					</tr>
				</thead>		
				<tbody>
					<%if not rs_sel.eof then%>
						<%do while not rs_sel.eof
						   If ( i Mod 2 )= 1 Then
								bgcolor = "#E0F2F7"
						   Else
								bgcolor = "#FFFFFF"
						   End If
						   i = i + 1
						%>
							<tr bgcolor="<%=bgcolor%>">
								<td class="table_cod"><%=trim(rs_sel("acc_codigo"))%></td>
								<td><%=trim(rs_sel("acc_nombre"))%></td>
								<td><%=trim(rs_sel("mod_nombre"))%></td>
								<td><%=trim(rs_sel("acc_fecha_registro"))%></td>
							</tr>

						<%rs_sel.movenext
						loop%>
					<%else%>

					<%end if%>
				</tbody>
				
			</table>


		</div>
	</div>
<%

elseif trim(accion_ajax) ="guardar_modulo" then
	nombre_modulo =request.querystring("nombre_modulo")

	sql_insert ="exec proc_modulo_ingresa_modulos '"&trim(nombre_modulo)&"'"
	Conn.execute(sql_insert)
	if err then 
		Response.write sql_insert & " / ERROR : " & err.description
		Response.end()
	end if


%>
	<span class="titulo_secundario nombre_modulo">Nombre módulo <input class="input" type="text" id="nombre_modulo" name="nombre_modulo" value=""></span><input class="boton_mantenedor" type="button" onclick="bt_guardar_modulo()" value="Guardar" >
	

	<%sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
	sql_sel = sql_sel &"FROM MODULO "
	set rs_sel =Conn.execute(sql_sel)
	if err then 
		Response.write sql_sel & " / ERROR : " & err.description
		Response.end()
	end if	
	%>

	
	<div id="contenedor_tabla">
	<table class="tabla_mantenedor" border="0">
		<thead>
			<tr>
				<th class="table_cod">Código</th>
				<th>Nombre</th>
				<th>Fecha registro</th>
			</tr>
		</thead>		
		<tbody>
			<%if not rs_sel.eof then%>
			<%do while not rs_sel.eof
			   If ( i Mod 2 )= 1 Then
					bgcolor = "#E0F2F7"
			   Else
					bgcolor = "#FFFFFF"
			   End If
			   i = i + 1

			%>
			<tr bgcolor="<%=bgcolor%>">
				<td class="table_cod">&nbsp;<%=trim(rs_sel("mod_codigo"))%></td>
				<td>&nbsp;<%=trim(rs_sel("mod_nombre"))%></td>
				<td>&nbsp;<%=trim(rs_sel("mod_fecha_registro"))%></td>
			</tr>
			<%rs_sel.movenext
			loop%>
			<%else%>
			<tr>
				<td colspan="3">Sin registros</td>
			</tr>
			<%end if%>
		</tbody>
		
	</table>
	</div>

<%	

elseif trim(accion_ajax)="guardar_accion" then

	nombre_accion 	=request.querystring("nombre_accion")
	ck_mod_codigo	=request.querystring("ck_mod_codigo")

	sql_insert ="exec proc_modulo_ingresa_acciones '"&trim(nombre_accion)&"','"&trim(ck_mod_codigo)&"'"
	Conn.execute(sql_insert)
	if err then 
		Response.write sql_insert & " / ERROR : " & err.description
		Response.end()
	end if	

elseif trim(accion_ajax)="filtrar_accion_modulo" then

	mod_codigo	=request.querystring("mod_codigo")

	sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
	sql_sel = sql_sel &"FROM MODULO "
	set rs_sel_mod =Conn.execute(sql_sel)
	if err then 
		Response.write sql_sel & " / ERROR : " & err.description
		Response.end()
	end if	
	


	sql_sel =" SELECT acc_codigo, acc_nombre, mod_nombre, convert(varchar,  acc_fecha_registro, 103) acc_fecha_registro "
	sql_sel = sql_sel & " FROM MODULO_ACCION moda "
	sql_sel = sql_sel & " INNER JOIN MODULO mod ON mod.mod_codigo=moda.mod_codigo "

	if trim(mod_codigo)<>"TODO" then
		sql_sel = sql_sel & " WHERE moda.mod_codigo="& trim(mod_codigo)
	end if

	sql_sel = sql_sel & " order by mod_nombre, acc_nombre desc "
	set rs_sel =Conn.execute(sql_sel)
	if err then 
		Response.write sql_sel & " / ERROR : " & err.description
		Response.end()
	end if	
	%>	
	<table class="tabla_mantenedor_accion" border="0">
		<caption> 
			<%if not rs_sel_mod.eof then%>
			<span>
				<select class="select_modulo" id="select_modulo_acccion" name="select_modulo_acccion" onchange="bt_select_modulo_acccion()">
					<option value="">Selecciona módulo</option>
					<option value="TODO">Todo</option>
					<%do while not rs_sel_mod.eof%>
						<option value="<%=trim(rs_sel_mod("mod_codigo"))%>"><%=trim(rs_sel_mod("mod_nombre"))%></option>
					<%rs_sel_mod.movenext
					loop%>
			 	</select>					 	
			</span>
			<%end if%>
		</caption>
		<thead>
			<tr>
				<th class="table_cod">Código</th>
				<th>Nombre acción</th>
				<th>Nombre módulo</th>
				<th>Fecha registro</th>
			</tr>
		</thead>		
		<tbody>
			<%if not rs_sel.eof then%>
				<%do while not rs_sel.eof
				   If ( i Mod 2 )= 1 Then
						bgcolor = "#E0F2F7"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1
				%>
					<tr bgcolor="<%=bgcolor%>">
						<td class="table_cod"><%=trim(rs_sel("acc_codigo"))%></td>
						<td><%=trim(rs_sel("acc_nombre"))%></td>
						<td><%=trim(rs_sel("mod_nombre"))%></td>
						<td><%=trim(rs_sel("acc_fecha_registro"))%></td>
					</tr>

				<%rs_sel.movenext
				loop%>
			<%else%>

			<%end if%>
		</tbody>
		
	</table>

<%

elseif trim(accion_ajax)="guardar_perfil" then
	nombre_perfil 	=request.querystring("nombre_perfil")
	ck_mod_codigo	=request.querystring("ck_mod_codigo")

	sql_insert ="exec proc_modulo_ingresa_perfil '"&trim(ck_mod_codigo)&"','"&trim(nombre_perfil)&"'"
 	Conn.execute(sql_insert)
	if err then 
		Response.write sql_insert & " / ERROR : " & err.description
		Response.end()
	end if	


elseif trim(accion_ajax)="per_accion" then
%>
	<div id="info_contenedor">	

		<div class="info_perfil" id="info_perfil">
			<div CLASS="titulo_principal Estilo13">Perfiles</div>

			<%
				sql_sel =" SELECT per_codigo, per_nombre, mod_nombre, convert(varchar,  per_fecha_registro, 103) per_fecha_registro "
				sql_sel = sql_sel & " FROM MODULO_PERFIL moda "
				sql_sel = sql_sel & " INNER JOIN MODULO mod ON mod.mod_codigo=moda.mod_codigo "
				sql_sel = sql_sel & " order by per_nombre desc "
				set rs_sel =Conn.execute(sql_sel)
				if err then 
					Response.write sql_sel & " / ERROR : " & err.description
					Response.end()
				end if	
				%>	
				<table class="tabla_mantenedor_perfil" border="0">
					<thead>
						<tr>
							<th></th>
							<th class="table_cod">Código</th>
							<th>Nombre Perfíl</th>
							
						</tr>
					</thead>		
					<tbody>
						<%if not rs_sel.eof then%>
							<%do while not rs_sel.eof
							   If ( i Mod 2 )= 1 Then
									bgcolor = "#E0F2F7"
							   Else
									bgcolor = "#FFFFFF"
							   End If
							   i = i + 1
							%>
								<tr bgcolor="<%=bgcolor%>">
									<td class="tr_perfil" width="5%">
										<input type="radio" name="seleccion_relacion_perfil" id="seleccion_relacion_perfil" value="<%=trim(rs_sel("per_codigo"))%>">
									</td>
									<td class="tr_perfil" width="20%" class="table_cod"><%=trim(rs_sel("per_codigo"))%></td>
									<td class="tr_perfil" width="75%"><%=trim(rs_sel("per_nombre"))%></td>									
								</tr>

							<%rs_sel.movenext
							loop%>
						<%else%>

						<%end if%>
					</tbody>
					
				</table>

		</div>

		<div class="info_perfil">

			<div CLASS="titulo_principal Estilo13">Acciones</div>
			<%sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
			sql_sel = sql_sel &"FROM MODULO "
			set rs_sel =Conn.execute(sql_sel)
			if err then 
				Response.write sql_sel & " / ERROR : " & err.description
				Response.end()
			end if	
			%>
			<div>				
				<select name="mod_codigo" id="mod_codigo" onchange="bt_refresca_info()">
					<option value="">Selecciona módulo</option>
					<%do while not rs_sel.eof%>
						<option value="<%=trim(rs_sel("mod_codigo"))%>"><%=trim(rs_sel("mod_nombre"))%></option>
					<%rs_sel.movenext
					loop%>				
				</select>
			</div>	
			<div  id="info_accion"></div>		
		</div>

	</div>


<%


elseif trim(accion_ajax)="filtrar_perfil_modulo" then
	mod_codigo =request.querystring("mod_codigo")

	sql_sel ="SELECT mod_codigo, mod_nombre, convert(varchar, mod_fecha_registro, 103) mod_fecha_registro "
	sql_sel = sql_sel &"FROM MODULO "
	set rs_sel_mod =Conn.execute(sql_sel)
	if err then 
		Response.write sql_sel & " / ERROR : " & err.description
		Response.end()
	end if	

	sql_sel =" SELECT per_codigo, per_nombre, mod_nombre, convert(varchar,  per_fecha_registro, 103) per_fecha_registro "
	sql_sel = sql_sel & " FROM MODULO_PERFIL moda "
	sql_sel = sql_sel & " INNER JOIN MODULO mod ON mod.mod_codigo=moda.mod_codigo "
	
	if trim(mod_codigo)<>"TODO" then
		sql_sel = sql_sel & " WHERE moda.mod_codigo = " & trim(mod_codigo)
	end if

	sql_sel = sql_sel & " order by mod_nombre, per_nombre desc "
	set rs_sel =Conn.execute(sql_sel)
	if err then 
		Response.write sql_sel & " / ERROR : " & err.description
		Response.end()
	end if	
	%>	
	<table class="tabla_mantenedor_accion" border="0">
		<caption> 
			<%if not rs_sel_mod.eof then%>
			<span>
				<select class="select_modulo" id="select_modulo_acccion" name="select_modulo_acccion" onchange="bt_select_modulo_perfil()">
					<option value="">Selecciona módulo</option>
					<option value="TODO">Todo</option>
					<%do while not rs_sel_mod.eof%>
						<option value="<%=trim(rs_sel_mod("mod_codigo"))%>" <%if trim(mod_codigo)=trim(rs_sel_mod("mod_codigo")) then response.write " selected " end if%> ><%=trim(rs_sel_mod("mod_nombre"))%></option>
					<%rs_sel_mod.movenext
					loop%>
			 	</select>
			</span>
			<%end if%>
		</caption>
		<thead>
			<tr>
				<th class="table_cod">Código</th>
				<th>Nombre perfil</th>
				<th>Nombre módulo</th>
				<th>Fecha registro</th>
			</tr>
		</thead>		
		<tbody>
			<%if not rs_sel.eof then%>
				<%do while not rs_sel.eof
				   If ( i Mod 2 )= 1 Then
						bgcolor = "#E0F2F7"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1
				%>
					<tr bgcolor="<%=bgcolor%>">
						<td class="table_cod"><%=trim(rs_sel("per_codigo"))%></td>
						<td><%=trim(rs_sel("per_nombre"))%></td>
						<td><%=trim(rs_sel("mod_nombre"))%></td>
						<td><%=trim(rs_sel("per_fecha_registro"))%></td>
					</tr>

				<%rs_sel.movenext
				loop%>
			<%else%>

			<%end if%>
		</tbody>
		
	</table>


<%


elseif trim(accion_ajax)="relacion_perfil_accion" then
	seleccion_relacion_accion =request.querystring("seleccion_relacion_accion")
	seleccion_relacion_perfil =request.querystring("seleccion_relacion_perfil")


	sql_insert ="exec proc_modulo_relaciona_perfil_accion '"&trim(seleccion_relacion_perfil)&"','"&trim(seleccion_relacion_accion)&"'"
	conn.execute(sql_insert)
	if err then
		response.write sql_insert &" / ERROR : "& err.description
		Response.end()
	end if

end if
%>
<%CerrarSCG()%>

