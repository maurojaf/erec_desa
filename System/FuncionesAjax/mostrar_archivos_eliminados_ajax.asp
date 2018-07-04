<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"


accion_ajax 	=request("accion_ajax")
nombre_archivo 	=request("nombre_archivo")
strRut 			=request("strRut")
IntId 			=request("IntId")

abrirscg()

if trim(accion_ajax)="mostrar_archivos_eliminados_bibioteca_deudores" then
	
'Response.write nombre_archivo&"<br>"&strRut&"<br>"&IntId
	SQL_SEL ="SELECT id_archivo, nombre_archivo, cod_cliente, rut, convert(varchar(10), FECHA_CARGA, 103) +' '+CONVERT(VARCHAR(5),FECHA_CARGA, 108) FECHA_CARGA, "
	SQL_SEL = SQL_SEL & "ID_USUARIO_CARGA, convert(varchar(10), FECHA_ELIMINACION, 103) +' '+CONVERT(VARCHAR(5),FECHA_ELIMINACION, 108) FECHA_ELIMINACION, ID_USUARIO_ELIMINACION,  "
	SQL_SEL = SQL_SEL & "isnull(usu.nombres_usuario,'')+' '+isnull(usu.apellido_paterno,'')+' '+isnull(usu.apellido_materno,'') nombre_usuario, "
	SQL_SEL = SQL_SEL & "isnull(usuelim.nombres_usuario,'')+' '+isnull(usuelim.apellido_paterno,'')+' '+isnull(usuelim.apellido_materno,'') nombre_usuario_eliminacion "
	SQL_SEL = SQL_SEL & "FROM CARGA_ARCHIVOS car " 
	SQL_SEL = SQL_SEL & "INNER JOIN USUARIO usu ON usu.ID_USUARIO=car.id_usuario_carga "
	SQL_SEL = SQL_SEL & "left JOIN USUARIO usuelim ON usuelim.ID_USUARIO=car.id_usuario_eliminacion "
	SQL_SEL = SQL_SEL & "WHERE car.activo =0 AND cod_cliente="&trim(IntId)&" AND rut ='"&trim(strRut)&"' " 
	SQL_SEL = SQL_SEL & "AND origen = 2 "
	SQL_SEL = SQL_SEL & " ORDER BY FECHA_ELIMINACION DESC"

	'Response.write SQL_SEL
	set rs_sql_sel = Conn.execute(SQL_SEL)

	IF not rs_sql_sel.eof then
%>
	<div>

		<table class="intercalado" style="width:100%;">
			<thead>
				<tr>
					<th>Nombre archivo</th>
					<th>Fecha carga</th>
					<th>Usuario carga</th>
					<th>Fecha eliminación</th>
					<th>Usuario eliminación</th>
				</tr>
			</thead>
			<tbody>
			<%
				do while not rs_sql_sel.eof

				   If ( i Mod 2 )= 1 Then
						bgcolor = "#F0F0F0"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1

			%>
					<tr class="td_hover" BGCOLOR="<%=bgcolor%>">

						<td align="left">&nbsp;<%=trim(rs_sql_sel("nombre_archivo"))%></td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_CARGA"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario"))%></td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_ELIMINACION"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario_eliminacion"))%></td>

					</tr>
			<%

				rs_sql_sel.movenext 
				loop
			%>
			</tbody>
		</table>

		</div>

<%
	else%>

		<div>
			<label style='font: 14px bold #000;'>Sin archivos eliminados</label>
		</div>	
		
	<%end if

elseif trim(accion_ajax)="mostrar_archivos_eliminados_carga_archivos_admin" then

'Response.write nombre_archivo&"<br>"&strRut&"<br>"&IntId
	SQL_SEL ="SELECT id_archivo, nombre_archivo, cod_cliente, rut, convert(varchar(10), FECHA_CARGA, 103) +' '+CONVERT(VARCHAR(5),FECHA_CARGA, 108) FECHA_CARGA, "
	SQL_SEL = SQL_SEL & "ID_USUARIO_CARGA, convert(varchar(10), FECHA_ELIMINACION, 103) +' '+CONVERT(VARCHAR(5),FECHA_ELIMINACION, 108) FECHA_ELIMINACION, ID_USUARIO_ELIMINACION,  "
	SQL_SEL = SQL_SEL & "isnull(usu.nombres_usuario,'')+' '+isnull(usu.apellido_paterno,'')+' '+isnull(usu.apellido_materno,'') nombre_usuario, "
	SQL_SEL = SQL_SEL & "isnull(usuelim.nombres_usuario,'')+' '+isnull(usuelim.apellido_paterno,'')+' '+isnull(usuelim.apellido_materno,'') nombre_usuario_eliminacion "
	SQL_SEL = SQL_SEL & "FROM CARGA_ARCHIVOS car " 
	SQL_SEL = SQL_SEL & "INNER JOIN USUARIO usu ON usu.ID_USUARIO=car.id_usuario_carga "
	SQL_SEL = SQL_SEL & "left JOIN USUARIO usuelim ON usuelim.ID_USUARIO=car.id_usuario_eliminacion "
	SQL_SEL = SQL_SEL & "WHERE car.activo =0 AND cod_cliente="&trim(IntId)&" AND origen = 4 "
	SQL_SEL = SQL_SEL & " ORDER BY FECHA_ELIMINACION DESC"

	'Response.write SQL_SEL
	set rs_sql_sel = Conn.execute(SQL_SEL)

	IF not rs_sql_sel.eof then
%>
	<div>

		<table class="intercalado" style="width:100%;">
			<thead>
				<tr>
					<th>Nombre archivo</th>
					<th>Fecha carga</th>
					<th>Usuario carga</th>
					<th>Fecha eliminación</th>
					<th>Usuario eliminación</th>
				</tr>
			</thead>
			<tbody>
			<%
				do while not rs_sql_sel.eof

				   If ( i Mod 2 )= 1 Then
						bgcolor = "#F0F0F0"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1

			%>
					<tr class="td_hover" BGCOLOR="<%=bgcolor%>">

						<td align="left">&nbsp;<%=trim(rs_sql_sel("nombre_archivo"))%></td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_CARGA"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario"))%></td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_ELIMINACION"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario_eliminacion"))%></td>

					</tr>
			<%

				rs_sql_sel.movenext 
				loop
			%>
			</tbody>
		</table>

		</div>

<%
	else%>

		<div>
			<label style='font: 14px bold #000;'>Sin archivos eliminados</label>
		</div>	
		
	<%end if

elseif trim(accion_ajax)="mostrar_archivos_eliminados_biblioteca_clientes" then

	SQL_SEL ="SELECT id_archivo, nombre_archivo, cod_cliente, rut, convert(varchar(10), FECHA_CARGA, 103) +' '+CONVERT(VARCHAR(5),FECHA_CARGA, 108) FECHA_CARGA, "
	SQL_SEL = SQL_SEL & "ID_USUARIO_CARGA, convert(varchar(10), FECHA_ELIMINACION, 103) +' '+CONVERT(VARCHAR(5),FECHA_ELIMINACION, 108) FECHA_ELIMINACION, ID_USUARIO_ELIMINACION,  "
	SQL_SEL = SQL_SEL & "isnull(usu.nombres_usuario,'')+' '+isnull(usu.apellido_paterno,'')+' '+isnull(usu.apellido_materno,'')  nombre_usuario, "
	SQL_SEL = SQL_SEL & "isnull(usuelim.nombres_usuario,'')+' '+isnull(usuelim.apellido_paterno,'')+' '+isnull(usuelim.apellido_materno,'') nombre_usuario_eliminacion "
	SQL_SEL = SQL_SEL & "FROM CARGA_ARCHIVOS car " 
	SQL_SEL = SQL_SEL & "INNER JOIN USUARIO usu ON usu.ID_USUARIO=car.id_usuario_carga "
	SQL_SEL = SQL_SEL & "left JOIN USUARIO usuelim ON usuelim.ID_USUARIO=car.id_usuario_eliminacion "
	SQL_SEL = SQL_SEL & "WHERE car.activo =0 AND cod_cliente="&trim(IntId)&" AND origen = 1 "
	SQL_SEL = SQL_SEL & " ORDER BY FECHA_ELIMINACION DESC"

	'Response.write SQL_SEL
	set rs_sql_sel = Conn.execute(SQL_SEL)

	IF not rs_sql_sel.eof then
%>
	<div>

		<table class="intercalado" style="width:100%;">
			<thead>
				<tr>
					<th>Nombre archivo</th>
					<th>Fecha carga</th>
					<th>Usuario carga</th>
					<th>Fecha eliminación</th>
					<th>Usuario eliminación</th>
				</tr>
			</thead>
			<tbody>
			<%
				do while not rs_sql_sel.eof

				   If ( i Mod 2 )= 1 Then
						bgcolor = "#F0F0F0"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1

			%>
					<tr class="td_hover" BGCOLOR="<%=bgcolor%>">

						<td align="left">&nbsp;<%=trim(rs_sql_sel("nombre_archivo"))%></td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_CARGA"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario"))%></td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_ELIMINACION"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario_eliminacion"))%></td>

					</tr>
			<%

				rs_sql_sel.movenext 
				loop
			%>
			</tbody>
		</table>

		</div>

<%
	else%>

		<div>
			<label style='font: 14px bold #000;'>Sin archivos eliminados</label>
		</div>	
		
	<%end if


elseif trim(accion_ajax)="mostrar_archivos_eliminados_informe_clientes" then

	SQL_SEL ="SELECT id_archivo, nombre_archivo, cod_cliente, rut, convert(varchar(10), FECHA_CARGA, 103) +' '+CONVERT(VARCHAR(5),FECHA_CARGA, 108) FECHA_CARGA, "
	SQL_SEL = SQL_SEL & "ID_USUARIO_CARGA, convert(varchar(10), FECHA_ELIMINACION, 103) +' '+CONVERT(VARCHAR(5),FECHA_ELIMINACION, 108) FECHA_ELIMINACION, ID_USUARIO_ELIMINACION,  "
	SQL_SEL = SQL_SEL & "isnull(usu.nombres_usuario,'')+' '+isnull(usu.apellido_paterno,'')+' '+isnull(usu.apellido_materno,'') nombre_usuario, "
	SQL_SEL = SQL_SEL & "isnull(usuelim.nombres_usuario,'')+' '+isnull(usuelim.apellido_paterno,'')+' '+isnull(usuelim.apellido_materno,'') nombre_usuario_eliminacion "
	SQL_SEL = SQL_SEL & "FROM CARGA_ARCHIVOS car " 
	SQL_SEL = SQL_SEL & "INNER JOIN USUARIO usu ON usu.ID_USUARIO=car.id_usuario_carga "
	SQL_SEL = SQL_SEL & "left JOIN USUARIO usuelim ON usuelim.ID_USUARIO=car.id_usuario_eliminacion "
	SQL_SEL = SQL_SEL & "WHERE car.activo =0 AND cod_cliente="&trim(IntId)&" AND origen = 5 "
	SQL_SEL = SQL_SEL & " ORDER BY FECHA_ELIMINACION DESC"

	'Response.write SQL_SEL
	set rs_sql_sel = Conn.execute(SQL_SEL)

	IF not rs_sql_sel.eof then
%>
	<div>

		<table class="intercalado" style="width:100%;">
			<thead>
				<tr>
					<th>Nombre archivo</th>
					<th>Fecha carga</th>
					<th>Usuario carga</th>
					<th>Fecha eliminación</th>
					<th>Usuario eliminación</th>
				</tr>
			</thead>
			<tbody>
			<%
				do while not rs_sql_sel.eof

				   If ( i Mod 2 )= 1 Then
						bgcolor = "#F0F0F0"
				   Else
						bgcolor = "#FFFFFF"
				   End If
				   i = i + 1

			%>
					<tr class="td_hover" BGCOLOR="<%=bgcolor%>">

						<td align="left">&nbsp;<%=trim(rs_sql_sel("nombre_archivo"))%></td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_CARGA"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario"))%></td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_ELIMINACION"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario_eliminacion"))%></td>

					</tr>
			<%

				rs_sql_sel.movenext 
				loop
			%>
			</tbody>
		</table>

		</div>

<%
	else%>

		<div>
			<label style='font: 14px bold #000;'>Sin archivos eliminados</label>
		</div>	
		
	<%end if



elseif trim(accion_ajax)="mostrar_archivos_eliminados_vacio" then
	Response.write "&nbsp;"
end if


cerrarscg()
%>

