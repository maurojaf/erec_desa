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

if trim(accion_ajax)="verifica_biblioteca_deudores" then
	
'Response.write nombre_archivo&"<br>"&strRut&"<br>"&IntId

	sql_sel	="SELECT id_archivo, nombre_archivo, cod_cliente, rut "
	sql_sel = sql_sel & "FROM CARGA_ARCHIVOS "
	sql_sel = sql_sel & "WHERE activo =1 AND cod_cliente="&trim(IntId)&" AND origen = 2 and rut = '"&trim(strRut)&"' and nombre_archivo='"&trim(nombre_archivo)&"'	"
	'response.write sql_sel	
	set rs_sel	=conn.execute(sql_sel)
	if err then
		response.Write sql_sel &" / ERROR : "& err.description
		response.End()
	end if
	
	'response.Write sql_sel
	
	if not rs_sel.eof then	
	%>	
		<input type="hidden" id="archivo_validado" name="archivo_validado" value="si_existe" />
	<%
	else
	%>	
		<input type="hidden" id="archivo_validado" name="archivo_validado" value="no_existe" />
	<%	
	end if

elseif trim(accion_ajax)="verifica_carga_archivos_admin" then

	sql_sel	="SELECT id_archivo, nombre_archivo, cod_cliente, rut "
	sql_sel = sql_sel & "FROM CARGA_ARCHIVOS "
	sql_sel = sql_sel & "WHERE activo =1 AND cod_cliente="&trim(IntId)&" AND origen = 4 "
	sql_sel = sql_sel & "AND nombre_archivo='"&trim(nombre_archivo)&"'	"
	'response.write sql_sel	
	set rs_sel	=conn.execute(sql_sel)
	if err then
		response.Write sql_sel &" / ERROR : "& err.description
		response.End()
	end if
	
	'response.Write sql_sel
	
	if not rs_sel.eof then	
	%>	
		<input type="hidden" id="archivo_validado" name="archivo_validado" value="si_existe" />
	<%
	else
	%>	
		<input type="hidden" id="archivo_validado" name="archivo_validado" value="no_existe" />
	<%	
	end if	


elseif trim(accion_ajax)="verifica_biblioteca_cliente" then

	sql_sel	="SELECT id_archivo, nombre_archivo, cod_cliente, rut "
	sql_sel = sql_sel & "FROM CARGA_ARCHIVOS "
	sql_sel = sql_sel & "WHERE activo =1 AND cod_cliente="&trim(IntId)&" AND origen = 1 "
	sql_sel = sql_sel & "AND nombre_archivo='"&trim(nombre_archivo)&"'	"
	'response.write sql_sel	
	set rs_sel	=conn.execute(sql_sel)
	if err then
		response.Write sql_sel &" / ERROR : "& err.description
		response.End()
	end if
	
	'response.Write sql_sel
	
	if not rs_sel.eof then	
	%>	
		<input type="hidden" id="archivo_validado" name="archivo_validado" value="si_existe" />
	<%
	else
	%>	
		<input type="hidden" id="archivo_validado" name="archivo_validado" value="no_existe" />
	<%	
	end if


elseif trim(accion_ajax)="verifica_informes_anexados"	then

	sql_sel	="SELECT id_archivo, nombre_archivo, cod_cliente, rut "
	sql_sel = sql_sel & "FROM CARGA_ARCHIVOS "
	sql_sel = sql_sel & "WHERE activo =1 AND cod_cliente="&trim(IntId)&" AND origen = 5 "
	sql_sel = sql_sel & "AND nombre_archivo='"&trim(nombre_archivo)&"'	"
	'response.write sql_sel	
	set rs_sel	=conn.execute(sql_sel)
	if err then
		response.Write sql_sel &" / ERROR : "& err.description
		response.End()
	end if
	
	'response.Write sql_sel
	
	if not rs_sel.eof then	
	%>	
		<input type="hidden" id="archivo_validado" name="archivo_validado" value="si_existe" />
	<%
	else
	%>	
		<input type="hidden" id="archivo_validado" name="archivo_validado" value="no_existe" />
	<%	
	end if
end if


cerrarscg()
%>

