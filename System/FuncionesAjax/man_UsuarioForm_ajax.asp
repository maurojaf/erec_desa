<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->
<%

Response.CodePage = 65001
Response.charset="utf-8"


accion_ajax 		=request.querystring("accion_ajax")

abrirscg()

if trim(accion_ajax)="guardar_usuario" then

	TELEFONO_CONTACTO 	=request.querystring("TELEFONO_CONTACTO")
	CORREO_ELECTRONICO 	=request.querystring("CORREO_ELECTRONICO")
	FECHA_NACIMIENTO 	=request.querystring("FECHA_NACIMIENTO")
	RUT_USUARIO 		=request.querystring("RUT_USUARIO")
	APELLIDO_MATERNO 	=request.querystring("APELLIDO_MATERNO")
	APELLIDO_PATERNO 	=request.querystring("APELLIDO_PATERNO")
	NOMBRES_USUARIO 	=request.querystring("NOMBRES_USUARIO")
	PERFIL_ADM         	=request.querystring("PERFIL_ADM")
	PERFIL_EMP 			=request.querystring("PERFIL_EMP")
	PERFIL_FULL 		=request.querystring("PERFIL_FULL")
	PERFIL_PROC 		=request.querystring("PERFIL_PROC")
	PERFIL_CAJA			=request.querystring("PERFIL_CAJA")
	PERFIL_COB 			=request.querystring("PERFIL_COB")
	PERFIL_SUP 			=request.querystring("PERFIL_SUP")
	ACTIVO 				=request.querystring("ACTIVO")
	COD_AREA     		=request.querystring("COD_AREA")
	ANEXO 				=request.querystring("ANEXO")
	OBSERVACIONES		=request.querystring("OBSERVACIONES")

    EsInterno		                    =request.querystring("EsInterno")
    EsExterno		                    =request.querystring("EsExterno")
    PuedenEscucharMisGrabaciones		=request.querystring("PuedenEscucharMisGrabaciones")
    PuedoEscucharGrabaciones        	=request.querystring("PuedoEscucharGrabaciones")
    CodigoAgenteElastix       =request.querystring("CodigoAgenteElastix")


	if PERFIL_ADM ="" then
		PERFIL_ADM ="NULL"
	end if	
	if PERFIL_EMP ="" then
		PERFIL_EMP ="NULL"
	end if
	if PERFIL_FULL ="" then
		PERFIL_FULL ="NULL"
	end if	
	if PERFIL_PROC ="" then
		PERFIL_PROC ="NULL"
	end if
	if PERFIL_CAJA ="" then
		PERFIL_CAJA ="NULL"
	end if
	if PERFIL_COB ="" then
		PERFIL_COB ="NULL"
	end if	
	if PERFIL_SUP ="" then
		PERFIL_SUP ="NULL"
	end if	
    if CodigoAgenteElastix ="" then
		CodigoAgenteElastix ="NULL"
	end if

	strSQLQuery = "exec proc_usuario_ingresa '" & ucase(trim(RUT_USUARIO)) & "','" &  ucase(trim(NOMBRES_USUARIO)) & "','" & ucase(trim(APELLIDO_PATERNO)) & "','" & ucase(trim(APELLIDO_MATERNO)) & "', '" & ucase(trim(FECHA_NACIMIENTO)) & "','"& ucase(trim(CORREO_ELECTRONICO))&"', '"&trim(ucase(TELEFONO_CONTACTO))&"','COB_GER'," & ucase(trim(PERFIL_SUP)) & "," & ucase(trim(PERFIL_ADM)) & "," & ucase(trim(PERFIL_COB)) & "," & ucase(trim(PERFIL_CAJA)) & "," & ucase(trim(PERFIL_PROC)) & "," & ucase(trim(PERFIL_FULL)) & "," & ucase(trim(PERFIL_EMP)) & "," & ucase(trim(ACTIVO)) & ", '"&ucase(trim(OBSERVACIONES))&"','"&trim(COD_AREA)&"','"&trim(ANEXO)&"',"	& trim(EsInterno) & "," &  trim(EsExterno)   & "," & trim(PuedenEscucharMisGrabaciones)   & "," & trim(PuedoEscucharGrabaciones) & "," & trim(CodigoAgenteElastix)

   	set rs_sel = Conn.execute(strSQLQuery)
       if err then
    	Response.write strSQLQuery & " / error : "& err.description
    	response.end()
    end if
    'Response.write strSQLQuery

    %>
    <input type="hidden" name="ID_USUARIO" id="ID_USUARIO" value="<%=trim(rs_sel("ID_USUARIO"))%>">

    <%

elseif trim(accion_ajax)="actualiza_usuario" then

	TELEFONO_CONTACTO 	=request.querystring("TELEFONO_CONTACTO")
	CORREO_ELECTRONICO 	=request.querystring("CORREO_ELECTRONICO")
	FECHA_NACIMIENTO 	=request.querystring("FECHA_NACIMIENTO")
	APELLIDO_MATERNO 	=request.querystring("APELLIDO_MATERNO")
	APELLIDO_PATERNO 	=request.querystring("APELLIDO_PATERNO")
	NOMBRES_USUARIO 	=request.querystring("NOMBRES_USUARIO")
	PERFIL_ADM         	=request.querystring("PERFIL_ADM")
	PERFIL_EMP 			=request.querystring("PERFIL_EMP")
	PERFIL_FULL 		=request.querystring("PERFIL_FULL")
	PERFIL_PROC 		=request.querystring("PERFIL_PROC")
	PERFIL_CAJA			=request.querystring("PERFIL_CAJA")
	PERFIL_COB 			=request.querystring("PERFIL_COB")
	PERFIL_SUP 			=request.querystring("PERFIL_SUP")
	ACTIVO 				=request.querystring("ACTIVO")
	ID_USUARIO 			=request.querystring("ID_USUARIO")
	COD_AREA 			=request.querystring("COD_AREA")
	ANEXO 				=request.querystring("ANEXO")
	OBSERVACIONES		=request.querystring("OBSERVACIONES")

    
    EsInterno		                    =request.querystring("EsInterno")
    EsExterno		                    =request.querystring("EsExterno")
    PuedenEscucharMisGrabaciones		=request.querystring("PuedenEscucharMisGrabaciones")
    PuedoEscucharGrabaciones        	=request.querystring("PuedoEscucharGrabaciones")
    CodigoAgenteElastix       =request.querystring("CodigoAgenteElastix")



	if PERFIL_ADM ="" then
		PERFIL_ADM ="NULL"
	end if	
	if PERFIL_EMP ="" then
		PERFIL_EMP ="NULL"
	end if
	if PERFIL_FULL ="" then
		PERFIL_FULL ="NULL"
	end if	
	if PERFIL_PROC ="" then
		PERFIL_PROC ="NULL"
	end if
	if PERFIL_CAJA ="" then
		PERFIL_CAJA ="NULL"
	end if
	if PERFIL_COB ="" then
		PERFIL_COB ="NULL"
	end if	
	if PERFIL_SUP ="" then
		PERFIL_SUP ="NULL"
	end if	
    if CodigoAgenteElastix ="" then
		CodigoAgenteElastix ="NULL"
	end if

	strSQLQuery = "exec proc_usuario_actualiza '" & ucase(trim(ID_USUARIO)) & "','" &  ucase(trim(NOMBRES_USUARIO)) & "','" & ucase(trim(APELLIDO_PATERNO)) & "','" & ucase(trim(APELLIDO_MATERNO)) & "', '" & ucase(trim(FECHA_NACIMIENTO)) & "','"& ucase(trim(CORREO_ELECTRONICO))&"', '"&ucase(trim(TELEFONO_CONTACTO))&"','COB_GER'," & ucase(trim(PERFIL_SUP)) & "," & ucase(trim(PERFIL_ADM)) & "," & ucase(trim(PERFIL_COB)) & "," & ucase(trim(PERFIL_CAJA)) & "," & ucase(trim(PERFIL_PROC)) & "," & ucase(trim(PERFIL_FULL)) & "," & ucase(trim(PERFIL_EMP)) & "," & ucase(trim(ACTIVO)) & ", '"&ucase(trim(OBSERVACIONES))&"','"&trim(COD_AREA)&"','"&trim(ANEXO)&"',"	& trim(EsInterno) & "," &  trim(EsExterno)   & "," & trim(PuedenEscucharMisGrabaciones)   & "," & trim(PuedoEscucharGrabaciones) & "," & ucase(trim(CodigoAgenteElastix))
   	set rs_sel = Conn.execute(strSQLQuery)
       if err then
    	Response.write strSQLQuery & " / error : "& err.description
    	response.end()
    end if
    'Response.write strSQLQuery

    %>
    <input type="hidden" name="ID_USUARIO" id="ID_USUARIO" value="<%=trim(rs_sel("ID_USUARIO"))%>">

    <%



elseif trim(accion_ajax)="guardar_usuario_cliente" then
	ID_USUARIO 			=request.querystring("ID_USUARIO")
    COD_CLIENTE 	    =request.querystring("COD_CLIENTE")
	COD_CLIENTE 	    =split(COD_CLIENTE, ",")

	strSql = "DELETE USUARIO_CLIENTE WHERE ID_USUARIO= " &trim(ID_USUARIO)
	set rs_sel = Conn.execute(strSql)
	if err then
	    Response.write strSql & " / error : "& err.description
	    response.end()
	end if 

    For i = 0 to ubound(COD_CLIENTE)
        strSql = "INSERT INTO USUARIO_CLIENTE (ID_USUARIO, COD_CLIENTE) "
			    strSql = strSql & "VALUES (" & trim(ID_USUARIO) & "," & trim(COD_CLIENTE(i)) & ")"
   	    set rs_sel = Conn.execute(strSql)

        if err then
    	    Response.write strSql & " / error : "& err.description
            response.end()
        end if
    Next


elseif trim(accion_ajax)="verifica_login" then
	LOGIN 	=request.querystring("LOGIN")

	SQL_SEL ="SELECT LOGIN "
   	SQL_SEL = SQL_SEL & " FROM USUARIO "
   	SQL_SEL = SQL_SEL & " WHERE LOGIN='"&TRIM(LOGIN)&"' " 
   	set rs_sel = Conn.execute(SQL_SEL)
       if err then
    	Response.write SQL_SEL & " / error : "& err.description
    	response.end()
    end if
	'Response.write SQL_SEL
	if rs_sel.eof then
	%>
		<input type="hidden" name="verifica_login" id="verifica_login" value="S">
		<INPUT TYPE="hidden" NAME="ID_USUARIO" ID="ID_USUARIO" VALUE="">
	<%
	else
	%>
		<input type="hidden" name="verifica_login" id="verifica_login" value="N">
		<INPUT TYPE="hidden" NAME="ID_USUARIO" ID="ID_USUARIO" VALUE="">
	<%
	end if


elseif trim(accion_ajax)="verifica_rut" then
	RUT_USUARIO	=request.querystring("RUT_USUARIO") 

	SQL_SEL ="SELECT RUT_USUARIO "
   	SQL_SEL = SQL_SEL & " FROM USUARIO "
   	SQL_SEL = SQL_SEL & " WHERE RUT_USUARIO='"&TRIM(RUT_USUARIO)&"' " 
   	set rs_sel = Conn.execute(SQL_SEL)
       if err then
    	Response.write SQL_SEL & " / error : "& err.description
    	response.end()
    end if

	if rs_sel.eof then
	%>
		<input type="hidden" name="verifica_rut" id="verifica_rut" value="S">
		<INPUT TYPE="hidden" NAME="ID_USUARIO" ID="ID_USUARIO" VALUE="">
	<%
	else
	%>
		<input type="hidden" name="verifica_rut" id="verifica_rut" value="N">
		<INPUT TYPE="hidden" NAME="ID_USUARIO" ID="ID_USUARIO" VALUE="">
	<%
	end if

elseif trim(accion_ajax)="verifica_contrasena" then

	CLAVE_OLD 	=request.querystring("CLAVE_OLD")
	CLAVE_NEW1 	=request.querystring("CLAVE_NEW1")
	ID_USUARIO  =request.querystring("ID_USUARIO")

	sql_sel ="EXEC proc_verifica_contraseña '"&trim(CLAVE_NEW1)&"', "&trim(ID_USUARIO)&", '"&trim(CLAVE_OLD)&"'"
   	set rs_sel = Conn.execute(sql_sel)
       if err then
    	Response.write sql_sel & " / error : "& err.description
    	response.end()
    end if
'Response.write sql_sel
    if not rs_sel.eof then
    	estado_clave= rs_sel("estado_clave")
    end if

%>
	<input type="hidden" name="estado_clave" id="estado_clave" value="<%=trim(estado_clave)%>">
<%


elseif trim(accion_ajax)="modifica_contrasena" then
	CLAVE_NEW1 	=request.querystring("CLAVE_NEW1")
	ID_USUARIO  =request.querystring("ID_USUARIO")

	sql_sel ="exec proc_usuario_cambia_contrasena '"&trim(CLAVE_NEW1)&"','"&trim(ID_USUARIO)&"'"
   	Conn.execute(sql_sel)
       if err then
    	Response.write sql_sel & " / error : "& err.description
    	response.end()
    end if



elseif trim(accion_ajax)="verifica_asignaciones" then

	ID_USUARIO  =request.querystring("ID_USUARIO")

	sql_sel ="select * "  
	sql_sel = sql_sel & " from CUOTA C INNER JOIN CLIENTE CL ON C.COD_CLIENTE = CL.COD_CLIENTE"
	sql_sel = sql_sel & " where estado_deuda in (select codigo from estado_deuda where activo = 1) "
	sql_sel = sql_sel & " and USUARIO_ASIG = " & trim(ID_USUARIO)
	sql_sel = sql_sel & " AND CL.ACTIVO = 1 "
	
	set rs_sel = conn.execute(sql_sel)
	if not rs_sel.eof then
		%>
		<input type="hidden" name="verifica_asignacion" id="verifica_asignacion" value="S">
		<%
	else
		%>
		<input type="hidden" name="verifica_asignacion" id="verifica_asignacion" value="N">
		<%
	end if

elseif trim(accion_ajax)="actuliza_datos_usuario" then

	CORREO_ELECTRONICO 	=request.querystring("CORREO_ELECTRONICO")
	TELEFONO_CONTACTO 	=request.querystring("TELEFONO_CONTACTO")
	COD_AREA 			=request.querystring("COD_AREA")	
	ID_USUARIO 			=request.querystring("ID_USUARIO")
	ANEXO 				=request.querystring("ANEXO")

	sql_update ="UPDATE USUARIO " 
	sql_update = sql_update & " SET CORREO_ELECTRONICO='"&ucase(TRIM(CORREO_ELECTRONICO))&"', TELEFONO_CONTACTO='"&TRIM(TELEFONO_CONTACTO)&"', COD_AREA='"&TRIM(COD_AREA)&"', ANEXO='"&trim(ANEXO)&"'"
	sql_update = sql_update & " WHERE ID_USUARIO="&TRIM(ID_USUARIO)
	conn.execute(sql_update)
	if err then
		Response.write sql_update &" / ERROR : "& err.description
		response.end()
	end if


elseif trim(accion_ajax)="verifica_correo" then
	CORREO_ELECTRONICO 	=request.querystring("CORREO_ELECTRONICO")	

	SQL_SEL ="SELECT CORREO_ELECTRONICO "
   	SQL_SEL = SQL_SEL & " FROM USUARIO "
   	SQL_SEL = SQL_SEL & " WHERE CORREO_ELECTRONICO='"&TRIM(CORREO_ELECTRONICO)&"' " 
   	set rs_sel = Conn.execute(SQL_SEL)
       if err then
    	Response.write SQL_SEL & " / error : "& err.description
    	response.end()
    end if

	if rs_sel.eof then
	%>
		<input type="hidden" name="verifica_correo" id="verifica_correo" value="S">
	<%
	else
	%>
		<input type="hidden" name="verifica_correo" id="verifica_correo" value="N">
	<%
	end if

end if


cerrarscg()
%>