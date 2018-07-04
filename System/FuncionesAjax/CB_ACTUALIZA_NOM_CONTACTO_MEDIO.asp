<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>


<!--#include file="../arch_utils.asp"-->
<%


	Response.CodePage = 65001
	Response.charset="utf-8"

	accion_ajax =request.querystring("accion_ajax")
	abrirscg()

	'response.write accion_ajax

if trim(accion_ajax)="actualiza_CB_FONO_AGEND" then
	rut		=request.querystring("rut")

%>
	<SELECT NAME="CB_FONO_AGEND" id="CB_FONO_AGEND">
		<OPTION VALUE="0" >SELECCIONE</OPTION>
		<%if fono_con="0" or fono_con="" then%>
		<%
		AbrirSCG1()
		ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & rut & "' AND ESTADO <> 2"
		set rsFON=Conn1.execute(ssql_)
		Do until rsFON.eof
		
			strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
			strSel=""
			if strFonoCB = strFonoAsociado Then strSel = "SELECTED"	%>
			<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
		
		<%rsFON.movenext
		Loop
		rsFON.close
		set rsFON=nothing
		CerrarSCG1()
		%>
		<%else%>
			<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
		<%end if %>
	</SELECT>
<%



elseif trim(accion_ajax)="actualiza_CB_FONO_GESTION" then
	rut		=request.querystring("rut")
%>

	<select name="CB_FONO_GESTION" id="CB_FONO_GESTION" onchange="set_CB_CONTACTO_ASOCIADO(this.value); return false;">
	<option value="0">SELECCIONE</option>
	<%if fono_con="0" or fono_con="" then%>
	  <%
		AbrirSCG1()
		ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & rut & "' AND ESTADO <> 2"
		set rsFON=Conn1.execute(ssql_)
		Do until rsFON.eof
			strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
			strSel=""
			if strFonoCB = strFonoAsociado Then strSel = "SELECTED"
			%>
			<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
			<%
				rsFON.movenext
		Loop
		rsFON.close
		set rsFON=nothing
		CerrarSCG1()
	 %>
	<%else%>
		<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
	<%end if %>
	</select>
<%



elseif trim(accion_ajax)="actualiza_td_CB_FONO_CP_RUTA" then
	rut		=request.querystring("rut")
	
%>

	<select name="CB_FONO_CP_RUTA" id="CB_FONO_CP_RUTA" onchange="set_CB_CONTACTO_ASOCIADO_CP_RUTA(this.value); return false;">

	<option value="0">SELECCIONE</option>
	<%if fono_con="0" or fono_con="" then%>
	  <%
		AbrirSCG1()
		ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & rut & "' AND ESTADO <> 2"
		set rsFON=Conn1.execute(ssql_)
		Do until rsFON.eof
			strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
			strSel=""
			if strFonoCB = strFonoAgend Then strSel = "SELECTED"
			%>
			<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
			<%
				rsFON.movenext
		Loop
		rsFON.close
		set rsFON=nothing
		CerrarSCG1()
	 %>
	<%else%>
		<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
	<%end if %>
	</select>

<%

elseif trim(accion_ajax)="actualiza_CB_CONTACTO_ASOCIADO" then

	contentVar =request.queryString("contentVar")

	if trim(contentVar)<>0 then

			strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & contentVar
			strSql = strSql & " UNION"
			strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "


			set rsContacto = Conn.execute(strSql)

			'response.write "strQuery == " & strSql
			%>
			<select name="CB_CONTACTO_ASOCIADO" id="CB_CONTACTO_ASOCIADO">
				<option value="0">SELECCIONE</option>
			<%
			Do While not rsContacto.eof
				strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
			%>
				<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
			<%
				rsContacto.moveNext
			Loop
			rsContacto.close
			set rsContacto=nothing

	%>
			</select>

	<%else%>

		<select name="CB_CONTACTO_ASOCIADO" id="CB_CONTACTO_ASOCIADO">
			<option value="0">SELECCIONE</option>
		</select>
	<%end if

elseif trim(accion_ajax)="actualiza_td_CB_CONTACTO_ASOCIADO_CP_RUTA" then
	
	CB_FONO_CP_RUTA =request.querystring("CB_FONO_CP_RUTA")

	if trim(CB_FONO_CP_RUTA)<>0 then

		AbrirSCG1()
		strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & CB_FONO_CP_RUTA
		strSql = strSql & " UNION"
		strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "
		set rsContacto = Conn1.execute(strSql)
	%>

		<select name="CB_CONTACTO_ASOCIADO_CP_RUTA" id="CB_CONTACTO_ASOCIADO_CP_RUTA" onchange="this.style.width=260">
			<option value="0">SELECCIONE</option>
			<%do while not rsContacto.eof%>
				<option value="<%=trim(rsContacto("ID_CONTACTO"))%>"><%=trim(rsContacto("CONTACTO"))%></option>
			<%rsContacto.movenext
			loop%>
		</select>

	<%CerrarSCG1()
		'response.write strSql

	else
	%>
		<select name="CB_CONTACTO_ASOCIADO_CP_RUTA" id="CB_CONTACTO_ASOCIADO_CP_RUTA"  onchange="this.style.width=260">
			<option value="0">SELECCIONE</option>
		</select>
	<%
	end if


elseif trim(accion_ajax)="actualiza_CB_EMAIL_GESTION" then
	rut		=request.querystring("rut")
%>

	<select name="CB_EMAIL_GESTION" id="CB_EMAIL_GESTION" onchange="set_CB_CONTACTO_ASOCIADO_EMAIL(this.value); return false;"  onchange="this.style.width=260">
		<option value="0">SELECCIONE</option>
	  <%
		AbrirSCG1()
		ssql_ = "SELECT ID_EMAIL, UPPER(EMAIL) AS EMAIL FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & rut & "' AND ESTADO <> 2"
		set rsEmail=Conn1.execute(ssql_)
		Do until rsEmail.eof
			strEmailCB = rsEmail("EMAIL")
			strSel=""
			if strEmailCB = strEmailAgestionar Then strSel = "SELECTED"
			%>
			<option value="<%=rsEmail("ID_EMAIL")%>" <%=strSel%>><%=rsEmail("EMAIL")%></option>
			<%
				rsEmail.movenext
		Loop
		rsEmail.close
		set rsEmail=nothing
		CerrarSCG1()
	 %>
	</select>


<%
elseif trim(accion_ajax)="actualiza_CB_CONTACTO_ASOCIADO_EMAIL" then
	contentVar =request.queryString("contentVar")

	if trim(contentVar)<>0 then

			strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM EMAIL_CONTACTO WHERE ID_EMAIL = " & contentVar
			strSql = strSql & " UNION"
			strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "


			set rsContacto = Conn.execute(strSql)

			'response.write "strQuery == " & strSql
			%>
			<select name="CB_CONTACTO_ASOCIADO_EMAIL" id="CB_CONTACTO_ASOCIADO_EMAIL">
				<option value="0">SELECCIONE</option>
			<%
			Do While not rsContacto.eof
				strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
			%>
				<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
			<%
				rsContacto.moveNext
			Loop
			rsContacto.close
			set rsContacto=nothing

	%>
			</select>

	<%else%>

		<select name="CB_CONTACTO_ASOCIADO_EMAIL" id="CB_CONTACTO_ASOCIADO_EMAIL">
			<option value="0">SELECCIONE</option>
		</select>

	<%end if





ELSEIf Trim(accion_ajax) = "refresca_CB_LUGARPAGO_CP_RUTA" Then

	cliente = request.querystring("cliente_")
	rut		=request.querystring("rut")	
%>
	<select name="CB_LUGARPAGO_CP_RUTA" id="CB_LUGARPAGO_CP_RUTA" onchange="this.style.width=330">
		<option value="0">SELECCIONE</option>
		<%
		AbrirSCG1()

		strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & trim(rut) & "' AND ESTADO <> 2"

		strSql = strSql & " UNION"

		strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(cliente) & "' ORDER BY ORDEN ASC"


		set rsDIR=Conn1.execute(strSql)
		do until rsDIR.eof
			direccion = rsDIR("ID")&"-"&rsDIR("TIPO")
			%>
			<option value="<%=direccion%>"
			<%if Trim(strLugarPago)=Trim(direccion) then
				response.Write("Selected")
			end if%>
			><%=trim(rsDIR("LUGAR_PAGO"))%></option>
			<%
			rsDIR.movenext
		loop
		rsDIR.close
		set rsDIR=nothing
		CerrarSCG1()
		%>
	</select>

<%

elseif trim(accion_ajax)="refresca_CB_DIRECCION_GESTION" then
	cliente = request.querystring("cliente_")
	rut		=request.querystring("rut")	

%>
	<select name="CB_DIRECCION_GESTION" id="CB_DIRECCION_GESTION" onchange="set_CB_CONTACTO_ASOCIADO_DIRECCION(this.value); return false;" onchange="this.style.width=260">

		<option value="0">SELECCIONE</option>
		<%
		AbrirSCG1()

		strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & trim(rut) & "' AND ESTADO <> 2"
		
		set rsDIR=Conn1.execute(strSql)
		do until rsDIR.eof
			direccion = rsDIR("ID")
			%>
			<option value="<%=direccion%>"
			<%if Trim(strLugarPago)=Trim(direccion) then
				response.Write("Selected")
			end if%>
			><%=trim(rsDIR("LUGAR_PAGO"))%></option>
			<%
			rsDIR.movenext
		loop
		rsDIR.close
		set rsDIR=nothing
		CerrarSCG1()
		%>
	</select>
<%

elseif trim(accion_ajax)="refresca_CB_DIRECCION_TERRENO" then
	cliente = request.querystring("cliente_")
	rut		=request.querystring("rut")	
%>
	<select name="CB_DIRECCION_TERRENO" id="CB_DIRECCION_TERRENO" onchange="this.style.width=330">

		<option value="0">SELECCIONE</option>
		<%
		AbrirSCG1()

		strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & trim(rut) & "' AND ESTADO <> 2"
		
		'strSql = strSql & " UNION"

		'strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(cliente) & "' ORDER BY ORDEN ASC"


		set rsDIR=Conn1.execute(strSql)
		do until rsDIR.eof
			direccion = rsDIR("ID")&"-"&rsDIR("TIPO")
			%>
			<option value="<%=direccion%>"
			<%if Trim(strLugarPago)=Trim(direccion) then
				response.Write("Selected")
			end if%>
			><%=trim(rsDIR("LUGAR_PAGO"))%></option>
			<%
			rsDIR.movenext
		loop
		rsDIR.close
		set rsDIR=nothing
		CerrarSCG1()
		%>
	</select>
<%

elseif trim(accion_ajax)="refresca_CB_LUGAR_NORM" then
	cliente = request.querystring("cliente_")
	rut		=request.querystring("rut")	
%>
	<select name="CB_LUGAR_NORM" id="CB_LUGAR_NORM" onchange="this.style.width=330">

		<option value="0">SELECCIONE</option>
		<%
		AbrirSCG1()

		strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & trim(rut) & "' AND ESTADO <> 2"
		
		strSql = strSql & " UNION"

		strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(cliente) & "' ORDER BY ORDEN ASC"


		set rsDIR=Conn1.execute(strSql)
		do until rsDIR.eof
			direccion = rsDIR("ID")&"-"&rsDIR("TIPO")
			%>
			<option value="<%=direccion%>"
			<%if Trim(strLugarPago)=Trim(direccion) then
				response.Write("Selected")
			end if%>
			><%=trim(rsDIR("LUGAR_PAGO"))%></option>
			<%
			rsDIR.movenext
		loop
		rsDIR.close
		set rsDIR=nothing
		CerrarSCG1()
		%>
	</select>
<%


elseif trim(accion_ajax)="refresca_CB_LUGARPAGO_CP" then
	cliente = request.querystring("cliente_")
	rut		=request.querystring("rut")	
%>
	<select name="CB_LUGARPAGO_CP" id="CB_LUGARPAGO_CP" onchange="this.style.width=330">

		<option value="0">SELECCIONE</option>
		<%
		AbrirSCG1()

		strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & trim(rut) & "' AND ESTADO <> 2"
		
		strSql = strSql & " UNION"

		strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(cliente) & "' ORDER BY ORDEN ASC"


		set rsDIR=Conn1.execute(strSql)
		do until rsDIR.eof
			direccion = rsDIR("ID")&"-"&rsDIR("TIPO")
			%>
			<option value="<%=direccion%>"
			<%if Trim(strLugarPago)=Trim(direccion) then
				response.Write("Selected")
			end if%>
			><%=trim(rsDIR("LUGAR_PAGO"))%></option>
			<%
			rsDIR.movenext
		loop
		rsDIR.close
		set rsDIR=nothing
		CerrarSCG1()
		%>
	</select>

<%

elseif trim(accion_ajax)="refresca_CB_LUGAR_NORM2" then
	cliente = request.querystring("cliente_")
	rut		=request.querystring("rut")	
%>
	<select name="CB_LUGAR_NORM2" id="CB_LUGAR_NORM2" onchange="this.style.width=330">
		<option value="0">SELECCIONE</option>
		<%
		AbrirSCG1()

		strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & trim(rut) & "' AND ESTADO <> 2"
		
		strSql = strSql & " UNION"

		strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(cliente) & "' ORDER BY ORDEN ASC"


		set rsDIR=Conn1.execute(strSql)
		do until rsDIR.eof
			direccion = rsDIR("ID")&"-"&rsDIR("TIPO")
			%>
			<option value="<%=direccion%>"
			<%if Trim(strLugarPago)=Trim(direccion) then
				response.Write("Selected")
			end if%>
			><%=trim(rsDIR("LUGAR_PAGO"))%></option>
			<%
			rsDIR.movenext
		loop
		rsDIR.close
		set rsDIR=nothing
		CerrarSCG1()
		%>
	</select>
<%


elseif trim(accion_ajax)="actualiza_CB_CONTACTO_ASOCIADO_DIRECCION" then
	contentVar =request.queryString("contentVar")

	if trim(contentVar)<>0 then

			strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM DIRECCION_CONTACTO WHERE id_direccion = " & contentVar
			strSql = strSql & " UNION"
			strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "


			set rsContacto = Conn.execute(strSql)

			'response.write "strQuery == " & strSql
			%>
			<select name="CB_CONTACTO_ASOCIADO_DIRECCION" id="CB_CONTACTO_ASOCIADO_DIRECCION">
				<option value="0">SELECCIONE</option>
			<%
			Do While not rsContacto.eof
				strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
			%>
				<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
			<%
				rsContacto.moveNext
			Loop
			rsContacto.close
			set rsContacto=nothing

	%>
			</select>

	<%else%>

		<select name="CB_CONTACTO_ASOCIADO_DIRECCION" id="CB_CONTACTO_ASOCIADO_DIRECCION">
			<option value="0">SELECCIONE</option>
		</select>

	<%end if


elseif trim(accion_ajax)="actualiza_CB_CONTACTO_ASOCIADO_TERRENO" then
	contentVar =request.queryString("contentVar")

	if trim(contentVar)<>0 then

			strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & contentVar
			strSql = strSql & " UNION"
			strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "


			set rsContacto = Conn.execute(strSql)

			'response.write "strQuery == " & strSql
			%>
			<select name="CB_CONTACTO_ASOCIADO_TERRENO" id="CB_CONTACTO_ASOCIADO_TERRENO">
				<option value="0">SELECCIONE</option>
			<%
			Do While not rsContacto.eof
				strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
			%>
				<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
			<%
				rsContacto.moveNext
			Loop
			rsContacto.close
			set rsContacto=nothing

	%>
			</select>

	<%else%>

		<select name="CB_CONTACTO_ASOCIADO_TERRENO" id="CB_CONTACTO_ASOCIADO_TERRENO">
			<option value="0">SELECCIONE</option>
		</select>

	<%end if



elseif trim(accion_ajax)="actualiza_td_CB_FONO_TERRENO" then
	rut		=request.querystring("rut")	
%>
	<select name="CB_FONO_TERRENO" id="CB_FONO_TERRENO" onchange="set_CB_CONTACTO_ASOCIADO_TERRENO(this.value); return false;">
	<option value="0">SELECCIONE</option>
	<%if fono_con="0" or fono_con="" then%>
	  <%
		AbrirSCG1()
		ssql_ = "SELECT ID_TELEFONO, TELEFONO, COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & rut & "' AND ESTADO <> 2"
		set rsFON=Conn1.execute(ssql_)
		Do until rsFON.eof
			strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
			strSel=""
			if strFonoCB = strFonoAgend Then strSel = "SELECTED"
			%>
			<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
			<%
				rsFON.movenext
		Loop
		rsFON.close
		set rsFON=nothing
		CerrarSCG1()
	 %>
	<%else%>
		<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
	<%end if %>
	</select>

<%
end if
%>