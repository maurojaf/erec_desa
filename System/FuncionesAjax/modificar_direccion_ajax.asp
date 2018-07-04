<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->


<%

Response.CodePage = 65001
Response.charset  ="utf-8"


accion_ajax 	= request.queryString("accion_ajax")

rut 			= request.querystring("rut")
cliente 		= request.queryString("cliente_")



If Trim(accion_ajax) = "refresca_CB_LUGARPAGO_CP_RUTA" Then
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
%>
	<select name="CB_DIRECCION_TERRENO" id="CB_DIRECCION_TERRENO" onchange="this.style.width=330">

		<option value="0">SELECCIONE</option>
		<%
		AbrirSCG1()

		strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & trim(rut) & "' AND ESTADO <> 2"
		
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

elseif trim(accion_ajax)="refresca_CB_DIRECCION_TERRENO" then
%>
	<select name="CB_DIRECCION_TERRENO" id="CB_DIRECCION_TERRENO" onchange="this.style.width=330">

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

elseif trim(accion_ajax)="refresca_CB_LUGAR_NORM" then
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
End If

%>

