<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->


<%

Response.CodePage = 65001
Response.charset  ="utf-8"


accion_ajax 	= request.queryString("accion_ajax")
strRutDeudor 	=session("session_RUT_DEUDOR")
'Response.write accion_ajax

AbrirSCG()

If Trim(accion_ajax) = "actualiza_td_CB_ID_CONTACTO_GESTION" Then

	ID_CONTACTO_GESTION 	=request.queryString("ID_CONTACTO_GESTION")
	MEDIO_ASOCIADO			=request.queryString("MEDIO_ASOCIADO")

%>

	<select name="CB_ID_CONTACTO_GESTION" id="CB_ID_CONTACTO_GESTION">
		<option value="">SELECCIONE</option>
		<%if trim(MEDIO_ASOCIADO)=1 then

			strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & ID_CONTACTO_GESTION
			strSql = strSql & " UNION"
			strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "
			set rsContacto = conn.execute(strSql)
			
			Do While not rsContacto.eof
				strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
			%>
				<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
			<%
			rsContacto.moveNext
			Loop

		end if

		if trim(MEDIO_ASOCIADO)=2 then
			
			strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM EMAIL_CONTACTO WHERE ID_EMAIL = " & ID_CONTACTO_GESTION
			strSql = strSql & " UNION"
			strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "
			set rsContacto = conn.execute(strSql)
			
			Do While not rsContacto.eof
				strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
			%>
				<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
			<%
			rsContacto.moveNext
			Loop

		end if%>

		<%if trim(MEDIO_ASOCIADO)=3 then

			strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM DIRECCION_CONTACTO WHERE Id_Direccion = " & ID_CONTACTO_GESTION
			strSql = strSql & " UNION"
			strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "
			set rsContacto = conn.execute(strSql)
			
			Do While not rsContacto.eof
				strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
			%>
				<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
			<%
			rsContacto.moveNext
			Loop

		end if%>

	</select>


<%elseif trim(accion_ajax)="actualiza_td_CB_ID_CONTACTO_FONO_COBRO" then
	ID_FONO_COBRO 	=request.queryString("ID_FONO_COBRO")
	
	
%>
	<select name="CB_ID_CONTACTO_FONO_COBRO" id="CB_ID_CONTACTO_FONO_COBRO">

	<option value="">SELECCIONE</option>
	<%if ID_FONO_COBRO<>"0" AND ID_FONO_COBRO<>"" then

		strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & ID_FONO_COBRO
		strSql = strSql & " UNION"
		strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "
		set rsContacto = conn.execute(strSql)
		
		Do While not rsContacto.eof
			strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
		%>
			<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
		<%
		rsContacto.moveNext
		Loop

	end if %>
	</select>

<%


elseif trim(accion_ajax)="actualiza_ID_MEDIO_AGENDAMIENTO" then
	strRutDeudor		=request.queryString("rut")
	MEDIO_ASOCIADO		=request.queryString("MEDIO_ASOCIADO")
%>
	<SELECT NAME="CB_ID_MEDIO_AGENDAMIENTO" id="CB_ID_MEDIO_AGENDAMIENTO">
		<OPTION VALUE="" >SELECCIONE</OPTION>
		<%AbrirSCG1()
		if trim(MEDIO_ASOCIADO)="1" then
			
			ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRutDeudor & "' AND ESTADO <> 2 "
			set rsFON=Conn1.execute(ssql_)
			Do While not rsFON.eof
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
		
		elseif trim(MEDIO_ASOCIADO)="2" then

			ssql_ = "SELECT ID_EMAIL, UPPER(EMAIL) AS EMAIL FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2 "
			set rsEmail=Conn1.execute(ssql_)
			Do While not rsEmail.eof
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

		elseif trim(MEDIO_ASOCIADO)="3" then

			strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2"

			set rsDIR=Conn1.execute(strSql)
			do While not  rsDIR.eof
				direccion = rsDIR("LUGAR_PAGO")
				%>
				<option value="<%=rsDIR("ID")%>"
				<%if Trim(strLugarPago)=Trim(direccion) then
					response.Write("Selected")
				end if%>
				><%=direccion%></option>
				<%
				rsDIR.movenext
			loop
			rsDIR.close
			set rsDIR=nothing

		end if
		CerrarSCG1()%>
	</SELECT>

<%elseif trim(accion_ajax)="actualiza_ID_MEDIO_GESTION" then
	rut					=request.queryString("rut")
	MEDIO_ASOCIADO		=request.queryString("MEDIO_ASOCIADO")
	Forma_Pago			=request.queryString("Forma_Pago")
%>
	<select name="CB_ID_MEDIO_GESTION" id="CB_ID_MEDIO_GESTION"  onchange="set_CB_ID_CONTACTO_GESTION(<%=trim(MEDIO_ASOCIADO)%>, this.value)">	
	<option value="">SELECCIONE</option>
		<%AbrirSCG1()
		if trim(MEDIO_ASOCIADO)="1" then
			
			ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRutDeudor & "' AND ESTADO <> 2 "
			set rsFON=Conn1.execute(ssql_)
			Do While not rsFON.eof
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
		
		elseif trim(MEDIO_ASOCIADO)="2" then

			ssql_ = "SELECT ID_EMAIL, UPPER(EMAIL) AS EMAIL FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2 "
			set rsEmail=Conn1.execute(ssql_)
			Do While not rsEmail.eof
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

		elseif trim(MEDIO_ASOCIADO)="3" then

			strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO <> 2 "

			set rsDIR=Conn1.execute(strSql)
			do While not rsDIR.eof
				direccion = rsDIR("LUGAR_PAGO")
				%>
				<option value="<%=rsDIR("ID")%>"
				<%if Trim(strLugarPago)=Trim(direccion) then
					response.Write("Selected")
				end if%>
				><%=direccion%></option>
				<%
				rsDIR.movenext
			loop
			rsDIR.close
			set rsDIR=nothing

		end if
		CerrarSCG1()%>
	</select>
<%
elseif trim(accion_ajax)="actualiza_ID_DIRECCION_COBRO_DEUDOR" then
	strRutDeudor		=request.queryString("rut")
	MEDIO_ASOCIADO		=request.queryString("MEDIO_ASOCIADO")
    strFormaPago		=request.queryString("Forma_Pago")
    strCodCliente       =request.queryString("strCodCliente") 
%>
<select name="CB_ID_DIRECCION_COBRO_DEUDOR" id="CB_ID_DIRECCION_COBRO_DEUDOR">
	<option value="">SELECCIONE</option>
	<%
	AbrirSCG1()

    strSql ="select top 0 1"

				if strFormaPago = "undefined" then
				
				   strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , " & _
                                 " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                 " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2 "
                 end if 							  
	
	
                IF  strFormaPago = "TR" or  strFormaPago ="DP" THEN '' TRANSFERENCIA/DEPOSITO
                     strSql = " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                              " ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                              " FROM FORMA_RECAUDACION " & _
                                              " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  = 'DEPOSITO' " & _
                                              " ORDER BY ORDEN ASC"
                END IF 

                IF  strFormaPago = "CF" or  strFormaPago ="CD"  or  strFormaPago ="EF" or strFormaPago = "CF" THEN  '' CHEQUE A FECHA/CHEQUE AL DIA/EFECTIVO/CF

                        strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , " & _
                                 " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                 " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2 "
                    
				        strSql = strSql & " UNION"

                        strSql = strSql  & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                      "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                      " FROM FORMA_RECAUDACION " & _
                                      " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  = 'PRESENCIAL' " & _
                                      " ORDER BY ORDEN ASC"
                    END IF 


                    IF  strFormaPago = "VV" or  strFormaPago ="PG"  or  strFormaPago ="LT" THEN '' VALE VISTA/PAGARE/LETRA

                                            strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , " & _
                                                     " 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo " & _
                                                     " FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2"
                    
				                            strSql = strSql & " UNION"

                                            strSql = strSql  & " SELECT 2 AS TIPO, NOMBRE + ' ' + UBICACION AS LUGAR_PAGO , ORDEN , " & _
                                                          "ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' TIPO " & _
                                                          " FROM FORMA_RECAUDACION " & _
                                                          " WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' AND TIPO  in ('PRESENCIAL','RETIRO') " & _
                                                          " ORDER BY ORDEN ASC"
                    END IF

	'strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRutDeudor) & "' AND ESTADO <> 2 "

	'strSql = strSql & " UNION"


	'strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' tipo FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(strCodCliente) & "' ORDER BY ORDEN ASC"

	set rsDIR=Conn1.execute(strSql)
	do While not  rsDIR.eof
		direccion = rsDIR("LUGAR_PAGO")
		%>
		<option value="<%=rsDIR("ID")&"-"&rsDIR("TIPO")%>"
		<%if Trim(strLugarPago)=Trim(direccion) then
			response.Write("Selected")
		end if%>
		><%=direccion%></option>
		<%
		rsDIR.movenext
	loop
	rsDIR.close
	set rsDIR=nothing
	CerrarSCG1()
	%>
	</select>
<%
elseif trim(accion_ajax)="actualiza_ID_FONO_COBRO" then

	strRutDeudor		=request.queryString("rut")
	MEDIO_ASOCIADO		=request.queryString("MEDIO_ASOCIADO")
%>	

	<select name="CB_ID_FONO_COBRO" id="CB_ID_FONO_COBRO" onchange="set_CB_ID_CONTACTO_FONO_COBRO(this.value);">

	<option value="">SELECCIONE</option>
	<%if fono_con="0" or fono_con="" then%>
	  <%
		AbrirSCG1()
		ssql_ = "SELECT ID_TELEFONO, TELEFONO, COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRutDeudor & "' AND ESTADO <> 2 "
		set rsFON=Conn1.execute(ssql_)
		Do While not rsFON.eof
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


END IF%>

