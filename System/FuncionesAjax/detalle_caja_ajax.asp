<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="../arch_utils.asp"-->

<%

Response.CodePage = 65001
Response.charset  ="utf-8"


accion_ajax 	= request.queryString("accion_ajax")
'Response.write accion_ajax

AbrirSCG()

If Trim(accion_ajax) = "refresca_campo_nro_boleta" Then
	intIdPago 		=request.queryString("id_pago")
	intNroBoleta 	=request.queryString("NRO_BOLETA")
%>
	<input type="text" maxlength="7" name="NRO_BOLETA" id="NRO_BOLETA" style="width:50px;" value="<%=trim(intNroBoleta)%>" onblur="actuliza_nro_boleta('<%=trim(intIdPago)%>', this.value,'<%=intNroBoleta%>')">
<%

elseif trim(accion_ajax)="actualiza_campo_nro_boleta" then
	intIdPago 		=request.queryString("id_pago")
	intNroBoleta 	=request.queryString("NRO_BOLETA")

	if trim(intNroBoleta)="" then
		intNroBoleta ="NULL"
	else
		intNroBoleta ="'"&intNroBoleta&"'"
	end if

	SQL_UPDATE ="UPDATE CAJA_WEB_EMP "
	SQL_UPDATE = SQL_UPDATE & " SET NRO_BOLETA ="&TRIM(intNroBoleta)
	SQL_UPDATE = SQL_UPDATE & " WHERE ID_PAGO='"&TRIM(intIdPago)&"'"
	conn.execute(SQL_UPDATE)

	strSql = "UPDATE EXP_CONTABILIDAD "
	strSql = strSql & " SET NRO_BOLETA ="&TRIM(intNroBoleta)
	strSql = strSql & " WHERE ID_PAGO = '"&TRIM(intIdPago)&"'"
	set strSql = Conn.execute(strSql)

	SQL_SEL ="SELECT COUNT(*) cantidad "	
	SQL_SEL = SQL_SEL & " FROM CAJA_WEB_EMP_DOC_PAGO "
	SQL_SEL = SQL_SEL & " WHERE TIPO_PAGO=1 AND ID_PAGO=" & TRIM(intIdPago)
	set rs_sel = conn.execute(SQL_SEL)

	if not rs_sel.eof then
		cantidad =rs_sel("cantidad")
	else
		cantidad =0
	end if

	if trim(intNroBoleta)="NULL" and cantidad>0 THEN
		intNroBoleta 	=""
		color 		="background-color:#F5A9A9;"
	else
		color 		=""
	END IF

	if trim(intNroBoleta)="NULL" THEN
		intNroBoleta 	=""
	END IF


	%>
	<div <%if cantidad>0  then%> onclick="refresca_nro_boleta('<%=intIdPago%>','<%=REPLACE(intNroBoleta,"'","")%>')" <%end if%> style="cursor:pointer; width:100%;height:20px; <%=trim(color)%>">
		<%=REPLACE(intNroBoleta,"'","")%>&nbsp;&nbsp;
	</div>

<%

elseif trim(accion_ajax)="modifica_numero_cheque" then
	intIdPago =request.queryString("id_pago")

	sql_sel ="SELECT correlativo, ID_PAGO, MONTO, RUT_CHEQUE, VENCIMIENTO, COD_BANCO, NRO_CHEQUE, CODIGO_PLAZA, " 
	sql_sel = sql_sel  & " OBSERVACIONES, NRO_CHEQUE " 
	sql_sel = sql_sel  & " FROM CAJA_WEB_EMP_DOC_PAGO CWP " 
	sql_sel = sql_sel  & " WHERE FORMA_PAGO IN ('CD','CF')  " 
	sql_sel = sql_sel  & " AND ID_PAGO=" & trim(intIdPago)
	sql_sel = sql_sel  & " order by correlativo asc " 
	set rs_sel = conn.execute(sql_sel)
	'Response.write sql_sel

		if not rs_sel.eof then%>
			<table border="0" align="center" class="intercalado" style="width:100%">
				<thead>
				<tr>
					<td align="left" width="80">ID_PAGO</td>
					<td align="left" width="80">MONTO</td>
					<td align="left" width="80">RUT_CHEQUE</td>
					<td align="left" width="80">VENCIMIENTO</td>
					<td align="left" width="80">COD_BANCO</td>
					<td align="left" width="80">CODIGO_PLAZA</td>
					<td align="left" width="100">OBSERVACIONES</td>
					<td align="left" WIDTH="200">NRO_CHEQUE</td>
				</tr>				
				</thead>
				<tbody>


			<%	do while not rs_sel.eof 
					correlativo 	=rs_sel("correlativo")
					intIdPago		=rs_sel("ID_PAGO")
					MONTO			=rs_sel("MONTO")
					RUT_CHEQUE		=rs_sel("RUT_CHEQUE")
					VENCIMIENTO		=rs_sel("VENCIMIENTO")
					COD_BANCO		=rs_sel("COD_BANCO")
					NRO_CHEQUE		=rs_sel("NRO_CHEQUE")
					CODIGO_PLAZA	=rs_sel("CODIGO_PLAZA")
					OBSERVACIONES	=rs_sel("OBSERVACIONES")
					NRO_CHEQUE		=rs_sel("NRO_CHEQUE")

		%>
					<tr>
						<td align="left"><%=intIdPago%></td>
						<td align="left"><%=MONTO%></td>
						<td align="left"><%=RUT_CHEQUE%></td>
						<td align="left"><%=VENCIMIENTO%></td>
						<td align="left"><%=COD_BANCO%></td>
						<td align="left"><%=CODIGO_PLAZA%></td>
						<td align="left"><%=OBSERVACIONES%></td>
						<td align="left">
							<input type="text" style="width:100px;" name="NRO_CHEQUE_<%=correlativo%>_<%=intIdPago%>" ID="NRO_CHEQUE_<%=correlativo%>_<%=intIdPago%>" maxlength="23" value="<%=NRO_CHEQUE%>">
							<input type="hidden" name="cheque_correlativo_<%=intIdPago%>" id="cheque_correlativo_<%=intIdPago%>" value="<%=trim(correlativo)%>">
						</td>
					</tr>				

				<%
				rs_sel.movenext
				loop
				%>
				<tr>
					<td align="right" colspan="8"><br><input type="button" class="fondo_boton_100" onclick="modificar_nro_cheque('<%=intIdPago%>')" value="Guardar"></td>
				</tr>
				</table>
				</tbody>
			<%
		else
			%>
				Sin datos
			<%
		end if%>
	
		<div id="guarda_numero_cheque_<%=intIdPago%>"></div>
	
	
<%
elseif trim(accion_ajax)="guarda_numero_cheque" then
	intIdPago 			=request.queryString("id_pago")
	valor_nro_cheque  	=request.queryString("valor_nro_cheque")
	contador 			=request.queryString("contador")

	'Response.write valor_nro_cheque&"<br>"

	registros 			=split(valor_nro_cheque,"*")
	total				=ubound(registros)

	For indice = 0 to contador 
		if trim(indice)<>0 then		

			SQL_UPDATE ="UPDATE CAJA_WEB_EMP_DOC_PAGO "
			SQL_UPDATE = SQL_UPDATE & " SET NRO_CHEQUE ='"&TRIM(registros(indice-1))&"'"
			SQL_UPDATE = SQL_UPDATE & " WHERE ID_PAGO='"&TRIM(intIdPago)&"' AND correlativo ="&indice	
		
			'Response.write SQL_UPDATE&"<br>"
			conn.execute(SQL_UPDATE)	

		end if

	next


END IF%>



