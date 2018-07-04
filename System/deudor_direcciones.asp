<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="sesion_inicio.asp"-->
	<link rel="stylesheet" href="../css/style_generales_sistema.css">

</head>
<body>
<%

Response.CodePage = 65001
Response.charset="utf-8"

rut = request("rut")
strCodCliente = request("strCodCliente") 

%>

<%


			abrirscg()
			ssql=""
			ssql="SELECT IdTipoContacto, HORA_DESDE, HORA_HASTA,DIAS_PAGO, ID_DIRECCION,Calle,Numero,Comuna,CORRELATIVO,Resto,Estado,FECHA_INGRESO FROM DEUDOR_DIRECCION WHERE ESTADO IN (0,1) AND RUT_DEUDOR='"&rut&"' ORDER BY FECHA_INGRESO"
			'Response.write "ssql=" & ssql
			set rsDIR=Conn.execute(ssql)
			if rsDIR.eof then
			%>
				<table width="100%" border="0">
					<tr bordercolor="#FFFFFF" bgcolor="#d0cfd7" height="25">
					<td align="center" class="Estilo10"><b>No existen direcciones válidas o sin auditar</b></td>

					 <td align="center" bgcolor="#<%=session("COLTABBG2")%>">
					 	<a href="#" onClick="envia_direccion('ND');" title="Nueva Dirección"><img src="../imagenes/brick_add.png" border="0"></a>
					 </td>
					 <td align="center" bgcolor="#<%=session("COLTABBG2")%>">
						<a href="#" onClick="envia_direccion('NV');" title="Ver No validas"><img src="../imagenes/brick_delete.png" border="0"></a>
					 </td>

					</tr>

				</table>
			<%
			Else
			 %>
		<input type="hidden" name="pagina_origen" id="pagina_origen" value="deudor_direcciones">
		  <table width="100%" border="0" bordercolor="#000000" class="intercalado" style="width:100%;">
		  	<thead>
			<tr bordercolor="#FFFFFF" >
				<td></td>
				<td ALIGN="CENTER" width="200">DIRECCION</td>
				<td align="center">RESTO</td>
				<td align = "center">TIPO DE CONTACTO</td>
				<td align="center">DIAS DE PAGO</td>
				<td align="center">HORARIO PAGO</td>
				<td align="center">ESTADO</td>
				<td>
					<a href="#" onClick="envia_direccion('AD');" title="Auditar Direcciones"><img src="../imagenes/brick.png" border="0"></a>
					<a href="#" onClick="envia_direccion('ND');" title="Nueva Dirección"><img src="../imagenes/brick_add.png" border="0"></a>
					<a href="#" onClick="envia_direccion('NV');" title="Ver No validas"><img src="../imagenes/brick_delete.png" border="0"></a>
			  	</td>
	        </tr>
	    	</thead>
			<%
			sinauditar=0
			novalida=0
			valida=0
			do until rsDIR.eof
				intId = rsDIR("ID_DIRECCION")
				FECHA_REVISION=rsDIR("FECHA_INGRESO")
				if isNULL(FECHA_REVISION) then
					FECHA_REVISION=""
				end if
				calle_deudor=rsDIR("Calle")
				numero_deudor=rsDIR("Numero")
				comuna_deudor=rsDIR("Comuna")
				correlativo_deudor=rsDIR("CORRELATIVO")

				strResto=UCASE(rsDIR("RESTO"))
				If Trim(strResto) = "" Then
					strLabelResto = "Sin Información"
				Else
					strLabelResto = strResto
				End If

				Estado=rsDIR("Estado")
				if estado="0" then
					estado_direccion="SIN AUDITAR"
				elseif estado="1" then
					estado_direccion="VALIDA"
				elseif estado="2" then
					estado_direccion="NO VALIDA"
				end if

				strHoraDesde=Trim(rsDIR("HORA_DESDE"))
				strHoraHasta=Trim(rsDIR("HORA_HASTA"))
				strDiasPago=Trim(rsDIR("DIAS_PAGO"))

				strDireccion = calle_deudor & " " & numero_deudor & " " & Comuna_deudor
				strDireccion = Trim(strDireccion)


				strDireccion_geo = replace(ucase(strDireccion),"CALLEJON","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"CALLE","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"POBLACION","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"POBLACIÓN","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PASAJE","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"AV.","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PJE.","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PSJE","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PGE","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIDA","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"CAYE","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"CALLLE","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIDAS","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIA","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"V.","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"AVDA","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PASAGE","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PARCELA","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PARC.","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PARCELAS","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PARSELA","")
				strDireccion_geo = replace(ucase(strDireccion_geo),"PARS.","")


			%>
			<input type="hidden" id="correlativo_deudor" name="correlativo_deudor" value="<%=correlativo_deudor%>">
	        <tr>
	        	<td width="20" align="center">
	        		<img width="20" style="cursor:pointer;" onclick="bt_geolocalizacion('<%=trim(strDireccion_geo)%>')" height="20" src="../Imagenes/map.png" title="Consulta dirección mapa <%=Mid(strDireccion,1,30)%>">
	        	</td>
	          	<td>
	          		<% j = 1 %>
	          		<span title="
					<% 	strLista = "SELECT CONTACTO FROM DIRECCION_CONTACTO WHERE RUT_DEUDOR = '"& RUT &"' AND ID_DIRECCION = '"& rsDIR("ID_DIRECCION") &"' ORDER BY Fecha_ingreso DESC"
						set rsLista = Conn.execute(strLista)
						if not rsLista.Eof then
							Do While Not rsLista.Eof %>
								<% response.write(j) %> - <%=rsLista("CONTACTO") %></br>
						<% 	rsLista.movenext
							j = j + 1 
							Loop
							else
								response.write("No hay contactos ingresados.")
							end if %>
					"><%=Mid(strDireccion,1,30)%></span>
	          	</td>

	          	<td title="<%=strLabelResto%>"><div align="CENTER"><input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=strResto%>" size="30" maxlength="50"></td>
				<%

				strChequedLu = ""
				strChequedMa = ""
				strChequedMi = ""
				strChequedJu = ""
				strChequedVi = ""
				strChequedSa = ""

				If instr(strDiasPago,"LU") > 0 Then strChequedLu = "CHECKED"
				If instr(strDiasPago,"MA") > 0 Then strChequedMa = "CHECKED"
				If instr(strDiasPago,"MI") > 0 Then strChequedMi = "CHECKED"
				If instr(strDiasPago,"JU") > 0 Then strChequedJu = "CHECKED"
				If instr(strDiasPago,"VI") > 0 Then strChequedVi = "CHECKED"
				If instr(strDiasPago,"SA") > 0 Then strChequedSa = "CHECKED"
				%>
				
			<td><select id="cbxTipoContacto_<%=correlativo_deudor%>" name="cbxTipoContacto_<%=correlativo_deudor%>">
				<option value="">Seleccione</option>
				<% 	strListaTipoContacto = "SELECT IdTipoContacto, Glosa, Descripcion FROM TipoContacto WHERE TipoDatoContacto = 'D' AND CodigoCliente = '"& strCodCliente &"'"					
					set rsListaTipoContacto = Conn.execute(strListaTipoContacto)
					i = 1
					
					Do While Not rsListaTipoContacto.Eof
					if(rsListaTipoContacto("IdTipoContacto") = rsDIR("IdTipoContacto")) THEN strSeleccionado = "selected" else strSeleccionado="" end if %>
						<option value="<%=rsListaTipoContacto("IdTipoContacto") %>" <%=strSeleccionado %> title="<%=rsListaTipoContacto("Descripcion") %>">
							<% response.write(i) %> - <%=rsListaTipoContacto("Glosa") %>
						</option>
				<% 	rsListaTipoContacto.movenext
					i = i + 1 
					Loop %>
			</select></td>
					
				<td>
				Lu
				<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="LU" <%=strChequedLu%>>
				Ma
				<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="MA" <%=strChequedMa%>>
				Mi
				<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>"value="MI" <%=strChequedMi%>>
				Ju
				<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="JU" <%=strChequedJu%>>
				Vi
				<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="VI" <%=strChequedVi%>>
				Sa
				<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" ID="CH_DIAS_<%=correlativo_deudor%>" value="SA" <%=strChequedSa%>>
	            </td>

	          	<td align = "center">
	          		Desde
	          		<input name="TX_DESDE_<%=correlativo_deudor%>" id="TX_DESDE_<%=correlativo_deudor%>" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
	          		Hasta
					<input name="TX_HASTA_<%=correlativo_deudor%>" id="TX_HASTA_<%=correlativo_deudor%>" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
				</td>

		      <td align="center">
	              <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="1"
				  <%if estado_direccion="VALIDA" then
				   Response.Write("checked")
				   valida=valida+1
				   end if%>>
	              VA
				  <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="2"
				  <%if estado_direccion="NO VALIDA" then
				  Response.Write("checked")
				  novalida=novalida+1
				  end if%>>
				  NV
	              <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="0"
				  <%if estado_direccion="SIN AUDITAR" then
				  Response.Write("checked")
				  sinauditar=sinauditar+1
				  end if%>>
	              SA
			   </td>
			   <td align="CENTER">
					<% i = 1 %>
			   		<img src="../imagenes/Agrega_contacto.png" style="cursor:pointer;" border="0" title="
					<% 	strLista = "SELECT CONTACTO FROM DIRECCION_CONTACTO WHERE RUT_DEUDOR = '"& RUT &"' AND ID_DIRECCION = '"& rsDIR("ID_DIRECCION") &"' ORDER BY Fecha_ingreso DESC"
						set rsLista = Conn.execute(strLista)
						if not rsLista.Eof then
							Do While Not rsLista.Eof %>
								<% response.write(i) %> - <%=rsLista("CONTACTO") %></br>
						<% 	rsLista.movenext
							i = i + 1 
							Loop
							else
								response.write("No hay contactos ingresados.")
							
							end if %>" onclick="agrega_direccion('deudor_direcciones','<%=rut%>','<%=rsDIR("ID_DIRECCION")%>')">
			   	</td>

	        </tr>
		<%
		rsDIR.movenext
		loop
		   %>


	        <tr class="totales">
	          <td colspan="2"><span class="">TOTALES :</span></td>
			  <td colspan="2">V&Aacute;LIDAS : <%=valida%></span></td>
	          <td colspan="1"><span class="">SIN AUDITAR : <%=sinauditar%></span></td>
	          <td colspan="2"><span class="">TOTAL DIRECCIONES : <%=(valida+novalida+sinauditar)%></span></td>

	          <td >
					<a href="#" onClick="envia_direccion('AD');" title="Auditar Direcciones"><img src="../imagenes/brick.png" border="0"></a>
					<a href="#" onClick="envia_direccion('ND');" title="Nueva Dirección"><img src="../imagenes/brick_add.png" border="0"></a>
					<a href="#" onClick="envia_direccion('NV');" title="Ver No validas"><img src="../imagenes/brick_delete.png" border="0"></a>
			  </td>


	        </tr>

		     <%
				end if
				rsDIR.close
				set rsDIR=nothing
				cerrarscg()

			  %>
		</table>
</body>
</html>


