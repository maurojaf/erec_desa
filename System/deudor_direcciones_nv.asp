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

<%

Response.CodePage 	=65001
Response.charset	="utf-8"

rut 			= request.querystring("rut")
strCodCliente 	= request.querystring("strCodCliente")

%>


</head>
<body>
<table width="100%" border="0">
	<tr>
	<td valign="top" colspan="2">
	  <%

		abrirscg()
		ssql=""
		ssql="SELECT IdTipoContacto,HORA_DESDE, HORA_HASTA,DIAS_PAGO, ID_DIRECCION,Calle,Numero,Comuna,CORRELATIVO,Resto,Estado,FECHA_INGRESO FROM DEUDOR_DIRECCION WHERE ESTADO = 2 AND RUT_DEUDOR='"&rut&"' ORDER BY FECHA_INGRESO"
		'Response.write "ssql=" & ssql
		set rsDIR=Conn.execute(ssql)

		if rsDIR.eof then
			%>
				<script>
					alert('No existen direcciones no validas');
					carga_funcion_direccion()
				</script>
			<%
				Response.End()
		Else
	  %>

	  <input type="hidden" name="pagina_origen" id="pagina_origen" value="deudor_direcciones_nv">
	  <table width="100%" border="0" bordercolor="#000000" class="intercalado" style="width:100%;">
	  	<thead>
		<tr bordercolor="#FFFFFF">
		  <td ALIGN = "CENTER" Width = "200">DIRECCION</td>
		  <td ALIGN = "CENTER">RESTO</td>
		  <td align = "center">TIPO CONTACTO</td>
		  <td ALIGN = "CENTER">DIAS DE PAGO</td>
		  <td colspan="1" ALIGN = "CENTER">HORARIO DE PAGO</td>
		  <td colspan="1" ALIGN = "CENTER">ESTADO</td>
		  <td align="center">
			<a href="#" onClick="envia_direccion('AD');" title="Auditar Direcciones"><img src="../imagenes/brick.png" border="0"></a>
			<a href="#" onClick="envia_direccion('ND');" title="Nueva Direccion"><img src="../imagenes/brick_add.png" border="0"></a>
			<a href="#" onClick="carga_funcion_direccion();" title="Volver"><img src="../imagenes/arrow_left.png" border="0"></a>
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

			strDireccion = calle_deudor & " " & numero_deudor & " " & comuna_deudor
			strDireccion = Trim(strDireccion)
		%>
		<input type="hidden" id="correlativo_deudor" name="correlativo_deudor" value="<%=correlativo_deudor%>">
		<tr>
			<td><acronym title="<%=strDireccion%>"><%=Mid(strDireccion,1,35)%></acronym></td>
          	<td title="<%=strLabelResto%>">
          		<div align="CENTER">
          			<input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=strResto%>" size="30" maxlength="50"></div>
          	</td>
			
			<td><select id="cbxTipoContacto_<%=correlativo_deudor%>" name="cbxTipoContacto_<%=correlativo_deudor%>">
				<% if(rsDIR("IdTipoContacto") <> "") THEN strSeleccionado = "selected" else strSeleccionado="" end if %>
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
			<td align="center">
			Lu
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="LU" <%=strChequedLu%>>
			Ma
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MA" <%=strChequedMa%>>
			Mi
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MI" <%=strChequedMi%>>
			Ju
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="JU" <%=strChequedJu%>>
			Vi
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="VI" <%=strChequedVi%>>
			Sa
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="SA" <%=strChequedSa%>>
			</td>

			<td align = "center">
				Desde
				<input name="TX_DESDE_<%=correlativo_deudor%>" id="TX_DESDE_<%=correlativo_deudor%>" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
				Hasta
				<input name="TX_HASTA_<%=correlativo_deudor%>" id="TX_HASTA_<%=correlativo_deudor%>" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
			</td>

			<td><div align="right"><span class="Estilo35">
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
			</span></div></td>

			<td align="CENTER">
				<img src="../imagenes/Agrega_contacto.png" style="cursor:pointer;" border="0" onclick="agrega_direccion('deudor_direcciones','<%=rut%>','<%=rsDIR("ID_DIRECCION")%>')">
			</td>

		</tr>
	<%
	rsDIR.movenext
	loop
	   %>


		<tr class="totales">
		  <td colspan="1"><span class="">TOTALES :</span></td>
		  <td colspan="1"><span class="">NO V&Aacute;LIDAS : <%=novalida%></span></td>
		  <td colspan="3"><span class="">&nbsp;</td>
		  <td colspan="1"><span class="" COLSPAN=2>&nbsp;</td>
		  <td >
			<a href="#" onClick="envia_direccion('AD');" title="Auditar Direcciones"><img src="../imagenes/brick.png" border="0"></a>
			<a href="#" onClick="envia_direccion('ND');" title="Nueva Direccion"><img src="../imagenes/brick_add.png" border="0"></a>
			<a href="#" onClick="carga_funcion_direccion();" title="Volver"><img src="../imagenes/arrow_left.png" border="0"></a>
		  </td>

		</tr>

	  </table>



	  <%
		end if

		rsDIR.close
		set rsDIR=nothing
		cerrarscg()
	%>
	</td>
	</tr>


</table>
</body>
</html>


