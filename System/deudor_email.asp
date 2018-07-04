<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"  LCID = 1034%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="sesion_inicio.asp"-->
	<link href="../css/style_generales_sistema.css" rel="stylesheet">	
<%

Response.CodePage = 65001
Response.charset="utf-8"

strRutDeudor 			= request("rut")
strCodCliente 			= request("strCodCliente")
muestra_envio_correo	= request("muestra_envio_correo")

'response.write "strRutDeudor : "& strRutDeudor &"<br>"
'response.write "strCodCliente : "& strCodCliente &"<br>"



%>
</head>
<body>
	  <%
	    abrirscg()
		ssql =" SELECT IdTipoContacto, ID_EMAIL, DM.FECHA_INGRESO, EMAIL, CORRELATIVO, ESTADO, FECHA_REVISION, ANEXO, NOMBRE_DEUDOR " 
		ssql = ssql & " FROM DEUDOR_EMAIL DM "
		ssql = ssql & " INNER JOIN DEUDOR D on D.RUT_DEUDOR=DM.RUT_DEUDOR "
		ssql = ssql & " WHERE DM.RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO IN (0,1)  and d.cod_cliente = " & strCodCliente
		ssql = ssql & " ORDER BY DM.FECHA_INGRESO "

		set rsDIR=Conn.execute(ssql)
		if rsDIR.eof then
		%>

		<input name="muestra_carga_funcion_email" 		id="muestra_carga_funcion_email" 		type="hidden" 	value="N">
		<table width="100%" border="0">

			<tr bordercolor="#FFFFFF" bgcolor="#d0cfd7" height="25">
			<td align="center" class="Estilo10"><b>No existen email válidos o sin auditar</b></td>
			<td align="center" bgcolor="#<%=session("COLTABBG2")%>">
				<a href="#" onClick="envia_email('NE');" title="Nuevo Email"><img src="../imagenes/email_add.png" border="0"></a>
			</td>
			<td align="center" bgcolor="#<%=session("COLTABBG2")%>">
				<a href="#" onClick="envia_email('NV');" title="Ver No válidos"><img src="../imagenes/brick_delete.png" border="0"></a>
			</td>
			</tr>

		</table>

		<%

		Else
	  %>
	<input type="hidden" name="pagina_origen" id="pagina_origen" value="deudor_email">
	&nbsp;
	  <table width="100%" border="0"  class="intercalado" style="width:100%;">
	  	<thead>
	    <tr bordercolor="#FFFFFF" class="Estilo13" bgcolor="#<%=session("COLTABBG")%>">
        	<td ALIGN="CENTER">EMAIL</td>
        	<td ALIGN="CENTER">ANEXO</td>
			<td align = "center">TIPO DE CONTACTO</td>
			<td ALIGN="CENTER">FECHA DE INGRESO</td>
			<td ALIGN="CENTER">FECHA DE AUDITORIA</td>
			<td WIDTH="125" ALIGN = "CENTER">ESTADO</td>
			<td align="center">
				<a href="#" onClick="envia_email('AE');" title="Auditar Email"><img src="../imagenes/email.png" border="0"></a>
				<a href="#" onClick="envia_email('NE');" title="Nuevo Email"><img src="../imagenes/email_add.png" border="0"></a>
				<a href="#" onClick="envia_email('NV');" title="Ver No validos"><img src="../imagenes/email_delete.png" border="0"></a>
			</td>
		</tr>
		</thead>
		<tbody>
		<%
		sinauditar=0
		novalida=0
		valida=0
		do until rsDIR.eof
			FECHA_INGRESO=rsDIR("FECHA_INGRESO")
			if isNULL(FECHA_INGRESO) then
			FECHA_INGRESO=""
			end if
			Email=rsDIR("Email")

			FECHA_REVISION=rsDIR("FECHA_REVISION")
			if isNULL(FECHA_REVISION) then
			FECHA_REVISION=""
			end if

			correlativo_deudor=rsDIR("CORRELATIVO")
			strEstado=Trim(rsDIR("Estado"))

			if strEstado="0" then
				estado_EMAIL="SIN AUDITAR"
			elseif strEstado="1" then
				estado_EMAIL="VALIDO"
			elseif strEstado="2" then
				estado_EMAIL="NO VALIDO"
			end if

			srtAnexo = UCASE(rsDIR("ANEXO"))

			If Trim(srtAnexo) <> "" Then
				srtAnexoMsg = srtAnexo
			Else
				srtAnexoMsg = "Sin información"
			End If


			'REsponse.Write "strEstado=" & strEstado
			'REsponse.Write "estado_EMAIL=" & estado_EMAIL
		%>
		<input type="hidden" id="correlativo_deudor" name="correlativo_deudor" value="<%=trim(correlativo_deudor)%>">
        <tr >
			<% j = 1 %>
			<td title="
					<% 	strLista = "SELECT CONTACTO FROM EMAIL_CONTACTO WHERE RUT_DEUDOR = '"& strRutDeudor &"' AND ID_EMAIL = '"& rsDIR("ID_EMAIL") &"' ORDER BY ID_CONTACTO DESC"
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
					"><%=Email%></td>

			<td title="<%=srtAnexoMsg%>"><div align="CENTER">
				<input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>"  type="text" value="<%=srtAnexo%>" size="30" maxlength="30">
			</td>
			
			<td><select id="cbxTipoContacto_<%=correlativo_deudor%>" name="cbxTipoContacto_<%=correlativo_deudor%>">
				<option value="">Seleccione</option>
				<% 	strListaTipoContacto = "SELECT IdTipoContacto, Glosa, Descripcion FROM TipoContacto WHERE TipoDatoContacto = 'E' AND CodigoCliente = '"& session("ses_codcli") &"'"					
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

          	<td align="CENTER"><%=MID(Cstr(FECHA_INGRESO),1,10)%></td>

        	<td align="CENTER"><%=MID(Cstr(FECHA_REVISION),1,10)%></td>

			<td><div align="right"><span class="Estilo35">
              <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="1"
			  <%if Trim(estado_EMAIL)="VALIDO" then
			   Response.Write("checked")
			   valida=valida+1
			   end if%>>
              VA
			  <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="2"
			  <%if Trim(estado_EMAIL)="NO VALIDO" then
			  Response.Write("checked")
			  novalida=novalida+1
			  end if%>>
			  NV
              <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="0"
			  <%if Trim(estado_EMAIL)="SIN AUDITAR" then
			  Response.Write("checked")
			  sinauditar=sinauditar+1
			  end if%>>
              SA
		    </span>
		    </div>
		   </td>
		   <td align="center">
				<% i = 1 %>
		   		<img src="../imagenes/Agrega_contacto.png" width="20" heigth="20" border="0" title="
					<% 	strLista = "SELECT CONTACTO FROM EMAIL_CONTACTO WHERE RUT_DEUDOR = '"& strRutDeudor &"' AND ID_EMAIL = '"& rsDIR("ID_EMAIL") &"' ORDER BY Fecha_ingreso DESC"
						set rsLista = Conn.execute(strLista)
						if not rsLista.Eof then
							Do While Not rsLista.Eof %>
								<% response.write(i) %> - <%=rsLista("CONTACTO") %></br>
						<% 	rsLista.movenext
							i = i + 1 
							Loop
							else
								response.write("No hay contactos ingresados.")
							
							end if %>" onclick="agrega_contacto_mail('deudor_telefonos','<%=strRutDeudor%>','<%=rsDIR("ID_EMAIL")%>')">

				&nbsp;
				&nbsp;

				<img style="cursor:pointer;" src="../imagenes/bt_email.png" width="20" alt="Envia correo" heigth="20" border="0" onclick="ventana_simulacion_convenio('<%=Email%>','<%=rsDIR("NOMBRE_DEUDOR")%>')">

		   </td>

        </tr>
	<%
	rsDIR.movenext
	loop
	%>
	
        <tr class="totales">
          <td ><span class="">TOTAL</span></td>
		  <td ></td>
		  <td ></td>
          <td ><span class=""></span> V&Aacute;LIDOS : <%=valida%></span></td>
          <td ><span class="">SIN AUDITAR : <%=sinauditar%></span></td>
          <td ><span class="">TOTAL CORREOS : <%=(valida+novalida+sinauditar)%></span></td>
          <td align="center">
			<a href="#" onClick="envia_email('AE');" title="Auditar Email"><img src="../imagenes/email.png" border="0"></a>
			<a href="#" onClick="envia_email('NE');" title="Nuevo Email"><img src="../imagenes/email_add.png" border="0"></a>
			<a href="#" onClick="envia_email('NV');" title="Ver No validos"><img src="../imagenes/email_delete.png" border="0"></a>
			</td>
        </tr>

      </table>
  	</tbody>
	  <%
		end if
		rsDIR.close
		set rsDIR=nothing
		cerrarscg()

	  %>

</body>
</html>

