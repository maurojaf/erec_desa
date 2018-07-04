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
	<link href="../css/style_generales_sistema.css" rel="stylesheet">		
<%

Response.CodePage = 65001
Response.charset="utf-8"

strRutDeudor 			= request("rut")
strCodCliente 			= request("strCodCliente")

%>
</head>
<body>
<table width="100%" border="0">
   <tr>
    <td valign="top" background="">
	  <%
	    abrirscg()
		strSql="SELECT IdTipoContacto,ID_EMAIL,FECHA_INGRESO,EMAIL,CORRELATIVO,ESTADO,FECHA_REVISION,ANEXO FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND ESTADO = 2 ORDER BY FECHA_INGRESO"

		'Response.write "<br>strSql=" & strSql
		set rsDIR=Conn.execute(strSql)
		if rsDIR.eof then
		%>
			<script>
				alert('No existen email no validos');
				carga_funcion_email()
			</script>
		<%
			Response.End
		Else
	  %>
	  <input type="hidden" name="pagina_origen" id="pagina_origen" value="deudor_email_nv">
	  <table width="100%" border="0" bordercolor="#FFFFFF" class="intercalado" style="width:100%;">
	  &nbsp;
	  	<thead>
        <tr bordercolor="#FFFFFF" class="Estilo13">
        	<td ALIGN="CENTER">DIRECCI&Oacute;N DE CORREO </td>
        	<td ALIGN="CENTER">ANEXO</td>
			<td align = "center">TIPO CONTACTO</td>
			<td ALIGN="CENTER">FECHA DE INGRESO</td>
			<td ALIGN="CENTER">FECHA DE AUDITORIA</td>
			<td WIDTH="125" ALIGN = "CENTER">ESTADO</td>
			<td align="center">
				<a href="#" onClick="envia_email('AE')" title="Auditar Email"><img src="../imagenes/email.png" border="0"></a>
				<a href="#" onClick="envia_email('NE')" title="Nuevo Email"><img src="../imagenes/email_add.png" border="0"></a>
				<a href="#" onClick="carga_funcion_email();" title="Volver"><img src="../imagenes/arrow_left.png" border="0"></a>
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

        <tr bordercolor="#FFFFFF">
			<td><%=Email%></td>

			<td title="<%=srtAnexoMsg%>"><div align="CENTER">
				<input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=srtAnexo%>" size="30" maxlength="30">
			</td>
			
			<td><select id="cbxTipoContacto_<%=correlativo_deudor%>" name="cbxTipoContacto_<%=correlativo_deudor%>">
				<% if(rsDIR("IdTipoContacto") <> "") THEN strSeleccionado = "selected" else strSeleccionado="" end if %>
				<option value="">Seleccione</option>
				
				<% 	strListaTipoContacto = "SELECT IdTipoContacto, Glosa, Descripcion FROM TipoContacto WHERE TipoDatoContacto = 'E' AND CodigoCliente = '"& strCodCliente &"'"					
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
	   
	   <td align="center"><img style="cursor:pointer;" src="../imagenes/Agrega_contacto.png" border="0" onclick="agrega_contacto_mail('deudor_telefonos','<%=strRutDeudor%>','<%=rsDIR("ID_EMAIL")%>')"></td>

        </tr>
	<%
	rsDIR.movenext
	loop
	   %>
        <tr class="totales">
          <td ><span class="">TOTAL</span></td>
          <td ><span class=""></span> NO V&Aacute;LIDOS : <%=novalida%></span></td>
          <td ><span class="">TOTAL CORREOS : <%=(novalida)%></span></td>
          <td COLSPAN=3>&nbsp;</td>
          <td align="center">
			<a href="#" onClick="envia_email('AE')" title="Auditar Email"><img src="../imagenes/email.png" border="0"></a>
			<a href="#" onClick="envia_email('NE')" title="Nuevo Email"><img src="../imagenes/email_add.png" border="0"></a>
			<a href="#" onClick="carga_funcion_email();" title="Volver"><img src="../imagenes/arrow_left.png" border="0"></a>
			</td>
        </tr>
    </tbody>
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