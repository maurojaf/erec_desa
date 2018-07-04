<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/style_generales_sistema.css">    
<%
Response.CodePage 	=65001
Response.charset	="utf-8"

rut = request("rut")
strOrigen = request("strOrigen")
%>

	<% If strOrigen = "" Then %>
		<!--#include file="sesion.asp"-->
	<% Else %>
		<!--#include file="sesion_inicio.asp"-->
	<% End If %>
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->


</head>
<body>
<input name="strOrigen" type="hidden" id="strOrigen" value="<%=strOrigen%>">

<table <% If strOrigen = "" Then %> width="90%" <%else%> width="100%" <%end if%>border="0" align="center">

	<% If strOrigen = "" Then %>
		<tr>
			<TD width="100%" ALIGN=LEFT class="titulo_informe">
				<B>Nuevo Email</B>
			</TD>
		</tr>
		<tr>
			<TD width="100%" ALIGN="LEFT">&nbsp;</TD>
		</tr>		
	<% End If %>

  <tr>

    <td valign="top">
    <table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">
    <thead>	
      <tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

        <td colspan = "2">CORREO</td>
        <td colspan = "1">ANEXO</td>
		<td colspan = "4">TIPO CONTACTO</td>

      </tr>
  	</thead>
      <tr bordercolor="#FFFFFF">

        <td colspan= "2"><input name="EMAIL" type="text" id="EMAIL" size="35" maxlength="50"></td>
		<td colspan= "1"><input name="TX_ANEXO" type="text" id="TX_ANEXO" size="35" maxlength="50"></td>
		<td colspan= "2"><select id="cbxTipoContacto" name="cbxTipoContacto">
			<option value="" selected>Seleccione</option>
			<% 	abrirscg()
				strListaTipoContacto = "SELECT IdTipoContacto, Glosa, Descripcion FROM TipoContacto WHERE TipoDatoContacto = 'E' AND CodigoCliente = '"& session("ses_codcli") &"'"					
				set rsListaTipoContacto = Conn.execute(strListaTipoContacto)
				i = 1
				
				Do While Not rsListaTipoContacto.Eof %>
					<option value="<%=rsListaTipoContacto("IdTipoContacto") %>" title="<%=rsListaTipoContacto("Descripcion") %>">
						<% response.write(i) %> - <%=rsListaTipoContacto("Glosa") %>
					</option>
			<% 	rsListaTipoContacto.movenext
				i = i + 1 
				Loop %>
				<% cerrarscg() %>
		</select></td>
        <input name="rut" type="hidden" id="rut" value="<%=rut%>">
	  </tr>
	  <thead>
        <tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

		<td>NOMBRE</td>
		<td>APELLIDO</td>
       	<td>CARGO</td>

       	<% If session("perfil_emp") <> "Verdadero" Then %>

       	<td>DEPARTAMENTO</td>

        <td >FUENTE </td>
		<td></td>

        <%Else%>

        <td>DEPARTAMENTO</td>

       	<% End If%>

       	<td></td>
		</tr>
	</thead>

		<tr bordercolor="#FFFFFF">
			<td><input name="TX_CONTACTO" type="text" id="TX_CONTACTO" size="35" maxlength="20"></td>
			<td><input name="TX_APELLIDO" type="text" id="TX_APELLIDO" size="35" maxlength="20"></td>
			<td><input name="TX_CARGO" type="text" id="TX_CARGO" size="35" maxlength="20"></td>

			<% If session("perfil_emp") <> "Verdadero" Then %>
				<td><input name="TX_DPTO" type="text" id="TX_DPTO" size="35" maxlength="20"></td>
				<td colspan="2"><select name="CB_FUENTE" id="CB_FUENTE">>
					<%
					abrirscg()
					ssql="SELECT * FROM FUENTE_UBICABILIDAD ORDER BY COD_FUENTE"
					set rsFuente= Conn.execute(ssql)
					do until rsFuente.eof%>
						<option value="<%=rsFuente("NOM_FUENTE")%>" selected><%=rsFuente("NOM_FUENTE")%></option>
						<%
							rsFuente.movenext
							loop
							rsFuente.close
							set rsFuente=nothing
							cerrarscg()
						%>
						</select>
				</td>
			<% Else%>
				<td ><input name="TX_CARGO" type="text" id="TX_CARGO" size="35" maxlength="35"></td>
			<% End If%>
			<TD align="right">
				<A HREF="#" onClick="ingresa_nuevo_mail('<%=strOrigen%>');">
					<img ID=ImgSave src="../imagenes/save_as.png" border="0">
				</A>
				&nbsp;&nbsp;
				<% If strOrigen = "" Then %>
					<A HREF="#" onClick="location.href='principal.asp'">
					<img src="../imagenes/arrow_left.png" border="0">
					</A>
				<%else%>
					<A HREF="#" onClick="carga_funcion_email();">
					<img src="../imagenes/arrow_left.png" border="0">
					</A>
				<%end if%>
			</TD>
		</tr>
    </table>
    </td>

  </tr>
</table>
<% If strOrigen = "" Then %>
	<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>

	<script type="text/javascript">

	function IsValidCamposContacto() {

	    var validaContactoTelefono = $('#validaContactoTelefono').val()
	    var validaContactoEmail = $('#validaContactoEmail').val()
	    var validaContactoDireccion = $('#validaContactoDireccion').val()

	    var TX_CONTACTO = $('#TX_CONTACTO').val()
	    var TX_CARGO = $('#TX_CARGO').val()
	    var TX_DPTO = $('#TX_DPTO').val()
	    var TX_APELLIDO = $('#TX_APELLIDO').val()

	    if (validaContactoTelefono == 'True' && validaContactoEmail == 'True' && validaContactoDireccion == 'True') {
	        if (TX_CONTACTO == '') {
	            alert('DEBE INGRESAR NOMBRE');
	            return false
	        }

	        if (TX_APELLIDO == '') {
	            alert('DEBE INGRESAR APELLIDO');
	            return false
	        }

	        if (TX_CARGO == '' && TX_DPTO == '') {
	            alert('DEBE INGRESAR AL MENOS UN CARGO O UN DEPARTAMENTO');
	            return false
	        }
	    }
	    return true
	}

	function ValidarCorreo(strmail) {
	    var Formato = /^(([^<>()[\]\.,;:\s@\"]+(\.[^<>()[\]\.,;:\s@\"]+)*)|(\".+\"))@(([^<>()[\]\.,;:\s@\"]+\.)+[^<>()[\]\.,;:\s@\"]{2,})$/i;
	    var Comparacion = Formato.test(strmail);
	    if (Comparacion == false) {
	        alert("El e-mail ingresado es invalido!");
	        return false;
	    }
	    return true;
	}


	function ingresa_nuevo_mail(strOrigen)
	{
		var EMAIL 			=$('#EMAIL').val()
		var TX_ANEXO 		=$('#TX_ANEXO').val()
		var TX_CONTACTO 	=$('#TX_CONTACTO').val()
		var TX_APELLIDO		=$('#TX_APELLIDO').val()
		var TX_CARGO 		=$('#TX_CARGO').val()
		var TX_DPTO 		=$('#TX_DPTO').val()
		var CB_FUENTE 		=$('#CB_FUENTE').val()
		var rut 			=$('#rut').val()
		var cbxTipoContacto	=$('#cbxTipoContacto option:selected').val()
		
		if(cbxTipoContacto=='')
		{
			alert('DEBE SELECCIONAR UNA OPCION DE TIPO DE CONTACTO');
			return
		}
		
		if (!IsValidCamposContacto()) {
		    return
		}
		
		if(EMAIL=='')
		{
			alert('DEBE INGRESAR UN CORREO');
			return
		}

		if (ValidarCorreo(EMAIL))
			{
				
				var criterios ="alea="+Math.random()+"&EMAIL="+encodeURIComponent(EMAIL)+"&TX_ANEXO="+encodeURIComponent(TX_ANEXO)+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&CB_FUENTE="+encodeURIComponent(CB_FUENTE)+"&rut="+rut+"&strOrigen="+strOrigen+"&strTipoContacto="+cbxTipoContacto

				$('#carga_funcion2').load('scg_cor.asp', criterios, function(){
					alert("¡Datos ingresado!")
				})
				location.href="principal.asp"

			}		
	}

	</script>

	<div id="carga_funcion2"></div>
<% End If %>
</body>
</html>