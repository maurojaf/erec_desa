<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<link rel="stylesheet" href="../css/style_generales_sistema.css">
<%

Response.CodePage 	=65001
Response.charset	="utf-8"

rut 		=request("rut")
strOrigen 	=request("strOrigen")
%>

<% If strOrigen = "" Then %>
	<!--#include file="sesion.asp"-->
<% Else %>
	<!--#include file="sesion_inicio.asp"-->
<% End If %>

</head>
<body>
<input name="rut" type="hidden" id="rut" value="<%=rut%>">
<input name="strOrigen" type="hidden" id="strOrigen" value="<%=strOrigen%>">
<table <% If strOrigen = "" Then %>width="90%"<%else%>width="100%"<%end if%> border="0" align="center">

	<% If strOrigen = "" Then %>
		<tr>
			<TD width="100%" ALIGN="LEFT" class="titulo_informe">
				Nueva Dirección
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
        <td >COMUNA</td>
        <td >CALLE</td>
        <td >NUMERO</td>
        <td colspan="2">RESTO</td>
        </tr>
    </thead>
      <tr bordercolor="#FFFFFF">

		<td>
		<select name="comuna" id="comuna">
				<option value="0">SELECCIONE</option>
				<%
				abrirscg()
				ssql="SELECT nombre_comuna,n_sadi FROM COMUNA WHERE codigo_comuna<>'0' ORDER BY nombre_comuna"
				set rsCOM= Conn.execute(ssql)
				 do until rsCOM.eof%>
					<option value="<%=rsCOM("n_sadi")%>"><%=rsCOM("n_sadi")%></option>
					<%
				  rsCOM.movenext
				  loop
				  rsCOM.close
				  set rsCOM=nothing
				  cerrarscg()
				  %>
		        </select>
		&nbsp;&nbsp;
		</td>

        <td><input name="calle" type="text" id="calle" size="40" maxlength="80"></td>
        <td><input name="numero" type="text" id="numero" size="10"></td>
        <td><input name="resto" type="text" id="resto" size="30" maxlength="80"></td>

        </tr>
		<thead>
      	<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
        	<td>DIAS PAGO</td>
        	<td>HORARIOS DE PAGO</td>
			<td colspan="3">TIPO CONTACTO</td>
        </tr>
    	</thead>
      <tr bordercolor="#FFFFFF">

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
			<td>
			Lu
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS" id="CH_DIAS" value="LU" <%=strChequedLu%>>
			Ma
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS" id="CH_DIAS" value="MA" <%=strChequedMa%>>
			Mi
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS" id="CH_DIAS" value="MI" <%=strChequedMi%>>
			Ju
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS" id="CH_DIAS" value="JU" <%=strChequedJu%>>
			Vi
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS" id="CH_DIAS" value="VI" <%=strChequedVi%>>
			Sa
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS" id="CH_DIAS" value="SA" <%=strChequedSa%>>
            </td>

        <td><input name="TX_DESDE" id="TX_DESDE" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
			<input name="TX_HASTA" id="TX_HASTA" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
		</td>
		
		<td><select id="cbxTipoContacto" name="cbxTipoContacto">
			<option value="" selected>Seleccione</option>
			<% 	abrirscg()
				strListaTipoContacto = "SELECT IdTipoContacto, Glosa, Descripcion FROM TipoContacto WHERE TipoDatoContacto = 'D' AND CodigoCliente = '"& session("ses_codcli") &"'"					
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
		
		<thead>
			<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td colspan=1>NOMBRE</td>
				<td colspan=1>APELLIDO</td>
				<td colspan=1>CARGO</td>
				<td colspan=2>DEPARTAMENTO</td>
			</tr>
		</thead>
		<tr bordercolor="#FFFFFF">
			<td colspan=1><input name="TX_CONTACTO" type="text" id="TX_CONTACTO" size="35" maxlength="20"></td>
			<td colspan=1><input name="TX_APELLIDO" type="text" id="TX_APELLIDO" size="35" maxlength="20"></td>
			<td colspan=1><input name="TX_DPTO" type="text" id="TX_DPTO" size="35" maxlength="20"></td>
			<td colspan=1><input name="TX_CARGO" type="text" id="TX_CARGO" size="35" maxlength="35"></td>
			<td align="right">
				<A HREF="#" onClick="guarda_nueva_direccion();">
				<img ID=ImgSave src="../imagenes/save_as.png" border="0">
				</A>
				&nbsp;&nbsp;
				<% If strOrigen = "" Then %>
					<A HREF="#" onClick="location.href='principal.asp'">
					<img src="../imagenes/arrow_left.png" border="0">
					</A>
				<%else%>
					<A HREF="#" onClick="carga_funcion_direccion();">
					<img src="../imagenes/arrow_left.png" border="0">
					</A>
				<%end if%>
			</td>
		</tr>

    </table>
    </td>
  </tr>
</table>
</body>
</html>
<% If strOrigen = "" Then %>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script type="text/javascript">

	function guarda_nueva_direccion(){

		var comuna  		=$('#comuna').val()
		var calle  			=$('#calle').val()
		var numero  		=$('#numero').val()
		var resto  			=$('#resto').val()
		var strDiasAtencion =""
		var rut 			=$('#rut').val()
		var TX_DESDE  		=$('#TX_DESDE').val()
		var TX_HASTA  		=$('#TX_HASTA').val()
		var TX_CONTACTO 	= $('#TX_CONTACTO').val()
		var TX_CARGO 		= $('#TX_CARGO').val()
		var TX_DPTO 		= $('#TX_DPTO').val()
		var TX_APELLIDO		= $('#TX_APELLIDO').val()
		
		var cbxTipoContacto	= $('#cbxTipoContacto option:selected').val()

		$("input[id='CH_DIAS']:checked").each(function () {
			strDiasAtencion =$(this).val()+","+strDiasAtencion
		})

		strDiasAtencion 	=strDiasAtencion.substring(0, strDiasAtencion.length-1)
		
		if(cbxTipoContacto=='')
		{
			alert('DEBE SELECCIONAR UNA OPCION DE TIPO DE CONTACTO');
			return
		}
		
		if(!IsValidCamposContacto()) {
			return
		}

		if (strDiasAtencion=="")
		{
			alert("Debe seleccionar al menos 1 día de pago");
			return
		}

		if (calle=="")
		{
			alert('Debe ingresar una calle');
			return
		}

		if (numero=="")
		{
			alert('Debe ingresar un numero');
			return
		}

		if (comuna=="")
		{
			alert('Debe seleccionar una comuna');
			return
		}


		var criterios ="alea="+Math.random()+"&strOrigen=deudor_direcciones&rut="+rut+"&comuna="+comuna+"&numero="+encodeURIComponent(numero)+"&calle="+encodeURIComponent(calle)+"&strDiasAtencion="+strDiasAtencion+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&TX_HASTA="+TX_HASTA+"&TX_DESDE="+TX_DESDE+"&resto="+encodeURIComponent(resto)+"&strTipoContacto="+cbxTipoContacto

	 	$('#carga_funcion2').load('scg_dir.asp', criterios, function(data){
			alert("¡Datos ingresado!")
		})
		location.href="principal.asp"

	}
function ValidaHora( ObjIng, strHora )
{
    var er_fh = /^(00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23)\:([0-5]0|[0-5][1-9])$/
    if( strHora == "" )
    {
            alert("Introduzca la hora.")
            return false
    }
    if ( !(er_fh.test( strHora )) )
    {
            alert("El dato en el campo hora no es válido.");
            ObjIng.value = '';
            ObjIng.focus();
            return false
    }

    //alert("¡Campo de hora correcto!")
    return true
}


</script>
<div id="carga_funcion2"></div>
<%end if%>
