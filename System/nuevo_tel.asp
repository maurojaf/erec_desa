<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/style_generales_sistema.css">
<%
Response.CodePage = 65001
Response.charset="utf-8"

rut 		= request("rut")
strOrigen 	= request("strOrigen")
%>

<% If strOrigen = "" Then %>
	<!--#include file="sesion.asp"-->
<% Else %>
	<!--#include file="sesion_inicio.asp"-->
<% End If %>

<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->
</head>
<body>

<input name="strOrigen" type="hidden" id="strOrigen" value="<%=strOrigen%>">
<table <% If strOrigen = "" Then %>width="90%"<%else%>width="100%"<%end if%> border="0" align="center">

<% If strOrigen = "" Then %>
	<tr>
		<TD width="100%" ALIGN=LEFT class="titulo_informe">
			Nuevo Teléfono
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
						<td>CODIGO AREA </td>
						<td>TELEFONO </td>
						<td>ANEXO</td>
						<td>TIPO CONTACTO</td>
						<td>DIAS DE ATENCION</td>
						<td Colspan ="5">HORA ATENCION</td>
					</tr>
				</thead>
				<tr>
					<td> 
						<select name="COD_AREA" id="COD_AREA" onchange="asigna_minimo_a_variable(this.value,0)" style="width:50px;">
							<%
							abrirscg()
							ssql="SELECT DISTINCT CODIGO_AREA FROM COMUNA WHERE ID_SADI<>0 UNION SELECT 9 AS CODIGO_AREA  ORDER BY CODIGO_AREA DESC"
							set rsCOM= Conn.execute(ssql)
							do until rsCOM.eof%>

							<option value="<%=rsCOM("codigo_area")%>" selected><%=rsCOM("codigo_area")%></option>

							<%
							rsCOM.movenext
							loop
							rsCOM.close
							set rsCOM=nothing
							cerrarscg()
							%>
							<option value="0" selected>--</option>
						 </select>
						 (CEL.9)
					</td>
					<td>
						<input name="numero" type="text" id="numero" size="10" maxlength="10" onKeyUp="numero.value=solonumero(numero)">
					</td>
					<td>
						<input name="TX_ANEXO" id="TX_ANEXO" type="text" value="<%=strAnexo%>" size="35" maxlength="50">
					</td>		
				
					<td>
						<select id="cbxTipoContacto" name="cbxTipoContacto">
							<option value="" selected>Seleccione</option>
							<% 	abrirscg()
								strListaTipoContacto = "SELECT IdTipoContacto, Glosa, Descripcion FROM TipoContacto WHERE TipoDatoContacto = 'T' AND CodigoCliente = '"& session("ses_codcli") &"'"					
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
						</select>
					</td>
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
							LU
							<INPUT TYPE="CHECKBOX" NAME="CH_DIAS"  id="CH_DIAS"  value="LU" <%=strChequedLu%>>
							Ma
							<INPUT TYPE="CHECKBOX" NAME="CH_DIAS" id="CH_DIAS" value="MA" <%=strChequedMa%>>
							Mi
							<INPUT TYPE="CHECKBOX" NAME="CH_DIAS" id="CH_DIAS" value="MI" <%=strChequedMi%>>
							Ju
							<INPUT TYPE="CHECKBOX" NAME="CH_DIAS" id="CH_DIAS" value="JU" <%=strChequedJu%>>
							Vi
							<INPUT TYPE="CHECKBOX" NAME="CH_DIAS" id="CH_DIAS" value="VI" <%=strChequedVi%>>
							Sa
							<INPUT TYPE="CHECKBOX" NAME="CH_DIAS" id="CH_DIAS" value="SA" <%=strChequedSa%>>

					</td>

					<td>
							<input name="TX_DESDE" id="TX_DESDE"  type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
							<input name="TX_HASTA" id="TX_HASTA" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">

					</td>
					<td>
								  <input name="rut" type="hidden" id="rut" value="<%=rut%>">
					</td>
				</tr>
		</td>
	</tr>
	<thead>
		<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td colspan=1>NOMBRE</td>
			<td colspan=1>APELLIDO</td>
			<td colspan=1>CARGO</td>
			
			<% If TraeSiNo(session("perfil_emp")) <> "Si" Then %>
				<td colspan=1>DEPARTAMENTO</td>
				<td colspan=2>FUENTE </td>
			<%Else%>
				<td colspan=1>DEPARTAMENTO</td>
			<% End If%>
		</tr>
	</thead>
	<tr bordercolor="#FFFFFF">
		<td colspan=1><input name="TX_CONTACTO" type="text" id="TX_CONTACTO" size="35" maxlength="20"></td>
		<td colspan=1><input name="TX_APELLIDO" type="text" id="TX_APELLIDO" size="35" maxlength="20"></td>
		<td colspan=1><input name="TX_CARGO" type="text" id="TX_CARGO" size="35" maxlength="20"></td>

		<% If TraeSiNo(session("perfil_emp")) <> "Si" Then %>

		<td colspan=1><input name="TX_DPTO" type="text" id="TX_DPTO" size="35" maxlength="20"></td>

		<td colspan=1>
			<select name="CB_FUENTE" id="CB_FUENTE">
				<option value="" selected>Seleccione</option>
			<%
			abrirscg()
			ssql="SELECT * FROM FUENTE_UBICABILIDAD ORDER BY COD_FUENTE"
			set rsFuente= Conn.execute(ssql)
			do until rsFuente.eof%>
				<option value="<%=rsFuente("NOM_FUENTE")%>" <%if trim(rsFuente("COD_FUENTE"))="16" THEN RESPONSE.WRITE " SELECTED " END IF%>><%=rsFuente("NOM_FUENTE")%></option>
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

		<td colspan=1><input name="TX_CARGO" type="text" id="TX_CARGO" size="35" maxlength="35">
		</td>

			<% End If%>

		<td align="right">

			<A HREF="#" onClick="nuevo_telefono();">
				<img ID="ImgSave" src="../imagenes/save_as.png" border="0">
			</A>
			&nbsp;&nbsp;
			<% If strOrigen = "" Then %>
				<A HREF="#" onClick="location.href='principal.asp'">
				<img src="../imagenes/arrow_left.png" border="0">
				</A>
			<%else%>
				<A HREF="#" onClick="carga_funcion_telefono();">
				<img src="../imagenes/arrow_left.png" border="0">
				</A>
			<%end if%>			
		</td>
	</tr>
</table>

<% If strOrigen = "" Then %>
<input name="num_min" 			id="num_min" 			type="hidden" 	value="10">

<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script type="text/javascript">

	//METODO QUE PERMITE VERIFICAR SI EL CLIENTE VALIDA CAMPOS DE CONTACTO
	function IsValidCamposContacto() {

		var validaContactoTelefono = $('#validaContactoTelefono').val()
		var validaContactoEmail = $('#validaContactoEmail').val()
		var validaContactoDireccion = $('#validaContactoDireccion').val()
		
		var TX_CONTACTO 	= $('#TX_CONTACTO').val()
		var TX_CARGO 		= $('#TX_CARGO').val()
		var TX_DPTO 		= $('#TX_DPTO').val()
		var TX_APELLIDO		= $('#TX_APELLIDO').val()
		
		if(validaContactoTelefono == 'True' && validaContactoEmail == 'True' && validaContactoDireccion == 'True') {
			if(TX_CONTACTO==''){
				alert('DEBE INGRESAR NOMBRE');
				return false
			}
			
			if(TX_APELLIDO==''){
				alert('DEBE INGRESAR APELLIDO');
				return false
			}
			
			if(TX_CARGO=='' && TX_DPTO==''){
				alert('DEBE INGRESAR AL MENOS UN CARGO O UN DEPARTAMENTO');
				return false
			}
		}	
		return true
	}
	
	function valida_largo_nuevo(campo, minimo){
		var numero = $('#numero').val();
		var codigoArea = $('#COD_AREA').val();
		var inicioNumero = numero.substring(0,3);
		
		if(codigoArea != 0) {
			if(campo.length != minimo) {
				alert("Fono debe tener " + minimo + " digitos");
				$('#numero').select();
				$('#numero').focus();
				return(true);
			}	
		} else {
			
			if (numero.length < 9) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 600 y de largo 10 o prefijo 800 o 197 y de largo 9.");
				return(true);
			}
			
			if (numero.length == 9 && inicioNumero != 800 && inicioNumero != 197) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 800 o 197 y de largo 9.");
				return(true);
			}
			
			if (numero.length == 10 && inicioNumero != 600) {
				alert("El numero ingresado no es valido, debe ser un numero con prefijo 600 y de largo 10.");
				return(true);
			}
		}
		
		return(false);
	}


	function solonumero(valor){
	     //Compruebo si es un valor numérico
	      if (isNaN(valor.value)) {
	            //entonces (no es numero) devuelvo el valor cadena vacia
	            valor.value=""
				return ""
	      }else{
	            //En caso contrario (Si era un número) devuelvo el valor
				valor.value
				return valor.value
	      }
	}
	function nuevo_telefono(){
		var rut 			 =$('#rut').val()


		var COD_AREA 		= $('#COD_AREA').val()
		var numero 			= $('#numero').val()
		var TX_CONTACTO 	= $('#TX_CONTACTO').val()
		var TX_CARGO 		= $('#TX_CARGO').val()
		var TX_DPTO 		= $('#TX_DPTO').val()
		var TX_APELLIDO		= $('#TX_APELLIDO').val()
		var CB_FUENTE 		= $('#CB_FUENTE').val()
		var TX_ANEXO		= $('#TX_ANEXO').val()
		var TX_DESDE 		= $('#TX_DESDE').val()
		var TX_HASTA 		= $('#TX_HASTA').val()
		var dias_atencion 	= "" 	
		var num_min 		= $('#num_min').val()
		var cbxTipoContacto	= $('#cbxTipoContacto option:selected').val()
		
		if(cbxTipoContacto=='')
		{
			alert('DEBE SELECCIONAR UNA OPCION DE TIPO DE CONTACTO');
			return
		}
		
		if(!IsValidCamposContacto()) {
			return
		}
		
		if(numero==''){
			alert('Debe ingresar un numero');

		}else if (valida_largo_nuevo(numero, num_min)){
			
		}else{
			
			$('input[name="CH_DIAS"]:checked').each(function () {

				dias_atencion =$(this).val()+","+dias_atencion
			})

			strDiasAtencion =dias_atencion.substring(0, dias_atencion.length-1)
/*
				alert(COD_AREA)
				alert(numero)
				alert(TX_ANEXO)
				alert(TX_DESDE)
				alert(TX_HASTA)
				alert(TX_CONTACTO)
				alert(TX_CARGO)
				alert(TX_DPTO)
				alert(CB_FUENTE)
				alert(strDiasAtencion)
*/
			var criterios ="alea="+Math.random()+"&strOrigen=deudor_telefonos&COD_AREA="+COD_AREA+"&numero="+numero+"&rut="+rut+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&CB_FUENTE="+encodeURIComponent(CB_FUENTE)+"&CH_DIAS="+strDiasAtencion+"&TX_HASTA="+TX_HASTA+"&TX_DESDE="+TX_DESDE+"&TX_ANEXO="+encodeURIComponent(TX_ANEXO)+"&cbxTipoContacto="+cbxTipoContacto

			$('#carga_funcion2').load('scg_tel.asp', criterios, function(data){
				alert("¡Datos ingresado!")
			})
			location.href="principal.asp"

		}

	}


	function  asigna_minimo_a_variable(COD_AREA, num_min)
	{
		$('#num_min').val(asigna_minimo_nuevo(COD_AREA,num_min))

	} 


	function asigna_minimo_nuevo(campo, minimo1){
		
		if (campo!=0)	{
			if(campo==41 || campo==32 || campo==45 || campo==57 || campo==55 || campo==72 || campo==71 || campo==73 || campo==75){
				minimo1=7;
			}else if(campo.length==1 || campo==2){
				minimo1=8;
			}else {
				minimo1=7;
			}
		}else{minimo1=10}
		return(minimo1)

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
<% End If %>

</body>
</html>








