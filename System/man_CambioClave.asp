<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
    <link href="../css/style_generales_sistema.css" rel="stylesheet">
	<!--#include file="arch_utils.asp"-->


<%
	Response.CodePage=65001
	Response.charset ="utf-8"

  	strUsuario 	=session("session_idusuario")
  	estado 		=request.querystring("estado")
  	if trim(estado)<>"primera_vez" then
%>
	<!--#include file="sesion.asp"-->	
<%
	end if

	AbrirSCG()
		sql_sel="SELECT ID_USUARIO, rut_usuario, nombres_usuario, apellido_paterno, " 
		sql_sel= sql_sel & " apellido_materno, fecha_nacimiento, correo_electronico, telefono_contacto, "
		sql_sel= sql_sel & " perfil, LOGIN, CLAVE, PERFIL_ADM, perfil_cob, ACTIVO, perfil_proc, perfil_sup, "
		sql_sel= sql_sel & " PERFIL_CAJA, perfil_emp, PERFIL_FULL, perfil_back, gestionador_preventivo, "
		sql_sel= sql_sel & " ANEXO, observaciones_usuario, COD_AREA "
		sql_sel= sql_sel & " FROM USUARIO "
		sql_sel= sql_sel & " WHERE ID_USUARIO= " & TRIM(strUsuario)

		set rs_sel = conn.execute(sql_sel)
		if err Then
			Response.Write sql_sel &" / ERROR : " & err.description
			Response.end()
		end if

		if not rs_sel.eof then 

			ID_USUARIO				=rs_sel("ID_USUARIO")
			rut_usuario 			=rs_sel("rut_usuario")
			nombres_usuario 		=rs_sel("nombres_usuario")
			apellido_paterno 		=rs_sel("apellido_paterno")
			apellido_materno 		=rs_sel("apellido_materno")
			fecha_nacimiento 		=rs_sel("fecha_nacimiento")
			correo_electronico 		=rs_sel("correo_electronico")
			telefono_contacto 		=rs_sel("telefono_contacto")
			perfil 					=rs_sel("perfil")
			LOGIN 					=rs_sel("LOGIN")
			CLAVE 					=rs_sel("CLAVE")
			PERFIL_ADM 				=rs_sel("PERFIL_ADM")
			perfil_cob 				=rs_sel("perfil_cob")
			ACTIVO 					=rs_sel("ACTIVO")
			perfil_proc 			=rs_sel("perfil_proc")
			perfil_sup 				=rs_sel("perfil_sup")
			PERFIL_CAJA 			=rs_sel("PERFIL_CAJA")
			perfil_emp 				=rs_sel("perfil_emp")
			PERFIL_FULL 			=rs_sel("PERFIL_FULL") 	
			perfil_back 			=rs_sel("perfil_back")
			gestionador_preventivo 	=rs_sel("gestionador_preventivo")
			ANEXO					=rs_sel("ANEXO")
			observaciones 			=rs_sel("observaciones_usuario")
			COD_AREA  				=rs_sel("COD_AREA")
		end if
%>


<TITLE>Cambio de Clave</TITLE>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">


<script type="text/javascript">
	$(document).ready(function(){
		$(document).tooltip();
		$('#num_min').val($('#TELEFONO_CONTACTO').val().length) 

		$('#imagen_seguridad').css('display', 'none')

		$('#CLAVE_OLD').keyup(function(){
			$('#actual').text("")
			$('#actual').css('display','')
			$('#CLAVE_OLD').css('border-color','')			
		})

		$('#CLAVE_NEW1').keyup(function(){
			$('#nueva').text("")
			$('#nueva').css('display','')
			$('#CLAVE_NEW1').css('border-color','')	

			var CLAVE_NEW1 	=$(this).val()
			var numeros 	="0123456789";
			var cont_num 	=0
			for(i=0; i<CLAVE_NEW1.length; i++){
				if (numeros.indexOf(CLAVE_NEW1.charAt(i),0)!=-1){
					cont_num = cont_num+1
				}
			}

			var letras 		="qwertyuiopasdfghjklñzxcvbnmQWERTYUIOPÑLKJHGFDSAZXCVBNMáéíóúÁÉÍÓÚ";
			var cont_letras =0
			for(i=0; i<CLAVE_NEW1.length; i++){
				if (letras.indexOf(CLAVE_NEW1.charAt(i),0)!=-1){
					cont_letras = cont_letras+1
				}
			}

			if($(this).val().length==0){
				$('#imagen_seguridad').css('display', 'none')
			}else{

				if(CLAVE_NEW1.length<=6 && cont_num>=0 && cont_letras<=6)
				{
					$('#imagen_seguridad').css('display', 'inline-block')				
					$('#imagen_seguridad').attr('src', '../Imagenes/seguridad_bajo.png')
				}
				if(CLAVE_NEW1.length>=6 && cont_num>0 && cont_letras>=5)
				{
					$('#imagen_seguridad').css('display', 'inline-block')
					$('#imagen_seguridad').attr('src', '../Imagenes/seguridad_medio.png')
				}			
				if(CLAVE_NEW1.length>=6 && cont_num>=6 && cont_letras>=6)
				{
					$('#imagen_seguridad').css('display', 'inline-block')
					$('#imagen_seguridad').attr('src', '../Imagenes/seguridad_alto.png')
				}	
			}


		})

		$('#CLAVE_NEW2').keyup(function(){
			$('#confirmada').text("")
			$('#confirmada').css('display','')
			$('#CLAVE_NEW2').css('border-color','')			
		})



		$('#CLAVE_NEW1').blur(function(){
			var RegExPattern 	=/^[a-zA-Z0-9_]+$/
			var CLAVE_NEW1		=$(this).val()
			var CLAVE_NEW2 		=$('#CLAVE_NEW2').val()
			var CLAVE_OLD 		=$('#CLAVE_OLD').val()

			if(CLAVE_NEW1!="")
			{
				if(CLAVE_NEW1.match(RegExPattern)){
					
				}else{			
					//alert("Contraseña nueva inválida\n ¡Ingresa solo numeros y letras!")
					$('#nueva').text("Contraseña nueva inválida\n ¡Ingresa solo numeros y letras!")
					$('#nueva').css('display','block')
					$('#CLAVE_NEW1').css('border-color','#FE2E2E')				
					return	
				}

				if(CLAVE_OLD==CLAVE_NEW1)
				{
					//alert("¡Contraseña nueva debe ser diferente a la actual!")
					$('#nueva').text("¡Contraseña nueva debe ser diferente a la actual!")
					$('#nueva').css('display','block')
					$('#CLAVE_NEW1').css('border-color','#FE2E2E')			
					return
				}

			
				if(CLAVE_NEW1.length<6)
				{
					$('#nueva').text("¡Contraseña nueva debe tener mínimo 6 caracteres!")
					$('#nueva').css('display','block')
					$('#CLAVE_NEW1').css('border-color','#FE2E2E')			
					return
				}
				if(CLAVE_NEW1.length>16)
				{
					$('#nueva').text("¡Contraseña nueva debe tener máximo 16 caracteres!")
					$('#nueva').css('display','block')
					$('#CLAVE_NEW1').css('border-color','#FE2E2E')			
					return
				}					
			}


		})	

		$('#CLAVE_NEW2').blur(function(){
			var RegExPattern 	=/^[a-zA-Z0-9_]+$/
			var CLAVE_NEW2		=$(this).val()
			var CLAVE_NEW1 		=$('#CLAVE_NEW1').val()

			if(CLAVE_NEW2!="")
			{
				if(CLAVE_NEW2.match(RegExPattern)){
					
				}else{			
					//alert("Contraseña confirmación inválida\n ¡Ingresa solo numeros y letras!")
					$('#confirmada').text("Contraseña confirmación inválida ¡Ingresa solo numeros y letras!")
					$('#confirmada').css('display','block')
					$('#CLAVE_NEW2').css('border-color','#FE2E2E')				
					return	
				}

				if(CLAVE_NEW1!=CLAVE_NEW2)
				{
					//alert("¡Confirmacion de contraseña incorrecta!")
					$('#confirmada').text("¡Confirmacion de contraseña incorrecta!")
					$('#confirmada').css('display','block')
					$('#CLAVE_NEW2').css('border-color','#FE2E2E')			
					return
				}

			}
				

		})		

		$('#CLAVE_OLD').blur(function(){
			var RegExPattern 	=/^[a-zA-Z0-9_]+$/
			var CLAVE_OLD		=$(this).val()
			var CLAVE_NEW1 		=$('#CLAVE_NEW1').val()
			var ID_USUARIO		=$('#ID_USUARIO').val()


			if(CLAVE_OLD!="")
			{
				if(CLAVE_OLD.match(RegExPattern)){
					
				}else{			
					//alert("Contraseña actual inválida\n ¡Ingresa solo numeros y letras!")
					$('#actual').text("Contraseña actual inválida ¡Ingresa solo numeros y letras!")
					$('#actual').css('display','block')
					$('#CLAVE_OLD').css('border-color','#FE2E2E')			
					return	
				}

				if(CLAVE_OLD==CLAVE_NEW1)
				{
					//alert("¡Contraseña nueva debe ser diferente a la actual!")
					$('#nueva').text("¡Contraseña nueva debe ser diferente a la actual!")
					$('#nueva').css('display','block')
					$('#CLAVE_NEW1').css('border-color','#FE2E2E')			
					return
				}


				var criterios ="alea="+Math.random()+"&accion_ajax=verifica_contrasena&CLAVE_OLD="+CLAVE_OLD+"&CLAVE_NEW1="+CLAVE_NEW1+"&ID_USUARIO="+ID_USUARIO
				$('#div_cambio_clave').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){

					var estado_clave =$('#estado_clave').val()

					if(estado_clave=="no_corresponde")
					{
						//alert("¡Contraseña actual invalida!")
						$('#actual').text("¡Contraseña no corresponde a la contraseña actual!")
						$('#actual').css('display','block')
						$('#CLAVE_OLD').css('border-color','#FE2E2E')
						return
					}	

				})				
			}


		})	

		 $('#TELEFONO_CONTACTO').keyup(function(){		
			if (isNaN($(this).val())) {
			    //entonces (no es numero) devuelvo el valor cadena vacia
			    $('#span_TELEFONO_CONTACTO').text("Formato télefono invalido")
				$(this).css('border-color','#FE2E2E')	
				return 
			}else{
			    $('#span_TELEFONO_CONTACTO').text("")
				$(this).css('border-color','')	
			}
		
		 }) 

		$('#ANEXO').keyup(function(){		
			if (isNaN($(this).val())) {
			    //entonces (no es numero) devuelvo el valor cadena vacia
			    $('#span_ANEXO').text("Formato anexo invalido")
				$(this).css('border-color','#FE2E2E')	
				return 
			}else{
			    $('#span_ANEXO').text("")
				$(this).css('border-color','')	
			}
		
		 }) 

		
		 $('#TELEFONO_CONTACTO').keyup(function(){
		 	var num_min = $('#num_min').val() 	

			if($(this).val().length != num_min) {
			    $('#span_TELEFONO_CONTACTO').text("Fono debe tener " + num_min + " digitos")
				$(this).css('border-color','#FE2E2E')
				return			
			}	 	

		 }) 
	 
		 $('#COD_AREA').change(function(){
		 	var num_min = $('#num_min').val() 

			if($('#TELEFONO_CONTACTO').val().length != num_min) {
			    $('#span_TELEFONO_CONTACTO').text("Fono debe tener " + num_min + " digitos")
				$('#TELEFONO_CONTACTO').css('border-color','#FE2E2E')							
			}else{
				$('#span_TELEFONO_CONTACTO').text("")
				$('#TELEFONO_CONTACTO').css('border-color','')
			}			 	
		 })

		 $('#TELEFONO_CONTACTO').blur(function(){
		 	var num_min = $('#num_min').val() 	

			if($(this).val().length != num_min) {
			    $('#span_TELEFONO_CONTACTO').text("Fono debe tener " + num_min + " digitos")
				$(this).css('border-color','#FE2E2E')
				return			
			}	 	

		 }) 


		 $('#CORREO_ELECTRONICO').blur(function(){
		 	var RegExPattern = /[\w-\.]{3,}@([\w-]{2,}\.)*([\w-]{2,}\.)[\w-]{2,4}/

			if($(this).val().match(RegExPattern)){
				
			}else{			
				$('#span_CORREO_ELECTRONICO').text("Formato correo electronico invalido")
				$(this).css('border-color','#FE2E2E')	
				return	
			}

		 }) 
	})

	function Continuar(){

		var CLAVE_OLD 			=$('#CLAVE_OLD').val()
		var CLAVE_NEW1 			=$('#CLAVE_NEW1').val()
		var CLAVE_NEW2 			=$('#CLAVE_NEW2').val()
		var ID_USUARIO  		=$('#ID_USUARIO').val()
		var CORREO_ELECTRONICO 	=$('#CORREO_ELECTRONICO').val()
		var TELEFONO_CONTACTO 	=$('#TELEFONO_CONTACTO').val()
		var COD_AREA 			=$('#COD_AREA').val()
		var ANEXO 				=$('#ANEXO').val()
		var estado 				=$('#estado').val()


	

	
		


		if(TELEFONO_CONTACTO=="")
		{
			$('#span_TELEFONO_CONTACTO').text("Ingresa telefono contacto")
			$('#TELEFONO_CONTACTO').css('border-color','#FE2E2E')
			return
		}
		if(COD_AREA==0)
		{
			$('#span_TELEFONO_CONTACTO').text("Ingresa Código de área")
			$('#COD_AREA').css('border-color','#FE2E2E')
			return
		}
		if(CORREO_ELECTRONICO=="")
		{
			$('#span_CORREO_ELECTRONICO').text("Ingresa correo electronico")
			$('#CORREO_ELECTRONICO').css('border-color','#FE2E2E')
			return
		}
		
		if (isNaN($('#TELEFONO_CONTACTO').val())) {
			return 
		}
	
	
		if (isNaN($('#ANEXO').val())) {
			return 
		}

		var criterios ="alea="+Math.random()+"&accion_ajax=actuliza_datos_usuario&TELEFONO_CONTACTO="+encodeURIComponent(TELEFONO_CONTACTO)+"&COD_AREA="+encodeURIComponent(COD_AREA)+"&CORREO_ELECTRONICO="+encodeURIComponent(CORREO_ELECTRONICO)+"&ID_USUARIO="+ID_USUARIO+"&ANEXO="+ANEXO

		$('#div_actualiza_datos').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){})


		if (estado=="primera_vez"){

			if(CLAVE_OLD=="")
			{
				$('#actual').text("Debe ingresar contraseña actual")
				$('#actual').css('display','block')
				$('#CLAVE_OLD').css('border-color','#FE2E2E')				
				return
			}

		}

		if(CLAVE_OLD==""){

			alert("Datos actualizados correctamente")
			window.location.href='principal.asp'
		}



		if(CLAVE_OLD!=""){

			if(CLAVE_OLD=="")
			{
				$('#actual').text("Debe ingresar contraseña actual")
				$('#actual').css('display','block')
				$('#CLAVE_OLD').css('border-color','#FE2E2E')				
				return
			}

			if(CLAVE_NEW1=="")
			{
				$('#nueva').text("Debe ingresar contraseña nueva")
				$('#nueva').css('display','block')
				$('#CLAVE_NEW1').css('border-color','#FE2E2E')
				return
			}

			if(CLAVE_NEW2=="")
			{
				$('#confirmada').text("Debe confirmar contraseña nueva")
				$('#confirmada').css('display','block')
				$('#CLAVE_NEW2').css('border-color','#FE2E2E')				
				return
			}

			var RegExPattern = /[\w-\.]{3,}@([\w-]{2,}\.)*([\w-]{2,}\.)[\w-]{2,4}/

			if(CORREO_ELECTRONICO.match(RegExPattern)){
				
			}else{			
				$('#span_CORREO_ELECTRONICO').text("Formato correo electronico invalido")
				$('#CORREO_ELECTRONICO').css('border-color','#FE2E2E')
				return	
			}


			var RegExPattern = /^[a-zA-Z0-9_]+$/

			if(CLAVE_OLD.match(RegExPattern)){
				
			}else{			
				//alert("Contraseña actual inválida\n ¡Ingresa solo numeros y letras!")
				$('#actual').text("Contraseña actual inválida ¡Ingresa solo numeros y letras!")
				$('#actual').css('display','block')
				$('#CLAVE_OLD').css('border-color','#FE2E2E')			
				return	
			}
			if(CLAVE_NEW1.match(RegExPattern)){
				
			}else{			
				//alert("Contraseña nueva inválida\n ¡Ingresa solo numeros y letras!")
				$('#nueva').text("Contraseña nueva inválida\n ¡Ingresa solo numeros y letras!")
				$('#nueva').css('display','block')
				$('#CLAVE_NEW1').css('border-color','#FE2E2E')				
				return	
			}

			if(CLAVE_NEW2.match(RegExPattern)){
				
			}else{			
				//alert("Contraseña confirmación inválida\n ¡Ingresa solo numeros y letras!")
				$('#confirmada').text("Contraseña confirmación inválida ¡Ingresa solo numeros y letras!")
				$('#confirmada').css('display','block')
				$('#CLAVE_NEW2').css('border-color','#FE2E2E')				
				return	
			}

			if(CLAVE_OLD==CLAVE_NEW1)
			{
				//alert("¡Contraseña nueva debe ser diferente a la actual!")
				$('#nueva').text("¡Contraseña nueva debe ser diferente a la actual!")
				$('#nueva').css('display','block')
				$('#CLAVE_NEW1').css('border-color','#FE2E2E')			
				return
			}

			if(CLAVE_NEW1!=CLAVE_NEW2)
			{
				//alert("¡Confirmacion de contraseña incorrecta!")
				$('#confirmada').text("¡Confirmacion de contraseña incorrecta!")
				$('#confirmada').css('display','block')
				$('#CLAVE_NEW2').css('border-color','#FE2E2E')			
				return
			}
			
			if(CLAVE_NEW1.length<6)
			{
				$('#nueva').text("¡Contraseña nueva debe tener mínimo 6 caracteres!")
				$('#nueva').css('display','block')
				$('#CLAVE_NEW1').css('border-color','#FE2E2E')			
				return
			}
			if(CLAVE_NEW1.length>16)
			{
				$('#nueva').text("¡Contraseña nueva debe tener máximo 16 caracteres!")
				$('#nueva').css('display','block')
				$('#CLAVE_NEW1').css('border-color','#FE2E2E')			
				return
			}		

			var criterios ="alea="+Math.random()+"&accion_ajax=verifica_contrasena&CLAVE_OLD="+CLAVE_OLD+"&CLAVE_NEW1="+CLAVE_NEW1+"&ID_USUARIO="+ID_USUARIO
			$('#div_cambio_clave').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){

				var estado_clave =$('#estado_clave').val()

				if(estado_clave=="existe_clave_new")
				{
					//alert("Contraseña nueva debe ser diferente a la enviada por correo")				
					$('#nueva').text("¡Contraseña nueva debe ser diferente a la enviada por correo!")
					$('#nueva').css('display','block')
					$('#CLAVE_NEW1').css('border-color','#FE2E2E')
					return
				}	
				if(estado_clave=="no_corresponde")
				{
					//alert("¡Contraseña actual invalida!")
					$('#actual').text("¡Contraseña no corresponde a la contraseña actual!")
					$('#actual').css('display','block')
					$('#CLAVE_OLD').css('border-color','#FE2E2E')
					return
				}	

				if(estado_clave=="igual_login")
				{
					//alert("¡Contraseña actual invalida!")
					$('#nueva').text("¡Contraseña no puede ser igual al LOGIN de usuario!")
					$('#nueva').css('display','block')
					$('#CLAVE_NEW1').css('border-color','#FE2E2E')
					return
				}
				

				var criterios ="alea="+Math.random()+"&accion_ajax=modifica_contrasena&CLAVE_NEW1="+CLAVE_NEW1+"&ID_USUARIO="+ID_USUARIO
				$('#div_cambio_clave').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){
					alert("Datos actualizados correctamente")
					if (estado=="primera_vez"){
						window.location.href='default.asp'
					}else{
						window.location.href='principal.asp'
					}
						
				})

			})

		}
		
		
	}

	function bt_volver(){
		window.location.href='principal.asp'
	}

	function asigna_minimo_nuevo(campo, minimo1){
		
		if (campo!=0)	{
			if(campo==41 || campo==32 || campo==45 || campo==57 || campo==55 || campo==72 || campo==71 || campo==73 || campo==75){
				minimo1=7;
			}else if(campo.length==1 || campo==2){
				minimo1=8;
			}else {
				minimo1=6;
			}
		}else{minimo1=0}
		return(minimo1)

	}


	function  asigna_minimo_a_variable(COD_AREA, num_min)
	{
		$('#num_min').val(asigna_minimo_nuevo(COD_AREA,num_min))

	} 

	function valida_largo_nuevo(campo, minimo){

			if(campo.length != minimo) {
				alert("Fono debe tener " + minimo + " digitos")
				$('#numero').select()
				$('#numero').focus()
				return(true)
			}

		return(false)
	}	
</script>

<style type="text/css">
	.aviso_rojo{
		color:#FE2E2E;
		font-size:12px ;
		font-weight: bold;
	}

	.span_aviso_rojo{
		color:#FE2E2E;
		font-size:12px;
	}
	.boton_man_usu{
		background-color: #08088A;
		color:#fff;
		border: 1px solid #08088A;
		padding: 5px;
		margin: 10px;
		width: 110px;

	}	
	.hdr_i{
		background-color: #A9E2F3; 		
		color: #000;
		font: 14px bold tahoma;	

	}
	
	.input_cambio_clave{
		width: 230px;
	}


</style>
</HEAD>

<BODY BGCOLOR='FFFFFF'>
<FORM NAME="mantenedorForm"  action="" method="">
<input name="num_min" 			id="num_min" 			type="hidden" 	value="0">
<input name="actualiza_datos" 	id="actualiza_datos" 	type="hidden" 	value="0">
<input name="actualiza_clave" 	id="actualiza_clave" 	type="hidden" 	value="0">
<input name="estado" 			id="estado" 			type="hidden" 	value="<%=trim(estado)%>">
<div class="titulo_informe">CAMBIO DE CLAVE</div>
<br>
<input type="hidden" name="ID_USUARIO" id="ID_USUARIO" value="<%=trim(strUsuario)%>" >
<table width="90%" BORDER="0" CELLPADDING="0" CELLSPACING="0" align="center" class="estilo_columnas" >
	<thead>
 	<tr>
		<td width="250" colspan="2" height="22" style="color:#FFFFFF;">Actualiza datos personales</td>
	</TR>
	</thead>
	<tbody>
 	<tr BGCOLOR="#FFFFFF">
		<td width="250" bgcolor="#<%=session("COLTABBG2")%>"><font class="aviso_rojo">* </font>Corre electrónico</td>
		<td class="td_t"><input type="text" class="input_cambio_clave" name="CORREO_ELECTRONICO" id="CORREO_ELECTRONICO" value="<%=trim(CORREO_ELECTRONICO)%>"><span id="span_CORREO_ELECTRONICO" class="span_aviso_rojo"></span></td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" bgcolor="#<%=session("COLTABBG2")%>"><font class="aviso_rojo">* </font>Teléfono contacto</Font></td>
		<td class="td_t">
		<select name="COD_AREA" id="COD_AREA" onchange="asigna_minimo_a_variable(this.value,0)" style="width:50px;">
			<%
			ssql="SELECT DISTINCT CODIGO_AREA FROM COMUNA WHERE ID_SADI<>0 UNION SELECT 9 AS CODIGO_AREA  ORDER BY CODIGO_AREA DESC"
			 set rsCOM= Conn.execute(ssql)
			do until rsCOM.eof%>

				<option value="<%=rsCOM("codigo_area")%>" <%if trim(COD_AREA)=trim(rsCOM("codigo_area")) then%> selected <%end if%>><%=rsCOM("codigo_area")%></option>

			<%
			rsCOM.movenext
			loop
			rsCOM.close
			set rsCOM=nothing

			%>

			<option value="0">--</option>
			</select>
			(CEL.9)			
			<input type="text" style="width:133px;" name="TELEFONO_CONTACTO" id="TELEFONO_CONTACTO" value="<%=trim(TELEFONO_CONTACTO)%>">
			<span id="span_TELEFONO_CONTACTO" class="span_aviso_rojo"></span>
		</td>
	</TR>
 	<tr BGCOLOR="#FFFFFF">
		<td width="250" bgcolor="#<%=session("COLTABBG2")%>">&nbsp;&nbsp;&nbsp;Anexo</td>
		<td class="td_t"><input <%if trim(ANEXO)<>"" then%> title="<%=trim(ANEXO)%>" <%end if%> type="text" class="input_cambio_clave" name="ANEXO" id="ANEXO" value="<%=trim(ANEXO)%>"><span id="span_ANEXO" class="span_aviso_rojo"></span></td>
	</TR>	
	<tr BGCOLOR="#FFFFFF">
		<td colspan="2">&nbsp;</td>
	</TR>
	</tbody>
	<thead>
 	<tr BGCOLOR="#FFFFFF">
		<td width="250" colspan="2" height="22" style="color:#FFFFFF;">Cambio de clave</td>
	</TR>
	</thead>
 	<tr BGCOLOR="#FFFFFF">
		<td width="250" bgcolor="#<%=session("COLTABBG2")%>"><font class="aviso_rojo">* </font>Contraseña Actual</td>
		<td class="td_t"><input type="password" class="input_cambio_clave" name="CLAVE_OLD" id="CLAVE_OLD" value=""><span id="actual" class="span_aviso_rojo"></span></td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" bgcolor="#<%=session("COLTABBG2")%>"><font class="aviso_rojo">* </font>Contraseña Nueva</Font></td>
		<td class="td_t"><input type="password" class="input_cambio_clave" name="CLAVE_NEW1" id="CLAVE_NEW1" value=""><img id="imagen_seguridad" src="" alt=""><span id="nueva" class="span_aviso_rojo" ></span></td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" bgcolor="#<%=session("COLTABBG2")%>"><font class="aviso_rojo">* </font>Confirmar Contraseña Nueva</Font></td>
		<td class="td_t"><input type="password" class="input_cambio_clave" name="CLAVE_NEW2" id="CLAVE_NEW2" value=""><span id="confirmada" class="span_aviso_rojo"></span></td>
	</TR>

</TABLE>
<br>
<table width="90%" border="0" align="center">
     <TR>
	  <td align="right">
	   <INPUT TYPE="BUTTON" class="fondo_boton_100" value="actualizar datos" name="BT_GUARDAR" ID="BT_GUARDAR" onClick="Continuar()">
	   </TD>
	  </TD>
    </TR>
</table>
<div id="div_cambio_clave"></div>
<div id="div_actualiza_datos"></div>



</FORM>


</BODY>
</HTML>

<%CerrarSCG()%>


