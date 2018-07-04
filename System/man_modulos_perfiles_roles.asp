<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

AbrirSCG()
%>

	<TITLE>Mantenedor módulos, perfiles y roles</TITLE>
	<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
	<link href="../css/style.css" rel="Stylesheet">

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>


<style type="text/css">
	
	#relacion_perfil_accion{
		float:right;
		margin-top: -25px ;
	}

	#botonera_flotante{

		position:fixed;
		top:0px;
		margin-top:0px;
		padding-top:0px;	

	}

	.titulo_principal{
		background-color:#380ACD; 
		height: 20px;				
		text-align: center;
	}
	.titulo_secundario{
		font-size: 12px;
		margin: 10px;

	}
	#opcion_mantenedor{
		width: 200px;
	}
	.seleccion{
		margin:20px;
	}

	#mantenedor{
		width: 100%;
		background-color:#FAFAFA ;
		margin-left:5%;
		margin-right:5%;
		height:70%;
		text-align: center;
		clear: both;
	}

	.boton_mantenedor{
		float: left;
		margin: 10px;
	}

	.nombre_modulo{
		float: left;
	}
	.tabla_mantenedor{
		
		width:400px;
		text-align: left;
		border-collapse:collapse;
		padding: 0; 		

	}

	.table_cod{
		text-align: center;	
	}

	 th{
		background-color: #A9E2F3; 		
		color: #000;
		font: 14px bold tahoma;		
	 }

	 #contenedor_tabla{
	 	margin-top:50px;
	 	border: 1px solid #ccc;
	 	width:400px;

	 }

	 .tabla_mantenedor_accion{
	 	width:700px;
		text-align: left;
		border-collapse:collapse;
		padding: 0; 		
	 }

	.tabla_mantenedor_accion caption{
		text-align: left;
		margin-bottom: 10px;
	}


	 .tabla_mantenedor_per_accion{
	 	width:100%;
		text-align: left;
		border-collapse:collapse;
		padding: 0; 
		margin-top: 10px;		
	 }

	 .tabla_mantenedor_perfil{
	 	width:100%;
		text-align: left;
		border-collapse:collapse;
		padding: 0; 

	 }

	 .div_datos_accion{
	 	margin-top:5px;
	 	width:700px;
	 	display: block;

	 }

	 .div_datos_modulo{
	 	border: 1px solid #ccc;
	 }

	 .info_perfil{
	 	border:1px solid #ccc;
	 	height: 50px;
	 	width:400px;
	 	float:left;
	 	margin:20px;
	 }
	 #buscar_perfil{
	 	text-align: center;
	 	width:180px;
	 	height: 20px;
	 	float: left;
	 	font-size: 14px;
	 	margin-botton:10px;
	 	display: block;
	 	cursor: pointer;
	 }

	 #div_datos_accion{
	 	margin-top: 20px;
	 }
</style>

<script type="text/javascript">
$(document).ready(function(){

	$('#relacion_perfil_accion').css('display','none')

	$('#opcion_mantenedor').change(function(){
		$('#relacion_perfil_accion').css('display','none')
		var opcion_mantenedor = $(this).val()
		var criterios = "alea="+Math.random()+"&accion_ajax="+opcion_mantenedor
		$('#mantenedor').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){

			if(opcion_mantenedor=="perfil" || opcion_mantenedor=="accion"){
				$('#muestra_perfil').css('display', 'none')
				$('#buscar_perfil').toggle(function(){
					$('#muestra_perfil').css('display', 'block')
				}, function(){
					$('#muestra_perfil').css('display', 'none')
				})		
			}
			
			if(opcion_mantenedor=="per_accion"){
				$('#relacion_perfil_accion').css('display','block')
			}


				

		})
	})

	$('.tr_perfil').click(function(){
		$(this).css('background-color','#81DAF5')
	})

})

function bt_select_modulo_acccion(){

	var select_modulo_acccion = $('#select_modulo_acccion').val()

	var criterios = "alea="+Math.random()+"&accion_ajax=filtrar_accion_modulo&mod_codigo="+select_modulo_acccion
	$('#div_datos_accion').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){})

}


function bt_select_modulo_perfil()
{
	var select_modulo_acccion = $('#select_modulo_acccion').val()

	var criterios = "alea="+Math.random()+"&accion_ajax=filtrar_perfil_modulo&mod_codigo="+select_modulo_acccion
	$('#div_datos_accion').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){})	
}
function bt_guardar_modulo(){
	var nombre_modulo =$('#nombre_modulo').val()

	if(nombre_modulo=="")
	{
		alert("Debe ingresar nombre modulo")
		return
	}

	var criterios = "alea="+Math.random()+"&accion_ajax=guardar_modulo&nombre_modulo="+encodeURIComponent(nombre_modulo)
	$('#mantenedor').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){})

}

function bt_guardar_accion()
{
	var nombre_accion =$('#nombre_accion').val()

	if(nombre_accion=="")
	{
		alert("Debe ingresar nombre accion")
		return
	}

	if($('input[id="ck_mod_codigo"]:checked').size()==0)
	{
		alert("Debe seleccionar un modulo")
		return
	}

	var cont =0
	

	$('input[id="ck_mod_codigo"]:checked').each(function() {  
		cont 	= cont +1      

		var criterios = "alea="+Math.random()+"&accion_ajax=guardar_accion&nombre_accion="+encodeURIComponent(nombre_accion)+"&ck_mod_codigo="+$(this).val()
        $('#mantenedor').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){
        	

        })              

    }); 		
	
	if($('input[id="ck_mod_codigo"]:checked').size()==cont)
	{
		var criterios = "alea="+Math.random()+"&accion_ajax=accion"
		$('#mantenedor').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){

				$('#muestra_perfil').css('display', 'none')
				$('#buscar_perfil').toggle(function(){
					$('#muestra_perfil').css('display', 'block')
				}, function(){
					$('#muestra_perfil').css('display', 'none')
				})	

		})
	}




}

function bt_guardar_perfil()
{
	var nombre_perfil =$('#nombre_perfil').val()

	if(nombre_perfil=="")
	{
		alert("Debe ingresar nombre perfil")
		return
	}

	if($('input[id="ck_mod_codigo"]:checked').size()==0)
	{
		alert("Debe seleccionar un modulo")
		return
	}

	var cont =0
	

	$('input[id="ck_mod_codigo"]:checked').each(function() {  
		cont 	= cont +1      

		var criterios = "alea="+Math.random()+"&accion_ajax=guardar_perfil&nombre_perfil="+encodeURIComponent(nombre_perfil)+"&ck_mod_codigo="+$(this).val()
        $('#mantenedor').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){})              

    }); 		
	
	if($('input[id="ck_mod_codigo"]:checked').size()==cont)
	{
		var criterios = "alea="+Math.random()+"&accion_ajax=perfil"
		$('#mantenedor').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){

				$('#muestra_perfil').css('display', 'none')
				$('#buscar_perfil').toggle(function(){
					$('#muestra_perfil').css('display', 'block')
				}, function(){
					$('#muestra_perfil').css('display', 'none')
				})	

		})
	}



}

function bt_refresca_info()
{
	var mod_codigo =$('#mod_codigo').val()

	if($('input[id="seleccion_relacion_perfil"]:checked').size()>0)
	{
		var seleccion_relacion_perfil = $('input[id="seleccion_relacion_perfil"]:checked').val()
	}else{
		var seleccion_relacion_perfil =0
	}


	var criterios = "alea="+Math.random()+"&accion_ajax=cagar_info_contenedor&mod_codigo="+mod_codigo+"&seleccion_relacion_perfil="+seleccion_relacion_perfil
	$('#info_accion').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){})



}

function bt_guardar_relacion_perfil_acciones(){
	if($('input[id="seleccion_relacion_perfil"]:checked').size()==0)
	{
		alert("Debe seleccionar un perfil")
		return
	}

	if($('input[id="seleccion_relacion_accion"]:checked').size()==0)
	{
		alert("Debe seleccionar una acción")
		return
	}

	$('input[id="seleccion_relacion_perfil"]:checked').each(function(){
		var seleccion_relacion_perfil =$(this).val()

		$('input[id="seleccion_relacion_accion"]:checked').each(function(){
			var seleccion_relacion_accion =$(this).val()

			var criterios ="alea="+Math.random()+"&accion_ajax=relacion_perfil_accion&seleccion_relacion_perfil="+seleccion_relacion_perfil+"&seleccion_relacion_accion="+seleccion_relacion_accion
			
			$('#acciones').load('FuncionesAjax/man_modulos_perfiles_roles_ajax.asp', criterios, function(){})

				


		})	


	})


}


</script>
</head>
<BODY BGCOLOR='FFFFFF'>
<div class="titulo_principal Estilo13">MANTENEDOR MODULOS, PERFILES, ROLES</div>
<div class="seleccion">
	<span class="titulo_secundario" for="">Mantenedor</span> 
	<select name="opcion_mantenedor" id="opcion_mantenedor">
		<option value="">Selecciona mantenedor</option>
		<option value="modulo">Modulos</option>	
		<option value="perfil">Perfiles</option>	
		<option value="accion">Acciones</option>	
		<option value="per_accion">Relacionar perfiles y acciones</option>	
	</select>
	<span id="botonera_flotante">
		<input type="button" name="relacion_perfil_accion" id="relacion_perfil_accion" value="Relacionar perfil y acciones" onclick="bt_guardar_relacion_perfil_acciones()">
	</span>
</div>
<div id="acciones"></div>
<div id="mantenedor"></div>

</BODY>
</HTML>
<%CerrarSCG()%>



