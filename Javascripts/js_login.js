document.createElement("nav");
document.createElement("header");
document.createElement("footer");
document.createElement("section");
document.createElement("article");
document.createElement("aside");
document.createElement("hgroup");
document.createElement("figure");


$(document).ready(function(){
	verifica_intentos_fallidos()

	$('#usuario_nombre').keyup(function(){
		$('.alert_usuario').css('display', 'none')
		$('#usuario_nombre').css('border-color','')
		$('#usuario_nombre').css('border','')
		$('#span_mensaje_error').text("")
	})



	$('#contrasena').keyup(function(){
		$('.alert_contrasena').css('display', 'none')
		$('#contrasena').css('border-color','')
		$('#contrasena').css('border','')
		$('#span_mensaje_error').text("")
	})

	$('#usuario_nombre').click(function(){
		if($('#span_mensaje_error').text()!="")
		{
			$('.alert_usuario').css('display', 'none')
			$('#usuario_nombre').css('border-color','')
			$('#usuario_nombre').css('border','')
			$('#span_mensaje_error').text("")
			$('.alert_contrasena').css('display', 'none')
			$('#contrasena').css('border-color','')
			$('#contrasena').css('border','')
			$('#span_mensaje_error').text("")	

		}else{
			$('.alert_usuario').css('display', 'none')
			$('#usuario_nombre').css('border-color','')
			$('#usuario_nombre').css('border','')
			$('#span_mensaje_error').text("")			
		}

	})


	$('#contrasena').click(function(){
		if($('#span_mensaje_error').text()!="")
		{
			$('.alert_usuario').css('display', 'none')
			$('#usuario_nombre').css('border-color','')
			$('#usuario_nombre').css('border','')
			$('#span_mensaje_error').text("")
			$('.alert_contrasena').css('display', 'none')
			$('#contrasena').css('border-color','')
			$('#contrasena').css('border','')
			$('#span_mensaje_error').text("")

		}else{

			$('.alert_contrasena').css('display', 'none')
			$('#contrasena').css('border-color','')
			$('#contrasena').css('border','')
			$('#span_mensaje_error').text("")
		}
	})



	$('#login_usuario').blur(function(){
		if($(this).val()!=""){
			var criterios ="alea="+Math.random()+"&accion_ajax=verifica_usuario&login_usuario="+encodeURIComponent($(this).val())
			$('#span_CORREO_ELECTRONICO').load('System/FuncionesAjax/index_ajax.asp', criterios, function(){})
		}
			
	})

	

	$('#login_usuario').keyup(function(){
		$('#span_CORREO_ELECTRONICO').text("")	
		$('#span_CORREO_ELECTRONICO_enviado').text("")	
				
	})

	$('#login_usuario').click(function(){		
		if($('#span_CORREO_ELECTRONICO').text()!="")
		{
			$('#span_CORREO_ELECTRONICO').text("")	
			$('#span_CORREO_ELECTRONICO_enviado').text("")				
		}				
	})

	

})

function verifica_intentos_fallidos(){
	var criterios = "alea="+Math.random()+"&accion_ajax=verifica_log"
	$('#verifica_intentos_fallidos').load('System/FuncionesAjax/index_ajax.asp', criterios, function(){
		var contador_fallidos =$('#contador_fallidos').val()
		if(contador_fallidos>=3)
		{
			
			$('#bloqueo_3_fallidos').dialog({
		   		show:"blind", 
		   		hide:"explode",	        	 
		    	width:450,
		    	height:210, 
		    	modal: true
			    	
			});

			setTimeout("$('#bloqueo_3_fallidos').dialog('close')",300000)


		}
	})

}
	

function bt_validar_usuario()
{
	var usuario_nombre 	=$('#usuario_nombre').val()
	var contrasena 		=$('#contrasena').val()

	if(usuario_nombre=="")
	{
		$('.alert_usuario').css('display', 'block')
		$('#usuario_nombre').css('border-color','#F5A9A9')
		$('#usuario_nombre').css('border-width','2px')

		
	}

	if(contrasena=="")
	{
		$('.alert_contrasena').css('display', 'block')
		$('#contrasena').css('border-color','#F5A9A9')
		$('#contrasena').css('border-width','2px')
		
	}	

	if(usuario_nombre=="" || contrasena==""){
		return
	}
	var criterios ="alea="+Math.random()+"&accion_ajax=verifica_login&usuario_nombre="+encodeURIComponent(usuario_nombre)+"&contrasena="+contrasena
	$('#action_section').load('System/FuncionesAjax/index_ajax.asp', criterios, function(){
 		verifica_intentos_fallidos()
 		var usuario_validado	=$('#usuario_validado').val()
 		var primer_ingreso 		=$('#primer_ingreso').val()


 		if(usuario_validado=="S")
 		{
 			if($('input[id="login_recordarme"]').is(':checked'))
			{
				var criterios ="alea="+Math.random()+"&accion_ajax=crear_cookie&usuario_nombre="+encodeURIComponent(usuario_nombre)+"&contrasena="+contrasena
				$('#action_section').load('System/FuncionesAjax/index_ajax.asp', criterios, function(){
					if(primer_ingreso=="N")
					{
						location.href='System/default.asp'	
					}else{
						location.href='System/man_CambioClave.asp?estado=primera_vez'	
					}
					
				})

			}else{

				var criterios ="alea="+Math.random()+"&accion_ajax=eliminar_cookie"
				$('#action_section').load('System/FuncionesAjax/index_ajax.asp', criterios, function(){
					if(primer_ingreso=="N")
					{
						location.href='System/default.asp'	
					}else{
						location.href='System/man_CambioClave.asp?estado=primera_vez'	
					}
				})	
			}		
 		}else{

 			var mensaje_error = $('#mensaje_error').val()
 			$('#span_mensaje_error').text(mensaje_error)
			$('#contrasena').css('border-color','#F5A9A9')
			$('#contrasena').css('border-width','2px')
			$('#usuario_nombre').css('border-color','#F5A9A9')
			$('#usuario_nombre').css('border-width','2px')	
 		}
	

	})



}

 $(document).keydown(function(tecla){
    if (tecla.keyCode == 13) { 
       bt_validar_usuario() 
    }
});


function bt_olvida_contrasena(){
	$('#contrasena_olvidada').dialog({
   		show:"blind", 
   		hide:"explode",
   		buttons: {
            Enviar: function() {
                bt_send_email()
            },
            Cerrar: function() {
                // Cerrar ventana de diálogo
                $('#span_CORREO_ELECTRONICO').text("")
                $('#span_CORREO_ELECTRONICO_enviado').text("")
				$('#login_usuario').css('border-color','')					
				$('#login_usuario').css('border','')
				$('#login_usuario').val('')
				$(this).dialog( "close" );

            }
        },
        	 
    	width:450,
    	height:210, 
    	modal: true
	    	
	});		 	
}

function bt_send_email(){
	var login_usuario 	=$('#login_usuario').val()

	if(login_usuario!="")
	{

		var criterios ="alea="+Math.random()+"&accion_ajax=verifica_usuario&login_usuario="+encodeURIComponent(login_usuario)
		$('#span_CORREO_ELECTRONICO').load('System/FuncionesAjax/index_ajax.asp', criterios, function(){
			var accion =$('#accion').val()
			if(accion==1)
			{
				
				var criterios ="alea="+Math.random()+"&accion_ajax=envia_correo_olvidado&login_usuario="+encodeURIComponent(login_usuario)
				$('#span_CORREO_ELECTRONICO_enviado').load('System/FuncionesAjax/index_ajax.asp', criterios, function(){
					setTimeout("$('#contrasena_olvidada').dialog('close')",2500)
				})

			}

		})
		
	}else{

		$('#span_CORREO_ELECTRONICO').text("Valida nombre usuario")
		$('#forget_email').css('border-color','#FE2E2E')					
		$('#forget_email').css('border','1')		
	}
 	


}