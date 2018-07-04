<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<link href="../css/style.css" rel="stylesheet" type="text/css">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"   
    
	pagina_origen 	= request.querystring("pagina_origen")

	IF trim(pagina_origen)="cambio_clave" then
		ID_USUARIO		= session("session_idusuario")
	else
		ID_USUARIO		= request.querystring("ID_USUARIO")
	end if

	AbrirSCG()

	if trim(ID_USUARIO)<>"" then

		sql_sel="SELECT ID_USUARIO, rut_usuario, nombres_usuario, apellido_paterno, " 
		sql_sel= sql_sel & " apellido_materno, fecha_nacimiento, correo_electronico, telefono_contacto, "
		sql_sel= sql_sel & " perfil, LOGIN, CLAVE, PERFIL_ADM, perfil_cob, ACTIVO, perfil_proc, perfil_sup, "
		sql_sel= sql_sel & " PERFIL_CAJA, perfil_emp, PERFIL_FULL, perfil_back, gestionador_preventivo, "
		sql_sel= sql_sel & " anexo, observaciones_usuario, COD_AREA,EsInterno,EsExterno,PuedenEscucharMisGrabaciones,PuedoEscucharGrabaciones, CodigoAgenteElastix"
		sql_sel= sql_sel & " FROM USUARIO "
		sql_sel= sql_sel & " WHERE ID_USUARIO= " & TRIM(ID_USUARIO)

         ' Response.Write sql_sel


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
			anexo					=rs_sel("anexo")
			observaciones 			=rs_sel("observaciones_usuario")
			COD_AREA  				=rs_sel("COD_AREA")

            EsInterno  				=rs_sel("EsInterno")
            EsExterno  				=rs_sel("EsExterno")
            PuedenEscucharMisGrabaciones  			=rs_sel("PuedenEscucharMisGrabaciones")
            PuedoEscucharGrabaciones  				=rs_sel("PuedoEscucharGrabaciones")
            CodigoAgenteElastix            =rs_sel("CodigoAgenteElastix")

            ''response.Write("-->" & EsInterno & "-->" & EsExterno & "-->" & PuedenEscucharMisGrabaciones & "-->" & ACTIVO)
            'response.end

		end if
		'Response.write perfil_proc
	Else

		ID_USUARIO				=""
		rut_usuario 			=""
		nombres_usuario 		=""
		apellidos_usuario 		=""
		apellido_paterno 		=""
		apellido_materno 		=""
		fecha_nacimiento 		=""
		correo_electronico 		=""
		telefono_contacto 		=""
		perfil 					=""
		LOGIN 					=""
		CLAVE 					=""
		PERFIL_ADM 				=""
		perfil_cob 				=""
		ACTIVO 					=""
		perfil_proc 			=""
		perfil_sup 				=""
		PERFIL_CAJA 			=""
		perfil_emp 				=""
		PERFIL_FULL 			=""	
		perfil_back 			=""
		gestionador_preventivo 	=""
		anexo					=""
		observaciones 			=""
		COD_AREA 				=""

        EsInterno  				=false
        EsExterno  				=false
        PuedenEscucharMisGrabaciones  			=false
        PuedoEscucharGrabaciones  				=false
        CodigoAgenteElastix            =""    

	end if


%>

	<TITLE>Mantenedor de Usuarios</TITLE>
	<link href="../css/style_multi_select.css" rel="stylesheet"> 
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<script src="../Componentes/jquery.multiselect.js"></script>
    <script src="../Componentes/jquery.numeric/jquery.numeric.js"></script>

	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
  	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<script type="text/javascript">
$(document).ready(function(){
    $(document).tooltip();

	var ID_USUARIO 			=$('#ID_USUARIO').val()

	if (ID_USUARIO!="")
	{
		$('#RUT_USUARIO').attr('disabled', true)
		$('#num_min').val($('#TELEFONO_CONTACTO').val().length) 		
	}

	$('#FECHA_NACIMIENTO').datepicker( {changeMonth: true,changeYear: true, yearRange: '-100:+0'})

	 $("#ch_cliente").multiselect();

	 $('#CORREO_ELECTRONICO').blur(function(){
	 	if($(this).val()!=""){
		 	var RegExPattern = /[\w-\.]{3,}@([\w-]{2,}\.)*([\w-]{2,}\.)[\w-]{2,4}/

			if($(this).val().match(RegExPattern)){
		
			}else{			
				$('#span_CORREO_ELECTRONICO').text("Formato correo electronico invalido")
				$(this).css('border-color','#FE2E2E')	
				return	
			}	 		
	 	}
	 }) 

	 $('.perfiles').click(function(){
	 	$('#span_PERFIL_ADM').text("")
	 })

	 $('#NOMBRES_USUARIO').keyup(function(){		
		$('#span_NOMBRES_USUARIO').text("")
		$(this).css('border-color','')	
	 }) 

	 $('#NOMBRES_USUARIO').click(function(){
	 	if($('#span_NOMBRES_USUARIO').text()!="")
	 	{
			$('#span_NOMBRES_USUARIO').text("")
			$(this).css('border-color','')	
			$(this).val("")		 		
	 	}
	 	
	 }) 

	 $('#APELLIDO_PATERNO').keyup(function(){		
		$('#span_APELLIDO_PATERNO').text("")
		$(this).css('border-color','')	
	 }) 

	 $('#APELLIDO_PATERNO').click(function(){
	 	if($('#span_APELLIDO_PATERNO').text()!="")
	 	{
			$('#span_APELLIDO_PATERNO').text("")
			$(this).css('border-color','')	
			$(this).val("") 		
	 	}		
	
	 }) 


	 $('#APELLIDO_MATERNO').keyup(function(){		
		$('#span_APELLIDO_MATERNO').text("")
		$(this).css('border-color','')	
	 }) 

	 $('#APELLIDO_MATERNO').click(function(){
	 	if($('#span_APELLIDO_MATERNO').text()!="")
	 	{
	 		$('#span_APELLIDO_MATERNO').text("")
			$(this).css('border-color','')	
			$(this).val("")
	 	}		
			
	 }) 

	 $('#CORREO_ELECTRONICO').keyup(function(){		
			var criterios ="alea="+Math.random()+"&accion_ajax=verifica_correo&CORREO_ELECTRONICO="+$('#CORREO_ELECTRONICO').val()
			$('#div_verifica_correo').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){
				var verifica_correo = $('#verifica_correo').val()

				if (verifica_correo=="N"){
					$('#span_CORREO_ELECTRONICO').text("Correo electronico ya existe en el sistema")
					$(this).css('border-color','#FE2E2E')	
				}else{
					$('#span_CORREO_ELECTRONICO').text("")
					$(this).css('border-color','')
				}

			})
	 }) 

	 $('#CORREO_ELECTRONICO').click(function(){		
	 	if($('#span_CORREO_ELECTRONICO').text()!=""){
	 		$('#span_CORREO_ELECTRONICO').text("")
			$(this).css('border-color','')	
			$(this).val("") 		
	 	}	

	 }) 


	 $('#RUT_USUARIO').keyup(function(){		
		$('#span_RUT_USUARIO').text("")
		$(this).css('border-color','')	
	 }) 

	 $('#RUT_USUARIO').click(function(){
		if($('#span_RUT_USUARIO').text()!=""){ 		
			$('#span_RUT_USUARIO').text("")
			$(this).css('border-color','')
			$(this).val("") 	
		}
	 }) 	 

	 $('#FECHA_NACIMIENTO').blur(function(){		
		$('#span_FECHA_NACIMIENTO').text("")
		$(this).css('border-color','')	
	 }) 

	 $('#FECHA_NACIMIENTO').click(function(){	
	 	if($('#span_FECHA_NACIMIENTO').text()!=""){ 	
			$('#span_FECHA_NACIMIENTO').text("")
			$(this).css('border-color','')	
			$(this).val("") 	
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

	$('#CodigoAgenteElastix').keyup(function () {
	    if (isNaN($(this).val())) {
	        $('#span_CodigoAgenteElastix').text(" Formato código agente discador Elastix invalido, favor solo ingresar datos numéricos.")
	        $(this).css('border-color', '#FE2E2E')
	        return
	    } else {
	        $('#span_CodigoAgenteElastix').text("")
	        $(this).css('border-color', '')
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


	 $('#TELEFONO_CONTACTO').click(function(){
		if($('#span_TELEFONO_CONTACTO').text()!=""){
			$('#span_TELEFONO_CONTACTO').text("")
			$(this).css('border-color','')	
			$(this).val("") 			
		}

	 }) 

	 $('#ANEXO').click(function(){
	 	if($('#span_ANEXO').text()!=""){
	 		$('#span_ANEXO').text("")
			$(this).css('border-color','')	
			$(this).val("") 
	 	}			 	

	 }) 


	 $('#RUT_USUARIO').blur(function(){	
	 	
	 	if($('#span_RUT_USUARIO').text()!="RUT invalido"){

			var criterios ="alea="+Math.random()+"&accion_ajax=verifica_rut&RUT_USUARIO="+$('#RUT_USUARIO').val()
			$('#div_mantenedor_usuario').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){
				var verifica_rut = $('#verifica_rut').val()

				if (verifica_rut=="N"){
					$('#span_RUT_USUARIO').text("RUT ya existe en el sistema")
					$('#RUT_USUARIO').css('border-color','#FE2E2E')	
				}else{
					$('#span_RUT_USUARIO').text("")
					$('#RUT_USUARIO').css('border-color','')
				}

			})
	 	}

	 }) 


	$('select[id="ch_cliente"]').change(function(){
		$('#span_CH_CLIENTE').text("")	
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

})

function bt_guardar_usuario()
{
	var ID_USUARIO 			=$('#ID_USUARIO').val()
	var TELEFONO_CONTACTO 	=$('#TELEFONO_CONTACTO').val()
	var CORREO_ELECTRONICO 	=$('#CORREO_ELECTRONICO').val()
	var FECHA_NACIMIENTO 	=$('#FECHA_NACIMIENTO').val()
	var RUT_USUARIO 		=$('#RUT_USUARIO').val()
	var APELLIDO_MATERNO 	=$('#APELLIDO_MATERNO').val()
	var APELLIDO_PATERNO 	=$('#APELLIDO_PATERNO').val()
	var NOMBRES_USUARIO 	=$('#NOMBRES_USUARIO').val()
	var PERFIL_ADM 			=$('input[id="PERFIL_ADM"]:checked').val()
	var PERFIL_EMP 			=$('input[id="PERFIL_EMP"]:checked').val()
	var PERFIL_FULL 		=$('input[id="PERFIL_FULL"]:checked').val()
	var PERFIL_PROC 		=$('input[id="PERFIL_PROC"]:checked').val()
	var PERFIL_CAJA			=$('input[id="PERFIL_CAJA"]:checked').val()
	var PERFIL_COB 			=$('input[id="PERFIL_COB"]:checked').val()
	var PERFIL_SUP 			=$('input[id="PERFIL_SUP"]:checked').val()
	var ACTIVO 				=$('input[id="ACTIVO"]:checked').val()   
	var COD_AREA            =$('#COD_AREA').val() 
	var ANEXO            	=$('#ANEXO').val()
	var OBSERVACIONES       = $('#OBSERVACIONES').val()


	var EsInterno = $('input[id="EsInterno"]:checked').val()
	var EsExterno = $('input[id="EsExterno"]:checked').val() 
	var PuedenEscucharMisGrabaciones =$('input[id="PuedenEscucharMisGrabaciones"]:checked').val()
	var PuedoEscucharGrabaciones = $('input[id="PuedoEscucharGrabaciones"]:checked').val()

	var CodigoAgenteElastix = $('#CodigoAgenteElastix').val()


	if ((EsInterno == EsExterno) && (EsInterno == 1))
    {
        alert("Usuario No Puede tener los mismos Valores en interno o externo");
        return;
    }


	if(PERFIL_ADM==null)
	{
		PERFIL_ADM=""
	}

	if(PERFIL_EMP==null)
	{
		PERFIL_EMP=""
	}

	if(PERFIL_FULL==null)
	{
		PERFIL_FULL=""
	}

	if(PERFIL_PROC==null)
	{
		PERFIL_PROC=""
	}

	if(PERFIL_CAJA==null)
	{
		PERFIL_CAJA=""
	}

	if(PERFIL_COB==null)
	{
		PERFIL_COB=""
	}

	if(PERFIL_SUP==null)
	{
		PERFIL_SUP=""
	}					


	if(NOMBRES_USUARIO=="")
	{
		$('#span_NOMBRES_USUARIO').text("Ingresa nombre de usuario")
		$('#NOMBRES_USUARIO').css('border-color','#FE2E2E')
	}
	if(APELLIDO_PATERNO=="")
	{
		$('#span_APELLIDO_PATERNO').text("Ingresa apellido paterno")
		$('#APELLIDO_PATERNO').css('border-color','#FE2E2E')
			}
	if(APELLIDO_MATERNO=="")
	{
		$('#span_APELLIDO_MATERNO').text("Ingresa apellido materno")
		$('#APELLIDO_MATERNO').css('border-color','#FE2E2E')
			}

	if(RUT_USUARIO=="")
	{
		$('#span_RUT_USUARIO').text("Ingresa RUT usuario")
		$('#RUT_USUARIO').css('border-color','#FE2E2E')
	}

	if(FECHA_NACIMIENTO=="")
	{
		$('#span_FECHA_NACIMIENTO').text("Inrgesa fecha nacimiento")
		$('#FECHA_NACIMIENTO').css('border-color','#FE2E2E')
		
	}


	if(TELEFONO_CONTACTO=="")
	{
		$('#span_TELEFONO_CONTACTO').text("Ingresa telefono contacto")
		$('#TELEFONO_CONTACTO').css('border-color','#FE2E2E')
		
	}

	if(COD_AREA==0)
	{
		$('#span_TELEFONO_CONTACTO').text("Ingresa Código de área")
		$('#COD_AREA').css('border-color','#FE2E2E')
		
	}


	if(CORREO_ELECTRONICO=="")
	{
		$('#span_CORREO_ELECTRONICO').text("Ingresa correo electronico")
		$('#CORREO_ELECTRONICO').css('border-color','#FE2E2E')
		
	}

	if($('select[id="ch_cliente"]').val()==null)
	{
		$('#span_CH_CLIENTE').text("seleccionar al menos 1 cliente")		
		
	}

	if(PERFIL_ADM=="" && PERFIL_EMP=="" && PERFIL_FULL=="" && PERFIL_PROC=="" && PERFIL_CAJA=="" && PERFIL_COB=="" && PERFIL_SUP=="")
	{
		$('#span_PERFIL_ADM').text("Debe seleccionar al menos 1 perfil de usuario")

	}else{
		$('#span_PERFIL_ADM').text("")
	}	

	var RegExPattern 	= /[\w-\.]{3,}@([\w-]{2,}\.)*([\w-]{2,}\.)[\w-]{2,4}/
	if(CORREO_ELECTRONICO.match(RegExPattern)){

	}else{			
		$('#span_CORREO_ELECTRONICO').text("Formato correo electronico invalido")
		$('#CORREO_ELECTRONICO').css('border-color','#FE2E2E')
		return	
	}


	if(NOMBRES_USUARIO=="" || APELLIDO_PATERNO=="" || APELLIDO_MATERNO=="" || RUT_USUARIO=="" || FECHA_NACIMIENTO=="" || TELEFONO_CONTACTO=="" || CORREO_ELECTRONICO=="" )
	{
		return
	}



	if(PERFIL_ADM=="" && PERFIL_EMP=="" && PERFIL_FULL=="" && PERFIL_PROC=="" && PERFIL_CAJA=="" && PERFIL_COB=="" && PERFIL_SUP=="")
	{
		return

	}	
	if($('select[id="ch_cliente"]').val()==null)
	{
		return
	}

	if (isNaN($('#TELEFONO_CONTACTO').val())) {
		return 
	}


	if (isNaN($('#ANEXO').val())) {
		return 
	}

	if (isNaN($('#CodigoAgenteElastix').val())) {
	    return
	}
		
	var verifica_asignacion = $('#verifica_asignacion').val()
	if(verifica_asignacion=="S"){
		$('#span_ACTIVO').text("El usuario que intenta dejar No Valido actualmente posee deuda asignada, favor reasignar deuda previa No Validación")
		$('#span_ACTIVO').css('border-color','#FE2E2E')
		return
	}else{
		$('#span_ACTIVO').text("")
		$('#span_ACTIVO').css('border-color','')
	}

 	var num_min = $('#num_min').val() 	

	if($('#TELEFONO_CONTACTO').val().length != num_min) {
	    $('#span_TELEFONO_CONTACTO').text("Fono debe tener " + num_min + " digitos")
		$('#TELEFONO_CONTACTO').css('border-color','#FE2E2E')
		return			
	}

	if (ID_USUARIO=="")
	{		
		var criterios ="alea="+Math.random()+"&accion_ajax=verifica_correo&CORREO_ELECTRONICO="+CORREO_ELECTRONICO
		$('#div_verifica_correo').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){
		})
		var verifica_correo = $('#verifica_correo').val()

		if(verifica_correo=="N"){
			$('#span_CORREO_ELECTRONICO').text("Correo electronico ya existe en el sistema")
			$('#CORREO_ELECTRONICO').css('border-color','#FE2E2E')	
			return
		}else{
			$('#span_CORREO_ELECTRONICO').text("")
			$('#CORREO_ELECTRONICO').css('border-color','')
		}

		var criterios ="alea="+Math.random()+"&accion_ajax=verifica_rut&RUT_USUARIO="+RUT_USUARIO
		$('#div_mantenedor_usuario').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){
			var verifica_rut =$('#verifica_rut').val()
			if(verifica_rut=="N")
			{
				$('#span_RUT_USUARIO').text("RUT ya existe en el sistema")
				$('#RUT_USUARIO').css('border-color','#FE2E2E')
				return

			}else{

			    var criteriosNuevos = "alea=" + Math.random() + "&accion_ajax=guardar_usuario&TELEFONO_CONTACTO=" + TELEFONO_CONTACTO + "&CORREO_ELECTRONICO=" + encodeURIComponent(CORREO_ELECTRONICO) + "&FECHA_NACIMIENTO=" + FECHA_NACIMIENTO + "&RUT_USUARIO=" + RUT_USUARIO + "&APELLIDO_MATERNO=" + encodeURIComponent(APELLIDO_MATERNO) + "&APELLIDO_PATERNO=" + encodeURIComponent(APELLIDO_PATERNO) + "&NOMBRES_USUARIO=" + encodeURIComponent(NOMBRES_USUARIO) + "&PERFIL_ADM=" + PERFIL_ADM + "&PERFIL_EMP=" + PERFIL_EMP + "&PERFIL_FULL=" + PERFIL_FULL + "&PERFIL_PROC=" + PERFIL_PROC + "&PERFIL_CAJA=" + PERFIL_CAJA + "&PERFIL_COB=" + PERFIL_COB + "&PERFIL_SUP=" + PERFIL_SUP + "&ACTIVO=" + ACTIVO + "&COD_AREA=" + COD_AREA + "&ANEXO=" + encodeURIComponent(ANEXO) + "&OBSERVACIONES=" + encodeURIComponent(OBSERVACIONES) + "&EsInterno=" + encodeURIComponent(EsInterno) + "&EsExterno=" + encodeURIComponent(EsExterno) + "&PuedenEscucharMisGrabaciones=" + encodeURIComponent(PuedenEscucharMisGrabaciones) + "&PuedoEscucharGrabaciones=" + encodeURIComponent(PuedoEscucharGrabaciones) + "&CodigoAgenteElastix=" + encodeURIComponent(CodigoAgenteElastix)
			
                $('#RUT_USUARIO').attr('disabled', true)
                
                guardar_usuario_cliente(criteriosNuevos);				
			}
		})
	}

	if (ID_USUARIO!="")
	{
	    var criteriosActualizar = "alea=" + Math.random() + "&accion_ajax=actualiza_usuario&TELEFONO_CONTACTO=" + TELEFONO_CONTACTO + "&CORREO_ELECTRONICO=" + encodeURIComponent(CORREO_ELECTRONICO) + "&FECHA_NACIMIENTO=" + FECHA_NACIMIENTO + "&APELLIDO_MATERNO=" + encodeURIComponent(APELLIDO_MATERNO) + "&APELLIDO_PATERNO=" + encodeURIComponent(APELLIDO_PATERNO) + "&NOMBRES_USUARIO=" + encodeURIComponent(NOMBRES_USUARIO) + "&PERFIL_ADM=" + PERFIL_ADM + "&PERFIL_EMP=" + PERFIL_EMP + "&PERFIL_FULL=" + PERFIL_FULL + "&PERFIL_PROC=" + PERFIL_PROC + "&PERFIL_CAJA=" + PERFIL_CAJA + "&PERFIL_COB=" + PERFIL_COB + "&PERFIL_SUP=" + PERFIL_SUP + "&ACTIVO=" + ACTIVO + "&ID_USUARIO=" + ID_USUARIO + "&COD_AREA=" + COD_AREA + "&ANEXO=" + ANEXO + "&OBSERVACIONES=" + encodeURIComponent(OBSERVACIONES) + "&EsInterno=" + encodeURIComponent(EsInterno) + "&EsExterno=" + encodeURIComponent(EsExterno) + "&PuedenEscucharMisGrabaciones=" + encodeURIComponent(PuedenEscucharMisGrabaciones) + "&PuedoEscucharGrabaciones=" + encodeURIComponent(PuedoEscucharGrabaciones) + "&CodigoAgenteElastix=" + encodeURIComponent(CodigoAgenteElastix)

	    guardar_usuario_cliente(criteriosActualizar);
	}
}

function guardar_usuario_cliente(criteriosInput)
{
    $('#div_mantenedor_usuario').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criteriosInput,
	function () {

		var codigoCliente = []

		var ID_USUARIO = $('#ID_USUARIO').val()

		$('select[id="ch_cliente"] option:checked').each(function () {

			codigoCliente.push($(this).val())
		})

		criteriosInput = "alea=" + Math.random() + "&accion_ajax=guardar_usuario_cliente&ID_USUARIO=" + ID_USUARIO + "&COD_CLIENTE=" + codigoCliente

		$('#div_mantenedor_usuario_cliente').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criteriosInput, function () { })

		alert("Datos modificado correctamente")

		location.href = 'man_Usuario.asp'
	})
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

function bt_verificar_asignacion(){
	var ID_USUARIO =$('#ID_USUARIO').val()

	var criterios = "alea="+Math.random()+"&accion_ajax=verifica_asignaciones&ID_USUARIO="+ID_USUARIO
	$('#div_verifica_asignaciones').load('FuncionesAjax/man_UsuarioForm_ajax.asp', criterios, function(){

		var verifica_asignacion = $('#verifica_asignacion').val()
		if(verifica_asignacion=="S"){
			$('#span_ACTIVO').text("El usuario que intenta dejar No Valido actualmente posee deuda asignada, favor reasignar deuda previa No Validación")
			$('#span_ACTIVO').css('border-color','#FE2E2E')
		}else{
			$('#span_ACTIVO').text("")
			$('#span_ACTIVO').css('border-color','')
		}
	})

}

function bt_verificar_asignacion2(){

	$('#span_ACTIVO').text("")
	$('#span_ACTIVO').css('border-color','')

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


	.hdr_i{
		background-color: #C9DEF2; 		
		color: #000;
		font-size: 12px;
		font-weight: normal;
		font-family: Verdana;	
	}


	input[type="text"]{
		font-size: 12px;
		padding: 1px;
		width:200px;
	}

	.aviso_rojo{
		color:#FE2E2E;
		font-size:12px ;
		font-weight: bold;
	}

	#ch_cliente{
		width:270px;
	}

	.input_usuario{
		width: 230px;

	}

	.input_usuario_tel{
		width: 144px;
	}

	.input_usuario_ob{
		width: 265px;
	}

</style>

</HEAD>

<BODY BGCOLOR='FFFFFF'>

<input name="num_min" 			id="num_min" 			type="hidden" 	value="0">
<DIV class="titulo_informe">MANTENCIÓN DE USUARIOS</DIV>
<FORM NAME="mantenedorForm"  action="man_UsuarioAction.asp" method="POST" >
<br>

<table width="90%" BORDER="0" CELLPADDING="0" CELLSPACING="0" align="center" CLASS="estilo_columnas">
	<thead>
 	<tr >
		<td colspan="2" height="22">&nbsp;&nbsp;Datos personales</td>
	</TR>
	</thead>
 	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i"><font class="aviso_rojo">* </font>Nombres</td>
		<td class="td_t"><input type="text" class="input_usuario" name="NOMBRES_USUARIO" id="NOMBRES_USUARIO" value="<%=trim(NOMBRES_USUARIO)%>"><span id="span_NOMBRES_USUARIO" class="span_aviso_rojo"></span></td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i"><font class="aviso_rojo">* </font>Apellido paterno</Font></td>
		<td class="td_t"><input type="text" class="input_usuario" name="APELLIDO_PATERNO" id="APELLIDO_PATERNO" value="<%=trim(APELLIDO_PATERNO)%>"><span id="span_APELLIDO_PATERNO" class="span_aviso_rojo"></span></td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i"><font class="aviso_rojo">* </font>Apellido materno</Font></td>
		<td class="td_t"><input type="text" class="input_usuario" name="APELLIDO_MATERNO" id="APELLIDO_MATERNO" value="<%=trim(APELLIDO_MATERNO)%>"><span id="span_APELLIDO_MATERNO" class="span_aviso_rojo"></span></td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i"><font class="aviso_rojo">* </font>Rut</Font></td>
		<td class="td_t"><input type="text" class="input_usuario" name="RUT_USUARIO" id="RUT_USUARIO" onblur="ValidaRut(this)" value="<%=trim(RUT_USUARIO)%>"><span id="span_RUT_USUARIO" class="span_aviso_rojo"></span></td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i"><font class="aviso_rojo">* </font>Fecha nacimiento</Font></td>
		<td class="td_t"><input type="text" class="input_usuario" readonly name="FECHA_NACIMIENTO" id="FECHA_NACIMIENTO" value="<%=trim(FECHA_NACIMIENTO)%>"><span id="span_FECHA_NACIMIENTO" class="span_aviso_rojo"></span></td>
	</TR>	
	</td>
</TR>
</TABLE>

<br>

<table width="90%" BORDER="0" CELLPADDING="0" CELLSPACING="0" align="center" CLASS="estilo_columnas">
	<thead>
 	<tr>
		<td colspan="2" height="22">&nbsp;&nbsp;Datos de empresa </td>
	</TR>
	</thead>	
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i"><font class="aviso_rojo">* </font>Correo electr&oacute;nico</Font></td>
		<td class="td_t"><input type="text" class="input_usuario" name="CORREO_ELECTRONICO" id="CORREO_ELECTRONICO" value="<%=trim(CORREO_ELECTRONICO)%>"><span id="span_CORREO_ELECTRONICO" class="span_aviso_rojo"></span></td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i"><font class="aviso_rojo">* </font>Tel&eacute;fono contacto</Font></td>
		<td class="td_t ">
			<select name="COD_AREA"  id="COD_AREA" onchange="asigna_minimo_a_variable(this.value,0)" style="width:50px;">
				<option value="0">--</option>
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

			
			</select>
			(CEL.9)			
			<input type="text" style="width:103px;" class="input_usuario_tel" name="TELEFONO_CONTACTO" id="TELEFONO_CONTACTO" value="<%=trim(TELEFONO_CONTACTO)%>">
			<span id="span_TELEFONO_CONTACTO" class="span_aviso_rojo"></span>
		</td>
	</TR>

	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;&nbsp;Anexo</Font></td>
		<td class="td_t"><input  <%if trim(ANEXO)<>"" then%> title="<%=trim(ANEXO)%>" <%end if%> type="text" class="input_usuario" name="ANEXO" id="ANEXO" title="<%=trim(ANEXO)%>" value="<%=trim(ANEXO)%>"><span id="span_ANEXO" class="span_aviso_rojo"></span></td>
	</TR>

	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;&nbsp;Activo</Font></td>
		<td class="td_t">

			<%if trim(ACTIVO)="" then%>
				<input type="radio" checked name="ACTIVO" id="ACTIVO" value="1">Si 			
			<%else%>
				<input type="radio" <%if trim(ACTIVO)<>"" then%> checked <%end if%> name="ACTIVO" id="ACTIVO" <%if ACTIVO=true then Response.write " checked "  end if%> value="1" onclick="bt_verificar_asignacion2()">Si 			
				<input type="radio" name="ACTIVO" id="ACTIVO" <%if ACTIVO=false then Response.write " checked "  end if%> value="0" onclick="bt_verificar_asignacion()">No
			<%end if%><span id="span_ACTIVO" class="span_aviso_rojo"></span>
		</td>
	</tr>
	<tr >
		<td width="250" class="hdr_i"><font class="aviso_rojo">* </font>Asignar cliente</Font></td>
		<td class="">
			<%
				strSql = "SELECT COD_CLIENTE, RAZON_SOCIAL FROM CLIENTE WHERE ACTIVO = 1"
				set rsEmpresa= Conn.execute(strSql)
			%>
			<select name="ch_cliente" id="ch_cliente" multiple style="font-size:12px;">
				<%Do While not rsEmpresa.eof
					if trim(ID_USUARIO)<>"" then
						strSql = "SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & ID_USUARIO & " AND COD_CLIENTE = " & rsEmpresa("COD_CLIENTE")
						set rsUsuarioEmpresa= Conn.execute(strSql)
						If Not rsUsuarioEmpresa.Eof Then
							strChecked = " selected "
						Else
							strChecked = ""
						End If
					END IF	
				%>
					<option <%=strChecked%> value="<%=rsEmpresa("COD_CLIENTE")%>"><%=rsEmpresa("RAZON_SOCIAL")%></option>
				<%rsEmpresa.movenext
				loop%>
			</select><span id="span_CH_CLIENTE" class="span_aviso_rojo"></span>
		</td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Observación y/o comentario</Font></td>
		<td class="td_t" >
			<textarea rows="2" <% if trim(OBSERVACIONES)<>"" then %> title="<%=trim(OBSERVACIONES)%>" <%end if%> class="input_usuario_ob" name="OBSERVACIONES" id="OBSERVACIONES" ><%=trim(OBSERVACIONES)%></textarea>
		</td>
	</tr>
</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Usuario interno&nbsp;</td>
		<td class="td_t" >
           <input type="radio" name="EsInterno" id="EsInterno" <%if EsInterno=true then Response.write " checked "  end if%> value="1">Si 			
           <input type="radio" name="EsInterno" id="EsInterno" <%if EsInterno=false then Response.write " checked "  end if%> value="0" >No
		</td>
	</tr>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Usuario externo&nbsp;</td>
		<td class="td_t" >
           <input type="radio"  name="EsExterno" id="EsExterno"  <%if EsExterno=true then Response.write " checked "  end if%> value="1">Si 			
           <input type="radio" name="EsExterno" id="EsExterno"   <%if EsExterno=false then Response.write " checked "  end if%> value="0" >No

		</td>
	</tr>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Pueden escuchar mis grabaciones&nbsp;</td>
		<td class="td_t" >
         <input type="radio"  name="PuedenEscucharMisGrabaciones" id="PuedenEscucharMisGrabaciones"  <%if PuedenEscucharMisGrabaciones=true then Response.write " checked "  end if%> value="1">Si 			
         <input type="radio" name="PuedenEscucharMisGrabaciones" id="PuedenEscucharMisGrabaciones"   <%if PuedenEscucharMisGrabaciones=false then Response.write " checked "  end if%> value="0" >No
		</td>
	</tr>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Puedo escuchar grabaciones&nbsp;</td>
		<td class="td_t" >
         <input type="radio"  name="PuedoEscucharGrabaciones" id="PuedoEscucharGrabaciones"  <%if PuedoEscucharGrabaciones=true then Response.write " checked "  end if%> value="1">Si 			
         <input type="radio" name="PuedoEscucharGrabaciones" id="PuedoEscucharGrabaciones"   <%if PuedoEscucharGrabaciones=false then Response.write " checked "  end if%> value="0" >No
		</td>
	</tr>
     <tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;C&oacute;digo agente discador Elastix</Font></td>
		<td class="td_t" >		
            <input type="text" <% if trim(CodigoAgenteElastix)<>"" then %> title="<%=trim(CodigoAgenteElastix)%>" <%end if%> class="input_usuario" name="CodigoAgenteElastix" id="CodigoAgenteElastix" value="<%=trim(CodigoAgenteElastix)%>"><span id="span_CodigoAgenteElastix" class="span_aviso_rojo"></span></td>
        </td>
	</tr>    
	</TR>
    
</TABLE>

<br>

<table width="90%" BORDER="0" CELLPADDING="0" CELLSPACING="0" align="center" CLASS="estilo_columnas">
	<thead>
 	<tr BGCOLOR="#FFFFFF">
		<td height="22" class="" colspan="2"><font class="aviso_rojo">* </font>Perfiles</td>
	</TR>	
	</thead>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Perfil Administrador</Font></td>
		<td class="td_t">
			<input type="radio" class="perfiles" name="PERFIL_ADM" id="PERFIL_ADM" <%if PERFIL_ADM= true then Response.write " checked " end if%> value="1">Si 
			<input type="radio" class="perfiles" name="PERFIL_ADM" id="PERFIL_ADM" <%if PERFIL_ADM= false then Response.write " checked " end if%> value="0">No <span id="span_PERFIL_ADM" class="span_aviso_rojo"></span>
		</td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Perfil Supervisor</Font></td>
		<td class="td_t">
			<input type="radio" class="perfiles" name="PERFIL_SUP" id="PERFIL_SUP" <%if PERFIL_SUP = true then Response.write " checked " end if%> value="1">Si 
			<input type="radio" class="perfiles" name="PERFIL_SUP" id="PERFIL_SUP" <%if PERFIL_SUP = false then Response.write " checked " end if%> value="0">No
		</td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Perfil Cobrador</Font></td>
		<td class="td_t">
			<input type="radio" class="perfiles" name="PERFIL_COB" id="PERFIL_COB" <%if PERFIL_COB=true then Response.write " checked " end if%> value="1">Si 
			<input type="radio" class="perfiles" name="PERFIL_COB" id="PERFIL_COB" <%if PERFIL_COB=false then Response.write " checked " end if%> value="0">No
		</td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Perfil Caja</Font></td>
		<td class="td_t">
			<input type="radio" class="perfiles" name="PERFIL_CAJA" id="PERFIL_CAJA" <%if PERFIL_CAJA=true then Response.write " checked " end if%> value="1">Si 
			<input type="radio" class="perfiles" name="PERFIL_CAJA" id="PERFIL_CAJA" <%if PERFIL_CAJA=false then Response.write " checked " end if%> value="0">No
		</td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Perfil Judicial</Font></td>
		<td class="td_t">
			<input type="radio" class="perfiles" name="PERFIL_PROC" id="PERFIL_PROC" <%if PERFIL_PROC=true then Response.write " checked " end if%> value="1">Si 
			<input type="radio" class="perfiles" name="PERFIL_PROC" id="PERFIL_PROC" <%if PERFIL_PROC=false then Response.write " checked " end if%> value="0">No
		</td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Perfil Full</Font></td>
		<td class="td_t">
			<input type="radio" class="perfiles" name="PERFIL_FULL" id="PERFIL_FULL" <%if PERFIL_FULL=true then Response.write " checked " end if%> value="1">Si 
			<input type="radio" class="perfiles" name="PERFIL_FULL" id="PERFIL_FULL" <%if PERFIL_FULL=false then Response.write " checked " end if%> value="0">No
		</td>
	</TR>
	<tr BGCOLOR="#FFFFFF">
		<td width="250" class="hdr_i">&nbsp;&nbsp;Perfil Cliente</Font></td>
		<td class="td_t">
			<input type="radio" class="perfiles" name="PERFIL_EMP" id="PERFIL_EMP" <%if PERFIL_EMP=true then Response.write " checked " end if%> value="1">Si 
			<input type="radio" class="perfiles" name="PERFIL_EMP" id="PERFIL_EMP" <%if PERFIL_EMP=false then Response.write " checked " end if%> value="0">No
		</td>
	</TR>



</table>
<br>

<table width="90%" border="0" align="center">
     <TR>
	  <td align="right">
	   <INPUT TYPE="BUTTON" class="fondo_boton_100" value="Guardar" name="BT_GUARDAR" ID="BT_GUARDAR" onClick="bt_guardar_usuario()">
	   </TD>
	  </TD>
    </TR>
</table>
<div id="div_verifica_correo"></div>
<div id="div_mantenedor_usuario_cliente"></div>

<div id="div_mantenedor_usuario">
	<INPUT TYPE="hidden" NAME="ID_USUARIO" ID="ID_USUARIO" VALUE="<%=trim(ID_USUARIO)%>">
</div>
<div id="div_verifica_asignaciones"></div>
</FORM>

<%CerrarSCG()%>

</BODY>
</HTML>

<!--#include file="../lib/comunes/rutinas/validarRut.inc"-->
