<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">

    <!--#include file="arch_utils.asp"-->
    <!--#include file="sesion.asp"-->
    <!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

    <title>Acceso e-Rec de Llacruz</title>
    <meta name="description" content="Acceso e-Rec de Llacruz">
    <meta name="author" content="Departamento desarrollo Llacruz">
    <link href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css" rel="stylesheet"> 
    <link href="../css/normalize.css" rel="stylesheet"> 
    <link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet"> 
    <link href="../css/style_generales_sistema.css" rel="stylesheet">
    <script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>  
    <script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>  
    <script src="../Componentes/Timepicker/jquery.timepickerinputmask.min.js"></script> 
	<link href="../css/style_Carga_Gestiones_Masiva.css" rel="stylesheet">
    <script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 

<%


Response.CodePage=65001
Response.charset ="utf-8"

    abrirscg()

			'response.write(strSql)
            strSql ="EXEC PROC_GET_DATOS_CLIENTE " &  session("ses_codcli")  
			set RsCli=Conn.execute(strSql)
			
            If not RsCli.eof then
                intUsaCobInterna = RsCli("USA_COB_INTERNA")
            End if
            RsCli.close
            set RsCli=nothing

    cerrarscg()

If TraeSiNo(session("perfil_emp")) = "Si" and strCobranza = "" and intUsaCobInterna = "1" Then
    strCobranza="INTERNA"
ElseIf TraeSiNo(session("perfil_emp")) = "No" and strCobranza = "" then
    strCobranza="EXTERNA"
End If
If TraeSiNo(session("perfil_emp")) = "Si" Then
    intVerEjecutivos="0"
End If
If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then
    sinCbUsario="0"
End If
intVerCobExt = "1"
%> 

    <script type="text/javascript">

        function bt_actualiza_usuario(id_campo, valor)
        {
            var var_accion  ="#accion_usuario_"+id_campo
            var CB_COBRANZA =$('#CB_COBRANZA').val()

            var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_campo&id_campo="+id_campo+"&valor_input="+valor+"&campo=usuario&CB_COBRANZA="+CB_COBRANZA
            $(var_accion).load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){})
        }

        function bt_edita(campo, id_campo, valor){
           
            var var_accion          ="#accion_"+campo+"_"+id_campo
            var CB_COBRANZA         =$('#CB_COBRANZA').val()
            var ck_dv_rut           =$('input[id="ck_dv_rut"]:checked').val()
            var tipo_gestion        =$('#tipo_gestion').val()                
            var var_tipo_gestion    =tipo_gestion.split("-")
            var cod_tipo_gestion    =var_tipo_gestion[0]
            var tipo                =var_tipo_gestion[1]            

            var criterios   ="alea="+Math.random()+"&accion_ajax=modifica_accion&id="+id_campo+"&campo="+campo+"&valor="+valor+"&CB_COBRANZA="+CB_COBRANZA+"&ck_dv_rut="+ck_dv_rut

            $(var_accion).load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){ 

                var input_focus ="#input_"+campo+"_"+id_campo

                $(input_focus).focus()    

                if(campo=="usuario")
                {
                     $('input[name="id_usuario"]').change(function(){

                        var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_campo&id_campo="+id_campo+"&valor_input="+$(this).val()+"&campo="+campo+"&CB_COBRANZA="+CB_COBRANZA
                        $(var_accion).load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){})

                    })

                }else if(campo=="fecha"){
                    $(input_focus).datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

                }else{

                     $('input[id*="input"]').blur(function(){

                        var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_campo&id_campo="+id_campo+"&valor_input="+$(this).val()+"&campo="+campo+"&CB_COBRANZA="+CB_COBRANZA+"&tipo="+tipo
                        $(var_accion).load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){})

                    })                
                }
    

            })

        }

        function bt_elimina(id_campo){
           // alert(campo+' '+id_campo)
            if(confirm("Se eliminará el registro de la gestion (Código "+id_campo+") ¿Desea eliminar el registro?")){
                var creiterios ="alea="+Math.random()+"&accion_ajax=elimina_registro&id_campo="+id_campo
                $('#elimina_registro_gestion').load('FuncionesAjax/asigna_masiva_ajax.asp', creiterios, function(){
                    var var_accion ="#tr_"+id_campo
                    //alert(var_accion)
                    $(var_accion).remove()
                })
            }
        }  
		

        $(document).ready(function(){

            $('#limpia_campo').css('display','none')
            $('#procesar_ingresa').css('display','none')
			$('#procesar_elimina').css('display','none')

            $.ajaxSetup({ cache:false });

            $('input[id="ck_activa_id_email"]').attr('disabled', true) 
            $.prettyLoader();

        	
			/*boton eliminar*/			
			$('#procesar_elimina').click(function()
			{
			
			
			})	
			
					
			$('#limpia_campo').click(function()
			{
			
				if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?"))
				{

                    var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                    $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

							$('#td_seleect_insert').text("")
							$('#td_seleect').text("")
							$('#textarea_id_usuario').removeAttr('disabled')
							$('#textarea_hora_ingreso').removeAttr('disabled')
							$('#textarea_observacion').removeAttr('disabled')
							$('#textarea_fecha_ingreso').removeAttr('disabled')
							$('#textarea_id_medio').removeAttr('disabled')
							$('#textarea_email').removeAttr('disabled')


							$('#procesar_ingresa').css('display','none')
							$('#procesar_elimina').css('display','none')
							$('#procesar').css('display','inline-block') 
							$('#limpia_campo').css('display','none')              
																				  
                    })
                }
			}
			
			)

            $('input[id*="text_observacion_"]').live('blur', function() {
                var error_observacion       =0
                var var_atributo            =$(this).attr("id").split("_")
                var contador_obs              =var_atributo[2]
                var var_div_observacion    ="#div_observacion_"+contador_obs

                if($(this).val().length>10){                                       
                    error_observacion =1                              
                    $(var_div_observacion).css('display','block')
                    $(var_div_observacion).text("largo máximo superado")
                    $(this).css('background-color', '#F5A9A9') 
                                            
                }else{
                    $(var_div_observacion).css('display','none')                    
                    $(var_div_observacion).text("") 
                    $(this).css('background-color', '')
                }

                if(error_observacion ==1 ){
                    
                    $('#td_observacion').css('border', '2px solid #F5A9A9') 
                    $('#textarea_observacion').css('background-color', '#F5A9A9')  
                       
                }else{
                    $('#td_observacion').css('border-color', '') 
                    $('#td_observacion').css('border-width', '') 
                    $('#td_observacion').css('border-style', '') 
                    $('#textarea_observacion').css('background-color', '')    
                }

            })

            $("#observaciones").keyup(function() {
                var limit   = $(this).attr("maxlength"); // Límite del textarea
                var value   = $(this).val();             // Valor actual del textarea
                var current = value.length;              // Número de caracteres actual
                if (limit < current) {                   // Más del límite de caracteres?
                    // Establece el valor del textarea al límite
                    $(this).css('background-color', '#F5A9A9') 
                }else{
                    $(this).css('background-color', '') 
                }
            });

            $('#CB_COBRANZA').change(function(){

                if ($('#td_seleect').text()!=""){

                    if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                        var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                            $('#td_seleect_insert').text("")
                            $('#td_seleect').text("")
                            $('#textarea_id_usuario').removeAttr('disabled')
                            $('#textarea_hora_ingreso').removeAttr('disabled')
                            $('#textarea_observacion').removeAttr('disabled')
                            $('#textarea_fecha_ingreso').removeAttr('disabled')
                            $('#textarea_id_medio').removeAttr('disabled')
                            $('#textarea_email').removeAttr('disabled')


                            $('#procesar_ingresa').css('display','none')
							$('#procesar_elimina').css('display','none')
							$('#procesar').css('display','inline-block')
                            $('#limpia_campo').css('display','none') 

                            var criterios = "alea="+Math.random()+"&accion_ajax=filtra_usuario&CB_COBRANZA="+$('#CB_COBRANZA').val()
                            $('#mostrar_id_usuario').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){
                                
                                $('#id_usuario').change(function(){
                                    if ($('#td_seleect').text()!=""){
                                        if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                                            var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                                            $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                                                $('#td_seleect_insert').text("")
                                                $('#td_seleect').text("")
                                                $('#textarea_id_usuario').removeAttr('disabled')
                                                $('#textarea_hora_ingreso').removeAttr('disabled')
                                                $('#textarea_observacion').removeAttr('disabled')
                                                $('#textarea_fecha_ingreso').removeAttr('disabled')
                                                $('#textarea_id_medio').removeAttr('disabled')
                                                $('#textarea_email').removeAttr('disabled')


                                                $('#procesar_ingresa').css('display','none')
												$('#procesar_elimina').css('display','none')
												$('#procesar').css('display','inline-block')
                                                $('#limpia_campo').css('display','none') 


                                            })

                                        }

                                    }
                                })                                
                            })                                                                                       
                        })
                    }   

                }else{

                    var criterios = "alea="+Math.random()+"&accion_ajax=filtra_usuario&CB_COBRANZA="+$(this).val()
                    $('#mostrar_id_usuario').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){
                        $('#id_usuario').change(function(){
                            if ($('#td_seleect').text()!=""){
                                if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                                    var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                                    $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){
                                                $('#td_seleect_insert').text("")
                                                $('#td_seleect').text("")
                                                $('#textarea_id_usuario').removeAttr('disabled')
                                                $('#textarea_hora_ingreso').removeAttr('disabled')
                                                $('#textarea_observacion').removeAttr('disabled')
                                                $('#textarea_fecha_ingreso').removeAttr('disabled')
                                                $('#textarea_id_medio').removeAttr('disabled')
                                                $('#textarea_email').removeAttr('disabled')


                                                $('#procesar_ingresa').css('display','none')
												$('#procesar_elimina').css('display','none')
												$('#procesar').css('display','inline-block')
                                                $('#limpia_campo').css('display','none') 

                                    })

                                }
                                
                            }
                        })


                    })
                }

    
            })

            $('.filtro_fecha_ingreso').css('display','none')
            $('.filtro_hora_ingreso').css('display','none')
            $('.filtro_observacion').css('display','none')
            $('.filtro_id_usuario').css('display','none')
            $('.filtro_email').css('display','none')
            $('.filtro_id_medio').css('display','none')

                    
            $('input[id="ck_dv_rut"]').click(function(){
                if ($('#td_seleect').text()!=""){
                    if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                        var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                            $('#td_seleect_insert').text("")
                            $('#td_seleect').text("")
                            $('#textarea_id_usuario').removeAttr('disabled')
                            $('#textarea_hora_ingreso').removeAttr('disabled')
                            $('#textarea_observacion').removeAttr('disabled')
                            $('#textarea_fecha_ingreso').removeAttr('disabled')
                            $('#textarea_id_medio').removeAttr('disabled')
                            $('#textarea_email').removeAttr('disabled')


                            $('#procesar_ingresa').css('display','none')
							$('#procesar_elimina').css('display','none')
							$('#procesar').css('display','inline-block')
                            $('#limpia_campo').css('display','none') 


                            if($('input[id="ck_id_usuario"]').is(':checked')){
                                $('.filtro_id_usuario').css('display','inline')
                                $('#id_usuario').attr('disabled',1)
                                $('#id_usuario').css('background-color','#E6E6E6')

                            }else{
                                $('.filtro_id_usuario').css('display','none')
                                $('#id_usuario').removeAttr('disabled')
                                $('#id_usuario').css('background-color','')
                            }                                                                                     
                        })

                        
                    }   
                }

            })

            $('input[id="ck_id_usuario"]').click(function(){

                if ($('#td_seleect').text()!=""){

                    if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                        var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                            $('#td_seleect_insert').text("")
                            $('#td_seleect').text("")
                            $('#textarea_id_usuario').removeAttr('disabled')
                            $('#textarea_hora_ingreso').removeAttr('disabled')
                            $('#textarea_observacion').removeAttr('disabled')
                            $('#textarea_fecha_ingreso').removeAttr('disabled')
                            $('#textarea_id_medio').removeAttr('disabled')
                            $('#textarea_email').removeAttr('disabled')


                            $('#procesar_ingresa').css('display','none')
							$('#procesar_elimina').css('display','none')
							$('#procesar').css('display','inline-block')
                            $('#limpia_campo').css('display','none') 


                            if($('input[id="ck_id_usuario"]').is(':checked')){
                                $('.filtro_id_usuario').css('display','inline')
                                $('#id_usuario').attr('disabled',1)
                                $('#id_usuario').css('background-color','#E6E6E6')

                            }else{
                                $('.filtro_id_usuario').css('display','none')
                                $('#id_usuario').removeAttr('disabled')
                                $('#id_usuario').css('background-color','')
                            }                                                                                     
                        })

                    }else{
                        $('input[id="ck_id_usuario"]').removeAttr('checked')
                    }   

                }else{

                    if($('input[id="ck_id_usuario"]').is(':checked')){
                        $('.filtro_id_usuario').css('display','inline')
                        $('#id_usuario').attr('disabled',1)
                        $('#id_usuario').css('background-color','#E6E6E6')

                    }else{
                        $('.filtro_id_usuario').css('display','none')
                        $('#id_usuario').removeAttr('disabled')
                        $('#id_usuario').css('background-color','')
                    }
                }



            })
            $('input[id="ck_hora_ingreso"]').click(function(){


                if ($('#td_seleect').text()!=""){

                    if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                        var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                            $('#td_seleect_insert').text("")
                            $('#td_seleect').text("")
                            $('#textarea_id_usuario').removeAttr('disabled')
                            $('#textarea_hora_ingreso').removeAttr('disabled')
                            $('#textarea_observacion').removeAttr('disabled')
                            $('#textarea_fecha_ingreso').removeAttr('disabled')
                            $('#textarea_id_medio').removeAttr('disabled')
                            $('#textarea_email').removeAttr('disabled')


                            $('#procesar_ingresa').css('display','none')
							$('#procesar_elimina').css('display','none')
							$('#procesar').css('display','inline-block')
                            $('#limpia_campo').css('display','none') 

                            if($('input[id="ck_hora_ingreso"]').is(':checked')){
                                $('.filtro_hora_ingreso').css('display','inline')
                                $('#hora_ingreso').attr('disabled',1)
                                $('#hora_ingreso').css('background-color','#E6E6E6')
                            }else{
                                $('.filtro_hora_ingreso').css('display','none')
                                $('#hora_ingreso').removeAttr('disabled')
                                $('#hora_ingreso').css('background-color','')
                            }                                                                                     
                        })
                    }else{
                        $('input[id="ck_hora_ingreso"]').removeAttr('checked')
                    }    

                }else{

                    if($('input[id="ck_hora_ingreso"]').is(':checked'))
                    {
                        $('.filtro_hora_ingreso').css('display','inline')
                        $('#hora_ingreso').attr('disabled',1)
                        $('#hora_ingreso').css('background-color','#E6E6E6')
                    }else{
                        $('.filtro_hora_ingreso').css('display','none')
                        $('#hora_ingreso').removeAttr('disabled')
                        $('#hora_ingreso').css('background-color','')
                    }
                }



            })
            $('input[id="ck_observaciones"]').click(function(){

                if ($('#td_seleect').text()!=""){

                    if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                        var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                            $('#td_seleect_insert').text("")
                            $('#td_seleect').text("")
                            $('#textarea_id_usuario').removeAttr('disabled')
                            $('#textarea_hora_ingreso').removeAttr('disabled')
                            $('#textarea_observacion').removeAttr('disabled')
                            $('#textarea_fecha_ingreso').removeAttr('disabled')
                            $('#textarea_id_medio').removeAttr('disabled')
                            $('#textarea_email').removeAttr('disabled')


                            $('#procesar_ingresa').css('display','none')
							$('#procesar_elimina').css('display','none')
							$('#procesar').css('display','inline-block')
                            $('#limpia_campo').css('display','none') 

                            if($('input[id="ck_observaciones"]').is(':checked')){
                                $('.filtro_observacion').css('display','inline')
                                $('#observaciones').attr('disabled',1)
                                $('#observaciones').css('background-color','#E6E6E6')
                            }else{
                                $('.filtro_observacion').css('display','none')
                                $('#observaciones').removeAttr('disabled')
                                $('#observaciones').css('background-color','')
                            }                                                                                    
                        })

                    }else{
                        $('input[id="ck_observaciones"]').removeAttr('checked')
                    }    

                }else{

                    if($('input[id="ck_observaciones"]').is(':checked')){
                        $('.filtro_observacion').css('display','inline')
                        $('#observaciones').attr('disabled',1)
                        $('#observaciones').css('background-color','#E6E6E6')
                    }else{
                        $('.filtro_observacion').css('display','none')
                        $('#observaciones').removeAttr('disabled')
                        $('#observaciones').css('background-color','')
                    }
                }


            })



            $('input[id="ck_fecha_ingreso"]').click(function(){

                if ($('#td_seleect').text()!=""){

                    if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                        var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                            $('#td_seleect_insert').text("")
                            $('#td_seleect').text("")
                            $('#textarea_id_usuario').removeAttr('disabled')
                            $('#textarea_hora_ingreso').removeAttr('disabled')
                            $('#textarea_observacion').removeAttr('disabled')
                            $('#textarea_fecha_ingreso').removeAttr('disabled')
                            $('#textarea_id_medio').removeAttr('disabled')
                            $('#textarea_email').removeAttr('disabled')


                            $('#procesar_ingresa').css('display','none')
							$('#procesar_elimina').css('display','none')
							$('#procesar').css('display','inline-block')
                            $('#limpia_campo').css('display','none') 

                            if($('input[id="ck_fecha_ingreso"]').is(':checked')){
                                $('.filtro_fecha_ingreso').css('display','inline')
                                $('#fecha_ingreso').attr('disabled',1)
                                $('#fecha_ingreso').css('background-color','#E6E6E6')
                            }else{
                                $('.filtro_fecha_ingreso').css('display','none')
                                $('#fecha_ingreso').removeAttr('disabled')
                                $('#fecha_ingreso').css('background-color','')
                            }                                                                                   
                        })
                    }else{
                        $('input[id="ck_fecha_ingreso"]').removeAttr('checked')
                    }     

                }else{

                    if($('input[id="ck_fecha_ingreso"]').is(':checked')){
                        $('.filtro_fecha_ingreso').css('display','inline')
                        $('#fecha_ingreso').attr('disabled',1)
                        $('#fecha_ingreso').css('background-color','#E6E6E6')
                    }else{
                        $('.filtro_fecha_ingreso').css('display','none')
                        $('#fecha_ingreso').removeAttr('disabled')
                        $('#fecha_ingreso').css('background-color','')
                    }
                }

    
            })
            $('input[id="ck_activa_id_email"]').click(function(){


                if ($('#td_seleect').text()!=""){

                    if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){

                        var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                            $('#td_seleect_insert').text("")
                            $('#td_seleect').text("")
                            $('#textarea_id_usuario').removeAttr('disabled')
                            $('#textarea_hora_ingreso').removeAttr('disabled')
                            $('#textarea_observacion').removeAttr('disabled')
                            $('#textarea_fecha_ingreso').removeAttr('disabled')
                            $('#textarea_id_medio').removeAttr('disabled')
                            $('#textarea_email').removeAttr('disabled')


                            $('#procesar_ingresa').css('display','none')
							$('#procesar_elimina').css('display','none')
							$('#procesar').css('display','inline-block')
                            $('#limpia_campo').css('display','none') 
                            
                            if($('input[id="ck_activa_id_email"]').is(':checked')){
                                $('.filtro_id_medio').css('display','inline')
                                $('.filtro_email').css('display','none')
                                $('#id_medio').text("ID e-mail asociado")

                            }else{

                                $('.filtro_id_medio').css('display','none')
                                $('.filtro_email').css('display','inline')
                                
                            }                                                                                  
                        })

                    }else{
                        $('input[id="ck_activa_id_email"]').removeAttr('checked')
                    }     

                }else{

                    if($('input[id="ck_activa_id_email"]').is(':checked')){
                        $('.filtro_id_medio').css('display','inline')
                        $('.filtro_email').css('display','none')
                        $('#id_medio').text("ID e-mail asociado")

                    }else{

                        $('.filtro_id_medio').css('display','none')
                        $('.filtro_email').css('display','inline')

                        
                    }
                }



            })            


            $('#tipo_gestion').change(function(){
                var tipo_gestion        =$(this).val()                
                var var_tipo_gestion    =tipo_gestion.split("-")
                var cod_tipo_gestion    =var_tipo_gestion[0]
                var tipo                =var_tipo_gestion[1]
                //alert(tipo_gestion)
                if ($('#td_seleect').text()!=""){

                    if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?")){
                        $('#td_seleect_insert').text("")
                        $('#td_seleect').text("")
                        $('#textarea_id_usuario').removeAttr('disabled')
                        $('#textarea_hora_ingreso').removeAttr('disabled')
                        $('#textarea_observacion').removeAttr('disabled')
                        $('#textarea_fecha_ingreso').removeAttr('disabled')
                        $('#textarea_id_medio').removeAttr('disabled')
                        $('#textarea_email').removeAttr('disabled')


                        $('#procesar_ingresa').css('display','none')
						$('#procesar_elimina').css('display','none')
                        $('#procesar').css('display','inline-block')
                        $('#limpia_campo').css('display','none') 

                       if (tipo=="1"){ 

                            $('.filtro_email').css('display','none')                   
                            $('.filtro_id_medio').css('display','inline')
                            $('#id_medio').text("ID teléfono asociado") 
                            $('input[id="ck_activa_id_email"]').removeAttr("checked"); 
                            $('input[id="ck_activa_id_email"]').attr('disabled', true)  

                        }else if (tipo=="2"){

                            if($('input[id="ck_activa_id_email"]').is(':checked')){
                                $('.filtro_id_medio').css('display','inline')      
                            }else{
                                $('.filtro_id_medio').css('display','none') 
                            } 

                            $('#id_medio').text("ID e-mail asociado") 
                            $('.filtro_email').css('display','inline')
                            $('input[id="ck_activa_id_email"]').removeAttr('disabled')  

                        }else if (tipo=="3"){

                            $('.filtro_email').css('display','none')                   
                            $('.filtro_id_medio').css('display','inline')
                            $('#id_medio').text("ID direccion asociado") 
                            $('input[id="ck_activa_id_email"]').removeAttr("checked"); 
                            $('input[id="ck_activa_id_email"]').attr('disabled', true)   

                        }else{

                            $('.filtro_email').css('display','none')                   
                            $('.filtro_id_medio').css('display','none')
                            $('input[id="ck_activa_id_email"]').removeAttr("checked"); 
                            $('input[id="ck_activa_id_email"]').attr('disabled', true) 


                        }

                    }

                }else{

                    if (tipo=="1"){ 

                        $('.filtro_email').css('display','none')                   
                        $('.filtro_id_medio').css('display','inline')
                        $('#id_medio').text("ID teléfono asociado") 
                        $('input[id="ck_activa_id_email"]').removeAttr("checked"); 
                        $('input[id="ck_activa_id_email"]').attr('disabled', true)  

                    }else if (tipo=="2"){

                        if($('input[id="ck_activa_id_email"]').is(':checked')){
                            $('.filtro_id_medio').css('display','inline')      
                        }else{
                            $('.filtro_id_medio').css('display','none') 
                        } 

                        $('#id_medio').text("ID e-mail asociado") 
                        $('.filtro_email').css('display','inline')
                        $('input[id="ck_activa_id_email"]').removeAttr('disabled')  

                    }else if (tipo=="3"){

                        $('.filtro_email').css('display','none')                   
                        $('.filtro_id_medio').css('display','inline')
                        $('#id_medio').text("ID direccion asociado") 
                        $('input[id="ck_activa_id_email"]').removeAttr("checked"); 
                        $('input[id="ck_activa_id_email"]').attr('disabled', true)   

                    }else{

                        $('.filtro_email').css('display','none')                   
                        $('.filtro_id_medio').css('display','none')
                        $('input[id="ck_activa_id_email"]').removeAttr("checked"); 
                        $('input[id="ck_activa_id_email"]').attr('disabled', true) 


                    }
          
                }


            })

            $('.input2').TimepickerInputMask({
                seconds: false
            }); 

            $('#procesar_ingresa').click(function(){
                 
                var tipo_gestion                =$('#tipo_gestion').val() 
                var var_tipo_gestion            =tipo_gestion.split("-")
                var cod_tipo_gestion            =var_tipo_gestion[0]
                var tipo                        =var_tipo_gestion[1]                
                var ck_activa_id_email          =$('input[id="ck_activa_id_email"]:checked').val()
                var estado_ingreso_gestiones    =$('#estado_ingreso_gestiones').val()
                var estado_ingreso              =""
                var cont                        =0
				var CB_COBRANZA					=$('#CB_COBRANZA').val()


                $('input[name="estado_ingreso_gestiones"]').each(function()
				{
				
					//alert($(this).val())
                    cont  = cont + 1
                    if($(this).val()=="RUT SIN ASIGNACION VIGENTE")
                    {
					     estado_ingreso ="S"
						return
                    }
					if($(this).val()=="FECHA INVALIDA")
                    {
                        estado_ingreso ="S"
						return
                    }
					if($(this).val()=="HORA INVALIDA")
                    {
                        estado_ingreso ="S"
						return
                    }
					if($(this).val()=="USUARIO INVALIDO")
                    {
                        estado_ingreso ="S"
						return
                    }
					if($(this).val()=="ID MEDIO INVALID0")
                    {
						//alert("adfasd")
                        estado_ingreso ="S"
						return
                    }
					if($(this).val()=="LARGO OBSERVACION SUPERADO")
                    {
                        estado_ingreso ="S"
						return
                    }
					if($(this).val()=="EMAIL INVALIDO")
                    {
                        estado_ingreso ="S"
						return
                    }
					if($(this).val()=="OBSERVACION INVALIDA")
                    {
                        estado_ingreso ="S"
						return
                    }

                }
			)

                
                if(estado_ingreso=="S"){
                   alert("No es posible ingresar gestiones con errores.")
                    return

                }else{
                    if (ck_activa_id_email==null){
                        ck_activa_id_email=""
                    }
					
					if(confirm("¿Está seguro de ingresar las gestiones procesadas?"))
					{
							var criterios ="alea="+Math.random()+"&accion_ajax=procesa_gestiones&cod_tipo_gestion="+cod_tipo_gestion+"&tipo="+tipo+"&ck_activa_id_email="+ck_activa_id_email+"&CB_COBRANZA="+CB_COBRANZA

							$('#td_seleect_procesa').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){                        
							   
								$('#td_seleect_insert').text("")
								$('#td_seleect').text("")
								$('#textarea_id_usuario').removeAttr('disabled')
								$('#textarea_hora_ingreso').removeAttr('disabled')
								$('#textarea_observacion').removeAttr('disabled')
								$('#textarea_fecha_ingreso').removeAttr('disabled')
								$('#textarea_id_medio').removeAttr('disabled')
								$('#textarea_email').removeAttr('disabled')
																
								$('#textarea_rut_deudor').text("Copia texto")
								$('#textarea_fecha_ingreso').text("Copia texto")
								$('#textarea_hora_ingreso').text("Copia texto")
								$('#textarea_observacion').text("Copia texto")
								$('#textarea_id_usuario').text("Copia texto")
								$('#textarea_id_medio').text("Copia texto")
								$('#textarea_email').text("Copia texto") 
								$('#td_seleect_insert').text("")
								

								$('#procesar_ingresa').css('display','none')
								$('#procesar_elimina').css('display','none')
								$('#procesar').css('display','inline-block')
								$('#limpia_campo').css('display','none') 

								if($('input[id="ck_hora_ingreso"]').is(':checked')){
									$('.filtro_hora_ingreso').css('display','inline')
									$('#hora_ingreso').attr('disabled',1)
									$('#hora_ingreso').css('background-color','#E6E6E6')
								}else{
									$('.filtro_hora_ingreso').css('display','none')
									$('#hora_ingreso').removeAttr('disabled')
									$('#hora_ingreso').css('background-color','')
								} 
								alert("Gestiones procesadas con exito")
							  

							})
					}

                }


            })
/*botn eliminar*/


            $('#procesar').click(function(){

                var cont                =0
                var CB_COBRANZA         =$('#CB_COBRANZA').val()
                var ck_dv_rut           =$('input[id="ck_dv_rut"]:checked').val()
                var ck_id_usuario       =$('input[id="ck_id_usuario"]:checked').val()
                var ck_hora_ingreso     =$('input[id="ck_hora_ingreso"]:checked').val()
                var ck_observaciones    =$('input[id="ck_observaciones"]:checked').val()
                var ck_fecha_ingreso    =$('input[id="ck_fecha_ingreso"]:checked').val()
                var ck_activa_id_email  =$('input[id="ck_activa_id_email"]:checked').val()
                var tipo_gestion        =$('#tipo_gestion').val() 
   
                if(ck_dv_rut==null){
                    ck_dv_rut =""
                }

                if(CB_COBRANZA=="")
                {
                    alert("Debe ingresar cobranza")
                    return
                }

                if(tipo_gestion=="")
                {
                    alert("Debe ingresar tipo gestion")
                    return  
                }

                var var_tipo_gestion    =tipo_gestion.split("-")
                var cod_tipo_gestion    =var_tipo_gestion[0]
                var tipo                =var_tipo_gestion[1]

                /*SE PREGUNTA SI EL INGRESO ES MASIVO O NO, DEPENDIENDO DE LA SELECCCION RESCATA LA VARIABLE*/

                if(ck_id_usuario=="1"){
                    var var_id_usuario      =$('#textarea_id_usuario').val()
                }else{
                    var var_id_usuario      =$('#id_usuario').val()                        
                }
                    
                if(ck_hora_ingreso=="1"){
                    var var_hora_ingreso    =$("#textarea_hora_ingreso").val()
                }else{
                    var var_hora_ingreso    =$("#hora_ingreso").val()
                }  


                if(ck_observaciones=="1"){
                    var var_observacion     =$('#textarea_observacion').val()
                }else{
                    var var_observacion     =$("#observaciones").val()
                } 


                if(ck_fecha_ingreso=="1"){
                    var var_fecha_ingreso   =$('#textarea_fecha_ingreso').val()
                }else{
                    var var_fecha_ingreso   =$("#fecha_ingreso").val()
                }

                
                var var_id_medio            =$('#textarea_id_medio').val()
                var var_email               =$('#textarea_email').val()
                var var_rut_deudor          =$("#textarea_rut_deudor").val()


                /* VALIDACIONES DE INGRESO */
                if(var_rut_deudor=="" || var_rut_deudor=="Copia texto")
                {
                    alert("Debe ingresar rut deudor")
                    return  
                }

                if(var_id_usuario=="" || var_id_usuario=="Copia texto")
                {
                    alert("Debe ingresar usuario ingreso")
                    return  
                }
                if(var_hora_ingreso=="" || var_hora_ingreso=="Copia texto")
                {
                    alert("Debe ingresar hora ingreso")
                    return  
                }
                if(var_fecha_ingreso=="" || var_fecha_ingreso=="Copia texto")
                {
                    alert("Debe ingresar fecha ingreso")
                    return  
                }



                if(tipo=="2"){
                    //alert("aca")
                    if(ck_activa_id_email=="1")
                    {
                        if(var_id_medio=="" || var_id_medio=="Copia texto")
                        {
                            alert("Debe ingresar id medio")
                            return  
                        }                
                    }

                    if(ck_activa_id_email!=1){
                        if(var_email=="" || var_email=="Copia texto")
                        {
                            alert("Debe ingresar descripcion email")
                            return  
                        } 
                    }
                                        
                }else if (tipo!="0" && tipo!="2"){

                    if(var_id_medio=="" || var_id_medio=="Copia texto")
                    {
                        alert("Debe ingresar id medio")
                        return  
                    }                
                    
                 }


                var criterios ="alea="+Math.random()+"&accion_ajax=guardar_valores&var_rut_deudor="+encodeURIComponent(var_rut_deudor)+"&var_fecha_ingreso="+encodeURIComponent(var_fecha_ingreso)+"&var_hora_ingreso="+encodeURIComponent(var_hora_ingreso)+"&var_observacion="+encodeURIComponent(var_observacion)+"&var_id_usuario="+encodeURIComponent(var_id_usuario)+"&var_id_medio="+encodeURIComponent(var_id_medio)+"&ck_id_usuario="+ck_id_usuario+"&ck_hora_ingreso="+ck_hora_ingreso+"&ck_observaciones="+ck_observaciones+"&ck_fecha_ingreso="+ck_fecha_ingreso+"&cod_tipo_gestion="+cod_tipo_gestion+"&var_email="+encodeURIComponent(var_email)+"&ck_activa_id_email="+ck_activa_id_email+"&tipo="+tipo+"&ck_dv_rut="+ck_dv_rut

                $.post('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(data) {

                    $('#td_seleect_insert').html(data)
                    var error_largo_areglo =$('#error_largo_areglo').val()
                    if(error_largo_areglo=="")
                    {
                        var criterios ="alea="+Math.random()+"&accion_ajax=select_tabla_temporal&CB_COBRANZA="+CB_COBRANZA+"&ck_activa_id_email="+ck_activa_id_email+"&tipo="+tipo+"&ck_dv_rut="+ck_dv_rut                       

                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){
                                $('#limpia_campo').css('display','block')
                                $('#textarea_id_usuario').attr('disabled','true')
                                $('#textarea_hora_ingreso').attr('disabled','true')
                                $('#textarea_observacion').attr('disabled','true')
                                $('#textarea_fecha_ingreso').attr('disabled','true')
                                $('#textarea_id_medio').attr('disabled','true')
                                $('#textarea_email').attr('disabled','true')

                                $('#procesar_ingresa').css('display','inline-block')
						
						
								var estado_ingreso  =""
						        $('#procesar').css('display','none')
        
								$('input[name="estado_ingreso_gestiones"]').each(function()
										{
											cont  = cont + 1
											if($(this).val()=="RUT SIN ASIGNACION VIGENTE")
											{
												 estado_ingreso ="S"
											}
											if($(this).val()=="FECHA INVALIDA")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="HORA INVALIDA")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="USUARIO INVALIDO")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="ID MEDIO INVALID0")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="LARGO OBSERVACION SUPERADO")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="EMAIL INVALIDO")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="OBSERVACION INVALIDA")
											{
												estado_ingreso ="S"
											}
										}
									)
							
								if (estado_ingreso =="")
								{
									$('#procesar_elimina').css('display','none')
								}
								else{
									$('#procesar_elimina').css('display','inline-block')
								}
		


                        })                   
                    }
                       

                })                   
            
            })
			
			$('#procesar_elimina').click(function(){

                var cont                =0
                var CB_COBRANZA         =$('#CB_COBRANZA').val()
                var ck_dv_rut           =$('input[id="ck_dv_rut"]:checked').val()
                var ck_id_usuario       =$('input[id="ck_id_usuario"]:checked').val()
                var ck_hora_ingreso     =$('input[id="ck_hora_ingreso"]:checked').val()
                var ck_observaciones    =$('input[id="ck_observaciones"]:checked').val()
                var ck_fecha_ingreso    =$('input[id="ck_fecha_ingreso"]:checked').val()
                var ck_activa_id_email  =$('input[id="ck_activa_id_email"]:checked').val()
                var tipo_gestion        =$('#tipo_gestion').val() 
   
                if(ck_dv_rut==null){
                    ck_dv_rut =""
                }

                if(CB_COBRANZA=="")
                {
                    alert("Debe ingresar cobranza")
                    return
                }

                if(tipo_gestion=="")
                {
                    alert("Debe ingresar tipo gestion")
                    return  
                }

                var var_tipo_gestion    =tipo_gestion.split("-")
                var cod_tipo_gestion    =var_tipo_gestion[0]
                var tipo                =var_tipo_gestion[1]

                /*SE PREGUNTA SI EL INGRESO ES MASIVO O NO, DEPENDIENDO DE LA SELECCCION RESCATA LA VARIABLE*/

                if(ck_id_usuario=="1"){
                    var var_id_usuario      =$('#textarea_id_usuario').val()
                }else{
                    var var_id_usuario      =$('#id_usuario').val()                        
                }
                    
                if(ck_hora_ingreso=="1"){
                    var var_hora_ingreso    =$("#textarea_hora_ingreso").val()
                }else{
                    var var_hora_ingreso    =$("#hora_ingreso").val()
                }  


                if(ck_observaciones=="1"){
                    var var_observacion     =$('#textarea_observacion').val()
                }else{
                    var var_observacion     =$("#observaciones").val()
                } 


                if(ck_fecha_ingreso=="1"){
                    var var_fecha_ingreso   =$('#textarea_fecha_ingreso').val()
                }else{
                    var var_fecha_ingreso   =$("#fecha_ingreso").val()
                }

                
                var var_id_medio            =$('#textarea_id_medio').val()
                var var_email               =$('#textarea_email').val()
                var var_rut_deudor          =$("#textarea_rut_deudor").val()


                /* VALIDACIONES DE INGRESO */
                if(var_rut_deudor=="" || var_rut_deudor=="Copia texto")
                {
                    alert("Debe ingresar rut deudor")
                    return  
                }

                if(var_id_usuario=="" || var_id_usuario=="Copia texto")
                {
                    alert("Debe ingresar usuario ingreso")
                    return  
                }
                if(var_hora_ingreso=="" || var_hora_ingreso=="Copia texto")
                {
                    alert("Debe ingresar hora ingreso")
                    return  
                }
                if(var_fecha_ingreso=="" || var_fecha_ingreso=="Copia texto")
                {
                    alert("Debe ingresar fecha ingreso")
                    return  
                }



                if(tipo=="2"){
                    //alert("aca")
                    if(ck_activa_id_email=="1")
                    {
                        if(var_id_medio=="" || var_id_medio=="Copia texto")
                        {
                            alert("Debe ingresar id medio")
                            return  
                        }                
                    }

                    if(ck_activa_id_email!=1){
                        if(var_email=="" || var_email=="Copia texto")
                        {
                            alert("Debe ingresar descripcion email")
                            return  
                        } 
                    }
                                        
                }else if (tipo!="0" && tipo!="2"){

                    if(var_id_medio=="" || var_id_medio=="Copia texto")
                    {
                        alert("Debe ingresar id medio")
                        return  
                    }                
                    
                 }

				
				if(confirm("Se Eliminaran los Registros Con Errores, ¿DESEA CONTINUAR?"))
				{
					var criterios ="alea="+Math.random()+"&accion_ajax=elimina_Registros_Errores&CB_COBRANZA="+CB_COBRANZA+"&ck_activa_id_email="+ck_activa_id_email+"&tipo="+tipo+"&ck_dv_rut="+ck_dv_rut 
				
                $.post('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(data) 
				{
					
                    $('#td_seleect_insert').html(data)
                    var error_largo_areglo =$('#error_largo_areglo').val()
					
                    if(error_largo_areglo=="")
                    {
                        var criterios ="alea="+Math.random()+"&accion_ajax=select_tabla_temporal&CB_COBRANZA="+CB_COBRANZA+"&ck_activa_id_email="+ck_activa_id_email+"&tipo="+tipo+"&ck_dv_rut="+ck_dv_rut                       

                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function()
						{
                                $('#limpia_campo').css('display','block')
                                $('#textarea_id_usuario').attr('disabled','true')
                                $('#textarea_hora_ingreso').attr('disabled','true')
                                $('#textarea_observacion').attr('disabled','true')
                                $('#textarea_fecha_ingreso').attr('disabled','true')
                                $('#textarea_id_medio').attr('disabled','true')
                                $('#textarea_email').attr('disabled','true')

                                $('#procesar_ingresa').css('display','inline-block')
								
																var estado_ingreso 
						        $('#procesar').css('display','none')
        
								$('input[name="estado_ingreso_gestiones"]').each(function()
										{
											cont  = cont + 1
											if($(this).val()=="RUT SIN ASIGNACION VIGENTE")
											{
												 estado_ingreso ="S"
											}
											if($(this).val()=="FECHA INVALIDA")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="HORA INVALIDA")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="USUARIO INVALIDO")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="ID MEDIO INVALID0")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="LARGO OBSERVACION SUPERADO")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="EMAIL INVALIDO")
											{
												estado_ingreso ="S"
											}
											if($(this).val()=="OBSERVACION INVALIDA")
											{
												estado_ingreso ="S"
											}
										}
									)
									
							
								if (estado_ingreso =="")
								{
									$('#procesar_elimina').css('display','none')
								}
								else{
									$('#procesar_elimina').css('display','inline-block')
								}
								
								
                                $('#procesar').css('display','none')
        


                        })                   
                    }else 
					{
						//alert("se eliminaron todos los registros")
						var criterios ="alea="+Math.random()+"&accion_ajax=select_tabla_temporal&CB_COBRANZA="+CB_COBRANZA+"&ck_activa_id_email="+ck_activa_id_email+"&tipo="+tipo+"&ck_dv_rut="+ck_dv_rut                       

                        $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function()
						{
									$('#textarea_rut_deudor').text("Copia texto")
									$('#textarea_fecha_ingreso').text("Copia texto")
									$('#textarea_hora_ingreso').text("Copia texto")
									$('#textarea_observacion').text("Copia texto")
									$('#textarea_id_usuario').text("Copia texto")
									$('#textarea_id_medio').text("Copia texto")
									$('#textarea_email').text("Copia texto") 
									$('#td_seleect_insert').text("")

									$('#textarea_id_usuario').removeAttr('disabled')
									$('#textarea_hora_ingreso').removeAttr('disabled')
									$('#textarea_observacion').removeAttr('disabled')
									$('#textarea_fecha_ingreso').removeAttr('disabled')
									$('#textarea_id_medio').removeAttr('disabled')
									$('#textarea_email').removeAttr('disabled')

									$('#procesar_ingresa').css('display','none')
									$('#procesar_elimina').css('display','none')
									$('#procesar').css('display','inline-block')
									$('#limpia_campo').css('display','none') 
                        })                   
						
						  
					
					
					}
                       

                })  
				
            }
            }
			)

			
            $('#textarea_rut_deudor').focus(function(){
                if($(this).val()=="Copia texto")
                {
                     $(this).val("")

                } 
            })   

           

            $('.fecha_validada').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

            $('.input2').TimepickerInputMask({
                seconds: false
            }); 

            $('#textarea_rut_deudor').click(function(){

                if($(this).val()!="Copia texto" && $(this).val()!="")
                {
                   limpia_consulta()
                }   
                    
            })


            $('#textarea_fecha_ingreso').focus(function(){
                if($(this).val()=="Copia texto")
                {
                    $(this).val("")   
                } 
            }) 


            $('#textarea_hora_ingreso').focus(function(){
                if($(this).val()=="Copia texto")
                {
                    $(this).val("")   
                } 
            })            


            $('#textarea_observacion').focus(function(){
                if($(this).val()=="Copia texto")
                {
                    $(this).val("")   
                } 
            })


            $('#textarea_id_usuario').focus(function(){
                if($(this).val()=="Copia texto")
                {
                    $(this).val("")   
                } 
            })  


            $('#textarea_id_medio').focus(function(){
                if($(this).val()=="Copia texto")
                {
                    $(this).val("")   
                } 
            })            

            $('#textarea_email').focus(function(){
                if($(this).val()=="Copia texto")
                {
                    $(this).val("")   
                } 
            }) 
 

        })

        function limpia_consulta(){

            if ($('#td_seleect').text()!=""){   
                if(confirm("Se eliminaran los registros ingresados, ¿DESEA CONTINUAR?"))
				{

                    var criterios ="alea="+Math.random()+"&accion_ajax=elimina_tabla_temp"
                    $('#td_seleect').load('FuncionesAjax/asigna_masiva_ajax.asp', criterios, function(){

                        $('#textarea_rut_deudor').text("Copia texto")
                        $('#textarea_fecha_ingreso').text("Copia texto")
                        $('#textarea_hora_ingreso').text("Copia texto")
                        $('#textarea_observacion').text("Copia texto")
                        $('#textarea_id_usuario').text("Copia texto")
                        $('#textarea_id_medio').text("Copia texto")
                        $('#textarea_email').text("Copia texto") 
                        $('#td_seleect_insert').text("")

                        $('#textarea_id_usuario').removeAttr('disabled')
                        $('#textarea_hora_ingreso').removeAttr('disabled')
                        $('#textarea_observacion').removeAttr('disabled')
                        $('#textarea_fecha_ingreso').removeAttr('disabled')
                        $('#textarea_id_medio').removeAttr('disabled')
                        $('#textarea_email').removeAttr('disabled')

                        $('#procesar_ingresa').css('display','none')
						$('#procesar_elimina').css('display','none')
                        $('#procesar').css('display','inline-block')
                        $('#limpia_campo').css('display','none')                                                            
                    })
                }    

            }
         
        }
    </script>

</head>
<body>
 <div class="titulo_informe">CARGA MASIVA GESTIONES</div>   
<input type="hidden" id="ses_codcli" name="ses_codcli" value="<%=trim(session("ses_codcli"))%>">
<div class="body_maasiva">
    <div id="valida_rut"></div>
    <div id="valida_ingreso"></div> 


    <!-- SUBIDA ARCHIVOS --> 
    <div class="div_filtrado">
        Cobranza 
        <span id="TD_CB_COBRANZA">
            <select name="CB_COBRANZA" id="CB_COBRANZA">
                <option value="">SELECCIONAR</option>
                    <%If Trim(intUsaCobInterna) = "1" Then%>
                        <option value="INTERNA" >INTERNA</option>
                    <%End If%>

                    <%If Trim(intVerCobExt) = "1" Then%>
                        <option value="EXTERNA" >EXTERNA</option>
                    <%End If%>				
            </select>
        </span>

        Tipo gestión  
        <select name="tipo_gestion" id="tipo_gestion">
            <option value="">SELECCIONAR</option>
            <%
            abrirscg()
			
          	strSql = "EXEC PROC_GET_TIPO_GESTIONES_POR_CLIENTE " &  session("ses_codcli") 
				
                set rsGest = Conn.execute(strSql)                

                Do While not rsGest.eof
					strCodigo = rsGest("CODIGO_GESTION")
					strGestionTotal = rsGest("DESC_GESTION")

                    
                %>
                    <option value="<%=Trim(strCodigo)%>"><%=strGestionTotal%></option>

                <%
                    rsGest.movenext
                Loop

            cerrarscg()
            %>


        </select> 
		
        <input type="button" style="float:right;" class="fondo_boton_100" name="procesar_ingresa" id="procesar_ingresa" value="Cargar">
        <input type="button" style="float:right;" class="fondo_boton_100" name="procesar" id="procesar" value="Procesar">
         
    </div>
    <!-- FIN SUBIDA ARCHIVOS -->  
    <BR>

        <div class="div_opcion_masiva" >
            <table class="table_opcion_masiva">
                <tr>
                    <td class="class_campos estilo_columna_individual" aling="center">
                        <input type="radio" style="visibility:hidden;">Formato RUT
                    </td>
                </tr>
                <tr>
                    <td class="columna_tipo1" style="font-size:12px;font-family:'Verdana'; height:50px;">
                        <div style="margin:3px;"><input type="radio" name="ck_dv_rut" id="ck_dv_rut" value="3" checked>Con guión</div>                        
                        <div style="margin:3px;"><input type="radio" name="ck_dv_rut" id="ck_dv_rut" value="2">Sin guión</div>
                        <div style="margin:3px;"><input type="radio" name="ck_dv_rut" id="ck_dv_rut" value="1">Sin Digito Verificador</div>
                  
                    </td>
                </tr>           
                <tr>
                    <td class="class_campos estilo_columna_individual">
                        <input type="checkbox" name="ck_fecha_ingreso" id="ck_fecha_ingreso" value="1">
 
                        Fecha ingreso
                    </td>
                </tr>
                <tr>    
                    <td class="">
                        <input type="text" class="fecha_validada" style="width:97%;" name="fecha_ingreso" readonly="readonly" id="fecha_ingreso" value="<%=date%>">
                    </td>
                </tr>
                <tr>
                    <td class="class_campos estilo_columna_individual">        
                        <input type="checkbox" name="ck_hora_ingreso" id="ck_hora_ingreso" value="1">
      
                        Hora ingreso
                    </td>
                </tr>
                <tr>    
                    <td class="">
                        <input type="text" class="input2" style="width:97%;" name="hora_ingreso" id="hora_ingreso" value="">
                    </td>
                </tr>
                <tr>
                    <td class="class_campos estilo_columna_individual" >        
                        <input type="checkbox" name="ck_id_usuario" id="ck_id_usuario" value="1">
                     
                        Usuario ingreso
                    </td>
                </tr>    
                <tr>
                    <td class="" id="mostrar_id_usuario">
                        <select name="id_usuario" id="id_usuario" style="width:100%;">
                            <option value="">SELECCIONAR</option>                            
                        </select> 
                    </td>
                </tr>
                <tr>
                    <td class="class_campos estilo_columna_individual">
                        <input type="checkbox" name="ck_observaciones" id="ck_observaciones" value="1">  
                        
                        Observaciones
                    </td>
                </tr>
                <tr>    
                    <td class="">
                        <textarea name="observaciones" style="width:97%;" maxlength="599" ROWS="5" id="observaciones"></textarea>
                    </td>
                </tr>
                <tr>
                    <td class="class_campos estilo_columna_individual" >        
                        <input type="checkbox" name="ck_activa_id_email" id="ck_activa_id_email" value="1">
                        
                        id e-mail asociado
                    </td>
                </tr>    
            </table>
        </div>

        <div class="div_masiva">
        	<table class="table_masiva estilo_columnas" >
               <thead>
                    <tr>
                        <th class="class_contador"></th>
                        <th class="class_campos">Rut deudor </th>
                        <th class="class_campos filtro_id_medio" id="id_medio">ID medio</th>
                        <th class="class_campos filtro_email">E-mail asociado</th>
                        <th class="class_campos filtro_fecha_ingreso">Fecha ingreso</th>
                        <th class="class_campos filtro_hora_ingreso">Hora ingreso</th>
                        <th class="class_campos filtro_observacion">Observación</th>
                        <th class="class_campos filtro_id_usuario">ID usuario</th>
                    </tr>
                
                </thead>
                <tbody>
                    <tr>
                        <td class="td_contador"></td>
                        <td class="class_campos"  ><textarea id="textarea_rut_deudor">Copia texto</textarea></td>
                        <td class="class_campos filtro_id_medio"><textarea id="textarea_id_medio">Copia texto</textarea></td>
                        <td class="class_campos filtro_email"><textarea id="textarea_email">Copia texto</textarea></td>
                        <td class="class_campos filtro_fecha_ingreso"><textarea id="textarea_fecha_ingreso">Copia texto</textarea></td>
                        <td class="class_campos filtro_hora_ingreso"><textarea id="textarea_hora_ingreso">Copia texto</textarea></td>
                        <td class="class_campos filtro_observacion"><textarea id="textarea_observacion">Copia texto</textarea></td>
                        <td class="class_campos filtro_id_usuario"><textarea id="textarea_id_usuario">Copia texto</textarea></td>
                      
                    </tr>            
                </tbody>
            </table>
            

            <div>
                <input type="button" class="fondo_boton_100" id="limpia_campo" name="limpia_campo" value="Limpia informe">
				
            </div>
			<div>
				<input type="button" class="fondo_boton_100" name="procesar_elimina" id="procesar_elimina" value="Quitar Errores" title="Eliminar Todos Los Registros Con Errores">
			</div>
      
            <div id="td_seleect_procesa"></div> 
            <div id="td_seleect_insert"></div> 
            <div id="td_seleect" class="td_seleect"></div>        
        </div>
    </div>
    <div  id="registro_hora_ingreso"></div>
    <div  id="registro_fecha_ingreso"></div>
    <div  id="registro_id_usuario"></div> 
    <div  id="registro_observaciones"></div> 
    <div  id="elimina_registro_gestion"></div> 

</body>
</html>
<!--#include file="../lib/comunes/rutinas/validarRut_masivo.inc"-->