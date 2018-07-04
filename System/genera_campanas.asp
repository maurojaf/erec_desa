<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/sbaeza.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->

	<link href="../css/jquery.alerts.css" rel="stylesheet"> 
	<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">	
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	
	
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 	
	<script src="../Componentes/jquery.alerts.mod.js"></script>
	<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>
	
	
	<%
		Cod_Cliente = session("ses_codcli")
		Id_Usuario = session("session_idusuario")
	%>
	
	
<script>

	$(document).ready(function(){
	$.prettyLoader();
	
	$('#Txt_Fecha_Inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#Txt_Fecha_Fin').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#Txt_Fecha_Inicio_Mod').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#Txt_Fecha_Fin_Mod').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	
	/*$("#table_tablesorter").tablesorter()*/
	
	$('#nueva_campana').toggle(
        function () { //Ocultar
               document.getElementById('abrir_cerrar').src = "../Imagenes/icono_menos.png";
			   $( "#info1" ).slideToggle( "clip", "", "" );
            },
        function () { //Mostrar
		        document.getElementById('abrir_cerrar').src = "../Imagenes/icono_mas.png";
				 $( "#info1" ).slideToggle( "clip", "", "" );	
            }
        );

		
		$('#Procesar').hover(function(){
				$(this).css('background-color','#CEE3F6');
			}, function(){
				$(this).css('background-color','');
		});
		
		/*cargo los datos de la campaña ya creada*/
		$('#CB_CAMPANA').change(function () {
        var parametros
        var CB_CAMPANA = $('#CB_CAMPANA').val();
        if (CB_CAMPANA != 0) {
			/*alert("CargoInfo")*/
			CargaCampanas(CB_CAMPANA);
			ResumenGeneral();
			Resumenejecutivo();
		
        } else {
           
					$("#Txt_Nom_Campana_Mod").val('');
					$('#Txt_Des_Campana_Mod').val('');
					$('#Txt_Fecha_Inicio_Mod').val('');
					$('#Txt_Fecha_Fin_Mod').val('');
					$('#Txt_Observacion_Mod').val('');
					$('#Txt_Rut_Mod').val('');
					
					
					document.getElementById("usuario").innerText = "";
					document.getElementById("FECHA_CREACION").innerText = "";
					document.getElementById("usuario_modifica").innerText = "";
					document.getElementById("fecha_modificacion").innerText = "";
					
					ResumenGeneral();
					Resumenejecutivo();
		   
        }
		});
		
		CargaCampanas(0);
		
		    $("#Btn_Crea_Campana").button().click(function () {
					IngresaCampana();
						
			});
			
			  $("#Btn_Elimina").button().click(function () {
				   EliminarCampana();
			});
		
			$("#Btn_Mod_Campana").button().click(function () {
				
				var ID_CAMPANA = $('#CB_CAMPANA').val();
				if (ID_CAMPANA==0)
				{
				jAlert("Seleccione Campaña","Advertencia!")		
				return;
				}
				ModificaCampana();
				
			});
		
    });
  

	function IngresaCampana()
	{
	
	var Cod_Cliente = $('#Cod_Cliente').val();
	var Id_Usuario =  $('#Id_Usuario').val();
	var Nom_Campana = $('#Txt_Nom_Campana').val();
	var Des_Campana = $('#Txt_Des_Campana').val();
	var Fecha_Inicio = $('#Txt_Fecha_Inicio').val();
	var Fecha_Fin = $('#Txt_Fecha_Fin').val();
	var Ruts = $("textarea#Txt_Rut").val() //$('#Txt_Rut').val();
	
	var Observacaion =  "" ;/*$('#Txt_Observacion').val();*/

	
	// 3000 + 1 enter registros como maximo
	if (Ruts.length >= 33020)
	{
	    $('#span_Rut_Campana_mod').css('border-color','#FE2E2E')	
		$('#span_Rut_Campana_mod').text("*")
		jAlert("Cantidad de Rut Ingresados Supera el Maximo de Datos 3000 Registros","Advertencia!")
		return;
	
	}
	
	
	if (Nom_Campana.length == 0)
	{
		$('#span_Nom_Campana').css('border-color','#FE2E2E')	
		$('#span_Nom_Campana').text("*")
		jAlert("Ingrese Nombre de la Campaña","Advertencia!")
		return;
	}else 
	{
	    $('#span_Nom_Campana').text("")
		$(this).css('border-color','')	
	}
	
	
	if (Fecha_Inicio.length == 0)
	{
		$('#span_Fecha_Inicio').css('border-color','#FE2E2E')	
		$('#span_Fecha_Inicio').text("*")
		jAlert("Ingrese Fecha Inicio de la Campaña","Advertencia!")
		return;
	}else
	{
		$('#span_Fecha_Inicio').text("")
		$(this).css('border-color','')	
	}
	
	var myDate=new Date();
	var dia = myDate.getDate()
	var mes = (myDate.getMonth() + 1)
	var agno =  myDate.getFullYear()
	
	if (mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5  || mes == 6 ||  mes == 7 || mes == 8  || mes == 9 )
	{ mes = "0" + mes; }

	if (dia == 1 || dia == 2 || dia == 3 || dia == 4 || dia == 5 || dia == 6 || dia == 7 || dia == 8 || dia == 9)
	{ dia = "0" + dia; }
	
	var x= (  dia + "/" + mes   + "/" + agno)

	
	
	if (x > Fecha_Inicio)
	{
	  jAlert("Fecha Inicio de la Campaña No Debe ser Menor a la Actual","Advertencia!")
	  $('#span_Fecha_Inicio').css('border-color','#FE2E2E')	
	  $('#span_Fecha_Inicio').text("*")
	  return false;
    }else
	{
		$('#span_Fecha_Inicio').text("")
		$(this).css('border-color','')	
	}

	if (Fecha_Fin.length == 0)
	{
		$('#span_Fecha_Fin').css('border-color','#FE2E2E')	
		$('#span_Fecha_Fin').text("*")
		jAlert("Ingrese Fecha Fin de la Campaña","Advertencia!")
		return;
	}else
	{
		$('#span_Fecha_Fin').text("")
		$(this).css('border-color','')	
	}
	
	
	if (Fecha_Fin < Fecha_Inicio)
	{
	$('#span_Fecha_Fin').css('border-color','#FE2E2E')	
	$('#span_Fecha_Fin').text("*")
	jAlert("Fecha Término debe ser igual o posterior a Fecha Inicio","Advertencia!")
	return ;
	}else
	{
		$('#span_Fecha_Fin').text("")
		$(this).css('border-color','')	
	}
	
	
			jConfirm("¿Esta seguro De Crear La Campaña?", "Advertencia!", function (r) 
			{
			if (r) {
					
					   $.prettyLoader.show();
					   $.ajax({url: "FuncionesAjax/Genera_Campana_ajax.asp?accion_ajax=Ingresa_Campana", 
						type: "POST", 
						data: {ID_CAMPANA: 0,COD_CLIENTE:Cod_Cliente,Nom_Campana:Nom_Campana,DESCRIPCION:Des_Campana,fecha_inicio:Fecha_Inicio,fecha_termino:Fecha_Fin,id_usuario:Id_Usuario,Observacaion:Observacaion,Ruts:Ruts},
						success: function(msg) {
							/*$('#refresca_resumen').append(msg)*/
						var Mensaje 	=  msg.split(','); 
						var estado  	= Mensaje[1];
						var Msj  	= Mensaje[2];
							
							if (estado.toUpperCase() != "OK")
							{
									jAlert(Msj,"Ingreso  Nueva Campaña");
									return ;
							}else
							{	
						
								/*Datos cargados ok*/
								jAlert(Msj,"Ingreso  Nueva Campañass");
								
								/* Limpiando */
								$("#Txt_Nom_Campana").val('');
								$('#span_Nom_Campana').css('border-color','#FE2E2E')	
								$('#span_Nom_Campana').text("*")
								$('#Txt_Des_Campana').val('');
								$('#Txt_Fecha_Inicio').val('');
								$('#span_Fecha_Inicio').css('border-color','#FE2E2E')	
								$('#span_Fecha_Inicio').text("*")
								$('#Txt_Fecha_Fin').val('');
								$('#span_Fecha_Fin').css('border-color','#FE2E2E')	
								$('#span_Fecha_Fin').text("*")
								$('#Txt_Rut').val('');
								
								
								/*Por si tienen alguna campaña seleccionada*/
								$("#Txt_Nom_Campana_Mod").val('');
								$('#Txt_Des_Campana_Mod').val('');
								$('#Txt_Fecha_Inicio_Mod').val('');
								$('#Txt_Fecha_Fin_Mod').val('');
								$('#Txt_Observacion_Mod').val('');
								$('#Txt_Rut_Mod').val('');
								
									
								document.getElementById("usuario").innerText = "";
								document.getElementById("FECHA_CREACION").innerText = "";
								document.getElementById("usuario_modifica").innerText = "";
								document.getElementById("fecha_modificacion").innerText = "";
								
								CargaCampanas(0)
								ResumenGeneral();
								Resumenejecutivo();
															 
								
							}
					
						
						},
						error: function(request, status, error){
						/*jAlert("Error al procesar los datos ('refresca_campana')");*/
						alert(error);
						/*jAlert("Error al procesar los datos ('"+errorThrown+"')")*/
						}
					   });
		   }else {
				 jAlert("Proceso Cancelado", "Advertencia!");
			} 
			}); 
	} 
  
  
  function ModificaCampana()
	{
	
	var Cod_Cliente = $('#Cod_Cliente').val();
	var Id_Usuario =  $('#Id_Usuario').val();
	var Nom_Campana = $('#Txt_Nom_Campana_Mod').val();
	var Des_Campana = $('#Txt_Des_Campana_Mod').val();
	var Fecha_Inicio = $('#Txt_Fecha_Inicio_Mod').val();
	var Fecha_Fin = $('#Txt_Fecha_Fin_Mod').val();
	var Ruts = $('#Txt_Rut_Mod').val();
	var Observacaion =  $('#Txt_Observacion_Mod').val();
	var ID_CAMPANA = $('#CB_CAMPANA').val();
	/* Validaciones */
	

	if (Nom_Campana.length == 0)
	{
		$('#span_Nom_Campana_mod').css('border-color','#FE2E2E')	
		$('#span_Nom_Campana_mod').text("*")
		jAlert("Ingrese Nombre de la Campaña","Advertencia!")
		return;
	}else 
	{
	    $('#span_Nom_Campana_mod').text("")
		$(this).css('border-color','')	
	}
	
	
	if (Ruts.length == 0)
	{
		$('#span_Rut_Campana_mod').css('border-color','#FE2E2E')	
		$('#span_Rut_Campana_mod').text("*")
		jAlert("Ingrese Rut a la  Campaña","Advertencia!")
		return;
	}else 
	{
	    $('#span_Rut_Campana_mod').text("")
		$(this).css('border-color','')	
	}
	
	// 3000 + 1 enter registros como maximo
	if (Ruts.length >= 33020)
	{
	
		$('#span_Rut_Campana_mod').css('border-color','#FE2E2E')	
		$('#span_Rut_Campana_mod').text("*")
		jAlert("Cantidad de Rut Ingresados Supera el Maximo de Datos 3000 Registros","Advertencia!")
		return;
	
	}
	

	
	if (Fecha_Inicio.length == 0)
	{
		$('#span_Fecha_Inicio_Mod').css('border-color','#FE2E2E')	
		$('#span_Fecha_Inicio_Mod').text("*")
		jAlert("Ingrese Fecha Inicio de la Campaña","Advertencia!")
		return;
	}else
	{
		$('#span_Fecha_Inicio_Mod').text("")
		$(this).css('border-color','')	
	}
	

		if (Fecha_Fin.length == 0)
	{
		$('#span_Fecha_Fin_mod').css('border-color','#FE2E2E')	
		$('#span_Fecha_Fin_mod').text("*")
		jAlert("Ingrese Fecha Fin de la Campaña","Advertencia!")
		return;
	}else
	{
		$('#span_Fecha_Fin_mod').text("")
		$(this).css('border-color','')	
	}
	
	
	if (Fecha_Fin < Fecha_Inicio)
	{
	$('#span_Fecha_Fin_mod').css('border-color','#FE2E2E')	
	$('#span_Fecha_Fin_mod').text("*")
	jAlert("Fecha Término debe ser igual o posterior a Fecha Inicio","Advertencia!")
	return ;
	}else
	{
		$('#span_Fecha_Fin_mod').text("")
		$(this).css('border-color','')	
	}
	
	
	
		if (Observacaion.length == 0)
	{
		$('#span_Observacion_Mod').css('border-color','#FE2E2E')	
		$('#span_Observacion_Mod').text("*")
		jAlert("Ingrese Observacaion de Modificación la Campaña","Advertencia!")
		return;
	}else
	{
	$('#span_Observacion_Mod').text("")
		$(this).css('border-color','')	
	
	}
	
	
	if (Fecha_Fin < Fecha_Inicio)
	{
	jAlert("Fecha Término debe ser igual o posterior a Fecha Inicio","Advertencia!")
	return ;
	}
	
	
	
	jConfirm("¿Esta seguro De Modificar La Campaña?", "Advertencia!", function (r) 
			{
				if (r) {
	
				$.prettyLoader.show();
				$.ajax({url: "FuncionesAjax/Genera_Campana_ajax.asp?accion_ajax=refresca_Campana_deudores", 
						method: "POST", 
						data: {ID_CAMPANA: ID_CAMPANA,COD_CLIENTE:Cod_Cliente},
						success: function(msg) {

						var Mensaje 	=  msg.split(','); 
						var estado  	= Mensaje[1];
						var Msj  	= Mensaje[2];
							
							if (estado.toUpperCase() != "OK")
							{
									jAlert(Msj,"Advertencia!")
									return ;
							}else
							{	
									
										   $.ajax({url: "FuncionesAjax/Genera_Campana_ajax.asp", 
											type: "POST", 
											data: {ID_CAMPANA: ID_CAMPANA,COD_CLIENTE:Cod_Cliente,Nom_Campana:Nom_Campana,DESCRIPCION:Des_Campana,fecha_inicio:Fecha_Inicio,fecha_termino:Fecha_Fin,id_usuario:Id_Usuario,Ruts:Ruts,Observacaion:Observacaion,accion_ajax:"Ingresa_Campana"},
											success: function(msg) {
											/*$('#refresca_resumen').append(msg)*/
											var Mensaje 	=  msg.split(','); 
											var estado  	= Mensaje[1];
											var Msj  	= Mensaje[2];
												
												if (estado.toUpperCase() != "OK")
												{
														jAlert(Msj,"Advertencia!")
														return ;
												}else
												{	
													jAlert(Msj,"Advertencia!")
													CargaCampanas(ID_CAMPANA)
													ResumenGeneral();
													Resumenejecutivo();													
												}
											},
											error: function(XMLHttpRequest, textStatus, errorThrown){
														jAlert("Error al procesar los datos refresca_resumen ('"+errorThrown+"')")
													}
										   });
										   
										   }
						},error: function(XMLHttpRequest, textStatus, errorThrown){
								/*jAlert("Error al procesar los datos ('refresca_campana')");*/
								jAlert("Error al procesar los datos refresca_Campana_deudores ('"+errorThrown+"')")
							}
					   });} 
			else {
				 jAlert("Proceso Cancelado", "Advertencia!");
			} 
			}); 
	} 
  
  
  function EliminarCampana()
  {
  
			var ID_CAMPANA 		 = $('#CB_CAMPANA').val();	
			var Cod_Cliente = $('#Cod_Cliente').val();
	
			if (ID_CAMPANA == 0){return;}
			jConfirm("¿Esta seguro De Eliminar La Campaña?", "Eliminar Campaña", function (r) 
			{
				if (r) {
						$.prettyLoader.show();
						$.ajax({url: "FuncionesAjax/Genera_Campana_ajax.asp?accion_ajax=Elimina_Campana", 
						method: "POST", 
						data: {ID_CAMPANA: ID_CAMPANA,COD_CLIENTE:Cod_Cliente,Mostrar:1},
						success: function(data) {
						/*	var campana 	=  data.split(','); 			*/
						
								$("#Txt_Nom_Campana_Mod").val('');
								$('#Txt_Des_Campana_Mod').val('');
								$('#Txt_Fecha_Inicio_Mod').val('');
								$('#Txt_Fecha_Fin_Mod').val('');
								$('#Txt_Observacion_Mod').val('');
								$('#Txt_Rut_Mod').val('');
								
								document.getElementById("usuario").innerText = "";
								document.getElementById("FECHA_CREACION").innerText = "";
								document.getElementById("usuario_modifica").innerText = "";
								document.getElementById("fecha_modificacion").innerText = "";
					
								
								
								CargaCampanas(0)
								ResumenGeneral();
								Resumenejecutivo();
								
						/*$('#refresca_resumen').append(data)*/
						},
						error: function(XMLHttpRequest, textStatus, errorThrown){
						/*jAlert("Error al procesar los datos ('refresca_campana')");*/
						jAlert("Error al procesar los datos ('"+errorThrown+"')")
					}
					   });
				   } 
				 else {
					 jAlert("Proceso Cancelado", "Eliminar Campaña");
				} 
			}); 
  }
  
  
function CargaCampanas(valor) 
{
	
	
if (valor==0)
{

	var cmb =   $('#CB_CAMPANA');
	$(cmb).html('');
    $(cmb).append("<option value='" +
                        0 + "'>" + "..Cargando"+ "</option>");

	var ID_CAMPANA 		 = 0;//$('#CB_CAMPANA').val();			
	var Cod_Cliente = $('#Cod_Cliente').val();
	
	 $.ajax({url: "FuncionesAjax/Genera_Campana_ajax.asp?accion_ajax=carga_Campana", 
			method: "POST", 
			data: {ID_CAMPANA: ID_CAMPANA,COD_CLIENTE:Cod_Cliente,Mostrar:valor},
			success: function(data) {
			
			//$('#refresca_resumen').append(data)
			//return;
			$(cmb).html('');
			$(cmb).append("<option value='" +
                        0 + "'>" + "Seleccione"+ "</option>");
						
			$(cmb).append(data)
			/*$('#refresca_resumen').append(data)*/
			},
			error: function(XMLHttpRequest, textStatus, errorThrown){
				/*jAlert("Error al procesar los datos ('carga_Campana')");*/
				 jAlert("Error al procesar los datos ('"+errorThrown+"')")
			}
           });
}else
{

			var ID_CAMPANA 		 = valor ;//$('#CB_CAMPANA').val();			
			var Cod_Cliente = $('#Cod_Cliente').val();
	
			/*alert(ID_CAMPANA)*/
	
	        $.ajax({url: "FuncionesAjax/Genera_Campana_ajax.asp?accion_ajax=carga_Campana", 
			method: "POST", 
			data: {ID_CAMPANA: ID_CAMPANA,COD_CLIENTE:Cod_Cliente,Mostrar:1},
			success: function(data) {
		
			/*$('#refresca_resumen').append(data)*/
			/*alert(data);*/
			
			var campana 	=  data.split(','); 
			var ID_CAMPANA  	= campana[1];
			var NOMBRE  	= campana[2];
			var DESCRIPCION  	= campana[3];
			var FECHA_CREACION  	= campana[4];
			var fecha_inicio  	= campana[5];
			var fecha_termino  	= campana[6];
			var fecha_modificacion  	= campana[7];
			var observacion  	= campana[8];
			var usuario  	= campana[9];
			var usuario_modifica  	= campana[10];
			var rut  	= campana[11];
			
			
				    $("#Txt_Nom_Campana_Mod").val(NOMBRE);
					$('#Txt_Des_Campana_Mod').val(DESCRIPCION);
					$('#Txt_Fecha_Inicio_Mod').val(fecha_inicio);
					$('#Txt_Fecha_Fin_Mod').val(fecha_termino);
					$('#Txt_Observacion_Mod').val(observacion);
					
					$('#Txt_Rut_Mod').val(rut);
					
					document.getElementById("usuario").innerText = "Usuario creador: "+ usuario;
					document.getElementById("FECHA_CREACION").innerText = "Creado: " + FECHA_CREACION;
					
					if (usuario_modifica.length == 0)
					{
					document.getElementById("usuario_modifica").innerText ="";
					document.getElementById("fecha_modificacion").innerText ="";
					}else 
					{
					document.getElementById("usuario_modifica").innerText 	="| Usuario Última Modificación: " + usuario_modifica;
					document.getElementById("fecha_modificacion").innerText ="| Última Actualización: " + fecha_modificacion;
					
					}
			},
			error: function(XMLHttpRequest, textStatus, errorThrown){
				 jAlert("Error al procesar los datos ('"+errorThrown+"')")
			}
           });

}

	
							
						
						
}
 
function ResumenGeneral()
{
			
			$('#Tbl_Resumen').html('');
			var ID_CAMPANA 		 = $('#CB_CAMPANA').val();	
			var Cod_Cliente = $('#Cod_Cliente').val();
	
	        $.ajax({url: "FuncionesAjax/Genera_Campana_ajax.asp?accion_ajax=refresca_Reporte_Asignacion", 
			method: "POST", 
			data: {ID_CAMPANA: ID_CAMPANA,COD_CLIENTE:Cod_Cliente},
			success: function(data) {
			
			$('#Tbl_Resumen').append(data)
			},
			error: function(XMLHttpRequest, textStatus, errorThrown){
				 jAlert("Error al procesar los datos ('"+errorThrown+"')")
			}
           });
  
}

function Resumenejecutivo()
{
			
			$('#Tbl_Ejecutivos').html('')
			
			var ID_CAMPANA 		 = $('#CB_CAMPANA').val();	
			var Cod_Cliente = $('#Cod_Cliente').val();

	        $.ajax({url: "FuncionesAjax/Genera_Campana_ajax.asp?accion_ajax=refresca_Reporte_Ejecutivos", 
			method: "POST", 
			data: {ID_CAMPANA: ID_CAMPANA,COD_CLIENTE:Cod_Cliente},
			success: function(data) {
			/*	var campana 	=  data.split(','); 			*/
				$('#Tbl_Ejecutivos').append(data)
			
			
			},
			error: function(XMLHttpRequest, textStatus, errorThrown){
				 jAlert("Error al procesar los datos ('"+errorThrown+"')")
			}
           });
  
}
	
</script>

<style>
.span_aviso_rojo{
		color:#FE2E2E;
		font-size:12px;
	}
</style>

</head>
<form name="datos" method="post">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div id="refresca_resumen"></div>
<div class="titulo1" id="administrados_campanas">
<h1>Administrador de Campañas</h1>
</div>

<div class="titulo2" id="nueva_campana">
<h2>> Nueva Campaña</h2> 
  <h3>Defina un nombre para la nueva campaña. Opcionalmente agregue una descripción, fechas de inicio y término e ingrese Rut.</h3>
<img id= "abrir_cerrar" src="../Imagenes/icono_mas.png">	
</div>
<!--Variables de session ocupados para enviarlos al ajax-->
<input type="hidden" id="Id_Usuario"  value ="<%=Id_Usuario%>"></input>
<input type="hidden" id="Cod_Cliente"  value ="<%=Cod_Cliente%>"></input>

<div class="info" id="info1">
  <table width="580" border="0">
    <tr>
      <td width="81" height="20" valign="middle"><h4>Nombre:</h4>
      </td>
      <td height="20" colspan="4" align="left" valign="middle" nowrap="nowrap">
		<input Id="Txt_Nom_Campana" type="text" value="" size="45" maxlength="20" >
		<span id="span_Nom_Campana" class="span_aviso_rojo">*</span>
      </td>
      <td width="6"></td>
      <td width="86"><h4>Ingreso Rut: </h4></td>
      <td width="149" rowspan="5" valign="top" class="tabla_rut">
		<textarea  cols="20" rows="10" id="Txt_Rut" maxlength="60000" ></textarea>   
		
	</td>
    </tr>
    <tr>
      <td width="81" height="25" valign="middle"><h4>Descripción:</h4></td>
      <td height="25" colspan="4" align="left" valign="middle" nowrap="nowrap">
		<input id="Txt_Des_Campana" type="text" value="" size="45" maxlength="120" onChange=""></td>
      <td width="6"></td>
      <td width="86"></td>
    </tr>
    <tr>
    	<td width="81">
    	<h4>Fecha Inicio:</h4>
    	</td>
    	<td width="60" align="left" valign="middle" nowrap="nowrap">
				<input type="text" id="Txt_Fecha_Inicio" readonly value="" size="10" maxlength="10">  
				<span id="span_Fecha_Inicio" class="span_aviso_rojo">*</span>				
        <td width="4" align="left" valign="middle" nowrap="nowrap">        
        <td width="100" align="left" valign="middle" nowrap="nowrap"><h4>Fecha Término:</h4>
        <td width="56" height="25" align="left" valign="middle" nowrap="nowrap">
				<input name="Txt_Fecha_Fin" type="text" id="Txt_Fecha_Fin" readonly value="" size="10" maxlength="10">
				<span id="span_Fecha_Fin" class="span_aviso_rojo">*</span>
        <td width="6">
        <td width="86">
        </td>
    </tr>
    <tr>
    	<td width="81">
    	<h4>&nbsp;</h4>
    	</td>
    	<td width="60" align="left" valign="middle" nowrap="nowrap">
		<input id="Btn_Crea_Campana" class="fondo_boton_100" type="button" value="Crear" >
  	    <td width="4" align="left" valign="middle" nowrap="nowrap">        
        <td width="100" align="left" valign="middle" nowrap="nowrap">        
        <td width="56" height="25" align="left" valign="middle" nowrap="nowrap">
        <td width="6">
        </td>
        <td width="86">
        </td>
    </tr>
  </table>
</div>
<br>

<div class="titulo2" id="informacion">
  <h2>> Información</h2>
  <h3>Seleccione una campaña para visualizar, editar y eliminar.</h3>
</div>

<div class="info" id="filtros">

<table width="100%" border="0">
  <tr>
    <td width="79" height="25" valign="middle"><h4>Campaña:</h4></td>
    <td width="586" height="25" valign="middle">
	
   	<select name="CB_CAMPANA" id="CB_CAMPANA">
						<option value="">Selecione</option>
					</select>
    </td>
  </tr>
</table>


</div>


<div class="info" id="info_basica">

	<div id="right">
    <p><label id="FECHA_CREACION"></label><label id="fecha_modificacion"></label></p>
    <p><label id="usuario"></label>  <label id="usuario_modifica"></label></p>
	
					
	</div>

  <table width="580" border="0">
    <tr>
      <td width="81" height="20" valign="middle"><h4>Nombre:</h4></td>
      <td height="20" colspan="4" align="left" valign="middle" nowrap="nowrap">
	  <input id="Txt_Nom_Campana_Mod" type="text" size="45" maxlength="20" >
	  <span id="span_Nom_Campana_mod" class="span_aviso_rojo">*</span>
	  </td>
      <td width="6"></td>
      <td width="86"><h4>Ingreso Rut: </h4></td>
      <td width="149" rowspan="5" valign="top" class="tabla_rut">
	  <textarea  cols="20" rows="10" id="Txt_Rut_Mod"></textarea>
	  <span id="span_Rut_Campana_mod" class="span_aviso_rojo">*</span></td>
	  
	  
    </tr>
    <tr>
      <td width="81" height="25" valign="middle"><h4>Descripción:</h4></td>
      <td height="25" colspan="4" align="left" valign="middle" nowrap="nowrap">
	  <input id="Txt_Des_Campana_Mod" type="text" value="" size="45" maxlength="40" onChange=""></td>
      <td width="6"></td>
      <td width="86"></td>
    </tr>
    <tr>
    	<td width="81">
    	<h4>Fecha Inicio:</h4>
    	</td>
    	<td width="60" align="left" valign="middle" nowrap="nowrap">
		<input type="text" id="Txt_Fecha_Inicio_Mod" readonly value="" size="10" maxlength="10">        
			<span id="span_Fecha_Inicio_Mod" class="span_aviso_rojo">*</span>
        <td width="4" align="left" valign="middle" nowrap="nowrap">        
        <td width="100" align="left" valign="middle" nowrap="nowrap"><h4>Fecha Término:</h4>
        <td width="56" height="25" align="left" valign="middle" nowrap="nowrap">
		<input  type="text" id="Txt_Fecha_Fin_Mod" readonly value="" size="10" maxlength="10">
			<span id="span_Fecha_Fin_Mod" class="span_aviso_rojo">*</span>
        <td width="6">
        <td width="86">
        </td>
    </tr>
	<tr>
    	<td width="81">
    	<h4>Observación:</h4>
    	</td>
    	<td width="86" height="25" align="left" valign="middle" nowrap="nowrap" colspan="4">
				<textarea  cols="45" rows="4" id="Txt_Observacion_Mod" maxlength="40" ></textarea>      
				<span id="span_Observacion_Mod" class="span_aviso_rojo">*</span>
		</td>
        <td width="6">
        <td width="86">
        </td>
    </tr>
    <tr>
    	<td width="81">
    	<h4>&nbsp;</h4>
    	</td>
    	<td height="25" colspan="4" align="left" valign="middle" nowrap="nowrap"><acronym title="REFRESCAR">
    	  <input id="Btn_Mod_Campana"class="fondo_boton_100" type="button"  value="Guardar">
    	  <input id="Btn_Elimina"class="fondo_boton_100" type="button" value="Eliminar">
    	</acronym>
    	<td width="6">
        </td>
        <td width="86">
        </td>
    </tr>
  </table>
</div>


<div class="informe" id="resumen_general">
<h2>> Resumen General - Deuda Activa</h2>
<br>
<!--div utilizado para cargar los datos del resumen general-->  
<DIV ID="Tbl_Resumen"></div>
</div>
<br>

<div class="informe" id="ejecutivos_asociados">
<h2>> Ejecutivos Asociados a Campaña</h2>
<br>
<!--div utilizado para cargar los datos del resumen ejecutivos-->  
<DIV ID="Tbl_Ejecutivos"></div>
</div>
<br>
</td>
</tr>
</table>
</form>
</body>
</html>


