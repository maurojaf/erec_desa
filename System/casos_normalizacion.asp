<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"  LCID = 1034%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<%
	AbrirSCG()
	Response.CodePage=65001
	Response.charset ="utf-8"

	strCodCliente 		=session("ses_codcli")
	intVerCobExt 		="1"
	intVerEjecutivos 	="1"

	strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
	strSql = strSql & " FROM CLIENTE CL"
	strSql = strSql & " WHERE CL.COD_CLIENTE = '" & strCodCliente & "'"

	set RsCli=Conn.execute(strSql)
	If not RsCli.eof then
		intUsaCobInterna = RsCli("USA_COB_INTERNA")
	End if
	RsCli.close
	set RsCli=nothing


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

	IF session("Ftro_TipoGestionCasosNorm")="" THEN

		strObEstadoTipoGestIP =" SELECTED "
		strObEstadoTipoGestPNA =" SELECTED "
		strObEstadoTipoGestCP =" SELECTED "

	ELSE
	
		IF instr(session("Ftro_TipoGestionCasosNorm"), "1")>0 then
			strObEstadoTipoGestIP =" SELECTED "
		END IF 

		IF instr(session("Ftro_TipoGestionCasosNorm"), "2")>0 then
			strObEstadoTipoGestCP =" SELECTED "
		END IF
		
		IF instr(session("Ftro_TipoGestionCasosNorm"), "3")>0 then
			strObEstadoTipoGestPNA =" SELECTED "
		END IF 

	END IF
	
	IF session("Ftro_EstadoProcesoCasosNorm")="" THEN

		strObEstadoProcesoNP =" SELECTED "
		strObEstadoProcesoNR =" SELECTED "
		strObEstadoProcesoNC =""

	ELSE

		IF instr(session("Ftro_EstadoProcesoCasosNorm"), "1")>0 then
			strObEstadoProcesoNP =" SELECTED "
		END IF 
		
		IF instr(session("Ftro_EstadoProcesoCasosNorm"), "2")>0 then
			strObEstadoProcesoNR =" SELECTED "
		END IF 

		IF instr(session("Ftro_EstadoProcesoCasosNorm"), "3")>0 then
			strObEstadoProcesoNC =" SELECTED "
		END IF

	END IF
	
	If session("Ftro_Cp_Adjunto") <> "" then
	
		CH_CP_ADJUNTO = session("Ftro_Cp_Adjunto") 
		
	End If
	
	'Response.write "valor : " & session("Ftro_Cp_Adjunto")

	sql_usuario ="SELECT DISTINCT U.ID_USUARIO, LOGIN "
	sql_usuario = sql_usuario & " FROM USUARIO U "
	sql_usuario = sql_usuario & " INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO "
	sql_usuario = sql_usuario & " AND UC.COD_CLIENTE = '"&trim(strCodCliente)&"' "
	sql_usuario = sql_usuario & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1 "

	If Trim(strCobranza) = "EXTERNA" Then 'externa
		sql_usuario = sql_usuario & " AND U.PERFIL_EMP=0 " 
	End If
	
	If Trim(strCobranza) = "INTERNA" Then 'interna
		sql_usuario = sql_usuario & " AND U.PERFIL_EMP=1 "
	End If
	set rs_usuario = conn.execute(sql_usuario)

%>

	<link href="../css/style_multi_select.css" rel="stylesheet"> 
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<script src="../Componentes/jquery.multiselect.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
	<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
	<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>
	<script src="../Componentes/Timepicker/jquery.timepickerinputmask.min.js"></script> 
	
	<script language="JavaScript " type="text/JavaScript">
	$(document).ready(function(){
 	$.prettyLoader();

	consulta_detalle()

	$('#TX_INICIO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_TERMINO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FECHA_CONSULTA').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})	
	$('#TX_FECHA_AGENDAMIENTO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})	


    $('#TX_HORA_AGENDAMIENTO').TimepickerInputMask({
        seconds: false
    }); 


	$(document).tooltip();
	$("#CMB_TIPO_GESTION").multiselect({minWidth:175});
	$("#CMB_ESTADO_PROCESO").multiselect({minWidth:175});

	$('.td_hover').hover(function(){
		$(this).css('background-color','#CEE3F6')
	}, function(){
		$(this).css('background-color','')
	})

	$('#marcar_todos').toggle(function(){

		$('input[id="CH_CASOS_APOYO"]').each(function(){
			$(this).attr('checked', true)
		})


		$(this).text("Desmarcar todos")

	}, function(){
		$('input[id="CH_CASOS_APOYO"]').each(function(){
			$(this).removeAttr('checked')
		})
		$(this).text("Marcar todos")
	})

	$('input[id="CH_CASOS_APOYO"]').live('click',function(){
		if($(this).is(':checked')){	
			$('td[id*="refresca_busca_cuotas_"]').text("")
		}
	})
	 

 
})

function CargaUsuarios(valor){

	var criterios ="alea="+Math.random()+"&accion_ajax=filtra_usuario&CB_COBRANZA="+valor
	$('#filtrado_ejecutivo').load('FuncionesAjax/casos_normalizacion_ajax.asp', criterios, function(){})
}

function consulta_detalle()
{
	var CB_COBRANZA 		=$('#CB_COBRANZA').val()
	var CMB_TIPO_GESTION 	=$('#CMB_TIPO_GESTION').val()
	var CMB_ESTADO_PROCESO 	=$('#CMB_ESTADO_PROCESO').val()
	var TX_INICIO 			=$('#TX_INICIO').val()
	var TX_TERMINO 			=$('#TX_TERMINO').val()
	var CB_EJECUTIVO 		=$('#CB_EJECUTIVO').val()
	var COD_CLIENTE 		=$('#COD_CLIENTE').val()
	var RUT_DEUDOR 			=$('#RUT_DEUDOR').val()
	var HORA_CONSULTA  		=$('#HORA_CONSULTA').val()
	var TX_FECHA_CONSULTA 	=$('#TX_FECHA_CONSULTA').val()
	var CH_CP_ADJUNTO 		=$('#CH_CP_ADJUNTO').val()

	if(CMB_TIPO_GESTION==null){
		CMB_TIPO_GESTION =""
		alert("Debe seleccionar tipo gestión")
		return
	}
	if(CMB_ESTADO_PROCESO==null){
		CMB_ESTADO_PROCESO =""
	}



	var criterios ="alea="+Math.random()+"&accion_ajax=refresa_normalizados&COD_CLIENTE="+COD_CLIENTE+"&CB_COBRANZA="+CB_COBRANZA+"&CMB_TIPO_GESTION="+CMB_TIPO_GESTION+"&CMB_ESTADO_PROCESO="+CMB_ESTADO_PROCESO+"&inicio="+TX_INICIO+"&termino="+TX_TERMINO+"&CB_EJECUTIVO="+CB_EJECUTIVO+"&inicia_contador=20&RUT_DEUDOR="+RUT_DEUDOR+"&HORA_CONSULTA="+HORA_CONSULTA+"&FECHA_CONSULTA="+TX_FECHA_CONSULTA+"&CH_CP_ADJUNTO="+CH_CP_ADJUNTO	
	//alert(criterios)
	$('#refresca').load('FuncionesAjax/casos_normalizacion_ajax.asp', criterios, function(){})
}



function consulta_resumen()
{
	var CB_COBRANZA 		=$('#CB_COBRANZA').val()
	var CMB_TIPO_GESTION 	=$('#CMB_TIPO_GESTION').val()
	var CMB_ESTADO_PROCESO 	=$('#CMB_ESTADO_PROCESO').val()
	var TX_INICIO 			=$('#TX_INICIO').val()
	var TX_TERMINO 			=$('#TX_TERMINO').val()
	var CB_EJECUTIVO 		=$('#CB_EJECUTIVO').val()
	var COD_CLIENTE 		=$('#COD_CLIENTE').val()
	var RUT_DEUDOR 			=$('#RUT_DEUDOR').val()
	var HORA_CONSULTA  		=$('#HORA_CONSULTA').val()
	var TX_FECHA_CONSULTA 	=$('#TX_FECHA_CONSULTA').val()
	var CH_CP_ADJUNTO 		=$('#CH_CP_ADJUNTO').val()

	if(CMB_TIPO_GESTION==null){
		CMB_TIPO_GESTION =""
		return
	}

	if(CMB_ESTADO_PROCESO==null){
		CMB_ESTADO_PROCESO =""
	}

	$.prettyLoader.show(2000);
	$('#boton_ver').attr('disabled', true)

	var criterios ="alea="+Math.random()+"&accion_ajax=refresca_resumen&COD_CLIENTE="+COD_CLIENTE+"&CB_COBRANZA="+CB_COBRANZA+"&CMB_TIPO_GESTION="+CMB_TIPO_GESTION+"&CMB_ESTADO_PROCESO="+CMB_ESTADO_PROCESO+"&inicio="+TX_INICIO+"&termino="+TX_TERMINO+"&CB_EJECUTIVO="+CB_EJECUTIVO+"&RUT_DEUDOR="+RUT_DEUDOR+"&HORA_CONSULTA="+HORA_CONSULTA+"&FECHA_CONSULTA="+TX_FECHA_CONSULTA+"&CH_CP_ADJUNTO="+CH_CP_ADJUNTO 
	//alert(criterios)
	$('#refresca_resumen').load('FuncionesAjax/casos_normalizacion_ajax.asp', criterios, function(){
		$('#boton_ver').attr('disabled', false)
		
	})
}


function exportar()
{
	var CB_COBRANZA 		=$('#CB_COBRANZA').val()
	var CMB_TIPO_GESTION 	=$('#CMB_TIPO_GESTION').val()
	var CMB_ESTADO_PROCESO 	=$('#CMB_ESTADO_PROCESO').val()
	var TX_INICIO 			=$('#TX_INICIO').val()
	var TX_TERMINO 			=$('#TX_TERMINO').val()
	var CB_EJECUTIVO 		=$('#CB_EJECUTIVO').val()
	var COD_CLIENTE 		=$('#COD_CLIENTE').val()
	var RUT_DEUDOR 			=$('#RUT_DEUDOR').val()
	var HORA_CONSULTA  		=$('#HORA_CONSULTA').val()
	var TX_FECHA_CONSULTA 	=$('#TX_FECHA_CONSULTA').val()
	var CH_CP_ADJUNTO 		=$('#CH_CP_ADJUNTO').val()

	$.prettyLoader.show(2000);

	$('#boton_exportar').attr('disabled', true)

	if(CMB_TIPO_GESTION==null){
		CMB_TIPO_GESTION =""
	}
	if(CMB_ESTADO_PROCESO==null){
		CMB_ESTADO_PROCESO =""
	}

	location.href="exp_casos_normalizacion.asp?accion_ajax=refresa_normalizados&COD_CLIENTE="+COD_CLIENTE+"&CB_COBRANZA="+CB_COBRANZA+"&CMB_TIPO_GESTION="+CMB_TIPO_GESTION+"&CMB_ESTADO_PROCESO="+CMB_ESTADO_PROCESO+"&inicio="+TX_INICIO+"&termino="+TX_TERMINO+"&CB_EJECUTIVO="+CB_EJECUTIVO+"&inicia_contador=20&RUT_DEUDOR="+RUT_DEUDOR+"&HORA_CONSULTA="+HORA_CONSULTA+"&FECHA_CONSULTA="+TX_FECHA_CONSULTA+"&CH_CP_ADJUNTO="+CH_CP_ADJUNTO

	setTimeout("$('#boton_exportar').attr('disabled', false)",3000)
}


function bt_mostrar_mas_registros(inicio_top){
	var CB_COBRANZA 		=$('#CB_COBRANZA').val()
	var CMB_TIPO_GESTION 	=$('#CMB_TIPO_GESTION').val()
	var CMB_ESTADO_PROCESO 	=$('#CMB_ESTADO_PROCESO').val()
	var TX_INICIO 			=$('#TX_INICIO').val()
	var TX_TERMINO 			=$('#TX_TERMINO').val()
	var CB_EJECUTIVO 		=$('#CB_EJECUTIVO').val()
	var COD_CLIENTE 		=$('#COD_CLIENTE').val()
	var RUT_DEUDOR 			=$('#RUT_DEUDOR').val()
	var HORA_CONSULTA  		=$('#HORA_CONSULTA').val()
	var TX_FECHA_CONSULTA 	=$('#TX_FECHA_CONSULTA').val()
	var CH_CP_ADJUNTO 		=$('#CH_CP_ADJUNTO').val()

	if(CMB_TIPO_GESTION==null){
		CMB_TIPO_GESTION =""
	}
	if(CMB_ESTADO_PROCESO==null){
		CMB_ESTADO_PROCESO =""
	}

	var criterios ="alea="+Math.random()+"&accion_ajax=refresa_normalizados&COD_CLIENTE="+COD_CLIENTE+"&CB_COBRANZA="+CB_COBRANZA+"&CMB_TIPO_GESTION="+CMB_TIPO_GESTION+"&CMB_ESTADO_PROCESO="+CMB_ESTADO_PROCESO+"&inicio="+TX_INICIO+"&termino="+TX_TERMINO+"&CB_EJECUTIVO="+CB_EJECUTIVO+"&inicia_contador="+inicio_top+"&RUT_DEUDOR="+RUT_DEUDOR+"&HORA_CONSULTA="+HORA_CONSULTA+"&FECHA_CONSULTA="+TX_FECHA_CONSULTA+"&CH_CP_ADJUNTO="+CH_CP_ADJUNTO	

	//alert(criterios)
	$('#refresca').load('FuncionesAjax/casos_normalizacion_ajax.asp', criterios, function(){})

}

function procesar(){

	var CH_ID_GESTION 			=""
	var CH_ID_CUOTA  			=""
	var concat_CH_ID_GESTION 	=""
	var concat_CH_ID_CUOTA  	=""	
	var concat_VARIABLE	 		=""
	var TX_OBSERVACION_CONSULTA =$('#TX_OBSERVACION_CONSULTA').val()
	var TX_FECHA_AGENDAMIENTO 	=$('#TX_FECHA_AGENDAMIENTO').val()
	var TX_HORA_AGENDAMIENTO 	=$('#TX_HORA_AGENDAMIENTO').val()

	if(TX_FECHA_AGENDAMIENTO==""){
		alert("Debe ingresar fecha agendamiento")
		return
	}

	if(TX_HORA_AGENDAMIENTO==""){
		alert("Debe ingresar hora agendamiento")
		return
	}

	$('input[id="CH_CASOS_APOYO"]:checked').each(function(){
		concat_CH_ID_GESTION 	=concat_CH_ID_GESTION+"*"+$(this).val() 
		
	})
	if (concat_CH_ID_GESTION==""){
		alert("Debe seleccionar una gestión")
		return
	}
	var criterios ="alea="+Math.random()+"&accion_ajax=proceso_casos&concat_CH_ID_GESTION="+concat_CH_ID_GESTION+"&TX_OBSERVACION_CONSULTA="+TX_OBSERVACION_CONSULTA+"&TX_FECHA_AGENDAMIENTO="+TX_FECHA_AGENDAMIENTO+"&TX_HORA_AGENDAMIENTO="+TX_HORA_AGENDAMIENTO

	$('#proceso_casos').load('FuncionesAjax/casos_normalizacion_ajax.asp', criterios, function(){

		$('#TX_OBSERVACION_CONSULTA').val("")
		consulta_resumen()
		consulta_detalle()	
		$('#ventana_procesa').dialog( "close" )	

	})
}


function ventana_procesa(){
	var concat_CH_ID_GESTION 	=""
	$('input[id="CH_CASOS_APOYO"]:checked').each(function(){
		concat_CH_ID_GESTION 	=concat_CH_ID_GESTION+"*"+$(this).val() 
		
	})
	if (concat_CH_ID_GESTION==""){
		alert("Debe seleccionar una gestión")
		return
	}

	$("#TX_FECHA_AGENDAMIENTO").datepicker("disable");
	
	$('#ventana_procesa').dialog({
   		show:"blind", 
   		hide:"explode",   		       	 
    	width:550,
    	height:370 ,
    	modal:true,	
	    buttons: {
            Si: function() {
                procesar()
            },
            No: function() {
				$(this).dialog( "close" );
            }
        }  	
	});	
	
	$("#TX_FECHA_AGENDAMIENTO").datepicker("enable");

}


function busca_cuotas(intIdGestion, RUT_DEUDOR){
	var strCodCliente 			=$('#COD_CLIENTE').val()
	var concat_refresca 		="#refresca_busca_cuotas_"+intIdGestion
	var concat_muestra_cuotas 	="#imagen_muestra_cuotas_"+intIdGestion
	var concat_oculta_cuotas 	="#imagen_oculta_cuotas_"+intIdGestion



	$("img[id*=imagen_muestra_cuotas_]").css('display','block')
	$("img[id*=imagen_oculta_cuotas_]").css('display','none')
	$("td[id*=refresca_busca_cuotas_]").text("")
	$(concat_muestra_cuotas).css('display','none')
	$(concat_oculta_cuotas).css('display','block')



	var criterios ="alea="+Math.random()+"&accion_ajax=refresca_busca_cuotas&intIdGestion="+intIdGestion+"&RUT_DEUDOR="+RUT_DEUDOR
	$(concat_refresca).load('FuncionesAjax/casos_normalizacion_ajax.asp', criterios, function(){
		var strIDCuotas =$('#strIDCuotas').val()
		//alert(strIDCuotas)
		if(strIDCuotas!=""){

			var criterios ="alea="+Math.random()+"&accion_ajax=mostrar_todos_cuotas&rut="+RUT_DEUDOR+"&strCodCliente="+strCodCliente+"&ID_GESTION="+intIdGestion+"&strIDCuotas="+strIDCuotas+"&CH_TODOS_CUOTA=1&pagina_origen=casos_normalizacion"

			$(concat_refresca).load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){
				$("#table_tablesorter").tablesorter(); 

				$('.cambio_flecha_ordenamiento').toggle(function(){
					$('.flecha_ordenamiento').attr('src', '../Imagenes/flecha_arriba_ordenamiento.png')
				}, function(){
					$('.flecha_ordenamiento').attr('src', '../Imagenes/flecha_abajo_ordenamiento.png')
				})

			})

		}
	})

}



function oculta_cuotas(intIdGestion, RUT_DEUDOR){
	var concat_muestra_cuotas 	="#imagen_muestra_cuotas_"+intIdGestion
	var concat_oculta_cuotas 	="#imagen_oculta_cuotas_"+intIdGestion

	$(concat_muestra_cuotas).css('display','block')
	$(concat_oculta_cuotas).css('display','none')

	var concat_refresca 		="#refresca_busca_cuotas_"+intIdGestion
	$(concat_refresca).text("")

}

function ventanaMas (URL){
	window.open(URL,"DATOS1","width=840, height=450, scrollbars=no, menubar=no, location=no, resizable=yes")
}

function ventanaGestionesPorDoc (URL){
	window.open(URL,"DATOS2","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}


 $('textarea').live('keyup blur', function() {
        // Store the maxlength and value of the field.
        var maxlength = $(this).attr('maxlength');
        var val = $(this).val();

        // Trim the field if it has content over the maxlength.
        if (val.length > maxlength) {
            $(this).val(val.slice(0, maxlength));
        }
});

function bt_ver_historial(ID_CUOTA)
{

	window.open('historial_documentos_biblioteca_deudor.asp?ID_CUOTA='+ID_CUOTA,"_new","width=900, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

	function ventanaBusqueda (URL){
		window.open(URL,"DATOS3","width=1050, height=700, scrollbars=yes, menubar=no, location=no, resizable=yes")
	}


function bt_trae_cuotas_vista(ID_GESTION,RUT_DEUDOR, COD_CLIENTE){
	var concat_refresca ="#refresca_busca_cuotas_"+ID_GESTION


	var criterios ="alea="+Math.random()+"&accion_ajax=refresca_busca_cuotas&intIdGestion="+ID_GESTION+"&RUT_DEUDOR="+RUT_DEUDOR

	$(concat_refresca).load('FuncionesAjax/casos_normalizacion_ajax.asp', criterios, function(){
		var strIDCuotas =$('#strIDCuotas').val()

		if(strIDCuotas!=""){
			window.location ="detalle_gestiones.asp?rut="+RUT_DEUDOR+"&cliente="+COD_CLIENTE+"&strIDCuotas="+strIDCuotas+"&strNuevaGestion=S&pagina_origen=casos_normalizacion"
			
		}
	})

}

function ValidaHora( ObjIng, strHora )
{
    var er_fh = /^(00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23)\:([0-5]0|[0-5][1-9])$/

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



<style type="text/css">

	.mas_registros{
		padding: 5px;
		width: 30%;
		font-size: 12px;
		font-weight: bold;
		font-family: "verdana";	
		cursor: pointer;
		margin:0 auto;
		text-align: center;
	}

	.td_hover{
		height: 22px;
	}
	.td_hover:nth-child(even) {
	    background: #F0F0F0; 
	    height: 22px;
	}
	.td_hover:nth-child(odd) {
	    background: #FFF;
	    height: 22px;
	}	
	.td_bordes{
		border: 1px solid #ccc;
	}
	#marcar_todos{
		margin-left: 5%;
		cursor:pointer;
		width:150px;
		color:#0040FF;
		text-decoration: underline;
	}
	.pregunta{
		margin-top: 15px;
		font-size: 16px;
		color:#1C1C1C;
	}

	.campo_proceso{
		font-size: 12px;
		color:#1C1C1C;		
	}
</style>
</head>
<body>
<input type="hidden" name="COD_CLIENTE" ID="COD_CLIENTE" value="<%=TRIM(strCodCliente)%>">
<div class="titulo_informe">MÓDULO NORMALIZACION</div>
<br>	

	<table width="90%" class="estilo_columnas" align="center">
		<thead>
	      <tr height="20" >
			<td>COBRANZA</td>
			<td>TIPO GESTIÓN</td>
			<td>ESTADO PROCESO</td>
			<td>FECHA DESDE</td>
			<td>FECHA HASTA</td>
		  	<% If sinCbUsario = "0" Then %>
				<td>EJECUTIVO</td>
		  	<% End If %>
			<td>RUT DEUDOR</td>
			<td>FECHA CONSULTA</td>
			<td>HORA CUNSULTA</td>
			<td>CP ADJUNTO</td>
		  </tr>
	    </thead>
		  <tr>
			<td>
				<select name="CB_COBRANZA" id="CB_COBRANZA" <%If sinCbUsario = "0" then%> onChange="CargaUsuarios(this.value);" <%End If%> >
						<%If Trim(intVerCobExt) = "1" and Trim(intUsaCobInterna) = "1" Then%>
							<option value="TODOS" <%If Trim(strCobranza) ="" Then Response.write "SELECTED"%>>TODOS</option>
						<%End If%>

						<%If Trim(intUsaCobInterna) = "1" Then%>
							<option value="INTERNA" <%If Trim(strCobranza) ="INTERNA" Then Response.write "SELECTED"%>>INTERNA</option>
						<%End If%>

						<%If Trim(intVerCobExt) = "1" Then%>
							<option value="EXTERNA" <%If Trim(strCobranza) ="EXTERNA" Then Response.write "SELECTED"%>>EXTERNA</option>
						<%End If%>
					
				</select>
			</td>

			<td>

				<SELECT NAME="CMB_TIPO_GESTION" id="CMB_TIPO_GESTION" style="background-color:#ccc;" multiple>
					<option value="1" <%=strObEstadoTipoGestIP%> >INDICA QUE PAGO</option>
					<option value="2" <%=strObEstadoTipoGestCP%> >COMPROMISO D y T</option>
					<option value="3" <%=strObEstadoTipoGestPNA%> >PAGO NO APLICADO</option>
				</SELECT>
			</td>
			
			<td>
				<SELECT NAME="CMB_ESTADO_PROCESO" id="CMB_ESTADO_PROCESO" multiple style="width:100px;background-color:#ccc;">
					<option value="1" <%=strObEstadoProcesoNP%> >NO PROCESADO</option>
					<option value="2" <%=strObEstadoProcesoNR%> >NO RESPONDIDO</option>
					<option value="3" <%=strObEstadoProcesoNC%> >EN CONSULTA</option>
				</SELECT>
			</td>

			<td>
				<input name="TX_INICIO" type="text" id="TX_INICIO" readonly="true" value="" size="10" maxlength="10">
			</td>

			<td>
				<input name="TX_TERMINO" type="text" id="TX_TERMINO" readonly="true" value="" size="10" maxlength="10">
			</td>

		<% If sinCbUsario="0" Then %>
			<td id="filtrado_ejecutivo">

				<select name="CB_EJECUTIVO" id="CB_EJECUTIVO">
					<option value="">TODOS</option>
					<%do while not rs_usuario.eof%>
						<option <%if trim(session("Ftro_EjecAsigCasosObj"))=trim(rs_usuario("ID_USUARIO")) then Response.write " selected " end if%> value="<%=rs_usuario("ID_USUARIO")%>"><%=rs_usuario("LOGIN")%></option>
					<%rs_usuario.movenext
					loop%>
				</select>
			</td>
		<% End If %>
			<td>
				<input name="RUT_DEUDOR" type="text" id="RUT_DEUDOR" value="" size="12" maxlength="11">
			</td>
			<td>
				<input name="TX_FECHA_CONSULTA" type="text" id="TX_FECHA_CONSULTA" value="" size="12" maxlength="11">
			</td>
			<td>
				<input name="HORA_CONSULTA" type="text" id="HORA_CONSULTA" onChange="return ValidaHora(this,this.value)" value="" size="12" maxlength="11">
			</td>
			<td>
				<select id="CH_CP_ADJUNTO" STYLE="WIDTH:70PX;" name="CH_CP_ADJUNTO">

					<option value="" <%If Trim(CH_CP_ADJUNTO) ="" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="1" <%If Trim(CH_CP_ADJUNTO) ="1" Then Response.write "SELECTED"%>>SI</option>
					<option value="2" <%If Trim(CH_CP_ADJUNTO) ="2" Then Response.write "SELECTED"%>>NO</option>
					
				 </select> 
			</td>						
	      </tr>
	      <tr>
	      	<td colspan="6" align="right">

	      	</td>
	      </tr>
    </table>
	<br>

	<div id="refresca_resumen">
	<table class="" style="width:78%; margin-left:5%;" align="left" cellSpacing="0" cellPadding="0" border="0">
		<tr>
			<td style="width:100px;" class=""></td>
			<td style="width:100px;" class="estilo_columna_individual td_bordes" align="center">NO RESPONDIDOS</td>
			<td style="width:100px;" class="estilo_columna_individual td_bordes" align="center">NO PROCESADOS</td>
			<td style="width:100px;" class="estilo_columna_individual td_bordes" align="center">EN CONSULTA</td>
			<td style="width:100px;background-color:#424242;" class="estilo_columna_individual td_bordes" align="center">TOTAL ACUM. GESTION</td>
		</tr>
		<tr>
			<td style="width:100px;" class="estilo_columna_individual" align="center">TIPO GESTIÓN</td>
			<td class="" align="center">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" height="25" align="center"><b>Total Doc</b></td>
					<td width="33%" height="25" align="center"><b>Total casos</b></td>
					<td width="33%" height="25" align="center"><b>Monto Doc</b></td>
				</tr>					
				</table>
			</td>
			<td class="" align="center">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" height="25" align="center"><b>Total Doc</b></td>
					<td width="33%" height="25" align="center"><b>Total casos</b></td>
					<td width="33%" height="25" align="center"><b>Monto Doc</b></td>
				</tr>					
				</table>
			</td>
			<td class="" align="center">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" height="25" align="center"><b>Total Doc</b></td>
					<td width="33%" height="25" align="center"><b>Total casos</b></td>
					<td width="33%" height="25" align="center"><b>Monto Doc</b></td>
				</tr>					
				</table>
			</td>
			<td class="" align="center">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" height="25" align="center"><b>Totales Doc</b></td>
					<td width="33%" height="25" align="center"><b>Totales casos</b></td>
					<td width="33%" height="25" align="center"><b>Monto tot Doc</b></td>
				</tr>					
				</table>
			</td>			
		</tr>
		<%

			'###Inicio Código Informe###'
				
			sql_resumen = " SELECT TIPO_GESTION = (CASE WHEN TIPO_MODULO = 2"
			sql_resumen = sql_resumen & " 							THEN 'INDICA QUE PAGO'" 
			sql_resumen = sql_resumen & " 							WHEN (TIPO_MODULO = 1 OR TIPO_MODULO = 11) AND FORMA_PAGO IN ('TR','DP')"
			sql_resumen = sql_resumen & " 							THEN 'COMPROMISO D & T'" 
			sql_resumen = sql_resumen & " 							ELSE 'PAGO NO APLICADO'"
			sql_resumen = sql_resumen & " 					   END),"
			
			sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'NO RESPONDIDO' THEN 1 ELSE 0 END),0) AS TOTAL_DOC_NO_RESPONDIDO, "
			sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'NO RESPONDIDO' THEN SALDO_CUOTA ELSE 0 END),0) AS TOTAL_SALDO_NO_RESPONDIDO, "
			sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'NO RESPONDIDO' THEN RUT_DEUDOR END)) AS TOTAL_CASOS_NO_RESPONDIDO, "
			sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'NO RESPONDIDO' THEN ID_GESTION END)) AS TOTAL_GESTIONES_NO_RESPONDIDO,"
			sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'NO PROCESADO' THEN 1 ELSE 0 END),0) AS TOTAL_DOC_NO_PROCESADO, "
			sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'NO PROCESADO' THEN SALDO_CUOTA ELSE 0 END),0) AS TOTAL_SALDO_NO_PROCESADO, "
			sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'NO PROCESADO' THEN RUT_DEUDOR END)) AS TOTAL_CASOS_NO_PROCESADO, "
			sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'NO PROCESADO' THEN ID_GESTION END)) AS TOTAL_GESTIONES_NO_PROCESADO,"


			sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'EN CONSULTA' THEN 1 ELSE 0 END),0) AS TOTAL_DOC_NO_CONSULTA, "
			sql_resumen = sql_resumen & " ISNULL(SUM(CASE WHEN PROCESO = 'EN CONSULTA' THEN SALDO_CUOTA ELSE 0 END),0) AS TOTAL_SALDO_NO_CONSULTA, "
			sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'EN CONSULTA' THEN RUT_DEUDOR END)) AS TOTAL_CASOS_NO_CONSULTA, "
			sql_resumen = sql_resumen & " COUNT(DISTINCT (CASE WHEN PROCESO = 'EN CONSULTA' THEN ID_GESTION END)) AS TOTAL_GESTIONES_NO_CONSULTA "
			sql_resumen = sql_resumen & " FROM VIEW_CASOS_GESTION_APOYO "
			
			sql_resumen = sql_resumen & " WHERE COD_CLIENTE ='"&TRIM(strCodCliente)&"'"
			
			sql_resumen = sql_resumen & " AND (TIPO_MODULO = 2 OR (TIPO_MODULO = 6 AND AGEND_VENC = 1) OR ((TIPO_MODULO = 1 OR TIPO_MODULO = 11) AND FORMA_PAGO IN ('TR','DP') AND DATEDIFF(DAY,FECHA_COMPROMISO,GETDATE()) >= 0))"
	
			if trim(strCobranza)="INTERNA" then
				sql_resumen = sql_resumen & " AND CUSTODIO IS NOT  NULL "

			ElseIf trim(strCobranza)="EXTERNA" then
				sql_resumen = sql_resumen & " AND CUSTODIO IS NULL  "
				
			end if
			
			If TraeSiNo(session("perfil_adm")) = "No" and TraeSiNo(session("perfil_sup")) = "No" Then
				sql_resumen = sql_resumen & " 	AND  ID_USUARIO_ASIG = " & session("session_idusuario")
			Else
				if trim(strEjecutivo)<>"" then
					sql_resumen = sql_resumen & " AND  ID_USUARIO_ASIG ='"&strEjecutivo&"'"
				End if
			End if

			sql_resumen = sql_resumen & " GROUP BY (CASE  WHEN TIPO_MODULO = 2"
			sql_resumen = sql_resumen & " 				THEN 'INDICA QUE PAGO'" 
			sql_resumen = sql_resumen & " 				WHEN (TIPO_MODULO = 1 OR TIPO_MODULO = 11) AND FORMA_PAGO IN ('TR','DP')"
			sql_resumen = sql_resumen & " 				THEN 'COMPROMISO D & T' "
			sql_resumen = sql_resumen & " 				ELSE 'PAGO NO APLICADO'"
			sql_resumen = sql_resumen & " 		   END)"
			
			''response.write "Informe" & sql_resumen
				
			set rs_resumen =conn.execute(sql_resumen)
			
			if not rs_resumen.eof then
				do while not rs_resumen.eof

				strtipoGestion 					=rs_resumen("TIPO_GESTION")
				IntTotalDocNoRespondido 		=rs_resumen("TOTAL_DOC_NO_RESPONDIDO")
				IntTotalSaldoNoRespondido 		=rs_resumen("TOTAL_SALDO_NO_RESPONDIDO")
				IntTotalCasosNoRespondido	 	=rs_resumen("TOTAL_CASOS_NO_RESPONDIDO")
				intTotalGestionesNoRespondido	=rs_resumen("TOTAL_GESTIONES_NO_RESPONDIDO")
				intTotalDocNoProcesado 			=rs_resumen("TOTAL_DOC_NO_PROCESADO")
				intTotalSaldoNoProcesado 		=rs_resumen("TOTAL_SALDO_NO_PROCESADO")
				intTotalCasosNoProcesado	 	=rs_resumen("TOTAL_CASOS_NO_PROCESADO")
				intTotalGestionesNoProcesado	=rs_resumen("TOTAL_GESTIONES_NO_PROCESADO")
				intTotalDocNoConsulta 			=rs_resumen("TOTAL_DOC_NO_CONSULTA")
				intTotalSaldoNoConsulta 		=rs_resumen("TOTAL_SALDO_NO_CONSULTA")
				intTotalCasosNoConsulta	 		=rs_resumen("TOTAL_CASOS_NO_CONSULTA")
				intTotalGestionesNoConsulta	 	=rs_resumen("TOTAL_GESTIONES_NO_CONSULTA") 
				
				IntTotalGeneralDocNoRespondido  = IntTotalGeneralDocNoRespondido + IntTotalDocNoRespondido
				IntTotalGeneralSaldoNoRespondido = IntTotalGeneralSaldoNoRespondido + IntTotalSaldoNoRespondido
				IntTotalGeneralCasosNoRespondido = IntTotalGeneralCasosNoRespondido + IntTotalCasosNoRespondido
	
				intTotalGeneralDocNoProcesado = intTotalGeneralDocNoProcesado + intTotalDocNoProcesado
				intTotalGeneralSaldoNoProcesado = intTotalGeneralSaldoNoProcesado + intTotalSaldoNoProcesado
				intTotalGeneralCasosNoProcesado = intTotalGeneralCasosNoProcesado + intTotalCasosNoProcesado
				
				intTotalGeneralDocNoConsulta = intTotalGeneralDocNoConsulta + intTotalDocNoConsulta
				intTotalGeneralSaldoNoConsulta = intTotalGeneralSaldoNoConsulta + intTotalSaldoNoConsulta
				intTotalGeneralCasosNoConsulta = intTotalGeneralCasosNoConsulta + intTotalCasosNoConsulta
		
		%>			 	
				<tr>
				
				<td style="width:100px;" class="estilo_columna_individual td_bordes" align="center" >
					<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
					<tr>
						<td width="100%" style="height:20px;" align="center"><%=strtipoGestion%></td>
					</tr>
					</table>
				</td>	
				<td class="td_bordes">
					<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
					<tr>
						<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalDocNoRespondido,0)%></td>
						<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalCasosNoRespondido,0)%></td>
						<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalSaldoNoRespondido,0)%></td>
					</tr>
					</table>
				</td>
				<td class="td_bordes">
					<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
					<tr>
						<td width="33%" style="height:20px;" align="center"><%=FN(intTotalDocNoProcesado,0)%></td>
						<td width="33%" style="height:20px;" align="center"><%=FN(intTotalCasosNoProcesado,0)%></td>
						<td width="33%" style="height:20px;" align="center"><%=FN(intTotalSaldoNoProcesado,0)%></td>
					</tr>
					</table>
				</td>
				<td class="td_bordes">
					<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
					<tr>
						<td width="33%" style="height:20px;" align="center"><%=FN(intTotalDocNoConsulta,0)%></td>
						<td width="33%" style="height:20px;" align="center"><%=FN(intTotalCasosNoConsulta,0)%></td>
						<td width="33%" style="height:20px;" align="center"><%=FN(intTotalSaldoNoConsulta,0)%></td>
					</tr>
					</table>
				</td>				
				<td>
					<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
					<tr>
						<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalDocNoRespondido)+(intTotalDocNoProcesado)+(intTotalDocNoConsulta),0)%></td>
						<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalCasosNoRespondido)+(intTotalCasosNoProcesado)+(intTotalCasosNoConsulta),0)%></td>
						<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalSaldoNoRespondido)+(intTotalSaldoNoProcesado)+(intTotalSaldoNoConsulta),0)%></td>		
					</tr>
					</table>
				</td>				
				</tr>
<%
			 response.Flush()
			 rs_resumen.movenext
			 Loop
			end if

%>

		<tr>
			<td style="width:100px; background-color:#424242;" class="estilo_columna_individual" align="center">TOTAL ACUM. PROCESOS</td>
			<td class="td_bordes" bgcolor="#D8D8D8">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalGeneralDocNoRespondido,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalGeneralCasosNoRespondido,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(IntTotalGeneralSaldoNoRespondido,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes" bgcolor="#D8D8D8">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalGeneralDocNoProcesado,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalGeneralCasosNoProcesado,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalGeneralSaldoNoProcesado,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes" bgcolor="#D8D8D8">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalGeneralDocNoConsulta,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalGeneralCasosNoConsulta,0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN(intTotalGeneralSaldoNoConsulta,0)%></td>
				</tr>
				</table>
			</td>
			<td class="td_bordes" bgcolor="#D8D8D8">
				<table style="width:100%;"  cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalGeneralDocNoRespondido)+(intTotalGeneralDocNoProcesado)+(intTotalGeneralDocNoConsulta),0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalGeneralCasosNoRespondido)+(intTotalGeneralCasosNoProcesado)+(intTotalGeneralCasosNoConsulta),0)%></td>
					<td width="33%" style="height:20px;" align="center"><%=FN((IntTotalGeneralSaldoNoRespondido)+(intTotalGeneralSaldoNoProcesado)+(intTotalGeneralSaldoNoConsulta),0)%></td>
				</tr>
				</table>
			</td>			
		</tr>

	</table>

	<table style="float:right; margin-right:5%;" border="0">
	<tr>
		<td>
			<input 	type="button"  	class="fondo_boton_100" STYLE="MARGIN-TOP:2PX;" value="Ver" id="boton_ver" onClick="consulta_resumen();consulta_detalle();"><BR>
			<input  type="button" id="boton_exportar" 	class="fondo_boton_100" STYLE="MARGIN-TOP:2PX;" Value="Exportar" onClick="exportar();"><BR>
			
			<%If TraeSiNo(session("perfil_adm")) = "Si" OR TraeSiNo(session("perfil_sup")) = "Si" Then%>
				<input  type="button" 	class="fondo_boton_100" STYLE="MARGIN-TOP:2PX;" Value="Procesar" onClick="ventana_procesa();">	
			<%End if%>
			
		</td>
	</tr>

	</table>
	<br>
	<br>			
	</div>
	<div id="ventana_procesa" title="Procesa cuotas" style="display:none;">	
	<table align="center" width="500" align="right" cellSpacing="0" cellPadding="0" border="0">
	<tr>		
		<td align="left" colspan="2" class="titulo_informe" width="200">> AGENDAMIENTO</td>		
	</tr>
	<tr>		
		<td align="left" class="estilo_columna_individual" width="200">FECHA</td>		
		<td align="left" class="estilo_columna_individual" width="200">HORA</td>
	</tr>	
	<tr>
		<td align="left" class="" width="">
			<input type="text" name="TX_FECHA_AGENDAMIENTO" readondly  id="TX_FECHA_AGENDAMIENTO" value="<%=date()+7%>">
		</td>		
		<td align="left" class="" width="">
			<input type="text" name="TX_HORA_AGENDAMIENTO" id="TX_HORA_AGENDAMIENTO" value="">
		</td>		
	</tr>	
	<tr><td colspan="2">&nbsp;</td></tr>	
	<tr>		
		<td align="left" colspan="2" class="titulo_informe">> OBSERVACIÓN</td>	
	</TR>
	<TR>
		<td align="left" colspan="2" >
			<TEXTAREA id="TX_OBSERVACION_CONSULTA" maxlength="199" style="width:350px; height:50px;" name="TX_OBSERVACION_CONSULTA"></TEXTAREA>
		</td>				
	</tr>
	</table>
	<div class="pregunta"> 
		¿Esta seguro que desea enviar a consulta los siguientes documentos?
	</div>
	</div>
	<br>	
	<br>
	<div id="proceso_casos"></div>
	<div id="marcar_todos">Marcar todos</div>
	<table class="estilo_columnas" style="width:90%;" align="center" border="1" cellSpacing="0" cellPadding="0">
		<thead>
		<tr>
			<td width="20">&nbsp;</td>
			<td width="20">&nbsp;</td>
			<td width="130">TIPO GESTIÓN</td>
			<td width="100">ESTADO PROCESO</td>
			<td width="80">CONSULTA</td>
			<td width="40">ACUM.</td>

			<td width="70">INGRESO</td>
			<td width="80">RUT DEUDOR</td>
			<td width="80">SALDO GEST.</td>
			<td width="70">DIA MORA</td>
			<td width="70">FECHA</td>
			<td width="70">MONTO</td>

			<td width="100">FORMA NORM</td>
			<td width="100">LUGAR</td>
			<td width="70">Nº CP</td>
			<td width="70">EJECUTIVO</td>
			<td width="30" align="center">OBS</td>
			<td width="30" align="center">MAS</td>
			<td width="50" align="center">CUOTAS</td>
		</tr>
		</thead>
		<tbody>
			<tr>
				<td colspan="19" id="refresca">

				</td>
			</tr>
		</tbody>
	</table>
	<br>
	<br>
</body>
</html>
