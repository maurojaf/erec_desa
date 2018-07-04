<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"



	strCodCliente 	= session("ses_codcli")
	intVerCobExt 	= "1"


%>

    <title>Acceso e-Rec de Llacruz</title>
    <meta name="description" content="Acceso e-Rec de Llacruz">
    <meta name="author" content="Departamento desarrollo Llacruz">
    <!--[if lt IE 9]>
        <script src="http://html5shim.googlecode.com/svn/trunk/html5.js"></script>
    <![endif]-->

	<link href="../Css/style_estatus_cartera.css" rel="stylesheet">  
    
    <link href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css" rel="stylesheet"> 
    <link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet"> 


    <script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>  
    <script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>  
    <script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
    <script src="../Componentes/jquery.multiselect.js"></script>

    <!--#include file="arch_utils.asp"-->



<script type="text/javascript">
	$(document).ready(function(){


		$('#opcion_busqueda').click(function(){
			$('.filtro').slideToggle()
			var mostrar_oculto =$('#mostrar_oculto').val()

			if(mostrar_oculto=="S"){
				$('#mostrar_oculto').val("N")
				$('#mostrar_ocultar_opcion').text("mostrar")
				$('#img_mostrar_ocultar_opcion').attr("src", "../Imagenes/simbolo_mas_ec.fw.png")
				

			}else{
				$('#mostrar_oculto').val("S")	
				$('#mostrar_ocultar_opcion').text("ocultar")
				$('#img_mostrar_ocultar_opcion').attr("src", "../Imagenes/simbolo_menos_ec.fw.png")
			}			

		})
		
		
		$.prettyLoader();
		$('#fecha_gestion_desde').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#fecha_gestion_hasta').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#fecha_asignacion_desde').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#fecha_asignacion_hasta').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		//$("#COD_CLIENTE").multiselect();

		$('.contenido_detalle_ges').css('display', 'none')
		$('.contenido_detalle_negativa').css('display', 'none')
	
		$('#id_pendientes').hover(function(){
			$(this).css('background-color','#E0F2F7')
		}, function(){
			$(this).css('background-color','')
		}) 
		$('#id_negativa').hover(function(){
			$(this).css('background-color','#E0F2F7')
		}, function(){
			$(this).css('background-color','')
		}) 
		$('#id_tercero').hover(function(){
			$(this).css('background-color','#E0F2F7')
		}, function(){
			$(this).css('background-color','')
		}) 
		$('#id_gestionables').hover(function(){
			$(this).css('background-color','#E0F2F7')
		}, function(){
			$(this).css('background-color','')
		}) 				

  
		$('#id_pendientes').click(function(){			
			$('.contenido_detalle_negativa').css('display', 'block')
			$('#sacar_foco_neg').focus()
		})
		
		$('#id_negativa').click(function(){			
			$('.contenido_detalle_negativa').css('display', 'block')
			$('#sacar_foco_neg').focus()
		})
		
		$('#id_tercero').click(function(){			
			$('.contenido_detalle_negativa').css('display', 'block')
			$('#sacar_foco_neg').focus()
		})
		$('#sacar_foco_neg').blur(function(){
			$('.contenido_detalle_negativa').css('display', 'none')
		})
		$('#id_gestionables').click(function(){			
			$('.contenido_detalle_ges').css('display', 'block')
			$('#sacar_foco').focus()
		})
		$('#sacar_foco').blur(function(){
			$('.contenido_detalle_ges').css('display', 'none')
		})

		$('#COD_CLIENTE').change(function(){

			var criterios ="alea="+Math.random()+"&accion_ajax=refresa_ejecutivo&COD_CLIENTE="+$(this).val()
			$('#td_ejecutivo').load('FuncionesAjax/EstadoCartera_ajax.asp', criterios, function(){})

			var criterios ="alea="+Math.random()+"&accion_ajax=refresa_cobranza&COD_CLIENTE="+$(this).val()
			$('#td_cobranza').load('FuncionesAjax/EstadoCartera_ajax.asp', criterios, function(){})

			var criterios ="alea="+Math.random()+"&accion_ajax=refresa_rubro&COD_CLIENTE="+$(this).val()
			$('#td_rubro').load('FuncionesAjax/EstadoCartera_ajax.asp', criterios, function(){})

			var criterios ="alea="+Math.random()+"&accion_ajax=refresa_tipo_doc&COD_CLIENTE="+$(this).val()
			$('#id_tipo_doc').load('FuncionesAjax/EstadoCartera_ajax.asp', criterios, function(){})


			
		})


		$('#bt_filtro_limpiar').click(function(){

			$('#ID_USUARIO').val("")
			$('#fecha_asignacion_desde').val("")
			$('#fecha_asignacion_hasta').val("")
			$('#COD_ESTADO_COBRANZA').val("")
			$('#fecha_gestion_desde').val("")
			$('#fecha_gestion_hasta').val("")
			$('#TIPO_COBRANZA').val("")
			$('#CB_CAMPANA').val("")
			$('#CB_RUBRO').val("")
			$('#CB_TIPODOC').val("")			
		})
		

		$('#bt_filtro').click(function(){

			var COD_CLIENTE 			=""
			var ID_USUARIO 				=$('#ID_USUARIO').val()
			var fecha_asignacion_desde 	=$('#fecha_asignacion_desde').val()
			var fecha_asignacion_hasta 	=$('#fecha_asignacion_hasta').val()
			var COD_ESTADO_COBRANZA 	=$('#COD_ESTADO_COBRANZA').val()
			var fecha_gestion_desde 	=$('#fecha_gestion_desde').val()
			var fecha_gestion_hasta 	=$('#fecha_gestion_hasta').val()
			var TIPO_COBRANZA 			=$('#TIPO_COBRANZA').val()
			var CB_CAMPANA 				=$('#CB_CAMPANA').val()
			var CB_RUBRO 				=$('#CB_RUBRO').val()
			var CB_TIPODOC 				=$('#CB_TIPODOC').val()

			$('select[id="COD_CLIENTE"] option:checked').each(function(){
				COD_CLIENTE = COD_CLIENTE+","+$(this).val()
			})

			var COD_CLIENTE_MULTI= COD_CLIENTE.substring(1, COD_CLIENTE.Length)
			
			var criterios ="alea="+Math.random()+"&accion_ajax=filtra_estado_cartera&COD_CLIENTE="+COD_CLIENTE_MULTI+"&ID_USUARIO="+ID_USUARIO+"&fecha_asignacion_desde="+fecha_asignacion_desde+"&fecha_asignacion_hasta="+fecha_asignacion_hasta+"&COD_ESTADO_COBRANZA="+COD_ESTADO_COBRANZA+"&fecha_gestion_desde="+fecha_gestion_desde+"&fecha_gestion_hasta="+fecha_gestion_hasta+"&TIPO_COBRANZA="+TIPO_COBRANZA+"&CB_CAMPANA="+CB_CAMPANA+"&CB_RUBRO="+CB_RUBRO+"&CB_TIPODOC="+CB_TIPODOC

			$('#contenido').load('FuncionesAjax/EstadoCartera_ajax.asp', criterios, function(){

					$('.filtro').slideUp()

					$('#mostrar_oculto').val("N")
					$('#mostrar_ocultar_opcion').text("mostrar")
					$('#img_mostrar_ocultar_opcion').attr("src", "../Imagenes/simbolo_mas_ec.fw.png")					

					$('.contenido_detalle_ges').css('display', 'none')
					$('.contenido_detalle_negativa').css('display', 'none')

					$('#id_pendientes').click(function(){			
						$('.contenido_detalle_negativa').css('display', 'block')
						$('#sacar_foco_neg').focus()
					})
					
					$('#id_negativa').click(function(){			
						$('.contenido_detalle_negativa').css('display', 'block')
						$('#sacar_foco_neg').focus()
					})
					
					$('#id_tercero').click(function(){			
						$('.contenido_detalle_negativa').css('display', 'block')
						$('#sacar_foco_neg').focus()
					})
					$('#sacar_foco_neg').blur(function(){
						$('.contenido_detalle_negativa').css('display', 'none')
					})
					$('#id_gestionables').click(function(){			
						$('.contenido_detalle_ges').css('display', 'block')
						$('#sacar_foco').focus()
					})
					$('#sacar_foco').blur(function(){
						$('.contenido_detalle_ges').css('display', 'none')
					})

					$('#id_pendientes').hover(function(){
						$(this).css('background-color','#E0F2F7')
					}, function(){
						$(this).css('background-color','')
					}) 
					$('#id_negativa').hover(function(){
						$(this).css('background-color','#E0F2F7')
					}, function(){
						$(this).css('background-color','')
					}) 
					$('#id_tercero').hover(function(){
						$(this).css('background-color','#E0F2F7')
					}, function(){
						$(this).css('background-color','')
					}) 
					$('#id_gestionables').hover(function(){
						$(this).css('background-color','#E0F2F7')
					}, function(){
						$(this).css('background-color','')
					}) 				


			
			})
		})
	})

</script>
</head>
<%
abrirscg()

'###### CLIENTE
If session("perfil_adm") = true or session("perfil_sup") = true Then
	ssql_cliente="SELECT COD_CLIENTE,RAZON_SOCIAL, NOMBRE_FANTASIA FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
Else
	ssql_cliente="SELECT COD_CLIENTE,RAZON_SOCIAL, NOMBRE_FANTASIA FROM CLIENTE WHERE COD_CLIENTE = '" & strCodCliente & "' AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
End If
set rs_cliente= Conn.execute(ssql_cliente)


'###### CLIETNE INTERNO EXTERNO
strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
strSql = strSql & " FROM CLIENTE CL"
strSql = strSql & " WHERE CL.COD_CLIENTE = '" & strCodCliente & "'"

set RsCli=Conn.execute(strSql)
If not RsCli.eof then
	intUsaCobInterna = RsCli("USA_COB_INTERNA")
End if


If TraeSiNo(session("perfil_emp")) = "Si" and strCobranza = "" and intUsaCobInterna = "1" Then

	strCobranza="INTERNA"

ElseIf TraeSiNo(session("perfil_emp")) = "No" and strCobranza = "" then

	strCobranza="EXTERNA"

End If


'###### ESTADO COBRANZA
ssql_estado_cobranza="SELECT COD_ESTADO_COBRANZA, NOM_ESTADO_COBRANZA FROM ESTADO_COBRANZA"
set rs_est_cob= Conn.execute(ssql_estado_cobranza)

'###### EJECUTIVO
sql_usuario= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
sql_usuario= sql_usuario & " FROM USUARIO U "
sql_usuario= sql_usuario & " INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & trim(strCodCliente)
sql_usuario= sql_usuario & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
set rsUsuario=Conn.execute(sql_usuario)



'######## SELECT DETALLE ESTADO CARTERA

	sql_det =" SELECT ISNULL(COUNT(PP2.RUT_DEUDOR),0) AS TOTAL_RUT, ISNULL(SUM(SALDO_RUT),0) AS SALDO_RUT, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR "
	sql_det = sql_det & " PP2.DIR_SA=1	) THEN 1 ELSE 0 END)),0) AS RUT_GESTIONABLES, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 "
	sql_det = sql_det & " OR PP2.DIR_SA=1) THEN SALDO_RUT ELSE 0 END)),0) AS MONTO_GESTIONABLE, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=1 OR PP2.TEL_SA=1) AND (PP2.EMAIL_VA=0 AND PP2.EMAIL_SA=0) "
	sql_det = sql_det & " AND (PP2.DIR_VA=0 AND PP2.DIR_SA=0)) THEN 1 ELSE 0 END)),0) AS RUT_GES_TEL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=0 AND PP2.TEL_SA=0) AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1) "
	sql_det = sql_det & " AND (PP2.DIR_VA=0 AND PP2.DIR_SA=0)) THEN 1 ELSE 0 END)),0) AS RUT_GES_MAIL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=0 AND PP2.TEL_SA=0) AND (PP2.EMAIL_VA=0 AND PP2.EMAIL_SA=0) "
	sql_det = sql_det & " AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS RUT_GES_DIR, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=1 OR PP2.TEL_SA=1) AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1) "
	sql_det = sql_det & " AND (PP2.DIR_VA=0 AND PP2.DIR_SA=0)) THEN 1 ELSE 0 END)),0) AS RUT_GES_TEL_MAIL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=1 OR PP2.TEL_SA=1) AND (PP2.EMAIL_VA=0 AND PP2.EMAIL_SA=0) "
	sql_det = sql_det & " AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS RUT_GES_TEL_DIR, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=0 AND PP2.TEL_SA=0) AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1) "
	sql_det = sql_det & " AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS RUT_GES_MAIL_DIR, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN ((PP2.TEL_VA=1 OR PP2.TEL_SA=1) AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1) "
	sql_det = sql_det & " AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS RUT_GES_TEL_MAIL_DIR, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS CASOS_GESTIONADOS, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN SALDO_RUT ELSE 0 END)),0) AS MONTO_GESTIONADOS, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL=0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_PENDIENTES_CON_FONO, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL=0 AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_PENDIENTES_CON_MAIL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEST_GENERAL=0 AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_PENDIENTES_CON_DIRECCION, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_POSITIVA, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN SALDO_RUT ELSE 0 END)),0) AS MONTO_GESTION_POSITIVA, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE=0 AND PP2.TT_GEST_GENERAL>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1)) "
	sql_det = sql_det & " THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_NEGATIVA_CON_FONO, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE=0 AND PP2.TT_GEST_GENERAL>0 AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1)) "
	sql_det = sql_det & " THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_NEGATIVA_CON_MAIL, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE=0 AND PP2.TT_GEST_GENERAL>0 AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) "
	sql_det = sql_det & " THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_NEGATIVA_CON_DIRECCION, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GTIT>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) AS CASOS_GESTION_POSITIVA_TITULAR, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GTIT>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 "
	sql_det = sql_det & " OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN SALDO_RUT ELSE 0 END)),0) AS MONTO_GESTION_POSITIVA_TITULAR, "

	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND PP2.TT_GTIT=0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_GESTION_POSITIVA_TERCERO_CON_FONO, "
	sql_det = sql_det & " ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND PP2.TT_GTIT=0 AND (PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_GESTION_POSITIVA_TERCERO_CON_MAIL, "
	sql_det = sql_det & "  ISNULL(SUM((CASE WHEN (PP2.TT_GEFE>0 AND PP2.TT_GTIT=0 AND (PP2.DIR_VA=1 OR PP2.DIR_SA=1)) THEN 1 ELSE 0 END)),0) "
	sql_det = sql_det & " AS CASOS_GESTION_POSITIVA_TERCERO_CON_DIRECCION "

	sql_det = sql_det & " FROM "

	sql_det = sql_det & " (SELECT PP.RUT_DEUDOR, "
	sql_det = sql_det & " (SELECT SUM(SALDO) FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO WHERE C.RUT_DEUDOR=PP. RUT_DEUDOR AND C.COD_CLIENTE=PP.COD_CLIENTE AND ED.ACTIVO=1 "

	If Trim(intVerCobExt) <> "1" and Trim(intUsaCobInterna) = "1" Then
		sql_det = sql_det & " AND C.CUSTODIO IS not NULL "

	End If

	If Trim(intVerCobExt) = "1" and Trim(intUsaCobInterna) <> "1" Then
		sql_det = sql_det & " AND C.CUSTODIO IS NULL "

	End If

	sql_det = sql_det & " /* DEJAR PARAMETRICO) */) AS SALDO_RUT, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_TEL_VA >0 THEN 1 ELSE 0 END) AS TEL_VA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_TEL_SA >0 THEN 1 ELSE 0 END) AS TEL_SA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_TEL_NV >0 THEN 1 ELSE 0 END) AS TEL_NV, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_EMAIL_VA >0 THEN 1 ELSE 0 END) AS EMAIL_VA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_EMAIL_SA >0 THEN 1 ELSE 0 END) AS EMAIL_SA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_EMAIL_NV >0 THEN 1 ELSE 0 END) AS EMAIL_NV, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_DIR_VA >0 THEN 1 ELSE 0 END) AS DIR_VA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_DIR_SA >0 THEN 1 ELSE 0 END) AS DIR_SA, "
	sql_det = sql_det & " (CASE WHEN PP.TOTAL_DIR_NV >0 THEN 1 ELSE 0 END) AS DIR_NV, "
	sql_det = sql_det & " SUM(PP.GEST_GENERAL) AS TT_GEST_GENERAL, "
	sql_det = sql_det & " SUM(PP.GEST_TEL) AS TT_GEST_TEL, "
	sql_det = sql_det & " SUM(PP.GEST_MAIL) AS TT_GEST_MAIL, "
	sql_det = sql_det & " SUM(PP.GEST_DIR) AS TT_GDIR, "
	sql_det = sql_det & " SUM(PP.GEST_EFE) AS TT_GEFE, "
	sql_det = sql_det & " SUM(PP.GEST_TIT) AS TT_GTIT "

	sql_det = sql_det & " FROM  "
	sql_det = sql_det & " (SELECT D.RUT_DEUDOR,D.COD_CLIENTE, "
	sql_det = sql_det & " (CASE WHEN G.ID_GESTION IS NOT NULL THEN 1 ELSE 0 END) AS GEST_GENERAL, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GTEL),0) AS GEST_TEL, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GMAIL),0) AS GEST_MAIL, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GDIR),0) AS GEST_DIR, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GEFE),0) AS GEST_EFE, "
	sql_det = sql_det & " ISNULL((GTG.PRIORIDAD_GTIT),0) AS GEST_TIT, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_TEL_VA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_TEL_SA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_TEL_NV, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_EMAIL_VA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_EMAIL_SA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_EMAIL_NV, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_DIR_VA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_DIR_SA, "
	sql_det = sql_det & " (SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_DIR_NV "
	sql_det = sql_det & " FROM CUOTA C INNER JOIN DEUDOR D ON C.RUT_DEUDOR = D.RUT_DEUDOR AND C.COD_CLIENTE = D.COD_CLIENTE "
	sql_det = sql_det & " 			 INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO "
	sql_det = sql_det & " 			 LEFT JOIN GESTIONES_CUOTA GC ON C.ID_CUOTA = GC.ID_CUOTA "
	sql_det = sql_det & " 			 LEFT JOIN GESTIONES G ON GC.ID_GESTION = G.ID_GESTION "
	sql_det = sql_det & " 			 LEFT JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA AND "
	sql_det = sql_det & " 					  G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA AND "
	sql_det = sql_det & " 					  G.COD_GESTION = GTG.COD_GESTION AND "
	sql_det = sql_det & " 					  G.COD_CLIENTE = GTG.COD_CLIENTE "
						  
	sql_det = sql_det & " WHERE ED.ACTIVO=1  "
	sql_det = sql_det & " AND D.COD_CLIENTE = " & trim(strCodCliente)

	if trim(strCobranza)="INTERNA" THEN
		sql_det = sql_det & " AND D.CUSTODIO IS not NULL "

	elseif trim(strCobranza)="EXTERNA" THEN
		sql_det = sql_det & " AND D.CUSTODIO IS NULL "

	end if

	sql_det = sql_det & " GROUP BY D.RUT_DEUDOR,D.COD_CLIENTE,G.ID_GESTION,GTG.PRIORIDAD_GTEL,GTG.PRIORIDAD_GMAIL,GTG.PRIORIDAD_GDIR,GTG.PRIORIDAD_GEFE,GTG.PRIORIDAD_GTIT "
	sql_det = sql_det & " ) AS PP "	 	 
	sql_det = sql_det & " GROUP BY PP.COD_CLIENTE,PP.RUT_DEUDOR,PP. TOTAL_TEL_VA,TOTAL_TEL_SA,TOTAL_TEL_NV,TOTAL_EMAIL_VA, "
	sql_det = sql_det & " TOTAL_EMAIL_SA,TOTAL_EMAIL_NV,TOTAL_DIR_VA,TOTAL_DIR_SA,TOTAL_DIR_NV "
	sql_det = sql_det & " ) AS PP2	"
				 
	set rs_det = conn.execute(sql_det)			 

	'response.write sql_det

	if not rs_det.eof then 

		TOTAL_RUT      									=rs_det("TOTAL_RUT")
		SALDO_RUT										=rs_det("SALDO_RUT")

		RUT_GESTIONABLES								=rs_det("RUT_GESTIONABLES")
		MONTO_GESTIONABLE								=rs_det("MONTO_GESTIONABLE")

		RUT_GES_TEL 									=rs_det("RUT_GES_TEL")
		RUT_GES_MAIL 									=rs_det("RUT_GES_MAIL")
		RUT_GES_DIR 									=rs_det("RUT_GES_DIR")
		RUT_GES_TEL_MAIL 								=rs_det("RUT_GES_TEL_MAIL")
		RUT_GES_TEL_DIR 								=rs_det("RUT_GES_TEL_DIR")
		RUT_GES_MAIL_DIR 								=rs_det("RUT_GES_MAIL_DIR")
		RUT_GES_TEL_MAIL_DIR 							=rs_det("RUT_GES_TEL_MAIL_DIR")
		
		CASOS_GESTIONADOS 								=rs_det("CASOS_GESTIONADOS")
		MONTO_GESTIONADOS 								=rs_det("MONTO_GESTIONADOS")


		CASOS_PENDIENTES_CON_FONO 						=rs_det("CASOS_PENDIENTES_CON_FONO")
		CASOS_PENDIENTES_CON_MAIL 						=rs_det("CASOS_PENDIENTES_CON_MAIL")
		CASOS_PENDIENTES_CON_DIRECCION 					=rs_det("CASOS_PENDIENTES_CON_DIRECCION")
		
		CASOS_GESTION_POSITIVA 							=rs_det("CASOS_GESTION_POSITIVA")
		MONTO_GESTION_POSITIVA 							=rs_det("MONTO_GESTION_POSITIVA")

		CASOS_GESTION_NEGATIVA_CON_FONO 				=rs_det("CASOS_GESTION_NEGATIVA_CON_FONO")
		CASOS_GESTION_NEGATIVA_CON_MAIL 				=rs_det("CASOS_GESTION_NEGATIVA_CON_MAIL")
		CASOS_GESTION_NEGATIVA_CON_DIRECCION 			=rs_det("CASOS_GESTION_NEGATIVA_CON_DIRECCION")
		
		CASOS_GESTION_POSITIVA_TITULAR 					=rs_det("CASOS_GESTION_POSITIVA_TITULAR")
		MONTO_GESTION_POSITIVA_TITULAR 					=rs_det("MONTO_GESTION_POSITIVA_TITULAR")

		CASOS_GESTION_POSITIVA_TERCERO_CON_FONO 		=rs_det("CASOS_GESTION_POSITIVA_TERCERO_CON_FONO")
		CASOS_GESTION_POSITIVA_TERCERO_CON_MAIL 		=rs_det("CASOS_GESTION_POSITIVA_TERCERO_CON_MAIL")
		CASOS_GESTION_POSITIVA_TERCERO_CON_DIRECCION 	=rs_det("CASOS_GESTION_POSITIVA_TERCERO_CON_DIRECCION")

		CASOS_NO_GESTIONABLES							=CINT(TOTAL_RUT)-CINT(RUT_GESTIONABLES)
		MONTO_NO_GESTIONABLES 							=(SALDO_RUT)-(MONTO_GESTIONABLE)

		CASOS_NO_GESTIONADOS 							=CINT(RUT_GESTIONABLES)-CINT(CASOS_GESTIONADOS)
		MONTO_NO_GESTIONADOS 							=MONTO_GESTIONABLE-MONTO_GESTIONADOS

		CASOS_GESTION_NEGATIVA 							=CINT(CASOS_GESTIONADOS)-CINT(CASOS_GESTION_POSITIVA)
		MONTO_GESTION_NEGATIVA 							=MONTO_GESTIONADOS-MONTO_GESTION_POSITIVA

		CASOS_GESTION_POSITIVA_TERCERO 					=CINT(CASOS_GESTION_POSITIVA)-CINT(CASOS_GESTION_POSITIVA_TITULAR)
		MONTO_GESTION_POSITIVA_TERCERO 					=MONTO_GESTION_POSITIVA-MONTO_GESTION_POSITIVA_TITULAR

	Else

		TOTAL_RUT      									=""
		SALDO_RUT										=""
		RUT_GESTIONABLES								=""
		MONTO_GESTIONABLE								=""
		RUT_GES_TEL 									=""
		RUT_GES_MAIL 									=""
		RUT_GES_DIR 									=""
		RUT_GES_TEL_MAIL 								=""
		RUT_GES_TEL_DIR 								=""
		RUT_GES_MAIL_DIR 								=""
		RUT_GES_TEL_MAIL_DIR 							=""
		CASOS_GESTIONADOS 								=""
		MONTO_GESTIONADOS 								=""
		CASOS_PENDIENTES_CON_FONO 						=""
		CASOS_PENDIENTES_CON_MAIL 						=""
		CASOS_PENDIENTES_CON_DIRECCION 					=""
		CASOS_GESTION_POSITIVA 							=""
		MONTO_GESTION_POSITIVA 							=""
		CASOS_GESTION_NEGATIVA_CON_FONO 				=""
		CASOS_GESTION_NEGATIVA_CON_MAIL 				=""
		CASOS_GESTION_NEGATIVA_CON_DIRECCION 			=""
		CASOS_GESTION_POSITIVA_TITULAR 					=""
		MONTO_GESTION_POSITIVA_TITULAR 					=""
		CASOS_GESTION_POSITIVA_TERCERO_CON_FONO 		=""
		CASOS_GESTION_POSITIVA_TERCERO_CON_MAIL 		=""
		CASOS_GESTION_POSITIVA_TERCERO_CON_DIRECCION 	=""
		CASOS_NO_GESTIONADOS 							=""
		MONTO_NO_GESTIONADOS 							=""

	end if

%>
<body>
	<input type="hidden" name="mostrar_oculto" id="mostrar_oculto" value="S" >
<div class="body">
	
	<div class="titulo_informe"><img src="../Imagenes/icono informe.png" alt=""> Estatus de Cartera</div>
	<div class="subtitulo_informe">
		> Opciones de búsqueda
		
		<div id="opcion_busqueda">(Presione <span class="linkeado">aqui</span> para <span id="mostrar_ocultar_opcion">ocultar</span> opciones de busqueda) <img id="img_mostrar_ocultar_opcion" src="../Imagenes/simbolo_menos_ec.fw.png" alt=""></div>
		
	</div>
	<div class="subtitulo_informe_seleccion">Seleccione al menos un filtro para generar informe</div>

	<div class="filtro">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" >
			<tr>
				<td width="170">Cliente</td>
				<td width="330">			
					<select style="width:243px;" name="COD_CLIENTE" id="COD_CLIENTE">
					<%if not rs_cliente.eof then
						do while not rs_cliente.eof%>
						<option value="<%=rs_cliente("COD_CLIENTE")%>"<%if Trim(strCodCliente)=rs_cliente("COD_CLIENTE") then response.Write("Selected") End If%>><%=ucase(rs_cliente("NOMBRE_FANTASIA"))%></option>
						<%
						rs_cliente.movenext
						loop
					end if%>
					</select>
				</td>
				<td>Tipo Documento </td>
				<td id="id_tipo_doc">
					<select style="width:240px;" name="CB_TIPODOC" id="CB_TIPODOC">
						<option value="">TODOS</option>
						<%
						strSql="SELECT DISTINCT COD_TIPO_DOCUMENTO, NOM_TIPO_DOCUMENTO"
						strSql=strSql & " FROM CUOTA LEFT JOIN TIPO_DOCUMENTO ON TIPO_DOCUMENTO = COD_TIPO_DOCUMENTO"
						strSql=strSql & " WHERE CUOTA.COD_CLIENTE = '" & strCodCliente & "' AND COD_TIPO_DOCUMENTO is not null "
						strSql=strSql & " ORDER BY NOM_TIPO_DOCUMENTO ASC"
	
						set rsTemp= Conn.execute(strSql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("COD_TIPO_DOCUMENTO")%>"<%if Trim(intTipoDoc)=Trim(rsTemp("COD_TIPO_DOCUMENTO")) then response.Write("Selected") End If%>><%=ucase(rsTemp("NOM_TIPO_DOCUMENTO"))%></option>
							<%
							rsTemp.movenext
							loop
						end if
						%>
					</select>
				</td>


			</tr>



			<tr>
				<td>
					Tipo Cobranza
				</td>
				<td id="td_cobranza">
					<select style="width:243px;"  name="TIPO_COBRANZA" id="TIPO_COBRANZA" >
						<%If Trim(intVerCobExt) = "1" and Trim(intUsaCobInterna) = "1" Then%>
							<option value="0" <%If Trim(strCobranza) ="" Then Response.write "SELECTED"%>>TODOS</option>
						<%End If%>
						
						<%If Trim(intUsaCobInterna) = "1" Then%>
							<option value="INTERNA" <%If Trim(strCobranza) ="INTERNA" Then Response.write "SELECTED"%>>INTERNA</option>
						<%End If%>
						
						<%If Trim(intVerCobExt) = "1" Then%>
							<option value="EXTERNA" <%If Trim(strCobranza) ="EXTERNA" Then Response.write "SELECTED"%>>EXTERNA</option>
						<%End If%>
					</select>		

				</td>				

				<td>Rubro</td>
				<td  id="td_rubro">
					<select style="width:240px;" name="CB_RUBRO" ID="CB_RUBRO">
						<option value="" <%if Trim(strRubro)="" then response.Write("Selected") end if%>>SELECCIONE</option>
						<%

						ssql="SELECT DISTINCT ISNULL(ADIC_2,'OTRO') AS ADIC_2 FROM DEUDOR  WHERE COD_CLIENTE = '" & strCodCliente & "' ORDER BY ADIC_2"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("ADIC_2")%>"<%if strRubro=rsTemp("ADIC_2") then response.Write("Selected") End If%>><%=ucase(rsTemp("ADIC_2"))%></option>
							<%
							rsTemp.movenext
							loop
						end if
	
						%>
					</select>					

				</td>				


			</tr>
			<tr>
				<td>Fecha asignación</td>
				<td>
					Desde <input style="width:70px;" type="text" name="fecha_asignacion_desde" id="fecha_asignacion_desde" value="" placeholder="">
					Hasta <input style="width:78px;" type="text" name="fecha_asignacion_hasta" id="fecha_asignacion_hasta" value="" placeholder="">

				</td>
				<td>Campaña</td>
				<td>
					<select style="width:240px;" name="CB_CAMPANA" id="CB_CAMPANA">
					<option value="">TODAS</option>
					<%
						strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE = '" & strCodCliente & "'"
						set rsCampana=Conn.execute(strSql)
						Do While not rsCampana.eof
							If Trim(intCodCampana)=Trim(rsCampana("ID_CAMPANA")) Then strSelCam = "SELECTED" Else strSelCam = ""
							%>
							<option value="<%=rsCampana("ID_CAMPANA")%>" <%=strSelCam%>><%=ucase(rsCampana("ID_CAMPANA")) & " - " & ucase(rsCampana("NOMBRE"))%></option>
							<%
						rsCampana.movenext
						Loop
					''Response.End
					%>
					</select>
				</td>

			</tr>
			<tr>
				<td>Etapa cobranza</td>
				<td>
					<select style="width:243px;"  name="COD_ESTADO_COBRANZA" id="COD_ESTADO_COBRANZA" >
						<option value="">TODOS</option>
						<%
						if not rs_est_cob.eof then
							do until rs_est_cob.eof%>
							<option value="<%=rs_est_cob("COD_ESTADO_COBRANZA")%>"<%if Trim(intEtapaCobranza)=Trim(rs_est_cob("COD_ESTADO_COBRANZA")) then response.Write("Selected") End If%>><%=ucase(rs_est_cob("NOM_ESTADO_COBRANZA"))%></option>
							<%
							rs_est_cob.movenext
							loop
						end if
						%>
					</select>
				</td>
				<td>Ejecutivo</td>
				<td id="td_ejecutivo">			
					<select style="width:240px;"  name="ID_USUARIO" id="ID_USUARIO" >
						<option value="">TODOS</option>
						<%if not rsUsuario.eof then%>
							<%do while not rsUsuario.eof%>
							<option value="<%=trim(rsUsuario("ID_USUARIO"))%>"><%=ucase(trim(rsUsuario("LOGIN")))%></option>
							<%rsUsuario.movenext
							loop%>
						<%end if%>
					</select>


				</td>


			</tr>
			<tr>
				<td></td>
				<td></td>				
				<td>Fecha gestión</td>
				<td >
					Desde <input style="width:70px;" type="text" name="fecha_gestion_desde" id="fecha_gestion_desde" value="" placeholder="">
					Hasta <input style="width:76px;" type="text" name="fecha_gestion_hasta" id="fecha_gestion_hasta" value="" placeholder="">			
				</td>
			</tr>
			<tr>
				<td></td>
				<td></td>
				<td>
					
				</td>
				<td class="td_boton_limpiar">
					<br>
					<input type="button"  class="fondo_boton_100" id="bt_filtro" value="Generar informe">
					<input type="button" class="fondo_boton_100" id="bt_filtro_limpiar" value="Limpiar">
				</td>
			</tr>
		</table>

	</div>

	<div class="subtitulo_informe_2">> Informe</div>

	<div class="contenido" id="contenido">
		<div class="cargados">
			<div class="titulo_carga">CASOS CARGADOS</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&COD_ESTADO_COBRANZA=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=0&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=cint(TOTAL_RUT)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(SALDO_RUT,0)%></div>
		</div>
		<div class="titulo_carga_linea_carga"></div>
		<div class="gestionables">
			<div class="titulo_carga">
				GESTIONABLES 
				<img class="iconos_detalle_gestiones" id="id_gestionables" src="../Imagenes/48px-Crystal_Clear_app_kexi.png" alt="Detalle carga">
			</div> 			
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&COD_ESTADO_COBRANZA=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=1&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(RUT_GESTIONABLES)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTIONABLE,0)%></div>
		</div>

		<div class="contenido_detalle_ges">
			<div class="titulo_carga_detalle">INFORMACIÓN ADICIONAL GESTIONABLES</div><br>
			<span class="titulo_carga_det">Con Teléfono: </span><%=trim(RUT_GES_TEL)%><br>
			<span class="titulo_carga_det">Con Email: </span><%=trim(RUT_GES_MAIL)%><br>
			<span class="titulo_carga_det">Con Dirección: </span><%=trim(RUT_GES_DIR)%><br><br>
			<span class="titulo_carga_det">Con Teléfono-Email: </span><%=trim(RUT_GES_TEL_MAIL)%><br>
			<span class="titulo_carga_det">Con Teléfono-Dirección: </span><%=trim(RUT_GES_TEL_DIR)%><br>
			<span class="titulo_carga_det">Con Email-Dirección: </span><%=trim(RUT_GES_MAIL_DIR)%><br>
			<span class="titulo_carga_det">Con Teléfono-Email-Dirección: </span><%=trim(RUT_GES_TEL_MAIL_DIR)%><br>	
			<input type="text" name="sacar_foco" readonly id="sacar_foco">		
		</div>

		<div class="no_gestionables">
			<div class="titulo_carga_negativa">INUBICABLES</div>
			<div class="cuerpo_carga_negativa">
				Carga: <a style="color:#E6E14C;" href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=2&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_NO_GESTIONABLES)%></a>
			</div>
			<div class="cuerpo_monto_negativa">Monto: $<%=FormatNumber(MONTO_NO_GESTIONABLES,0)%></div>
		</div>	
		<div class="titulo_carga_linea">&nbsp;</div>
		<div class="gestionados">
			
			<div class="titulo_carga">GESTIONADOS</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=3&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_GESTIONADOS)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTIONADOS,0)%></div>	
		</div>	

		<div class="pendientes">
			<div class="titulo_carga">
				PENDIENTES
				<img class="iconos_detalle_gestiones" id="id_pendientes" src="../Imagenes/48px-Crystal_Clear_app_kexi.png" alt="Detalle carga">
			</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=4&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_NO_GESTIONADOS)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_NO_GESTIONADOS,0)%></div>
		</div>
		<div></div>
		<div class="titulo_carga_linea_pos">&nbsp;</div>
		<div class="gestion_positiva">
			<div class="titulo_carga">GESTIÓN POSITIVA</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=5&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_GESTION_POSITIVA)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTION_POSITIVA,0)%></div>	
		</div>

		<div class="gestion_negativa">
			<div class="titulo_carga">
				GESTIÓN NEGATIVA 
				<img class="iconos_detalle_gestiones" id="id_negativa" src="../Imagenes/48px-Crystal_Clear_app_kexi.png" alt="Detalle carga">
			</div>			
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=6&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_GESTION_NEGATIVA)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTION_NEGATIVA,0)%></div>
		</div>

		<div class="contenido_detalle_negativa">
			<div class="titulo_carga_detalle">INFORMACIÓN ADICIONAL GESTIONABLES</div><br>
			<span class="titulo_carga_det">Con Télefono: </span><%=trim(CASOS_GESTION_NEGATIVA_CON_FONO)%><br>
			<span class="titulo_carga_det">Con Email: </span><%=trim(CASOS_GESTION_NEGATIVA_CON_MAIL)%><br>
			<span class="titulo_carga_det">Con Dirección: </span><%=trim(CASOS_GESTION_NEGATIVA_CON_DIRECCION)%><br>			
			<input type="text" name="sacar_foco_neg" readonly id="sacar_foco_neg">		
		</div>
		<div class="titulo_carga_linea_titular">&nbsp;</div>
		<div class="titular">
			<div class="titulo_carga">TITULAR</div>
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=7&intVerCobExt=<%=trim(intVerCobExt)%>&intUsaCobInterna=<%=trim(intUsaCobInterna)%>"><%=CINT(CASOS_GESTION_POSITIVA_TITULAR)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTION_POSITIVA_TITULAR,0)%></div>
		</div>

		<div class="tercero">
			<div class="titulo_carga">
				TERCERO
				<img class="iconos_detalle_gestiones" id="id_tercero" src="../Imagenes/48px-Crystal_Clear_app_kexi.png" alt="Detalle carga">
			</div>
			
			<div class="cuerpo_carga">
				Carga: <a href="cartera_asignada.asp?COD_CLIENTE=<%=trim(COD_CLIENTE)%>&CB_EJECUTIVO=<%=trim(ID_USUARIO)%>&fecha_asignacion_desde=<%=trim(fecha_asignacion_desde)%>&fecha_asignacion_hasta=<%=trim(fecha_asignacion_hasta)%>&CB_TIPOCOB=<%=COD_ESTADO_COBRANZA%>&fecha_gestion_desde=<%=trim(fecha_gestion_desde)%>&fecha_gestion_hasta=<%=trim(fecha_gestion_hasta)%>&TIPO_COBRANZA=<%=trim(TIPO_COBRANZA)%>&CB_CAMPANA=<%=trim(CB_CAMPANA)%>&CB_RUBRO=<%=trim(CB_RUBRO)%>&CB_TIPODOC=<%=trim(CB_TIPODOC)%>&tipo_busqueda=8"><%=CINT(CASOS_GESTION_POSITIVA_TERCERO)%></a>
			</div>
			<div class="cuerpo_monto">Monto: $<%=FormatNumber(MONTO_GESTION_POSITIVA_TERCERO,0)%></div>	
		</div>
	
	</div>
</div>
</body>
</html>
<%

cerrarscg()
%>