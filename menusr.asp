<!--#include file="lib/asp/comunes/general/rutinasBooleano.inc"-->
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>Slashdot's Menu</title>
	<link rel="stylesheet" type="text/css" href="css/sdmenu.css" />
	<script type="text/javascript" src="javascripts/sdmenu.js">
		/***********************************************
		* Slashdot Menu script- By DimX
		* Submitted to Dynamic Drive DHTML code library: http://www.dynamicdrive.com
		* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
		***********************************************/
	</script>
	<script type="text/javascript">
	// <![CDATA[
	var myMenu;
	window.onload = function() {
		myMenu = new SDMenu("my_menu");
		myMenu.init();
	};
	// ]]>
	</script>
</head>
  <body>
    <div style="float: left" id="my_menu" class="sdmenu">


	<div>
		<span>Mod. Gestion</span>
		<a href="EstatusCartJud.asp" target='Contenido'>Estatus Cartera</a>
		<a href="principal.asp" target='Contenido'>Principal</a>
		<a href="cartera_asignada.asp" target='Contenido'>Cartera Asignada</a>
		<a href="scg_ingreso.asp?intNuevo=1&strRutDeudor=<%=rut%>" target='Contenido'>Nuevo Cliente - Deuda</a>
		<a href="busqueda.asp" target='Contenido'>Busqueda Deudor</a>
	</div>

	<div>
		<span>Mod.Informes</span>
		<a href="mis_gestiones.asp?codcob=<%=LTRIM(RTRIM(session("session_login")))%>" target='Contenido'>Gestiones</a>
		<a href="Informe_Gestiones_Jud.asp" target='Contenido'>Gestiones Por Dia</a>
		<a href="informe_metas.asp" target='Contenido'>Metas</a>
		<a href="informe_recupero.asp" target='Contenido'>Recuperaci&oacute;n</a>
	</div>

	<div>
		<span>Mod. Pagos</span>
		<a href="caja/caja_web.asp" target='Contenido'>Ingreso de Pagos</a>
		<a href="rendicion_caja.asp" target='Contenido'>Informe Rendiciones</a>
		<a href="rendicion_caja_inf2.asp" target='Contenido'>Nuevo Informe Rend.</a>
		<a href="caja/detalle_caja.asp" target='Contenido'>Listado de Pagos</a>
	</div>

	<div>
		<span>Mod. Administ.</span>
		<a href="MenuAdm.asp" target='Contenido'>Administracion</a>
		<a href="Asigna_masiva.asp" target='Contenido'>Asignacion Masiva</a>
		<a href="Asigna_manual.asp" target='Contenido'>Asignacion Individual</a>
		<a href="cbdd02.asp" target='Contenido'>Cerrar Sesion</a>
	</div>

	<div>
		<span>Mod.Convenios</span>
		<a href="simulacion_convenio.asp" target='Contenido'>Convenios</a>
		<a href="detalle_convenio.asp" target='Contenido'>Listado Convenios</a>
		<a href="convenios_vencidos.asp" target='Contenido'>Convenios Vencidos</a>
		<% If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_caja"))="Si" or TraeSiNo(session("perfil_sup"))="Si" or TraeSiNo(session("perfil_supterr"))="Si" Then %>
			<!--a href="caja/caja_web.asp?CB_TIPOPAGO=CO" target='Contenido'>Pago de Convenio</a-->
			<a href="caja/detalle_caja.asp?CB_TIPOPAGO=CO" target='Contenido'>Listado de Pagos</a>
		<% End If%>
	</div>

   </div>

 </body>
</html>
