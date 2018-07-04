<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">   
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	<link href="../css/style_multi_select.css" rel="stylesheet"> 
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
	
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib2.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/Minimo.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->
	
<%
	Response.CodePage 	=65001
	Response.charset 	="utf-8"
	intCodCliente=session("ses_codcli")
	intEtapaCobranza = Request("CB_ETAPA_COB")
	intIdFocoAT=Request("CB_FOCOAT")
	intCodCampana = Request("CB_CAMPANA")
	intHorarioAtencion = Request("CB_ATENCION")
	intEjeAsig = Request("CB_EJECUTIVO") 
	strTramoMonto = Request("CB_TRMONTO")
	strTramoVenc = Request("CB_TRVENC")
	strTramoAsig = Request("CB_TRASIG")
	strSucursal = Request("CB_SUCURSAL")
	strNombreDeudor = Request ("TX_NOMBRE_DEUDOR")
	strRutDeudor = Request ("TX_RUT_DEUDOR")
	intUltGest = Request ("CB_ULT_GEST")
	intUltMejorGest = Request ("CB_ULT_MEJOR_GEST")
	intTipoSF = Request ("CB_TIPO_SF")
	intContactabilidad = Request ("CB_CONTACTABILIDAD")
	dtmFecGestInicio = Request ("TX_G_INICIO")
	dtmFecGestTermino = Request ("TX_G_TERMINO")
	dtmFecAgendInicio = Request ("TX_AGEND_INICIO")
	dtmFecAgendTermino = Request ("TX_AGEND_TERMINO")
	
	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_ETAPACOB") = intEtapaCobranza
		session("Ftro_FOCOGF") = intIdFocoAT
		session("Ftro_Campana") = intCodCampana
		session("Ftro_CbAtencion") = intHorarioAtencion
		session("Ftro_Ejecutivo") = intEjeAsig
		session("Ftro_TramoMonto") = strTramoMonto
		session("Ftro_TramoVenc") = strTramoVenc
		session("Ftro_TramoAsig") = strTramoAsig
		session("Ftro_Sucursal") = strSucursal
		session("Ftro_UltGest") = intUltGest
		session("Ftro_UltMejorGest") = intUltMejorGest
		session("Ftro_TipoSF") = intTipoSF
		session("Ftro_Contactabilidad") = intContactabilidad
		session("Ftro_FecGestIni") = dtmFecGestInicio
		session("Ftro_FecGestTer") = dtmFecGestTermino
		session("Ftro_FecAgendIni") = dtmFecAgendInicio
		session("Ftro_FecAgendTer") = dtmFecAgendTermino
	End If
	
	If intEtapaCobranza = "" Then intEtapaCobranza = session("Ftro_ETAPACOB")
	If intIdFocoAT = "" Then intIdFocoAT = session("Ftro_FOCOGF")
	If intCodCampana = "" Then intCodCampana = session("Ftro_Campana")
	If intHorarioAtencion = "" Then intHorarioAtencion = session("Ftro_CbAtencion")
	If intEjeAsig = "" Then intEjeAsig = session("Ftro_Ejecutivo")
	If strTramoMonto = "" Then strTramoMonto = session("Ftro_TramoMonto")
	If strTramoVenc = "" Then strTramoVenc = session("Ftro_TramoVenc")
	If strTramoAsig = "" Then strTramoAsig = session("Ftro_TramoAsig")
	If strSucursal = "" Then strSucursal = session("Ftro_Sucursal")
	If intUltGest = "" Then intUltGest = session("Ftro_UltGest")
	If intUltMejorGest = "" Then intUltMejorGest = session("Ftro_UltMejorGest")
	If intTipoSF = "" Then intTipoSF = session("Ftro_TipoSF")
	If intContactabilidad = "" Then intContactabilidad = session("Ftro_Contactabilidad")
	If dtmFecGestInicio = "" Then dtmFecGestInicio = session("Ftro_FecGestIni")
	If dtmFecGestTermino = "" Then dtmFecGestTermino = session("Ftro_FecGestTer")
	If dtmFecAgendInicio = "" Then dtmFecAgendInicio = session("Ftro_FecAgendIni")
	If dtmFecAgendTermino = "" Then dtmFecAgendTermino = session("Ftro_FecAgendTer")
	
	If Trim(Request("strLimpiar")) = "S" Then
		session("Ftro_FOCOGF") = ""
		session("Ftro_Campana") = ""
		session("Ftro_CbAtencion") = ""
		session("Ftro_Ejecutivo") = ""
		session("Ftro_TramoMonto") = ""
		session("Ftro_TramoVenc") = ""
		session("Ftro_TramoAsig") = ""
		session("Ftro_Sucursal") = ""
		session("Ftro_UltGest") = ""
		session("Ftro_UltMejorGest") = ""
		session("Ftro_TipoSF") = ""
		session("Ftro_Contactabilidad") = ""
		session("Ftro_FecGestIni") = ""
		session("Ftro_FecGestTer") = ""
		session("Ftro_FecAgendIni") = ""
		session("Ftro_FecAgendTer") = ""
	End If

	If Trim(Request("strLimpiar")) = "S" Then
		intEtapaCobranza =""
		intIdFocoAT =""
		intCodCampana =""
		intHorarioAtencion =""
		intEjeAsig = ""
		strTramoMonto = ""
		strTramoVenc = ""
		strTramoAsig = ""
		strSucursal = ""
		strNombreDeudor = ""
		strRutDeudor = ""
		intUltGest = ""
		intUltMejorGest = ""
		intTipoSF = ""
		intContactabilidad = ""
		dtmFecGestInicio = ""
		dtmFecGestTermino = ""
		dtmFecAgendInicio = ""
		dtmFecAgendTermino = ""
	End If

	abrirscg()

			strSql = "SELECT ISNULL(TIPO_CARTERA_ASIGNADA,0) AS TIPO_COBRADOR"
			strSql = strSql & " FROM USUARIO U"
			strSql = strSql & " WHERE U.ID_USUARIO = " & session("session_idusuario")
		
			set RsCli=Conn.execute(strSql)
			If not RsCli.eof then
				intTipoCobrador = RsCli("TIPO_COBRADOR")
			End if
			RsCli.close
			set RsCli=nothing

	cerrarscg()
	
	If intEtapaCobranza = "" and intTipoCobrador = 2 then 
	
	intEtapaCobranza = 2 
	
	ElseIf intEtapaCobranza = "" and intTipoCobrador = 4 then 
	
	intEtapaCobranza = 4

	ElseIf intEtapaCobranza = ""  then 
	
	intEtapaCobranza = 10
	
	End If
	
	If intIdFocoAT = "" Then intIdFocoAT = 0
	If intCodCampana = "" Then intCodCampana = 0
	If intHorarioAtencion = "" Then intHorarioAtencion = 0
	
	If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then
		perfilEjecutivo="0"
	Else
		perfilEjecutivo="1"
		intEjeAsig=session("session_idusuario")
	End If
	
	'Response.write "strTramoVenc=" & strTramoVenc

%>

<title>MODULO GESTIÓN CAMPAÑAS</title>

<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
<link href="../css/style_generales_sistema.css" rel="stylesheet">

<script language='javascript' src="../javascripts/popcalendar.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 

<script type="text/javascript">
    $(document).ready(function(){
        
        $("#myTable").tablesorter(); 
		
        $("#Buscar1").click(function () {
            $("[name='divCNC']").toggle();
		});	
    })
</script>
<script language="JavaScript1.2">
$(document).ready(function(){
	$.prettyLoader();
	$('#TX_G_TERMINO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_G_INICIO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})	
	$('#TX_AGEND_TERMINO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_AGEND_INICIO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})	
})
    
function buscar(){
	$.prettyLoader.show(2000);
	datos.action='modulo_agendamiento_tactico.asp?strBuscar=S';
	datos.submit();
}
function limpiar(){
	$.prettyLoader.show(2000);
	datos.Limpiar.disabled = true;
	datos.action='modulo_agendamiento_tactico.asp?strLimpiar=S';
	datos.submit();
}
function ventanaAgenda (URL){
window.open(URL,"DETALLEDEUDA","width=1400, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ventana_procesa(rut,nombreDeudor){

//alert(nombreDeudor)

		var criterios ="alea="+Math.random()+"&accion_ajax=Carga_Fonos_AT&rut="+rut+"&nombreDeudor="+nombreDeudor

		$('#ventana_procesa').load('FuncionesAjax/deudor_telefonos_ajax.asp', criterios, function(data){

		})
		
	$('#ventana_procesa').dialog({
   		show:"blind", 
   		hide:"explode",   		       	 
    	width:900,
    	height:500 ,
    	modal:true,	
	    buttons: {
            Cerrar: function() {
				$(this).dialog( "close" );
            }
        }  	
	});	
	

}

</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">

<div id="ventana_procesa" title="Telefonos Deudor" style="display:none;"></div>
	
<div class="titulo_informe">MÓDULO AGENDAMIENTO TÁCTICO</div>
<br>
	<table width="90%" align="CENTER" border="0" bgcolor="#f6f6f6" class="estilo_columnas">
		<thead>
			<tr>
				<td align="left">ETAPA COBRANZA</td>
				<td align="left">FOCO</td>
				<td align="left">CAMPAÑA</td>
				<td align="left">TRAMO VENCIMIENTO</td>
				<td align="left">TRAMO MONTO</td>
				<td align="left" colspan=3>TRAMO ASIGNACIÓN</td>				
			</tr>
		</thead>
			<tr>			
				<td>
					<select name="CB_ETAPA_COB" id="CB_ETAPA_COB" onChange="buscar()">
						<option value=10>TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT COD_ESTADO_COBRANZA,NOM_ESTADO_COBRANZA FROM ESTADO_COBRANZA" 
							strSql= strSql & " WHERE COD_ESTADO_COBRANZA <> 0"

							If intTipoCobrador = 2 then
							
							strSql= strSql & " AND COD_ESTADO_COBRANZA = 2"
							
							End If
							
							If intTipoCobrador = 4 then
							
							strSql= strSql & " AND COD_ESTADO_COBRANZA IN (4,5)"
							
							End If
							
							strSql= strSql & " ORDER BY COD_ESTADO_COBRANZA ASC"
							
							set rsFocos=Conn.execute(strSql)
							Do While not rsFocos.eof
								If Trim(intEtapaCobranza)=Trim(rsFocos("COD_ESTADO_COBRANZA")) Then strSelFoco = "SELECTED" Else strSelFoco = ""
								%>
								<option value="<%=rsFocos("COD_ESTADO_COBRANZA")%>" <%=strSelFoco%>> <%=rsFocos("NOM_ESTADO_COBRANZA")%></option>
								<%
								rsFocos.movenext
							Loop
							rsFocos.close
							set rsFocos=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>
				<td>
					<select name="CB_FOCOAT" id="CB_FOCOAT" onChange="buscar()">
						<option value=0>TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_FOCO,NOMBRE_FOCO FROM FOCOS WHERE TIPO_FOCO=1 ORDER BY ID_FOCO ASC"
							set rsFocos=Conn.execute(strSql)
							Do While not rsFocos.eof
								If Trim(intIdFocoAT)=Trim(rsFocos("ID_FOCO")) Then strSelFoco = "SELECTED" Else strSelFoco = ""
								%>
								<option value="<%=rsFocos("ID_FOCO")%>" <%=strSelFoco%>> <%=rsFocos("NOMBRE_FOCO")%></option>
								<%
								rsFocos.movenext
							Loop
							rsFocos.close
							set rsFocos=nothing
						CerrarSCG()
						''Response.End
						%>
						<option value=100 <%if Trim(intIdFocoAT)=100 then response.Write("Selected") end if%>>SIN FOCO</option>
					</select>
				</td>
				<td>
					<select name="CB_CAMPANA" id="CB_CAMPANA" onChange="buscar()">
						<option value=0>TODAS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_CAMPANA,NOMBRE FROM CAMPANA WHERE COD_CLIENTE IN ('" & intCodCliente & "')"
							set rsCampana=Conn.execute(strSql)
							Do While not rsCampana.eof
								If Trim(intCodCampana)=Trim(rsCampana("ID_CAMPANA")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsCampana("ID_CAMPANA")%>" <%=strSelCam%>> <%=rsCampana("NOMBRE")%></option>
								<%
								rsCampana.movenext
							Loop
							rsCampana.close
							set rsCampana=nothing
						CerrarSCG()
						''Response.End
						%>
						<option value=1 <%if Trim(intCodCampana)=1 then response.Write("Selected") end if%>>SIN CAMPAÑA</option>
					</select>
				</td>
				<td>
					<select name="CB_TRVENC" id="CB_TRVENC">
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_SVENC = ID,NOMBRE_SEGMENTO=LTRIM(NOMBRE) FROM SEGMENTACION_VENCIMIENTO WHERE COD_CLIENTE = '" & intCodCliente & "' AND GESTIONABLE=1 AND ESTADO=1 ORDER BY ORDEN ASC"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(strTramoVenc)=Trim(rsSel("ID_SVENC")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("ID_SVENC")%>" <%=strSelCam%>> <%=rsSel("NOMBRE_SEGMENTO")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>				
				<td>
					<select name="CB_TRMONTO" id="CB_TRMONTO" >
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_SMONTO = ID,NOMBRE_SEGMENTO=LTRIM(NOMBRE) FROM SEGMENTACION_MONTO WHERE COD_CLIENTE = ('" & intCodCliente & "') AND GESTIONABLE=1 AND ESTADO=1 ORDER BY ORDEN ASC"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(strTramoMonto)=Trim(rsSel("ID_SMONTO")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("ID_SMONTO")%>" <%=strSelCam%>> <%=rsSel("NOMBRE_SEGMENTO")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>
				<td>
					<select name="CB_TRASIG" id="CB_TRASIG" >
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_SASIG = ID,NOMBRE_SEGMENTO=LTRIM(NOMBRE) FROM SEGMENTACION_ASIGNACION WHERE ESTADO=1 ORDER BY ORDEN ASC"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(strTramoAsig)=Trim(rsSel("ID_SASIG")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("ID_SASIG")%>" <%=strSelCam%>> <%=rsSel("NOMBRE_SEGMENTO")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>				
				<td align="center"><input name="Limpiar" class="fondo_boton_100" type="button" value="Limpiar"  onClick="limpiar();"></td>
				<td align="center"><input class="fondo_boton_100" name="Buscar" type="button" value="Buscar"  onClick="buscar();"></td>				
			</tr>
		<thead>
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td>ULTIMA GESTIÓN CALL</td>
			<td>ULTIMA MEJOR GESTIÓN</td>
			<td>TIPO SUB FOCO</td>
			<td>CONTACTABILIDAD</td>
			<td>ATENCIÓN</td>
			<% If perfilEjecutivo = "0" Then %>
				<td>EJECUTIVO</td>
			<% Else%>
				<td>&nbsp;</td>  
			<% End If %>
			<td Colspan=3>&nbsp;</td>			
		</tr>
		</thead>		
		<tr bgcolor="#f6f6f6" class="Estilo8">
			<td>
				<select name="CB_ULT_GEST">
					<option value="0" <%if Trim(intUltGest)="0" then response.Write("Selected") end if%>>TODOS</option>
					<option value="1" <%if Trim(intUltGest)="1" then response.Write("Selected") end if%>>COMPROMISO PAGO</option>
					<option value="2" <%if Trim(intUltGest)="2" then response.Write("Selected") end if%>>NEGOCIANDO</option>
					<option value="3" <%if Trim(intUltGest)="3" then response.Write("Selected") end if%>>TITULAR</option>
					<option value="4" <%if Trim(intUltGest)="4" then response.Write("Selected") end if%>>TERCERO</option>
					<option value="5" <%if Trim(intUltGest)="5" then response.Write("Selected") end if%>>CORTE DE LLAMADA SC</option>
					<option value="6" <%if Trim(intUltGest)="6" then response.Write("Selected") end if%>>SIN GESTION</option>
				</select>
			</td>
			<td>
				<select name="CB_ULT_MEJOR_GEST">
					<option value="0" <%if Trim(intUltMejorGest)="0" then response.Write("Selected") end if%>>TODOS</option>
					<option value="1" <%if Trim(intUltMejorGest)="1" then response.Write("Selected") end if%>>COMPROMISO PAGO</option>
					<option value="2" <%if Trim(intUltMejorGest)="2" then response.Write("Selected") end if%>>NEGOCIANDO</option>
					<option value="3" <%if Trim(intUltMejorGest)="3" then response.Write("Selected") end if%>>TITULAR</option>
					<option value="4" <%if Trim(intUltMejorGest)="4" then response.Write("Selected") end if%>>TERCERO</option>
				</select>
			</td>
			<td>
				<select name="CB_TIPO_SF">
					<option value="0" <%if Trim(intTipoSF)="0" then response.Write("Selected") end if%>>TODOS</option>
					<option value="8" <%if Trim(intTipoSF)="8" then response.Write("Selected") end if%>>COMPROMISO & NEGOCIANDO</option>
					<option value="1" <%if Trim(intTipoSF)="1" then response.Write("Selected") end if%>>COMPROMISO PAGO</option>
					<option value="2" <%if Trim(intTipoSF)="2" then response.Write("Selected") end if%>>NEGOCIANDO</option>
					<option value="3" <%if Trim(intTipoSF)="3" then response.Write("Selected") end if%>>TITULAR</option>
					<option value="4" <%if Trim(intTipoSF)="4" then response.Write("Selected") end if%>>CPEF - EFECTIVA</option>
					<option value="5" <%if Trim(intTipoSF)="5" then response.Write("Selected") end if%>>SCL - TITULAR</option>
					<option value="6" <%if Trim(intTipoSF)="6" then response.Write("Selected") end if%>>SCL - TERCERO</option>
					<option value="7" <%if Trim(intTipoSF)="7" then response.Write("Selected") end if%>>SCL</option>
				</select>
			</td>
			<td>
				<select name="CB_CONTACTABILIDAD">
					<option value="0" <%if Trim(intContactabilidad)="0" then response.Write("Selected") end if%>>TODOS</option>
					<option value="1" <%if Trim(intContactabilidad)="1" then response.Write("Selected") end if%>>CONTACTO TITULAR</option>
					<option value="2" <%if Trim(intContactabilidad)="2" then response.Write("Selected") end if%>>CONTACTO TERCERO</option>
					<option value="3" <%if Trim(intContactabilidad)="3" then response.Write("Selected") end if%>>SIN CONTACTO</option>
				</select>
			</td>

			<td>
				<select name="CB_ATENCION" id="CB_ATENCION">
					<option value="">TODOS</option>
					<%
					AbrirSCG()
						strSql= " SELECT ID,NOMBRE"
						strSql= strSql & " FROM HORARIO_ATENCION_DEUDOR"
						
						set rsAtencion=Conn.execute(strSql)
						Do While not rsAtencion.eof
							If Trim(intHorarioAtencion)=Trim(rsAtencion("ID")) Then strHorarioAtencion = "SELECTED" Else strHorarioAtencion = ""
							%>
							<option value="<%=rsAtencion("ID")%>" <%=strHorarioAtencion%>> <%=rsAtencion("NOMBRE")%></option>
							<%
							rsAtencion.movenext
						Loop
						rsAtencion.close
						set rsAtencion=nothing
					CerrarSCG()
					
					%>
				</select>
			</td>
			
			<% If perfilEjecutivo="0" Then %>
			<td colspan=1>
				<select name="CB_EJECUTIVO" id="CB_EJECUTIVO" onChange="buscar()">
					<option value="">TODOS</option>
					<%
					AbrirSCG()
						strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
						strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE IN ('" & intCodCliente & "')"

						strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
						strSql = strSql & " AND U.PERFIL_EMP=0"
						
						set rsEjecutivo=Conn.execute(strSql)
						Do While not rsEjecutivo.eof
							If Trim(intEjeAsig)=Trim(rsEjecutivo("ID_USUARIO")) Then strSelEjecutivo = "SELECTED" Else strSelEjecutivo = ""
							%>
							<option value="<%=rsEjecutivo("ID_USUARIO")%>" <%=strSelEjecutivo%>> <%=rsEjecutivo("LOGIN")%></option>
							<%
							rsEjecutivo.movenext
						Loop
						rsEjecutivo.close
						set rsEjecutivo=nothing
					CerrarSCG()
					
					%>
				</select>
			</td>
			<% Else%>
				<td colspan=1>&nbsp;</td>  
			<% End If %>
			
			<%If dtmFecGestInicio="" and dtmFecGestTermino="" and dtmFecAgendInicio="" and dtmFecAgendTermino="" and strNombreDeudor="" and strRutDeudor="" then%>
				<td align="center">
				<img src="../imagenes/Filtro_Agenda_Desactivado.png" id="Buscar1" style="cursor:pointer;" border="0">
				</td>								
				
			<% Else%>
				<td align="center">
				<img src="../imagenes/Filtro_Agenda_Activado.png" id="Buscar1" style="cursor:pointer;" border="0">
				</td>			
			<% End If %>
			
			<td align="center"><a href="javascript:ventanaAgenda('agenda.asp?')">AGENDA</a></td>

			
		</tr>
			<thead name="divCNC" style="display: none">
			<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td>SEDE</td>
				<td>GESTIÓN DESDE</td>
				<td>GESTIÓN HASTA</td>
				<td>AGENDAMIENTO DESDE</td>
				<td>AGENDAMIENTO HASTA</td>
				<td>NOMBRE O RAZON SOCIAL</td>
				<td Colspan=3>RUT</td>
			</tr>
			</thead>		
			<tr name="divCNC" bgcolor="#f6f6f6" class="Estilo8" style="display: none">		
				<td>
					<select name="CB_SUCURSAL" id="CB_SUCURSAL" >
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT DISTINCT C.SUCURSAL FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"
							strSql= strSql & " WHERE C.COD_CLIENTE IN ('" & intCodCliente & "') AND ED.ACTIVO=1"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(strSucursal)=Trim(rsSel("SUCURSAL")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("SUCURSAL")%>" <%=strSelCam%>> <%=rsSel("SUCURSAL")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>			
				<td>
					<input name="TX_G_INICIO" readonly="true" type="text" id="TX_G_INICIO" value="<%=dtmFecGestInicio%>" size="10">
				</td>
				<td>
					<input name="TX_G_TERMINO" readonly="true" type="text" id="TX_G_TERMINO" value="<%=dtmFecGestTermino%>" size="10">
				</td>
				<td>
					<input name="TX_AGEND_INICIO" readonly="true" type="text" id="TX_AGEND_INICIO" value="<%=dtmFecAgendInicio%>" size="10">
				</td>
				<td>
					<input name="TX_AGEND_TERMINO" readonly="true" type="text" id="TX_AGEND_TERMINO" value="<%=f%>" size="10">
				</td>	
				<td>
					<input name="TX_NOMBRE_DEUDOR" type="text" value="<%=strNombreDeudor%>" size="27" maxlength="77">
				</td>
				<td>
					<input name="TX_RUT_DEUDOR" type="text" value="<%=strRutDeudor%>" size="15" maxlength="15">
				</td>
			
			</tr>
	</table>	
	<br/>			
	<table id="myTable" class="tablesorter intercalado" width="90%"  align="CENTER" border="0" bgcolor="#f6f6f6" bordercolor="#000000">
		<thead>
			<tr class="Estilo34">
				<th>&nbsp;</th>
				<th>&nbsp;</th>
				<th id="rut" align="left">RUT</th>
				<th>NOMBRE O RAZON SOCIAL </th>
				<th id="SALDO" align="left">SALDO</th>
				<th align="left">DOC.</th>
				<th align="center">TRAMO VENC.</th>
				<th align="center">DIA MORA</th>
				<th align="center">FECHA VENC.</th>
				<th align="center">TRAMO MONTO.</th>
				<th align="center">F.AGEND.</th>
				<th align="center">H.AGEND. </th>
				<th align="center">ATENCIÓN</th>
				<th align="center">ORDEN</th>
				<th align="center">PR</th>
				<th align="center" width="18">&nbsp;</th>
				<th align="center">EJECUTIVO</th>
			</tr>
		</thead>
		<tbody>		
		<%
			AbrirSCG1()	
			ssql="EXEC uspGestionAgendamientotacticoSelect " & intEtapaCobranza &","& intTipoCobrador &",'"&TRIM(intCodCliente)&"'," & intIdFocoAT &","& intCodCampana &","&intHorarioAtencion&",'"&intEjeAsig&"','"&strTramoMonto&"','"&strTramoVenc&"','"&strTramoAsig&"','"&strSucursal&"','"&strNombreDeudor&"','"&strRutDeudor&"','"&intUltGest&"','"&intUltMejorGest&"','"&intTipoSF&"','"&intContactabilidad&"','"&dtmFecGestInicio&"','"&dtmFecGestTermino&"','"&dtmFecAgendInicio&"','"&dtmFecAgendTermino&"'," & perfilEjecutivo
			
			'Response.write "ssql=" & ssql
			
			intNumReg=0

			set rsCam=Conn1.execute(ssql)
			if not rsCam.eof then
			
			Do While Not rsCam.Eof

				intNumReg= intNumReg + 1
				strRutDeudor = rsCam("RUT_DEUDOR")
				strNombreDeudor = rsCam("NOMBRE_DEUDOR")
				intSaldo = rsCam("SALDO")
				intTotalDoc = rsCam("TOTAL_DOC")
				intDiaMora = rsCam("DM")
				dtmFechaVenc = rsCam("VENC_INFERIOR")
				dtmFecAgend = rsCam("FEC_AGEND")
				strHoraAgend = rsCam("HORA_AGEND")	
				intOrdenCampana = rsCam("ORDEN_CAMPANA")
				strUsuarioAsig = rsCam("USUARIO_ASIG")	
				intUltGtt = rsCam("ULT_GEST_TT")
				intUltGefe = rsCam("ULT_GEST_EFE")
				strTramoVenc = rsCam("TRAMO_VENC")
				strTramoMonto = rsCam("TRAMO_MONTO")
				intPriridadCuota = rsCam("PRIORIDAD_CUOTA")
				intEstadoPriorizacion = rsCam("ESTADO_PRIORIZACION")
				strHorarioAtencion = rsCam("HA_DEUDOR")
				strTextoPrioF = "Caso priorizado. Ir pantalla de ingreso de gestión"
				

				if strHoraAgend = "00:00" or strHoraAgend = "08:59" then strHoraAgend = "" 
				
				if intUltGtt = 1 Then
					strContactado = "mod_telefono_va.png"
				Elseif intUltGefe = 1 then
					strContactado = "mod_telefono_sa.png"
				Else
					strContactado = "mod_telefono_nv.png"
				End If
				
		%>
			<tr>
				<td><%=FN(intNumReg,0)%></td>
				<td ALIGN="center"><img src="../imagenes/<%=strContactado%>" border="0" onClick="ventana_procesa('<%=strRutDeudor%>','<%=strNombreDeudor%>');"></td>
				<td><%=strRutDeudor%></td>
				<td><%=Mid(strNombreDeudor,1,30)%></td>
				<td><%=FN(intSaldo,0)%></td>				
				<td><%=FN(intTotalDoc,0)%></td>
				<td><%=strTramoVenc%></td>
				<td ALIGN="center"><%=intDiaMora%></td>
				<td ALIGN="center"><%=dtmFechaVenc%></td>
				<td><%=strTramoMonto%></td>
				<td ALIGN="center"><%=dtmFecAgend%></td>
				<td ALIGN="center"><%=strHoraAgend%></td>
				<td ALIGN="center"><%=strHorarioAtencion%></td>				
				<td ALIGN="center"><%=intOrdenCampana%></td>
				<td ALIGN="center"><%=intPriridadCuota%></td>			
				<td ALIGN="center">
					<% If intEstadoPriorizacion = "0" then %>
						<a href="detalle_gestiones.asp?rut=<%=strRutDeudor%>&cliente=<%=intCodCliente%>" onclick="javascript:SetCustomer('<%=intCodCliente%>');">
								<abbr title="<%=strTextoPrioF%>">
							<img src="../imagenes/priorizar_urgente.png" border="0">
							<abbr>						
							
							</acronym>
						</a>
					<% End If %>
				</td>
				<td ALIGN="center"><%=strUsuarioAsig%></td>				
			</tr>				
		<%	
			rsCam.movenext
			Loop
			end if
			rsCam.close
			set rsCam=nothing
			CerrarSCG1()
		
		If intNumReg=0 then				
			
			%>
			<tr bgcolor="<%=strbgcolor%>" class="Estilo8">																					
				<td colspan="17" align = "center"><h3>No Existen Casos Pendientes a Gestionar</h3></td>	
			</tr>
<%		end if%>
			
		</tbody>
			<tr class="totales">
				<td Colspan = "17" >&nbsp;</td>
			</tr>
	</table>		
</form>
</body>
</html>
