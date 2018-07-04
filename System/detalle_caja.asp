<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<!--#include file="../lib/lib.asp"-->

	<!--#include file="../lib/comunes/rutinas/chkFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/sondigitos.inc"-->
	<!--#include file="../lib/comunes/rutinas/formatoFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/validarFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	'cod_caja=110
	Response.CodePage=65001
	Response.charset ="utf-8"

	cod_caja=Session("intCodUsuario")
	AbrirSCG()

	sucursal = request("cmb_sucursal")
	intTipoPago = request("CB_TIPOPAGO")

	'response.write(perfil)
	intCodPago = request("TX_PAGO")
	strRut=request("TX_RUT")
	if sucursal="" then sucursal="0"
	'response.write(sucursal)
	usuario = request("cmb_usuario")
	if usuario = "" then usuario = "0"
	termino = request("termino")
	inicio = request("inicio")
	resp = request("resp")
	if Trim(inicio) = "" Then
		inicio = date()
	End If
	if Trim(termino) = "" Then
		termino = date()
	End If
	CLIENTE = REQUEST("CLIENTE")
	'hoy=date


	if CLIENTE ="" then
		CLIENTE = session("ses_codcli")
	end if
	resp ="si"
	'response.write(hoy)
%>
	<title>CRM Cobros</title>
	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo27 {color: #FFFFFF}
	-->
	</style>

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>
	<script src="../javascripts/SelCombox.js"></script>
	<script src="../javascripts/OpenWindow.js"></script>

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

	<script language="JavaScript " type="text/JavaScript">
	$(document).ready(function(){

		$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('.actualiza_boleta').hover(function(){
			$(this).css('background-color','#81DAF5')
		}, function(){
			$(this).css('background-color','')
		})


	 
	})
	function Refrescar()
	{
		resp='no'
		datos.action = "detalle_caja.asp?resp="+ resp +"";
		datos.submit();
	}

	function Ingresa()
	{
		with( document.datos )
		{
			action = "detalle_caja.asp";
			submit();
		}
	}

	function Reversar(cod_pago)
	{
		with( document.datos )
		{
			//alert("Opción deshabilitada");


		if (confirm("¿ Está seguro de reversar el pago ? El pago se eliminará completamente y la deuda será reversada, volviendo a su estado original antes del pago."))
			{
				action = "reversar_pago.asp?cod_pago=" + cod_pago;
				submit();
			}
		else
			alert("Reverso del pago cancelado");
		}
	}

	function ImpBoleta(intCompPago)
	{
		window.open("imprime_boleta.asp?intNroComp=" + intCompPago,"INFORMACION","width=800, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes");
	}

	function envia()
	{
		//datos.TX_RUT.value='';
		//datos.TX_PAGO.value='';
		resp='si'
		document.datos.action = "detalle_caja.asp?resp="+ resp +"";
		document.datos.submit();
	}

	function imprimir()
	{
		datos.action = "imprime_comprobantes.asp";
		datos.submit();
	}


	function envia_excel(URL){

	window.open(URL,"INFORMACION","width=200, height=200, scrollbars=yes, menubar=yes, location=yes, resizable=yes")
	}


	function VerCompPago(URL){
		window.open(URL,"INFORMACION","width=800, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes");
	}

	$(document).ready(function(){	

		$(document).tooltip();
	})

	function refresca_nro_boleta(id_pago, NRO_BOLETA){
		var concat_td ="#td_actualizar_nro_boleta_"+id_pago
		var criterios ="alea="+Math.random()+"&accion_ajax=refresca_campo_nro_boleta&id_pago="+id_pago+"&NRO_BOLETA="+NRO_BOLETA

		$(concat_td).load('FuncionesAjax/detalle_caja_ajax.asp', criterios, function(){
			$('#NRO_BOLETA').focus()
		})


	}

	function actuliza_nro_boleta(id_pago, valor, valor_original){
		if(valor!=""){

			if(valor==0){
				alert("N° boleta debe ser distinto de 0")
				$('#NRO_BOLETA').val("")
				$('#NRO_BOLETA').focus()
				return
			}

			if(isNaN(valor)){
				alert("N° boleta invalido")
				$('#NRO_BOLETA').val("")
				$('#NRO_BOLETA').focus()
				return	
			}

		}		

		if(valor!=valor_original){
			if(confirm("¿Esta seguro de modificar numero boleta?")){
				var concat_td ="#td_actualizar_nro_boleta_"+id_pago
				var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_campo_nro_boleta&id_pago="+id_pago+"&NRO_BOLETA="+valor

				$(concat_td).load('FuncionesAjax/detalle_caja_ajax.asp', criterios, function(){})

			}else{

				var concat_td ="#td_actualizar_nro_boleta_"+id_pago
				var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_campo_nro_boleta&id_pago="+id_pago+"&NRO_BOLETA="+valor_original

				$(concat_td).load('FuncionesAjax/detalle_caja_ajax.asp', criterios, function(){})	

			}			
		}else{
			var concat_td ="#td_actualizar_nro_boleta_"+id_pago
			var criterios ="alea="+Math.random()+"&accion_ajax=actualiza_campo_nro_boleta&id_pago="+id_pago+"&NRO_BOLETA="+valor

			$(concat_td).load('FuncionesAjax/detalle_caja_ajax.asp', criterios, function(){})			
		}

	}

	function modificar_cheque(id_pago){
		var contat_formulario	="#formulario_cheque_"+id_pago

		var criterios ="alea="+Math.random()+"&accion_ajax=modifica_numero_cheque&id_pago="+id_pago
		$('#modificar_cheque').load('FuncionesAjax/detalle_caja_ajax.asp', criterios, function(){
			$('#modificar_cheque').dialog({
		   		show:"blind", 
		   		hide:"explode",   		       	 
		    	width:800,
		    	height:300 ,
		    	modal:true				    	
			});	

		})
	}

	function modificar_nro_cheque(id_pago){
		//alert(id_pago)
		var valor_nro_cheque 	=""
		var contador 			=0
		var contat_td 			=""
		var	contat_formulario 	=""
		var concat_nro_cheque 	=""
		var concat_cheque_corre	="cheque_correlativo_"+id_pago
		var NRO_CHEQUE 			=""
		var concat_imagen 		=""
		var valor_vacio 		="NO"

		$('input[id="'+concat_cheque_corre+'"]').each(function(){
			contador 			= contador +1
			concat_nro_cheque 	="#NRO_CHEQUE_"+$(this).val()+"_"+id_pago 
			
			NRO_CHEQUE 			=$(concat_nro_cheque).val()
			valor_nro_cheque 	=valor_nro_cheque+"*"+NRO_CHEQUE
			if(NRO_CHEQUE==""){
				valor_vacio ="SI"
			}

		})
		concat_imagen 		="#imagen_"+id_pago
		contat_td  			="#guarda_numero_cheque_"+id_pago
		contat_formulario	="#formulario_cheque_"+id_pago
		valor_nro_cheque 	 	=valor_nro_cheque.substring(1, valor_nro_cheque.length)

		if(confirm("¿Esta seguro modificar número cheque?")){

			var criterios ="alea="+Math.random()+"&accion_ajax=guarda_numero_cheque&id_pago="+id_pago+"&valor_nro_cheque="+valor_nro_cheque+"&contador="+contador
			$('#modificar_cheque').load('FuncionesAjax/detalle_caja_ajax.asp', criterios, function(){
				if(valor_vacio=="SI"){
					$(concat_imagen).attr('src','../Imagenes/48px-Crystal_Clear_mimetype_document_rojo.png')
				}else{
					$(concat_imagen).attr('src','../Imagenes/48px-Crystal_Clear_mimetype_document2.png')
				}
				$('#modificar_cheque').dialog('close')				
			})
			
		}
	}

	</script>


</head>
<body>
<form name="datos" method="post">
<div class="titulo_informe">LISTADO DE PAGOS</div>	
<table width="90%" height="500" border="0" align="center">
  <tr>
    <td valign="top">
	<table width="100%" border="0" bordercolor="#999999" align="center">
		<tr height="20" class="Estilo8">
	        <td>RUT: </td>
			<td><INPUT TYPE="TEXT" NAME="TX_RUT" value="<%=strRut%>" onchange=""></td>
		    <td>CODIGO PAGO: </td>
			<td><INPUT TYPE="TEXT" NAME="TX_PAGO" value="<%=intCodPago%>" onchange=""></td>
			<td align="right"><INPUT TYPE="BUTTON" NAME="Imprimir" class="fondo_boton_100" VALUE="Imprimir" onClick="imprimir();"></td>
		</tr>

	</table>
	<div id="modificar_cheque" style="display:none;" title="Modificar número cheque"></div>
	<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
		<thead>
	      <tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
		  <td>CLIENTE</td>
		  <%if perfil="caja_modif" or perfil = "caja_listado" then%>
	        <td>SUCURSAL</td>
			<%end if%>
			<td>TIPO DE PAGO</td>
			<td>USUARIO</td>
			<td>DESDE</td>
			<td>HASTA</td>
			<td></td>
	      </tr>
	     </thead>
		  <tr bordercolor="#999999" class="Estilo8">
		  <td>

		<select name="CLIENTE" ID = "CLIENTE" width="15" onchange="">
		<% If TraeSiNo(session("perfil_emp")) <> "Si" Then %>
		<option value="0">SELECCIONAR</option>
		<% End If%>
		<%
		ssql="SELECT * FROM CLIENTE WHERE ACTIVO=1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ")"
		If TraeSiNo(session("perfil_emp")) = "Si" Then
			ssql = ssql & "AND COD_CLIENTE = " & session("ses_codcli")
		End If

		set rsCLI=Conn.execute(ssql)
		if not rsCLI.eof then
			do until rsCLI.eof
			%>
			<option value="<%=rsCLI("COD_CLIENTE")%>"
			<%if Trim(cliente)=Trim(rsCLI("COD_CLIENTE")) then
				response.Write("Selected")
			end if%>
			><%=ucase(rsCLI("DESCRIPCION"))%></option>

			<%rsCLI.movenext
			loop
		end if
		rsCLI.close
		set rsCLI=nothing
		%>
        </select>
        </td>

        <td>
			<select name="CB_TIPOPAGO">
				<%
				ssql="SELECT * FROM CAJA_TIPO_PAGO	"
				If Trim(intTipoPago)="RP" Then
					strTipoCompArch = "comp_pago_repactacion.asp"
					ssql = ssql & " WHERE ID_TIPO_PAGO = 'RP'"
				Else
					strTipoCompArch = "comp_pago.asp"
					ssql = ssql & " WHERE  ID_TIPO_PAGO IN( 'AB','PTC','PTE','CO')"
				%>
					<option value="">TODOS</option>
				<%
				End If
				set rsCLI=Conn.execute(ssql)
				if not rsCLI.eof then
					do until rsCLI.eof
					%>
					<option value="<%=rsCLI("ID_TIPO_PAGO")%>"
					<%if Trim(intTipoPago)=Trim(rsCLI("ID_TIPO_PAGO")) then Response.Write("SELECTED") end if%> WIDTH="10"
					><%=MID(rsCLI("DESC_TIPO_PAGO"),1,19)%></option>
					<%rsCLI.movenext
					loop
				end if
				rsCLI.close
				set rsCLI=nothing
				%>
			</select>
		</td>
			<td>
				<SELECT NAME="cmb_usuario" id="cmb_usuario" onchange="Refrescar();">
					<option value="0">TODOS</option>
					<%
					If Trim(intTipoPago)="CO" OR Trim(intTipoPago)="CC" Then
						stsSql="SELECT DISTINCT ID_USUARIO, LOGIN FROM USUARIO U, CONVENIO_ENC C WHERE U.ID_USUARIO = C.USR_INGRESO"
					ElseIf Trim(intTipoPago)="RP" Then

					Else
						stsSql="SELECT DISTINCT ID_USUARIO, LOGIN FROM USUARIO U, CAJA_WEB_EMP C WHERE U.ID_USUARIO = C.USR_INGRESO"
					End if
					set rsUsu=Conn.execute(stsSql)
					if not rsUsu.eof then
						do until rsUsu.eof
						%>
						<option value="<%=rsUsu("ID_USUARIO")%>"
						<%if Trim(usuario)=Trim(rsUsu("ID_USUARIO")) then
							response.Write("Selected")
						end if%>
						><%=ucase(rsUsu("LOGIN"))%></option>

						<%rsUsu.movenext
						loop
					end if
					rsUsu.close
					set rsUsu=nothing
					%>
				</SELECT>
			</td>
			<td><input name="inicio" readonly="true" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
			</td>
			<td>
				<input name="termino" readonly="true" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
			</td>
			<td align="right">
				<input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();">
			</td>
	      </tr>
    </table>
    <br>
	<table class="intercalado" style="width:100%;">
		<thead>
		<tr>
			<td>COD. PAGO</td>
			<td>USU.ING.</td>
			<td>FEC.PAGO</td>
			<td>CLIENTE</td>
			<td>RUT</td>
			<td>M.CAPIT.</td>
			<td>HONOR.</td>
			<td>INTERES</td>
			<td>G.ADM.</td>
			<td>G.OPE.</td>
			<td>T.PAGO</td>
			<td>REVERSA</td>
			<td>BOLETA</td>
			<td align="center">N° BOLETA</td>
			<td>COMP.</td>
			<td>CHEQUE</td>
		</tr>
		</thead>

	<%
	strSql = ""
	If resp="si" then
		strSql = "SELECT ID_PAGO,RENDIDO, MONTO_CAPITAL, CAJA_WEB_EMP.GASTOS_ADMINISTRATIVOS, GASTOS_OTROS, INDEM_COMP, INTERES_PLAZO, MONTO_EMP, CONVERT(VARCHAR(10), FECHA_PAGO, 103) AS FECHA_PAGO,COMP_INGRESO, CLIENTE.DESCRIPCION, RUT_DEUDOR, TOTAL_CLIENTE, TOTAL_EMP, CAJA_TIPO_PAGO.DESC_TIPO_PAGO,ISNULL(USR_INGRESO,0) AS USR_INGRESO, GASTOS_JUDICIALES, NRO_BOLETA, ( SELECT COUNT(*) FROM CAJA_WEB_EMP_DOC_PAGO CWP WHERE FORMA_PAGO IN ('CD','CF') AND CWP.ID_PAGO=CAJA_WEB_EMP.ID_PAGO ) CANTIDAD_FORMA_PAGO, (SELECT COUNT(*) cantidad FROM CAJA_WEB_EMP_DOC_PAGO cwec WHERE TIPO_PAGO=1 AND cwec.id_pago=CAJA_WEB_EMP.id_pago) cantidad_doc_pago, (SELECT top 1 case when  NRO_CHEQUE='' OR NRO_CHEQUE is null then 'vacio' else '' end NRO_CHEQUE_VACIO FROM CAJA_WEB_EMP_DOC_PAGO CWP WHERE FORMA_PAGO IN ('CD','CF') AND CWP.ID_PAGO=CAJA_WEB_EMP.ID_PAGO AND (NRO_CHEQUE='' OR NRO_CHEQUE is null) ) NRO_CHEQUE_VACIO  FROM CAJA_WEB_EMP,CLIENTE,CAJA_TIPO_PAGO WHERE CLIENTE.COD_CLIENTE = CAJA_WEB_EMP.COD_CLIENTE AND CAJA_TIPO_PAGO.ID_TIPO_PAGO = CAJA_WEB_EMP.TIPO_PAGO " ' " and sucursal.cod_suc = caja_web_emp.sucursal "
		'IF sucursal <> "0" THEN
		''	strSql = strSql & "and sucursal='" & sucursal & "' "
		'END IF
		IF CLIENTE <> "0" THEN
			strSql = strSql & "and caja_web_emp.cod_cliente = '" & CLIENTE & "'"
		END IF
		IF usuario <> "0" THEN
			strSql = strSql & "and  USR_INGRESO=" & usuario & " "
		END IF
		If Trim(strRut) <> "" Then
			strSql = strSql & " AND RUT_DEUDOR = '" & strRut & "'"
		End If
		IF intTipoPago <> "" THEN
			strSql = strSql & " AND TIPO_PAGO = '" & intTipoPago & "' "
		'Else
		'	strSql = strSql & " AND TIPO_PAGO NOT IN ('CO','CC')"
		END IF

		IF intCodPago <> "" THEN
			strSql = strSql & " AND ID_PAGO = " & intCodPago & " "
		END IF

		strSql = strSql & " AND CAJA_WEB_EMP.COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ")"

		strSql = strSql & "and fecha_pago between '" & inicio & " 00:00:00' and '" & termino & " 23:59:59' order by id_pago desc"
		'strSql = strSql & "and rendido between '" & inicio & "' and '" & termino & "' and caja_web_emp.estado_caja='A' order by rendido"

		'Response.write "strSql = " & strSql & "<br>.." & CLIENTE
	End if


		'Response.write "strSql = " & strSql
		'Response.End
	if strSql <> "" then
		set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
		%>
		<tbody>
		<%
			do while not rsDet.eof
				ssql="SELECT * FROM USUARIO WHERE ID_USUARIO = " & rsDet("USR_INGRESO")
				'Response.write ssql
				'Response.End

				set rsUsuIng=Conn.execute(ssql)
				if not rsUsuIng.eof then
					USR_INGRESO= rsUsuIng("login")
				end if

			%>

			<tr >
				<td><A HREF="det_caja.asp?cod_pago=<%=rsDet("id_pago")%>"><%=rsDet("id_pago")%></A></td>
				<td><%=USR_INGRESO%></td>
				<td><%=rsDet("fecha_pago")%></td>
				<td><%=rsDet("DESCRIPCION")%></td>
				<td>
					<A HREF="principal.asp?rut=<%=rsDet("RUT_DEUDOR")%>"><acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym></A>
				</td>
				<td><%=rsDet("monto_capital")%></td>
				<td align="right"><%=rsDet("monto_emp")%></td>
				<td align="right"><%=rsDet("interes_plazo")%></td>
				<td align="right"><%=rsDet("GASTOS_ADMINISTRATIVOS")%></td>
				<td align="right"><%=rsDet("GASTOS_OTROS")%></td>
				<td><%=rsDet("desc_tipo_pago")%></td>

				<%'if perfil="caja_modif" then%>


					<!--td><A HREF="reversar_pago.asp?cod_pago=<%=rsDet("id_pago")%>">Rever.</A></td-->

					<% If TraeSiNo(session("perfil_caja"))="Si" Then %>
						<% If Trim(intTipoPago)="CO" OR Trim(intTipoPago)="CC" Then%>
							<td align="center" title="<%=rsDet("id_pago")%>"><img style="cursor:pointer;" width="20" height="20" src="../Imagenes/limpia_campo.png" onClick="Reversar(<%=rsDet("id_pago")%>)"></td>
						<% ElseIf Trim(intTipoPago)="RP" Then%>
							<td align="center" title="<%=rsDet("id_pago")%>"><img style="cursor:pointer;" width="20" height="20" src="../Imagenes/limpia_campo.png" onClick="Reversar(<%=rsDet("id_pago")%>)"></td>
						<% Else%>
							<td align="center" title="<%=rsDet("id_pago")%>"><img style="cursor:pointer;" width="20" height="20" src="../Imagenes/limpia_campo.png" onClick="Reversar(<%=rsDet("id_pago")%>)"></td>
						<% End If%>

					<% Else%>
						<td>&nbsp;</td>
					<% End If

					strUrlCompPago = strTipoCompArch & "?strImprime=S&intNroComp=" & rsDet("COMP_INGRESO")
					%>


					<td align="center">
						<img style="cursor:pointer;" width="20" height="20" src="../Imagenes/48px-Crystal_Clear_app_kword.png" onClick="ImpBoleta(<%=rsDet("comp_ingreso")%>)">
					</td>

					<%if isnull(rsDET("NRO_BOLETA")) AND rsDET("cantidad_doc_pago")>0  then
						color ="background-color:#F5A9A9;"
					else
						color =""
					end if%>					


					<td align="center"  class="actualiza_boleta" id="td_actualizar_nro_boleta_<%=trim(rsDET("id_pago"))%>">
						<div <%if rsDET("cantidad_doc_pago")>0  then%> onclick="refresca_nro_boleta('<%=rsDET("id_pago")%>','<%=rsDET("NRO_BOLETA")%>')" <%end if%> style="cursor:pointer; width:100%;height:20px; <%=trim(color)%>">
							<%=rsDET("NRO_BOLETA")%>&nbsp;&nbsp;												
						</div>
					</td>

					<td align="center" title="<%=rsDet("comp_ingreso")%>">
						
						<img style="cursor:pointer;" width="20" height="20" src="../Imagenes/48px-Crystal_Clear_app_kspread.png" onClick="VerCompPago('<%=strUrlCompPago%>')" >
						
					</td>
				<%'end if
				%>
				<td align="center">
					<%IF rsDET("CANTIDAD_FORMA_PAGO")>0  THEN%>

						<%if isnull(rsDet("NRO_CHEQUE_VACIO")) then%>			
							<img width="20" height="20" style="cursor:pointer;" id="imagen_<%=rsDET("id_pago")%>" src="../Imagenes/48px-Crystal_Clear_mimetype_document2.png" onclick="modificar_cheque('<%=rsDET("id_pago")%>')">
						<%else%>
							<img width="20" height="20" style="cursor:pointer;" id="imagen_<%=rsDET("id_pago")%>" src="../Imagenes/48px-Crystal_Clear_mimetype_document_rojo.png" onclick="modificar_cheque('<%=rsDET("id_pago")%>')">
						<%end if%>

					<%END IF%>
				</td>
			</tr>
			<INPUT TYPE="HIDDEN" NAME="HD_COMP" VALUE="<%=rsDET("id_pago")%>"></td>
			<%

			rsDet.movenext
			loop
		end if
		%>
		</tbody>
	<%end if%>
		<thead>
			<tr>
				<td colspan=5>TOTAL</td>
				<td><%=intTotCapital%></td>
				<td><%=intTotHonorarios%></td>
				<td><%=intTotIC%></td>
				<td><%=intTotGastosJud%></td>
				<td><%=intTotInteres%></td>
				<td><%=intTotGastosAdmin%></td>
				<td><%=intTotGastosOtros%></td>
				<td colspan=6>&nbsp;</td>
			</tr>
		</thead>
	</table>
	</td>
   </tr>
  </table>

</form>

</body>
</html>