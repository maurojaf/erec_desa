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

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">	
<%

	Response.CodePage=65001
	Response.charset ="utf-8"

	AbrirSCG()
	intCodUsuario=session("session_idusuario")
	strCuadrar = Trim(request("strCuadrar"))

	intUsuario = Trim(request("CB_USUARIO"))

	intEtapa = Trim(request("CB_ETAPA"))

	'intCodUsuario = 110
	strsql="SELECT * FROM USUARIO WHERE ID_USUARIO = " & intCodUsuario & ""
	set rsUsu=Conn.execute(strsql)
	if not rsUsu.eof then

	end if
	'response.write(perfil)

	codpago = request("TX_PAGO")
	strrut=request("TX_RUT")
	if sucursal="" then sucursal="0"
	'response.write(sucursal)
	usuario = request("cmb_usuario")
	if usuario = "" then usuario = "0"
	termino = request("termino")
	inicio = request("inicio")
	resp = request("resp")
	GRABA = request("GRABA")
	if Trim(inicio) = "" Then
		inicio = TraeFechaMesActual(Conn,0)
		'inicio = "01" & Mid(inicio,3,10)
	End If
	if Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If

	strCOD_CLIENTE = session("ses_codcli")

	'Response.write "GRABA=" & GRABA
	'Response.End


	strsql = "SELECT ISNULL(SUM(ASIGNACION),0) AS ASIG FROM CAJA_WEB_EMP_CIERRE "
	strsql = strsql & "WHERE COD_USUARIO = " & intCodUsuario & " AND FECHA_APERTURA >= '" & inicio & " 00:00 ' AND FECHA_APERTURA <= '" & termino & " 23:59'"
	set rsFechas=Conn.execute(strsql)
	If not rsFechas.eof then
		intValorAsigCaja = rsFechas("ASIG")
		intValorCajaAsigCaja = intValorAsigCaja
	Else
		intValorAsigCaja = 0
		intValorCajaAsigCaja = 0
	End If

	'hoy=date

%>
	<title>Empresa</title>
	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo13n {color: #000000}
	.Estilo27 {color: #FFFFFF}
	-->
	</style>

    
    <link href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css" rel="stylesheet">
    <script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>  
    <script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script> 

	<script language="JavaScript " type="text/JavaScript">
	$(document).ready(function(){
		$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

	})
	function muestra_dia(){
	//alert(getCurrentDate())
	//alert("hola")
		var diferencia=DiferenciaFechas(datos.termino.value)
		//alert(diferencia)
		if(datos.termino.value!=''){
			if ((diferencia<=0)) {
				//alert('Ok')
				return true
			}else{
				alert('La fecha de cierre no puede ser posterior a la fecha actual')
				datos.termino.value = getCurrentDate();
				datos.termino.focus();
				return false;
			}
		}
	}


	function DiferenciaFechas (CadenaFecha1) {
	   var fecha_hoy = getCurrentDate() //hoy


	   //Obtiene dia, mes y año
	   var fecha1 = new fecha( CadenaFecha1 )
	   var fecha2 = new fecha(fecha_hoy)

	   //Obtiene objetos Date
	   var miFecha1 = new Date( fecha1.anio, fecha1.mes, fecha1.dia )
	   var miFecha2 = new Date( fecha2.anio, fecha2.mes, fecha2.dia )

	   //Resta fechas y redondea
	   var diferencia = miFecha1.getTime() - miFecha2.getTime()
	   var dias = Math.floor(diferencia / (1000 * 60 * 60 * 24))
	   var segundos = Math.floor(diferencia / 1000)
	   //alert ('La diferencia es de ' + dias + ' dias,\no ' + segundos + ' segundos.')

	   return dias //false
	}

	function fecha( cadena ) {

	   //Separador para la introduccion de las fechas
	   var separador = "/"

	   //Separa por dia, mes y año
	   if ( cadena.indexOf( separador ) != -1 ) {
	        var POSI_1 = 0
	        var POSI_2 = cadena.indexOf( separador, POSI_1 + 1 )
	        var POSI_3 = cadena.indexOf( separador, POSI_2 + 1 )
	        this.dia = cadena.substring( POSI_1, POSI_2 )
	        this.mes = cadena.substring( POSI_2 + 1, POSI_3 )
	        this.anio = cadena.substring( POSI_3 + 1, cadena.length )
	   } else {
	        this.dia = 0
	        this.mes = 0
	        this.anio = 0
	   }
	}

	function Refrescar()
	{
		GRABA='no'
		resp='no'

		datos.action = "cerrar_caja_web.asp?GRABA="+ GRABA +"&resp="+ resp +"";
		datos.submit();
	}



	function Ingresa()
	{
		GRABA='si'
		resp='si'
		strCuadrar='no'
		if (!muestra_dia()) return;
		if (confirm("¿Está seguro de cerrar la caja para el usuario : " + datos.CB_USUARIO.options[datos.CB_USUARIO.selectedIndex].text + " , no podrá seguir ingresando pagos para el día de hoy."))

			{
			datos.action = "cerrar_caja_web.asp?strCuadrar="+ strCuadrar +"&GRABA="+ GRABA +"&resp="+ resp +"";
			datos.submit();
			}
		else
			alert("Caja no será cerrada");

	}


	function Cuadrar()
	{
		//datos.TX_RUT.value='';
		//datos.TX_pago.value='';
		GRABA='no'
		resp='si'
		strCuadrar='si'
		datos.action = "cerrar_caja_web.asp?strCuadrar="+ strCuadrar +"&GRABA="+ GRABA +"&resp="+ resp +"";
		datos.submit();
	}


	function envia()
	{
		//datos.TX_RUT.value='';
		//datos.TX_pago.value='';
		GRABA='no'
		resp='si'
		strCuadrar='no'
		datos.action = "cerrar_caja_web.asp?strCuadrar="+ strCuadrar +"&GRABA="+ GRABA +"&resp="+ resp +"";
		datos.submit();
	}

	function envia_excel(URL){

	window.open(URL,"INFORMACION","width=200, height=200, scrollbars=yes, menubar=yes, location=yes, resizable=yes")
	}
	</script>


</head>
<body>
<form name="datos" method="post">
	<div class="titulo_informe">CERRAR CAJA</div>
<table width="90%" height="500" border="0" align="center">
  <tr>
    <td valign="top">
	<table width="100%" border="0" bordercolor="#999999">
	</table>
	<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
		<thead>
	      <tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
		  <%if perfil="caja_modif" or perfil = "caja_listado" then%>
	        <td>SUCURSAL</td>
			<%end if%>
			<td>DESDE</td>
			<td>HASTA</td>
			<td>USUARIO</td>
			<td>ETAPA</td>
			<td>&nbsp;</td>

	      </tr>
	    </thead>
		  <tr bordercolor="#999999" class="Estilo8">
			<td><input name="inicio" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
			</td>
			<td><input name="termino" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
			</td>
			<td>
			<select name="CB_USUARIO">
				<%If Trim(TraeSiNo(session("perfil_sup"))="Si") or Trim(TraeSiNo(session("perfil_emp"))="Si") Then%>
					<option value="">TODOS</option>
				<%End If%>
				<%
					strSql="SELECT * FROM USUARIO WHERE PERFIL_CAJA = 1 AND ACTIVO = 1"
					If Trim(TraeSiNo(session("perfil_sup"))="No") Then
						strSql=strSql & " AND ID_USUARIO = " & intCodUsuario
					End if
					If Trim(intUsuario) <> "" Then
						strSql=strSql & " AND ID_USUARIO = " & intUsuario
					End if
					set rsUsuario=Conn.execute(strSql)
					if Not rsUsuario.eof then
						do until rsUsuario.eof
						%>
						<option value="<%=rsUsuario("ID_USUARIO")%>"
						<%if Trim(intUsuario)=Trim(rsUsuario("ID_USUARIO")) then Response.Write("SELECTED") end if%> WIDTH="10"
						><%=MID(rsUsuario("LOGIN"),1,19)%></option>
						<%rsUsuario.movenext
						loop
					end if
					rsUsuario.close
					set rsUsuario=nothing
					%>
			</select>
			</td>
			<td>
				<select name="CB_ETAPA">
					<option value="">TODOS</option>
					<option value="1">ETAPA 1</option>
					<option value="2">ETAPA 2</option>
					<option value="3">ETAPA 3</option>
				</select>
			</td>
			<td>
			<input type="Button" class="fondo_boton_100" name="Ver" value="Ver" onClick="envia();">
			<input type="Button" name="Cerrar" class="fondo_boton_100" value="Cerrar" onClick="Ingresa();" DISABLED>
			</td>

	      </tr>
    </table>

	<% if resp="si" then

	if strCuadrar = "si" then


		intValorChequeDia = ValNulo(Request("TX_CHEQUEDIA"),"N")
		intValorChequeFecha = ValNulo(Request("TX_CHEQUEFECHA"),"N")
		intValorEfectivo = ValNulo(Request("TX_EFECTIVO"),"N")
		intValorDeposito = ValNulo(Request("TX_DEPOSITO"),"N")
		intValorOCanales = ValNulo(Request("TX_OCANALES"),"N")
		intValorCuota = ValNulo(Request("TX_CUOTA"),"N")
		intValorTotal = ValNulo(Request("TX_TOTALHABER"),"N")



		intValorCajaChequeDia = ValNulo(Request("hd_valChequeDia"),"N")
		intValorCajaChequeFecha = ValNulo(Request("hd_valChequeFecha"),"N")
		intValorCajaEfectivo = ValNulo(Request("hd_valEfectivo"),"N")
		intValorCajaDeposito = ValNulo(Request("hd_valDeposito"),"N")
		intValorCajaCuota = ValNulo(Request("hd_valCuota"),"N")
		intValorCajaOCanales = ValNulo(Request("hd_valOCanales"),"N")
		intValorCajaTotal = ValNulo(Request("hd_valTotal"),"N")



		strDescChequeDia = ValNulo(Request("TX_CHEQUEDIA"),"N") - ValNulo(Request("hd_valChequeDia"),"N")
		strDescChequeFecha = ValNulo(Request("TX_CHEQUEFECHA"),"N") - ValNulo(Request("hd_valChequeFecha"),"N")
		strDescEfectivo = ValNulo(Request("TX_EFECTIVO"),"N") - ValNulo(Request("hd_valEfectivo"),"N")
		strDescDeposito = ValNulo(Request("TX_DEPOSITO"),"N") - ValNulo(Request("hd_valDeposito"),"N")
		strDescCuota = ValNulo(Request("TX_CUOTA"),"N") - ValNulo(Request("hd_valCuota"),"N")
		strDescOCanales = ValNulo(Request("TX_OCANALES"),"N") - ValNulo(Request("hd_valOCanales"),"N")



		strDescTotal = ValNulo(Request("TX_TOTALHABER"),"N") - ValNulo(Request("hd_valTotal"),"N")


	End If

	%>
	<table width="100%" class="intercalado" style="width:100%;">
		<thead>
		<tr >
			<td>FORMA DE PAGO</td>
			<td>MONTO CLIENTE</td>
			<td>MONTO EMPRESA</td>
			<td>TOTAL</td>
		</tr>
		</thead>
		<tbody>
	<%
	'dim vectormontos
			SQL = "SELECT * FROM CAJA_FORMA_PAGO WHERE ID_FORMA_PAGO NOT IN ('AB')"
			set rsDet=Conn.execute(SQL)
		if not rsDet.eof then


			Do while not rsDet.eof
				forma_pago = rsDet("id_forma_pago")
				nom_forma_pago = rsDet("desc_forma_pago")
				strSql = "SELECT  SUM(CWDP.MONTO) AS MONTO FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, CLIENTE CLI, DEUDOR DEU "
				strSql = strSql & " WHERE CWC.ID_PAGO = CWDP.ID_PAGO AND CLI.COD_CLIENTE = CWC.COD_CLIENTE AND CWC.RUT_DEUDOR = DEU.RUT_DEUDOR AND CLI.COD_CLIENTE = DEU.COD_CLIENTE "
				strSql = strSql & " AND CWDP.FORMA_PAGO = '" & forma_pago & "' AND CWDP.TIPO_PAGO = 0 AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(day,FECHA_PAGO,'" & TERMINO & "') >= 0 "

				If Trim(intEtapa) <> "" Then
					strSql = strSql & " AND DEU.ETAPA_COBRANZA = " & Trim(intEtapa)
				End If


				If Trim(TraeSiNo(session("perfil_sup"))="No") Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
				End if
				If Trim(intUsuario) <> "" Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intUsuario
				End if

				strSql=strSql & " GROUP BY CWDP.TIPO_PAGO"

				'Response.write strSql
				set rsPago=Conn.execute(strSql)
				monto_cliente = 0
				if not rsPago.eof then

					do while not rsPago.eof
						monto_cliente = monto_cliente + rsPago("MONTO")
						rsPago.movenext
					loop

				end if
				total_cliente = total_cliente + monto_cliente

				strSql="SELECT  SUM(CWDP.MONTO) AS MONTO_EMPR FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, CLIENTE CLI, DEUDOR DEU "
				strSql= strSql & " WHERE CWC.ID_PAGO = CWDP.ID_PAGO AND CLI.COD_CLIENTE = CWC.COD_CLIENTE AND CWC.RUT_DEUDOR = DEU.RUT_DEUDOR AND CLI.COD_CLIENTE = DEU.COD_CLIENTE "
				strSql= strSql & " AND CWDP.FORMA_PAGO = '" & forma_pago & "' AND CWDP.TIPO_PAGO = 1 AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(day,FECHA_PAGO,'" & TERMINO & "')>=0 "

				If Trim(intEtapa) <> "" Then
					strSql = strSql & " AND DEU.ETAPA_COBRANZA = " & Trim(intEtapa)
				End If

				If Trim(TraeSiNo(session("perfil_sup"))="No") Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
				End if
				If Trim(intUsuario) <> "" Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intUsuario
				End if

				strSql=strSql & " GROUP BY CWDP.TIPO_PAGO"

				'Response.write strSql
				set rsPago2=Conn.execute(strSql)
				monto_empresa = 0
				if not rsPago2.eof then

					do while not rsPago2.eof
						monto_empresa = monto_empresa + rsPago2("MONTO_EMPR")
						rsPago2.movenext
					loop

				end if
				total_intercapital = total_intercapital + monto_empresa


				If Trim(forma_pago) = "CD" Then
				%>
				<INPUT TYPE="hidden" NAME="hd_valChequeDia" value="<%=monto_cliente + monto_empresa%>">
				<INPUT TYPE="hidden" NAME="hd_valChequeDiaCli" value="<%=monto_cliente%>">
				<INPUT TYPE="hidden" NAME="hd_valChequeDiaEmp" value="<%=monto_empresa%>">
	       		<%
				End If
				If Trim(forma_pago) = "CF" Then
				%>
				<INPUT TYPE="hidden" NAME="hd_valChequeFecha" value="<%=monto_cliente + monto_empresa%>">
				<INPUT TYPE="hidden" NAME="hd_valChequeFechaCli" value="<%=monto_cliente%>">
				<INPUT TYPE="hidden" NAME="hd_valChequeFechaEmp" value="<%=monto_empresa%>">
				<%
				End If
				If Trim(forma_pago) = "EF" Then
				%>
				<INPUT TYPE="hidden" NAME="hd_valEfectivo" value="<%=monto_cliente + monto_empresa%>">
				<INPUT TYPE="hidden" NAME="hd_valEfectivoCli" value="<%=monto_cliente%>">
				<INPUT TYPE="hidden" NAME="hd_valEfectivoEmp" value="<%=monto_empresa%>">
				<%
				End If
				If Trim(forma_pago) = "DP" Then
				%>
				<INPUT TYPE="hidden" NAME="hd_valDeposito" value="<%=monto_cliente + monto_empresa%>">
				<INPUT TYPE="hidden" NAME="hd_valDepositoCli" value="<%=monto_cliente%>">
				<INPUT TYPE="hidden" NAME="hd_valDepositoEmp" value="<%=monto_empresa%>">
				<%
				End If
				If Trim(forma_pago) = "VV" Then
				%>
				<INPUT TYPE="hidden" NAME="hd_valOCanales" value="<%=monto_cliente + monto_empresa%>">
				<INPUT TYPE="hidden" NAME="hd_valOCanalesCli" value="<%=monto_cliente%>">
				<INPUT TYPE="hidden" NAME="hd_valOCanalesEmp" value="<%=monto_empresa%>">
				<%
				End If
				If Trim(forma_pago) = "CU" Then
				%>
				<INPUT TYPE="hidden" NAME="hd_valCuota" value="<%=monto_cliente + monto_empresa%>">
				<INPUT TYPE="hidden" NAME="hd_valCuotaCli" value="<%=monto_cliente%>">
				<INPUT TYPE="hidden" NAME="hd_valCuotaEmp" value="<%=monto_empresa%>">
				<%
				End If
				If Trim(forma_pago) = "OC" Then
				%>
				<INPUT TYPE="hidden" NAME="hd_valOtrosCanales" value="<%=monto_cliente + monto_empresa%>">
				<INPUT TYPE="hidden" NAME="hd_valOtrosCanalesCli" value="<%=monto_cliente%>">
				<INPUT TYPE="hidden" NAME="hd_valOtrosCanalesEmp" value="<%=monto_empresa%>">
				<%
				End If



			%>
			<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
				<td align="left"><%= UCASE(nom_forma_pago)%></td>
				<td align="right"><%=formatnumber(monto_cliente,0)%></td>
				<td align="right"><%=formatnumber(monto_empresa,0)%></td>
				<td align="right"><%=formatnumber(monto_cliente + monto_empresa,0)%></td>
			</tr>
			<%
			rsDet.movenext
			Loop
		end if
		%>
		</tbody>
		<thead>
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td align="left">TOTAL</td>
				<td align="right"><%= formatnumber(total_cliente,0)%></td>
				<td align="right"><%= formatnumber(total_intercapital,0)%></td>
				<td align="right" >
				<A HREF="detalle_cuadratura.asp?resp=SI&INICIO=<%=INICIO%>&TERMINO=<%=TERMINO%>"><%= formatnumber(total_cliente + total_intercapital + intValorCajaAsigCaja,0)%></A>
				</td>
		</tr>
		</thead>
		</table>
		<INPUT TYPE="hidden" NAME="hd_valTotal" value="<%=total_cliente + total_intercapital + intValorCajaAsigCaja%>">

	<BR>
	<table width="100%" border="0" bordercolor="#000000" class="intercalado"  style="width:100%;">
		<thead>
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<TD WIDTH="30%">CLIENTE</td>
			<TD WIDTH="10%">EFECTIVO</td>
			<td WIDTH="20%">BANCO</td>
			<td WIDTH="10%">CHEQUE AL DIA</td>
			<td WIDTH="20%">BANCO</td>
			<td WIDTH="10%">CHEQUE A FECHA</td>
		</tr>
		</thead>
		<tbody>

		<%
		X = 1
		DO WHILE X < 2
			intEfectivoEmpresa = 0
			strBancoEmpresaEfec = ""
			intEmpresaCD = 0
			strBancoEmpresaCD = ""
			intEmpresaCFC = 0


			'************************************************************
			'* 					EFECTIVO 			 					*
			'*															*
			'************************************************************


			strSql = "SELECT CWDP.TIPO_PAGO,ISNULL(SUM(CWDP.MONTO),0) AS MONTO_EMPR FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, CLIENTE CLI, DEUDOR DEU "
			strSql = strSql & " WHERE CWC.ID_PAGO = CWDP.ID_PAGO AND CLI.COD_CLIENTE = CWC.COD_CLIENTE AND CWC.RUT_DEUDOR = DEU.RUT_DEUDOR AND CLI.COD_CLIENTE = DEU.COD_CLIENTE "


			If Trim(intEtapa) <> "" Then
				strSql = strSql & " AND DEU.ETAPA_COBRANZA = " & Trim(intEtapa)
			End If


			If Trim(TraeSiNo(session("perfil_sup"))="No") Then
				strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
			End if
			If Trim(intUsuario) <> "" Then
				strSql=strSql & " AND CWC.USR_INGRESO = " & intUsuario
			End if




			strSql= strSql & " AND CWDP.FORMA_PAGO = 'EF' AND CWDP.TIPO_PAGO in (" & X & ") AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(day,FECHA_PAGO,'" & TERMINO & "')>=0 GROUP BY CWDP.TIPO_PAGO"

			'Response.write strSql
			set rsEmpresa=Conn.execute(strSql)
			if not rsEmpresa.eof then
				intEfectivoEmpresa = rsEmpresa("MONTO_EMPR")
				''strBancoEmpresaEfec = rsEmpresa("NOMBRE_B")
				intTotalEfectivo = intTotalEfectivo + intEfectivoEmpresa
			END IF


			'************************************************************
			'* 					CHEQUE AL DIA EMPRESA					*
			'*															*
			'************************************************************

			strSql = "SELECT CWDP.TIPO_PAGO,ISNULL(SUM(CWDP.MONTO),0) AS MONTO_EMPR,NOMBRE_B FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, BANCOS B, CLIENTE CLI, DEUDOR DEU "
			strSql = strSql & " WHERE CWC.ID_PAGO = CWDP.ID_PAGO AND CWDP.COD_BANCO *= B.CODIGO AND CLI.COD_CLIENTE = CWC.COD_CLIENTE AND CWC.RUT_DEUDOR = DEU.RUT_DEUDOR AND CLI.COD_CLIENTE = DEU.COD_CLIENTE "

			If Trim(TraeSiNo(session("perfil_sup"))="No") Then
				strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
			End if
			If Trim(intUsuario) <> "" Then
				strSql=strSql & " AND CWC.USR_INGRESO = " & intUsuario
			End if

			If Trim(intEtapa) <> "" Then
				strSql = strSql & " AND DEU.ETAPA_COBRANZA = " & Trim(intEtapa)
			End If

			strSql = strSql & "AND CWDP.FORMA_PAGO = 'CD' AND CWDP.TIPO_PAGO in (" & X & ") AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(day,FECHA_PAGO,'" & TERMINO & "')>=0 GROUP BY CWDP.TIPO_PAGO,NOMBRE_B"
			set rsEmpresaCD=Conn.execute(strSql)

			if not rsEmpresaCD.eof then
				'intEmpresaCD = rsEmpresaCD("MONTO_EMPR")
				'strBancoEmpresaCD = rsEmpresaCD("NOMBRE_B")
				'intTotalCDEmpresa = intTotalCDEmpresa + intEmpresaCD
			End If



			'************************************************************
			'* 					CHEQUE A FECHA EMPRESA					*
			'*															*
			'************************************************************

			strSql = "SELECT CWDP.TIPO_PAGO,ISNULL(SUM(CWDP.MONTO),0) AS MONTO_EMPR,NOMBRE_B FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, BANCOS B, CLIENTE CLI, DEUDOR DEU "
			strSql=strSql & " WHERE CWC.ID_PAGO = CWDP.ID_PAGO AND CWDP.COD_BANCO *= B.CODIGO AND CLI.COD_CLIENTE = CWC.COD_CLIENTE AND CWC.RUT_DEUDOR = DEU.RUT_DEUDOR AND CLI.COD_CLIENTE = DEU.COD_CLIENTE "

			If Trim(TraeSiNo(session("perfil_sup"))="No") Then
				strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
			End if
			If Trim(intUsuario) <> "" Then
				strSql=strSql & " AND CWC.USR_INGRESO = " & intUsuario
			End if
			strSql = strSql + "AND CWDP.FORMA_PAGO = 'CF' AND CWDP.TIPO_PAGO in (" & X & ") AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(day,FECHA_PAGO,'" & TERMINO & "')>=0 GROUP BY CWDP.TIPO_PAGO,NOMBRE_B"

			'Response.write "<br>" & strSql
			'Response.End
			set rsEmpresaCF=Conn.execute(strSql)
			if not rsEmpresaCF.eof then
				'intEmpresaCF = rsEmpresaCF("MONTO_EMPR")
				'intTotalCF = intTotalCF + intEmpresaCF
				'strBancoEmpresaCD = rsEmpresaCF("NOMBRE_B")
			END IF



			IF X = 1 THEN EMPRESA = "Empresa"
			IF X = 2 THEN EMPRESA = "Costas"
			%>
			<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
				<TD><%=EMPRESA%></TD>
				<TD ALIGN="RIGHT"><b>
				<A HREF="detalle_cuadratura.asp?resp=SI&INICIO=<%=INICIO%>&TERMINO=<%=TERMINO%>&CLIENTE=0&CB_FORMA_PAGO=EF&CB_DESTINO=1"><%=FORMATNUMBER(intEfectivoEmpresa,0)%></A>
				</b>
				</td>

				<td colspan="2">
				<table width="100%" border="0" bordercolor="#000000">
				<%
				Do While Not rsEmpresaCD.Eof
					intEmpresaCD = rsEmpresaCD("MONTO_EMPR")
					strBancoEmpresaCD = rsEmpresaCD("NOMBRE_B")
					intTotalEmpresaCD = intTotalEmpresaCD + intEmpresaCD
				%>
				<tr>
				<td align="LEFT"><%=strBancoEmpresaCD%></td>
				<td ALIGN="RIGHT"><%=FORMATNUMBER(intEmpresaCD,0)%></td>
				</tr>

				<%
					rsEmpresaCD.movenext
				Loop
				%>

				<tr>
					<td align="LEFT"><B>TOTAL</B></td>
					<td ALIGN="RIGHT"><B><%=FORMATNUMBER(intTotalEmpresaCD,0)%></B></td>
				</tr>

				</table>
				</TD>



				<td colspan="2">
					<table width="100%" border="0" bordercolor="#000000">
					<%
					Do While Not rsEmpresaCF.Eof
						intEmpresaCF = rsEmpresaCF("MONTO_EMPR")
						strBancoEmpresaCF = rsEmpresaCF("NOMBRE_B")
						intTotalEmpresaCF = intTotalEmpresaCF + intEmpresaCF
					%>
					<tr>
					<td align="LEFT"><%=strBancoEmpresaCF%></td>
					<td ALIGN="RIGHT"><%=FORMATNUMBER(intEmpresaCF,0)%></td>
					</tr>

					<%
						rsEmpresaCF.movenext
					Loop
					%>

					<tr>
						<td align="LEFT"><B>TOTAL</B></td>
						<td ALIGN="RIGHT"><B><%=FORMATNUMBER(intTotalEmpresaCF,0)%></B></td>
					</tr>

					</table>
				</TD>
			</tr>
		<%
			X = X + 1
		LOOP
		strSql = "SELECT * FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE <> '999'"
		set rsCliente=Conn.execute(strSql)
		if not rsCliente.eof then
			do while not rsCliente.eof
				EF = 0
				CD = 0
				CF = 0
				BANCOEF = ""
				BANCOCD = ""

				CLIENTE = rsCliente("COD_CLIENTE")

				strSql = "SELECT  ISNULL(SUM(CWDP.MONTO),0) AS MONTO FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, CLIENTE CLI, DEUDOR DEU WHERE CWC.ID_PAGO = CWDP.ID_PAGO AND CLI.COD_CLIENTE = CWC.COD_CLIENTE AND CWC.RUT_DEUDOR = DEU.RUT_DEUDOR AND CLI.COD_CLIENTE = DEU.COD_CLIENTE "
				If Trim(TraeSiNo(session("perfil_sup"))="No") Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
				End if
				If Trim(intUsuario) <> "" Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intUsuario
				End if
				If Trim(intEtapa) <> "" Then
					strSql = strSql & " AND DEU.ETAPA_COBRANZA = " & Trim(intEtapa)
				End If
				strSql = strSql & " AND CWDP.FORMA_PAGO in ('EF')  AND CWDP.TIPO_PAGO = 0 AND CWC.COD_CLIENTE = '" & CLIENTE & "' AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(day,FECHA_PAGO,'" & TERMINO & "')>=0 GROUP BY CWDP.TIPO_PAGO"

				'Response.write "<br>strSql=" & strSql
				set rsDetalle=Conn.execute(strSql)

				IF NOT rsDetalle.EOF THEN
					EF = rsDetalle("MONTO")
					intTotalEfectivo = intTotalEfectivo  + EF
				END IF



				strSql = "SELECT  ISNULL(SUM(CWDP.MONTO),0) AS MONTO ,NOMBRE_B FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, BANCOS B, CLIENTE CLI, DEUDOR DEU WHERE CWC.ID_PAGO = CWDP.ID_PAGO AND CWDP.COD_BANCO = B.CODIGO AND CLI.COD_CLIENTE = CWC.COD_CLIENTE AND CWC.RUT_DEUDOR = DEU.RUT_DEUDOR AND CLI.COD_CLIENTE = DEU.COD_CLIENTE "
				If Trim(TraeSiNo(session("perfil_sup"))="No") Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
				End if
				If Trim(intUsuario) <> "" Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intUsuario
				End if
				If Trim(intEtapa) <> "" Then
					strSql = strSql & " AND DEU.ETAPA_COBRANZA = " & Trim(intEtapa)
				End If
				strSql = strSql & "AND CWDP.FORMA_PAGO in ('CD')  AND CWDP.TIPO_PAGO = 0 AND CWC.COD_CLIENTE = '" & CLIENTE & "' AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(day,FECHA_PAGO,'" & TERMINO & "')>=0 GROUP BY CWDP.TIPO_PAGO,NOMBRE_B"
				set rsDetalleCD=Conn.execute(strSql)

				If Not rsDetalleCD.EOF THEN
					CD = rsDetalleCD("MONTO")
					BANCOCD = rsDetalleCD("NOMBRE_B")
					intTotalCD = intTotalCD + CD
				End if



				strSql = "SELECT  ISNULL(SUM(CWDP.MONTO),0) AS MONTO ,NOMBRE_B FROM CAJA_WEB_EMP CWC,CAJA_WEB_EMP_DOC_PAGO CWDP, BANCOS B, CLIENTE CLI, DEUDOR DEU WHERE CWC.ID_PAGO = CWDP.ID_PAGO AND CWDP.COD_BANCO = B.CODIGO AND CLI.COD_CLIENTE = CWC.COD_CLIENTE AND CWC.RUT_DEUDOR = DEU.RUT_DEUDOR AND CLI.COD_CLIENTE = DEU.COD_CLIENTE "
				If Trim(TraeSiNo(session("perfil_sup"))="No") Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intCodUsuario
				End if
				If Trim(intUsuario) <> "" Then
					strSql=strSql & " AND CWC.USR_INGRESO = " & intUsuario
				End if
				If Trim(intEtapa) <> "" Then
					strSql = strSql & " AND DEU.ETAPA_COBRANZA = " & Trim(intEtapa)
				End If
				strSql = strSql & " AND CWDP.FORMA_PAGO in ('CF')  AND CWDP.TIPO_PAGO = 0 AND CWC.COD_CLIENTE = '" & CLIENTE & "' AND DATEDIFF(DAY,'" & INICIO & "',FECHA_PAGO)>=0 AND DATEDIFF(day,FECHA_PAGO,'" & TERMINO & "')>=0 GROUP BY CWDP.TIPO_PAGO,NOMBRE_B"
				'Response.write "<br>strSql=" & strSql
				set rsDetalleCF=Conn.execute(strSql)
				If not rsDetalleCF.EOF THEN
					CF = rsDetalleCF("MONTO")
					'CHEQUESBANCO = rsDetalleCF("NOMBRE_B")
					intTotalCF = intTotalCF + CF
				End if
				BANCO = ""

					%>
					<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
						<TD><%=rsCliente("DESCRIPCION")%></TD>
						<TD ALIGN="RIGHT"><b>
						<A HREF="detalle_cuadratura.asp?resp=SI&INICIO=<%=INICIO%>&TERMINO=<%=TERMINO%>&CLIENTE=<%=CLIENTE%>&CB_FORMA_PAGO=EF&CB_DESTINO=0"><%=formatnumber(EF,0)%></A>
						</b></td>

						<td COLSPAN=2>
							<table width="100%" border="0" bordercolor="#000000">
								<%
								''intTotalCDia = 0
								intTotalCDiaSubTotal = 0
								Do while not rsDetalleCD.eof
									CD = rsDetalleCD("MONTO")
									BANCOCD = rsDetalleCD("NOMBRE_B")
									intTotalCDia = intTotalCDia + CD
									intTotalCDiaSubTotal = intTotalCDiaSubTotal + CD

									''rESPONSE.WRITE "<BR>CD=" & CD
									%>

									<TR>
										<TD ALIGN="LEFT" width="58%"><%=BANCOCD%></TD>
										<TD ALIGN="RIGHT" width="42%"><%=formatnumber(CD,0)%></TD>
									</TR>

									<%
									rsDetalleCD.movenext
								Loop

								%>

								<TR>
									<TD ALIGN="LEFT" width="58%"><B>TOTAL</B></TD>
									<TD ALIGN="RIGHT" width="42%"><B><%=formatnumber(intTotalCDiaSubTotal,0)%></B></TD>
								</TR>


							</TABLE>
						</TD>



						<td COLSPAN=2>
							<TABLE width="100%" border = 0 VALIGN="TOP">
								<%
								''intTotalCFecha = 0
								intTotalCFechaSubTotal = 0
								Do while not rsDetalleCF.eof
									CF = rsDetalleCF("MONTO")
									BANCOCF = rsDetalleCF("NOMBRE_B")
									intTotalCFecha = intTotalCFecha + CF
									intTotalCFechaSubTotal = intTotalCFechaSubTotal + CF
									%>

									<TR>
										<TD ALIGN="LEFT" width="58%"><%=BANCOCF%></TD>
										<TD ALIGN="RIGHT" width="42%"><%=formatnumber(CF,0)%></TD>
									</TR>

									<%
									rsDetalleCF.movenext
								Loop
								%>
								<TR>
									<TD ALIGN="LEFT" width="58%"><B>TOTAL</B></TD>
									<TD ALIGN="RIGHT" width="42%"><B><%=formatnumber(intTotalCFechaSubTotal,0)%></B></TD>
								</TR>
							</TABLE>
						</td>

					</tr>
					<%
				rsCliente.movenext
			loop
		end if
		%>
		</tbody>
		<thead>
		<TR class="totales">
		<thead>	
			<TD>TOTAL</TD>
			<TD ALIGN = "RIGHT" ><b><A HREF="detalle_cuadratura.asp?resp=SI&INICIO=<%=INICIO%>&TERMINO=<%=TERMINO%>&CLIENTE=0&CB_FORMA_PAGO=EF&CB_DESTINO="><%=FORMATNUMBER(intTotalEfectivo,0)%></A></b></TD>
			<TD></TD>
			<TD ALIGN = "RIGHT"><b><%=FORMATNUMBER(intTotalCDia + intTotalEmpresaCD,0)%></b></TD>
			<TD></TD>
			<TD ALIGN = "RIGHT"><b><%=FORMATNUMBER(intTotalCFecha + intTotalEmpresaCF,0)%></b></TD>
		</TR>

		<TR class="Estilo13">
			<TD COLSPAN="6" ALIGN = "RIGHT">&nbsp;</TD>
		</TR>



		<TR bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<TD><b>TOTAL CHEQUES EMPRESA</b></TD>
			<TD ALIGN = "RIGHT"><b></b></TD>
			<TD></TD>
			<TD ALIGN = "RIGHT">&nbsp;</b></TD>
			<TD></TD>
			<TD ALIGN = "RIGHT"><b><%=FORMATNUMBER(intTotalEmpresaCF + intTotalEmpresaCD,0)%></b></TD>
		</TR>

		<TR bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<TD><b>TOTAL CHEQUES CLIENTE</b></TD>
			<TD ALIGN = "RIGHT"><b></b></TD>
			<TD></TD>
			<TD ALIGN = "RIGHT">&nbsp;</b></TD>
			<TD></TD>
			<TD ALIGN = "RIGHT"><b><%=FORMATNUMBER(intTotalCDia + intTotalCFecha,0)%></b></TD>
		</TR>

		<TR bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<TD><b>TOTAL CHEQUES</b></TD>
			<TD ALIGN = "RIGHT"><b></b></TD>
			<TD></TD>
			<TD ALIGN = "RIGHT">&nbsp;</b></TD>
			<TD></TD>
			<TD ALIGN = "RIGHT"><b><%=FORMATNUMBER(intTotalCDia + intTotalEmpresaCD + intTotalCFecha + intTotalEmpresaCF,0)%></b></TD>

		</TR>
		</thead>
	</table>
<br>
<br>

	<table width="100%" height="500" border="0" align="center">

	<tr>
		<td height="20" bordercolor="#999999" class="subtitulo_informe" ALIGN="CENTER">
		> CUADRATURA CAJA
		</td>
	</tr>

	<tr>
		<td height="20" class="estilo_columna_individual" ALIGN="LEFT">
		HABER EXISTENTE
		</td>
	</tr>
  <tr>
    <td valign="top">
	<table width="100%" border="0" bordercolor="#999999">
		<tr height="20" bordercolor="#999999" >
			<td width="25%">&nbsp;</td>
			<td width="15%">&nbsp;</td>
			<td width="20%" ALIGN="RIGHT"><B>Total Caja</B></td>
			<td width="20%" ALIGN="RIGHT"><B>Total Sistema</B></td>
			<td width="20%" ALIGN="RIGHT"><B>Diferencia</B></td>
		</tr>
		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG2")%>" class="Estilo13n">
			<td class="estilo_columna_individual">ASIGNACION DE CAJA</td>
			<td><input name="TX_ASIGCAJA" type="text" value="<%=intValorAsigCaja%>" size="10" maxlength="10" DISABLED onchange=""></td>
			<td ALIGN="RIGHT"><%=FN(intValorAsigCaja,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intValorCajaAsigCaja,0)%></td>
			<td ALIGN="RIGHT"><%=FN(strDescAsigCaja,0)%></td>
		</tr>
		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG2")%>" class="Estilo13n">
			<td class="estilo_columna_individual">CHEQUES AL DIA</td>
			<td><input name="TX_CHEQUEDIA" type="text" value="<%=intValorChequeDia%>" size="10" maxlength="10" onchange="solonumero(TX_CHEQUEDIA);suma_total_general(1);"></td>
			<td ALIGN="RIGHT"><%=FN(intValorChequeDia,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intValorCajaChequeDia,0)%></td>
			<td ALIGN="RIGHT"><%=FN(strDescChequeDia,0)%></td>
		</tr>
		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG2")%>" class="Estilo13n">
			<td class="estilo_columna_individual">CHEQUES A FECHA</td>
			<td><input name="TX_CHEQUEFECHA" type="text" value="<%=intValorChequeFecha%>" size="10" maxlength="10" onchange="solonumero(TX_CHEQUEFECHA);suma_total_general(1);"></td>
			<td ALIGN="RIGHT"><%=FN(intValorChequeFecha,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intValorCajaChequeFecha,0)%></td>
			<td ALIGN="RIGHT"><%=FN(strDescChequeFecha,0)%></td>
		</tr>
		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG2")%>" class="Estilo13n">
			<td class="estilo_columna_individual">EFECTIVO</td>
			<td><input name="TX_EFECTIVO" type="text" value="<%=intValorEfectivo%>" size="10" maxlength="10" onchange="solonumero(TX_EFECTIVO);suma_total_general(1);"></td>
			<td ALIGN="RIGHT"><%=FN(intValorEfectivo,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intValorCajaEfectivo,0)%></td>
			<td ALIGN="RIGHT"><%=FN(strDescEfectivo,0)%></td>
		</tr>
		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG2")%>" class="Estilo13n">
			<td class="estilo_columna_individual">DEPOSITO</td>
			<td><input name="TX_DEPOSITO" type="text" value="<%=intValorDeposito%>" size="10" maxlength="10" onchange="solonumero(TX_DEPOSITO);suma_total_general(1);"></td>
			<td ALIGN="RIGHT"><%=FN(intValorDeposito,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intValorCajaDeposito,0)%></td>
			<td ALIGN="RIGHT"><%=FN(strDescDeposito,0)%></td>
		</tr>

		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG2")%>" class="Estilo13n">
			<td class="estilo_columna_individual">CUOTA</td>
			<td><input name="TX_CUOTA" type="text" value="<%=intValorCuota%>" size="10" maxlength="10" onchange="solonumero(TX_CUOTA);suma_total_general(1);"></td>
			<td ALIGN="RIGHT"><%=FN(intValorCuota,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intValorCajaCuota,0)%></td>
			<td ALIGN="RIGHT"><%=FN(strDescCuota,0)%></td>
		</tr>

		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG2")%>" class="Estilo13n">
			<td class="estilo_columna_individual">OTROS CANALES</td>
			<td><input name="TX_OCANALES" type="text" value="<%=intValorOCanales%>" size="10" maxlength="10" onchange="solonumero(TX_OCANALES);suma_total_general(1);"></td>
			<td ALIGN="RIGHT"><%=FN(intValorOCanales,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intValorCajaOCanales,0)%></td>
			<td ALIGN="RIGHT"><%=FN(strDescOCanales,0)%></td>
		</tr>

		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG2")%>" class="Estilo13n">
			<td class="estilo_columna_individual">TOTAL HABER EXISTENTE</td>
			<td><input name="TX_TOTALHABER" type="text" value="<%=intValorTotal%>" size="10" maxlength="10"></td>
			<td ALIGN="RIGHT"><%=FN(intValorTotal,0)%></td>
			<td ALIGN="RIGHT"><%=FN(intValorCajaTotal,0)%></td>
			<td ALIGN="RIGHT"><%=FN(strDescTotal,0)%></td>
		</tr>
		<tr height="20" bordercolor="#999999" class="Estilo13n">
			<td colspan=5 ALIGN="center"><input type="Button" class="fondo_boton_100" name="Submit" value="Cuadrar" onClick="Cuadrar();"></td>
		</tr>

    </table>
	</td>
	</tr>

</table>
	<%end if%>
	<%IF GRABA = "si" THEN
		dtmFechaCierre = inicio

		If Trim(intUsuario) <> "" Then
			strSqlQry="SELECT FECHA_CIERRE FROM CAJA_WEB_EMP_CIERRE WHERE COD_USUARIO = '" & intUsuario & "' AND FECHA_CIERRE IS NULL"
			intUsuarioCierre = intUsuario
		Else
			strSqlQry="SELECT FECHA_CIERRE FROM CAJA_WEB_EMP_CIERRE WHERE COD_USUARIO = '" & intCodUsuario & "' AND FECHA_CIERRE IS NULL"
			intUsuarioCierre = intCodUsuario
		End If

		'Response.write "<br>strSqlQry=" & strSqlQry
		'Response.End

		strSql="SELECT LOGIN FROM USUARIO WHERE ID_USUARIO = '" & intUsuarioCierre & "'"
		set rsNomUsuario=Conn.execute(strSql)

		If Not rsNomUsuario.EOF Then
			strNomUsuarioCierre = rsNomUsuario("LOGIN")
		Else
			strNomUsuarioCierre = ""
		End if

		'Response.End
		set rsCIERRE=Conn.execute(strSqlQry)

		If Not rsCIERRE.EOF  THEN

			If Trim(rsCIERRE("FECHA_CIERRE")) = "" or IsNull(rsCIERRE("FECHA_CIERRE")) then

				strSql = "UPDATE CAJA_WEB_EMP_CIERRE SET FECHA_CIERRE = '" & dtmFechaCierre & "', FECHA_HORA_CIERRE = GETDATE()"
				strSql = strSql & " WHERE COD_USUARIO = " & intUsuarioCierre
				strSql = strSql & " AND FECHA_CIERRE IS NULL AND CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "'"
				'Response.write (strsql)

				set rsGraba=Conn.execute(strsql)

				CD_CLIENTE = ValNulo(request("hd_valChequeDiaCli"),"N")
				CD_EMPRESA = ValNulo(request("hd_valChequeDiaEmp"),"N")
				CF_CLIENTE = ValNulo(request("hd_valChequeFechaCli"),"N")
				CF_EMPRESA = ValNulo(request("hd_valChequeFechaEmp"),"N")
				EF_CLIENTE = ValNulo(request("hd_valEfectivoCli"),"N")
				EF_EMPRESA = ValNulo(request("hd_valEfectivoEmp"),"N")
				DP_CLIENTE = ValNulo(request("hd_valDepositoCli"),"N")
				DP_EMPRESA = ValNulo(request("hd_valDepositoEmp"),"N")
				OC_CLIENTE = ValNulo(request("hd_valOtrosCanalesCli"),"N")
				OC_EMPRESA = ValNulo(request("hd_valOtrosCanalesEmp"),"N")
				CU_CLIENTE = ValNulo(request("hd_valCuotaCli"),"N")
				CU_EMPRESA = ValNulo(request("hd_valCuotaEmp"),"N")
				VV_CLIENTE = ValNulo(request("hd_valOCanalesCli"),"N")
				VV_EMPRESA = ValNulo(request("hd_valOCanalesEmp"),"N")

				strSql = "INSERT INTO CIERRE_USUARIO (ID_USUARIO, FECHA_CIERRE, CD_CLIENTE, CD_EMPRESA, CF_CLIENTE, CF_EMPRESA, EF_CLIENTE, EF_EMPRESA, DP_CLIENTE, DP_EMPRESA, CU_CLIENTE, CU_EMPRESA, OC_CLIENTE, OC_EMPRESA)"
				strSql = strSql & " VALUES(" & intUsuarioCierre & ",'" & dtmFechaCierre & "'," & CD_CLIENTE & "," & CD_EMPRESA & "," & CF_CLIENTE & "," & CF_EMPRESA & "," & EF_CLIENTE & "," & EF_EMPRESA & "," & DP_CLIENTE & "," & DP_EMPRESA & "," & CU_CLIENTE & "," & CU_EMPRESA & "," & OC_CLIENTE & "," & OC_EMPRESA & ")"
				'Response.write (strsql)

				set rsInserta=Conn.execute(strsql)
				%>
					<script>alert("Cierre de caja realizado correctamente para el usuario <%=strNomUsuarioCierre%>. No podra seguir ingresando pagos para esta fecha (<%=dtmFechaCierre%>)");</script>
				<%
			Else
				%>
					<script>alert("Cierre de caja ya fue realizado para el usuario <%=strNomUsuarioCierre%> y para esta fecha (<%=dtmFechaCierre%>). No puede volver a cerrar");</script>
				<%
			End if
		Else
			%>
				<script>alert("Cierre no realizado ya que no se a abierto la caja para el día <%=dtmFechaCierre%> y el usuario <%=strNomUsuarioCierre%>")</script>
			<%
		End if
		''fecha = dateadd("d",1,fecha)

		Response.Write ("<script language = ""Javascript"">" & vbCrlf)
		Response.Write (vbTab & "location.href='cerrar_caja_web.asp?rut=" & rut & "&tipo=1'" & vbCrlf)
		Response.Write ("</script>")
	  END IF
	%>
	</td>
   </tr>
  </table>

</form>
</body>
</html>

<script language="JavaScript" type="text/JavaScript">

function suma_total_general(){
	datos.TX_TOTALHABER.value = eval(datos.TX_CHEQUEDIA.value) + eval(datos.TX_CHEQUEFECHA.value) + eval(datos.TX_OCANALES.value) + eval(datos.TX_CUOTA.value) + eval(datos.TX_EFECTIVO.value) + eval(datos.TX_DEPOSITO.value) + eval(datos.TX_ASIGCAJA.value)
}

function solonumero(valor){
     //Compruebo si es un valor numérico

 if (valor.value.length >0){
    if (isNaN(valor.value))
    	{
            //entonces (no es numero) devuelvo el valor cadena vacia
            ////valor.value="0";
			//alert(valor.value)
			//valor.focus();
			return ""
    	}
    else
    	{
            //En caso contrario (Si era un número) devuelvo el valor
			valor.value
			return valor.value
    	}
	}
}

<% if resp="si" then %>

datos.TX_CHEQUEDIA.value = 0;
datos.TX_CHEQUEFECHA.value = 0;
datos.TX_CUOTA.value = 0;
datos.TX_OCANALES.value = 0;
datos.TX_EFECTIVO.value = 0;
datos.TX_DEPOSITO.value = 0;
datos.TX_TOTALHABER.value = 0;
datos.Cerrar.disabled = false;
datos.Cerrar.visible = true;

<% End if %>


</script>

