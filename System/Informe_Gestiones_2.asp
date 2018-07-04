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

	<!--#include file="../lib/comunes/js_css/top_tooltip.inc" -->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	strFechaInicio = request("TX_FECINICIO")
	strFechaTermino = request("TX_FECTERMINO")
	
	If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

		strEjeAsig = Request("CB_EJECUTIVO")

	Else
	
		strEjeAsig =  session("session_idusuario")
	
	End If
	
	

	AbrirSCG()

	if Trim(strFechaInicio) = "" Then
		strFechaInicio = TraeFechaActual(Conn)
	End If

	if Trim(strFechaTermino) = "" Then
		strFechaTermino = TraeFechaActual(Conn)

	Else strFechaTermino = strFechaTermino

	End If

	CerrarSCG()

	If Trim(Request("CB_ESTADOCOB")) = "" Then intCodEstadoCob = "0" else intCodEstadoCob = Trim(Request("CB_ESTADOCOB")) End If

	'Response.write "strEjeAsig = " & strEjeAsig

		AbrirSCG()

	strSql = " SELECT COD_CLIENTE FROM CLIENTE WHERE COD_CLIENTE IN"
	strSql= strSql & " (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"

	set rsTemp= Conn.execute(strSql)
	'Response.write "strSql = " & strSql

	strTClienteUsu = "0"

	if not rsTemp.eof then
			do while not rsTemp.eof

			strClientesUsu = rsTemp("COD_CLIENTE")

			strTClienteUsu = strTClienteUsu + "," + strClientesUsu

			rsTemp.movenext
			Loop
		rsTemp.close
		set rsTemp=nothing
	End If

	CerrarSCG()

	If Request("CB_CLIENTE") <> "0" then
		strCodCliente = session("ses_codcli")
	Else
		strCodCliente = mid(strTClienteUsu,3,len(strTClienteUsu))
	End If

'--Calcula Objetos relacionados al tipo de cobranza (Interna, Externa), CB_COBRANZA Y CB_EJECUTIVO--'

strCobranza = Request("CB_COBRANZA")

abrirscg()

		strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
		strSql = strSql & " FROM CLIENTE CL"
		strSql = strSql & " WHERE CL.COD_CLIENTE IN ('" & strCodCliente & "')"

		set RsCli=Conn.execute(strSql)
		If not RsCli.eof then
			intUsaCobInterna = RsCli("USA_COB_INTERNA")
		End if
		RsCli.close
		set RsCli=nothing

cerrarscg()

intVerCobExt = "1"
intVerEjecutivos = "1"

If TraeSiNo(session("perfil_emp")) = "Si" and strCobranza = "" and intUsaCobInterna = "1" Then

	strCobranza="INTERNA"

ElseIf TraeSiNo(session("perfil_emp")) = "No" and strCobranza = "" then

	strCobranza="EXTERNA"

End If

If TraeSiNo(session("perfil_emp")) = "Si" Then

	intVerEjecutivos="0"
	intVerCobExt = "1"

End If

If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

	sinCbUsario="0"

End If

'---Fin codigo tipo de cobranza---'
'Response.write "strCobranza = " & strCobranza
%>

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

	$('#TX_FECTERMINO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	$('#TX_FECINICIO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
  
})

function envia()
{
	resp='si'
	document.datos.action = "Informe_Gestiones_2.asp?strBuscar=S&resp="+ resp +"";
	document.datos.submit();
}

</script>

<style type="text/css">
<!--
.Estilo13 {color: #FFFFFF}
.Estilo28 {color: #FFFFFF}
.Estilo27 {color: #FFFFFF}
-->

.uno a {
	text-align:center;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	text-decoration: none;
	color: #FFFFFF;
}
.uno a:hover {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	text-decoration: none;
	color: #FFFFFF;
}
.contenedor{
	width: 90%;
	border:2px solid #ccc;
	margin:0 auto;
}
</style>
</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<form name="datos" method="post">
<div class="titulo_informe">INFORME GESTIONES CATEGORIZADAS</div>
<br>
<div class="contenedor">

			<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
				<thead>
				  <tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					<td>CLIENTE</td>
					<td>COBRANZA</td>
					<td>ETAPA COBRANZA</td>
					<td>FECHA INICIO</td>
					<td>FECHA TERMINO</td>

				<% If sinCbUsario = "0" Then %>
					<td>EJECUTIVO</td>
				<%Else%>
					<td>&nbsp;</td>
				<%End If%>

					<td>&nbsp;</td>

				  </tr>
				</thead>
				  <tr bordercolor="#999999" class="Estilo8">

					<td>

					<SELECT NAME="CB_CLIENTE" id="CB_CLIENTE" onChange="envia();">

						<option value="0">TODOS</option>
						<%
						AbrirSCG()

							ssql="SELECT COD_CLIENTE,RAZON_SOCIAL, NOMBRE_FANTASIA FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
							set rsTemp= Conn.execute(ssql)
							if not rsTemp.eof then
								do until rsTemp.eof%>
									<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCodcliente)=Trim(rsTemp("COD_CLIENTE")) then response.Write("Selected") End If%>><%=rsTemp("NOMBRE_FANTASIA")%></option>
										<%
									rsTemp.movenext
								loop
							end if
							rsTemp.close
							set rsTemp=nothing

						CerrarSCG()
						%>
					</SELECT>
					</td>

					<td>
						<select name="CB_COBRANZA" <%If sinCbUsario = "0" then%> onChange="CargaUsuarios(this.value,CB_CLIENTE.value);" <%End If%> >

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

					<td>
						<SELECT NAME="CB_ESTADOCOB" id="CB_ESTADOCOB" onChange="envia();">

							<option value="0">TODOS</option>
							<%
						AbrirSCG()

								ssql="SELECT COD_ESTADO_COBRANZA, NOM_ESTADO_COBRANZA FROM ESTADO_COBRANZA"
								set rsTemp= Conn.execute(ssql)
								if not rsTemp.eof then
									do until rsTemp.eof%>
										<option value="<%=rsTemp("COD_ESTADO_COBRANZA")%>"<%if Trim(intCodEstadoCob)=Trim(rsTemp("COD_ESTADO_COBRANZA")) then response.Write("Selected") End If%>><%=rsTemp("NOM_ESTADO_COBRANZA")%></option>
											<%
										rsTemp.movenext
									loop
								end if
								rsTemp.close
								set rsTemp=nothing

						CerrarSCG()
							%>
						</SELECT>
					</td>


					<td>	<input name="TX_FECINICIO"  id="TX_FECINICIO"  readonly="true" type="text" value="<%=strFechaInicio%>" size="10" maxlength="10">
					</td>

					<td>	<input name="TX_FECTERMINO" id="TX_FECTERMINO" readonly="true" type="text" value="<%=strFechaTermino%>" size="10" maxlength="10">
					</td>


				<% If sinCbUsario="0" Then %>

					<td>
						<select name="CB_EJECUTIVO"  id="CB_EJECUTIVO" >
						</select>
					</td>

				<% End If %>

					<td align = "CENTER" >
						<input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();">
					</td>

				  </tr>
			</table>
	    </td>
	   </tr>


	   <tr>
		<td style="vertical-align: top;">
		<table width="100%" border="0" bordercolor="#000000" class="intercalado" style="width:100%;">

		<%

		AbrirSCG()

		strSql = "proc_Inf_Gestiones_Rpt '" & strCodcliente & "','" & strFechaInicio & "','" & strFechaTermino & "','" & strEjeAsig & "'," & intCodEstadoCob & ",'" & strCobranza & "'"

		'Response.write "strSql = " & strSql

		strTClienteSel = "0"

		if strSql <> "" then
			set rsDet=Conn.execute(strSql)

			if not rsDet.eof then

			strHoraMax = rsDet("ULT_HORA")

			strHoraMax = mid(strHoraMax,1,len(strHoraMax)-3)%>
			<thead>
			<tr >
				<td class="subtitulo_informe">> CASOS</td>
				<td class="subtitulo_informe" COLSPAN = "9">&nbsp;</td>
				<td align = "center" COLSPAN = "1" height = "20"><%=strHoraMax%></td>
			</tr>

			<tr>

				<td>&nbsp;</td>
				<td>CLIENTE</td>
				<td>EJECUTIVO</td>
				<td>COMPROMISOS</td>
				<td>TITULAR</td>
				<td>TERCERO</td>
				<td>NO COMUNICA</td>
				<td>ENV. CORREO</td>
				<td>BUSQUEDA SR</td>
				<td>OTRO</td>
				<td>TOTAL</td>

			</tr>
			</thead>
			<tbody>
			<%	intReg = 0
				do while not rsDet.eof
					intReg = intReg + 1

					strSql1 = "SELECT LOGIN FROM USUARIO WHERE ID_USUARIO = " & rsDet("GESTIONADOR")

					'Response.write "strSql1 = " & strSql1

					AbrirSCG1()

					set rsUsu=Conn1.execute(strSql1)

					strUsuario = rsUsu("LOGIN")

					CerrarSCG1()

					strClienteSel = rsDet("COD_CLIENTE")

					strTClienteSel = strTClienteSel + "," + strClienteSel

					intTotalCasosCP = intTotalCasosCP + ValNulo(rsDet("CASOS_COMPROMISO"),"N")
					intTotalCasosTT = intTotalCasosTT + ValNulo(rsDet("CASOS_TITULAR"),"N")
					intTotalCasosTE = intTotalCasosTE + ValNulo(rsDet("CASOS_TERCERO"),"N")
					intTotalCasosNC = intTotalCasosNC + ValNulo(rsDet("CASOS_NO_COMUNICA"),"N")
					intTotalCasosEC = intTotalCasosEC + ValNulo(rsDet("CASOS_ENV_CORREO"),"N")
					intTotalCasosBS = intTotalCasosBS + ValNulo(rsDet("CASOS_BUS_SIN_RES"),"N")
					intTotalCasosOtro = intTotalCasosOtro + ValNulo(rsDet("CASOS_OTRO"),"N")
					intTotal = intTotalCasosCP+intTotalCasosTT+intTotalCasosTE+intTotalCasosNC+intTotalCasosEC+intTotalCasosBS+intTotalCasosOtro

					%>

					<tr >

						<td><%=intReg%></td>
						<td><%=Mid(rsDet("DESCCLIENTE"),1,30)%></td>
						<td><%=Mid(strUsuario,1,15)%></td>

						<td align="right" width = "80">
							<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=1&strLogin=<%=rsDet("GESTIONADOR")%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=rsDet("COD_CLIENTE")%>">
							<%=rsDet("CASOS_COMPROMISO")%>
							</A>
						</td>

						<td align="right" width = "80">
							<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=2&strLogin=<%=rsDet("GESTIONADOR")%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=rsDet("COD_CLIENTE")%>">
							<%=rsDet("CASOS_TITULAR")%>
							</A>
						</td>

						<td align="right" width = "80">
							<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=3&strLogin=<%=rsDet("GESTIONADOR")%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=rsDet("COD_CLIENTE")%>">
							<%=rsDet("CASOS_TERCERO")%>
							</A>
						</td>

						<td align="right" width = "80">
							<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=4&strLogin=<%=rsDet("GESTIONADOR")%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=rsDet("COD_CLIENTE")%>">
							<%=rsDet("CASOS_NO_COMUNICA")%>
							</A>
						</td>

						<td align="right" width = "80">
							<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=5&strLogin=<%=rsDet("GESTIONADOR")%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=rsDet("COD_CLIENTE")%>">
							<%=rsDet("CASOS_ENV_CORREO")%>
							</A>
						</td>

						<td align="right" width = "80">
							<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=6&strLogin=<%=rsDet("GESTIONADOR")%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=rsDet("COD_CLIENTE")%>">
							<%=rsDet("CASOS_BUS_SIN_RES")%>
							</A>
						</td>

						<td align="right" width = "80">
							<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=9&strLogin=<%=rsDet("GESTIONADOR")%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=rsDet("COD_CLIENTE")%>">
							<%=rsDet("CASOS_OTRO")%>
							</A>
						</td>

						<td align="right" width = "80">
							<A HREF="Detalle_informe_gestiones_2.asp?strLogin=<%=rsDet("GESTIONADOR")%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=rsDet("COD_CLIENTE")%>">
							<%=rsDet("CASOS_COMPROMISO")+rsDet("CASOS_TITULAR")+rsDet("CASOS_TERCERO")+rsDet("CASOS_NO_COMUNICA")+rsDet("CASOS_ENV_CORREO")+rsDet("CASOS_BUS_SIN_RES")+rsDet("CASOS_OTRO")%></td>
						</td>

					</tr>

					<%
					rsDet.movenext
				loop
				rsDet.close
				set rsDet=nothing

		CerrarSCG()

			strTClienteSel = mid(strTClienteSel,3,len(strTClienteSel))

		%>
			</tbody>
			<thead>
				<tr class="totales">

				<td colspan="3">TOTAL CASOS</td>


				<td ALIGN="RIGHT" >
				<div class="uno">
					<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=1&strLogin=<%=strEjeAsig%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strTClienteSel%>">
					<%=FN(intTotalCasosCP,0)%>
				<div>
				</td>
				<td ALIGN="RIGHT">
				<div class="uno">
					<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=2&strLogin=<%=strEjeAsig%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strTClienteSel%>">
					<%=FN(intTotalCasosTT,0)%>
				<div>
				</td>

				<td ALIGN="RIGHT" >
				<div class="uno">
					<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=3&strLogin=<%=strEjeAsig%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strTClienteSel%>">
					<%=FN(intTotalCasosTE,0)%>
				<div>
				</td>

				<td ALIGN="RIGHT" >
				<div class="uno">
					<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=4&strLogin=<%=strEjeAsig%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strTClienteSel%>">
					<%=FN(intTotalCasosNC,0)%>
				<div>
				</td>

				<td ALIGN="RIGHT" >
				<div class="uno">
					<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=5&strLogin=<%=strEjeAsig%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strTClienteSel%>">
					<%=FN(intTotalCasosEC,0)%>
				<div>
				</td>

				<td ALIGN="RIGHT">
				<div class="uno">
					<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=6&strLogin=<%=strEjeAsig%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strTClienteSel%>">
					<%=FN(intTotalCasosBS,0)%>
				<div>
				</td>

				<td ALIGN="RIGHT" >
				<div class="uno">
					<A HREF="Detalle_informe_gestiones_2.asp?intCodTipoGes=9&strLogin=<%=strEjeAsig%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strTClienteSel%>">
					<%=FN(intTotalCasosOtro,0)%>
				<div>
				</td>

				<td ALIGN="RIGHT" >
				<div class="uno">
					<A HREF="Detalle_informe_gestiones_2.asp?strFechaInicio=<%=strFechaInicio%>&strLogin=<%=strEjeAsig%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strTClienteSel%>">
					<%=FN(intTotal,0)%>
				<div>
				</td>

				</tr>
			</thead>
		<%Else%>
				<thead class="estilo_columnas">
				<thead>
				<tr >
					<td height = "20">&nbsp;</td>
				</tr>
				<tr >
					<td ALIGN="CENTER" height = "20">NO EXISTEN GESTIONES SEGUN LOS PARAMETROS SELECCIONADOS</td>
				</tr>
				<tr >
					<td height = "20">&nbsp;</td>
				</tr>
				</thead>
				</thead>
		<%End if%>

		</table>
		<br>
		<table width="100%" border="0" bordercolor="#999999" align="center" class="intercalado" style="width:100%;">
		
	<%
		AbrirSCG()
			set rsDet=Conn.execute(strSql)%>

			<%if not rsDet.eof then%>
				<thead>	
				<tr>
					<td COLSPAN = "11" class="subtitulo_informe">> DOCUMENTOS</td>
				</tr>
				<tr>

					<td>&nbsp;</td>
					<td>CLIENTE</td>
					<td>EJECUTIVO</td>
					<td>COMPROMISOS</td>
					<td>TITULAR</td>
					<td>TERCERO</td>
					<td>NO COMUNICA</td>
					<td>ENV. CORREO</td>
					<td>BUSQUEDA SR</td>
					<td>OTRO</td>
					<td>TOTAL</td>

				</tr>
				</thead>
				<tbody>
				<%intReg = 0
				  do while not rsDet.eof
					intReg = intReg + 1

						strSql1 = "SELECT LOGIN FROM USUARIO WHERE ID_USUARIO = " & rsDet("GESTIONADOR")

						'Response.write "strSql1 = " & strSql1

						AbrirSCG1()

						set rsUsu=Conn1.execute(strSql1)

						strUsuario = rsUsu("LOGIN")

						CerrarSCG1()

					%>
					<tr >

						<td><%=intReg%></td>
						<td><%=Mid(rsDet("DESCCLIENTE"),1,30)%></td>
						<td><%=Mid(strUsuario,1,15)%></td>
						<td align="right" width = "80"><%=rsDet("DOC_COMPROMISO")%></td>
						<td align="right" width = "80"><%=rsDet("DOC_TITULAR")%></td>
						<td align="right" width = "80"><%=FN(rsDet("DOC_TERCERO"),0)%></td>
						<td align="right" width = "80"><%=FN(rsDet("DOC_NO_COMUNICA"),0)%></td>
						<td align="right" width = "80"><%=FN(rsDet("DOC_ENV_CORREO"),0)%></td>
						<td align="right" width = "80"><%=FN(rsDet("DOC_BUS_SIN_RES"),0)%></td>
						<td align="right" width = "80"><%=FN(rsDet("DOC_OTRO"),0)%></td>
						<td align="right" width = "80"><%=rsDet("DOC_COMPROMISO")+rsDet("DOC_TITULAR")+rsDet("DOC_TERCERO")+rsDet("DOC_NO_COMUNICA")+rsDet("DOC_ENV_CORREO")+rsDet("DOC_BUS_SIN_RES")+rsDet("DOC_OTRO")%></td>

					</tr>
					<%

					intTotalDocCP = intTotalDocCP + ValNulo(rsDet("DOC_COMPROMISO"),"N")
					intTotalDocTT = intTotalDocTT + ValNulo(rsDet("DOC_TITULAR"),"N")
					intTotalDocTE = intTotalDocTE + ValNulo(rsDet("DOC_TERCERO"),"N")
					intTotalDocNC = intTotalDocNC + ValNulo(rsDet("DOC_NO_COMUNICA"),"N")
					intTotalDocEC = intTotalDocEC + ValNulo(rsDet("DOC_ENV_CORREO"),"N")
					intTotalDocBS = intTotalDocBS + ValNulo(rsDet("DOC_BUS_SIN_RES"),"N")
					intTotalDocOtro = intTotalDocOtro + ValNulo(rsDet("DOC_OTRO"),"N")
					intTotal = intTotalDocCP+intTotalDocTT+intTotalDocTE+intTotalDocNC+intTotalDocEC+intTotalDocBS+intTotalDocOtro

					rsDet.movenext
				loop
				rsDet.close
				set rsDet=nothing

		CerrarSCG()

		%>
				</tbody>
				<thead>
				<tr class="totales">

				<td colspan="3" >TOTAL DOC.</td>
				<td ALIGN="RIGHT" ><%=FN(intTotalDocCP,0)%></td>
				<td ALIGN="RIGHT" ><%=FN(intTotalDocTT,0)%></td>
				<td ALIGN="RIGHT" ><%=FN(intTotalDocTE,0)%></td>
				<td ALIGN="RIGHT" ><%=FN(intTotalDocNC,0)%></td>
				<td ALIGN="RIGHT" ><%=FN(intTotalDocEC,0)%></td>
				<td ALIGN="RIGHT" ><%=FN(intTotalDocBS,0)%></td>
				<td ALIGN="RIGHT" ><%=FN(intTotalDocOtro,0)%></td>
				<td ALIGN="RIGHT" ><%=FN(intTotal,0)%></td>

				</tr>
				</thead>

		<%
			End if

		End if%>

		</table>

		<table width="100%" border="1" cellSpacing="0" cellPadding="0">

				<tr>
					<td colspan = "10">Los Casos se cuentan según un criterio de importancia de izquierda a derecha, COMPROMISOS, TITULAR, TERCERO, NO COMUNICA, ENV. CORREO, BUSQUEDA SR, por ejemplo si un caso tiene gestión compromiso y a la vez titular, se contará como compromiso</td>
				</tr>

				<tr>
					<td colspan = "10">Las gestiónes son contadas por deudor independientemente de que el caso se gestiones varias veces la cuenta será = 1</td>
				</tr>
		</table>

		<br>

		<table width="100%" border="1" cellSpacing="0" cellPadding="0">

				<tr >
					<td colspan = "10">COMPROMISOS:</td>
					<td colspan = "10">Gestión compromisos de pago, compromisos de ruta y de retiro de pago</td>
				</tr>
				<tr >
					<td colspan = "10">TITULAR:</td>
					<td colspan = "10">Gestión con contactos titular a través de una gestión telefónica o de recepción  mail</td>
				</tr>
				<tr >
					<td colspan = "10">TERCERO:</td>
					<td colspan = "10">Gestión con contacto tercero a través de una gestión telefónica</td>
				</tr>
				<tr >
					<td colspan = "10">NO COMUNICA:</td>
					<td colspan = "10">Gestión sin comunicación; NO CONTESTA, OCUPADO, FUERA DE SERVICIO, entre otras.</td>
				</tr>
				<tr >
					<td colspan = "10">ENVIO CORREO:</td>
					<td colspan = "10">Gestión de envío de correo electrónico</td>
				</tr>
				<tr >
					<td colspan = "10">BUSQUEDA SR:</td>
					<td colspan = "10">Gestión de búsqueda de datos de contactabilidad sin resultados</td>
				</tr>
		</table>

</div>
</form>

</body>
</html>


<script type="text/javascript">

function CargaUsuarios(subCat,cat)
{
	//alert(subCat);
	//alert(cat);

	var comboBox = document.getElementById('CB_EJECUTIVO');
	switch (cat)
	{
		<%
		  AbrirSCG()
			strSql="SELECT COD_CLIENTE FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE"
			set rsGestCat=Conn.execute(strSql)
			Do While not rsGestCat.eof
		%>
		case '<%=rsGestCat("COD_CLIENTE")%>':

			comboBox.options.length = 0;

				if (subCat=='INTERNA') {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

					strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
					strSql = strSql & " AND U.PERFIL_EMP=1"

					'Response.write "<br>strSql=" & strSql

					set rsUsuario=Conn2.execute(strSql)
					If Not rsUsuario.Eof Then
						Do While Not rsUsuario.Eof
							%>
								var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
								comboBox.options[comboBox.options.length] = newOption;
							<%
							rsUsuario.movenext
						Loop
					Else
					%>
						var newOption = new Option('SIN USUARIO', '');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					CerrarSCG2()
					%>
					break;
				}

				if (subCat=='EXTERNA' && (<%=intVerEjecutivos%>=='1')) {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

					strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
					strSql = strSql & " AND U.PERFIL_EMP=0"

					'Response.write "<br>strSql=" & strSql

					set rsUsuario=Conn2.execute(strSql)
					If Not rsUsuario.Eof Then
						Do While Not rsUsuario.Eof
							%>
								var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
								comboBox.options[comboBox.options.length] = newOption;
							<%
							rsUsuario.movenext
						Loop
					Else
					%>
						var newOption = new Option('SIN USUARIO', '');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					CerrarSCG2()
					%>
					break;
				}
				else if ((subCat=='EXTERNA') && (<%=intVerEjecutivos%>=='0')) {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					break;
				}
				else {
					var newOption = new Option('TODOS', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
					strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = " & rsGestCat("COD_CLIENTE")

					strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
					''strSql = strSql & " AND U.PERFIL_EMP=0"

					'Response.write "<br>strSql=" & strSql

					set rsUsuario=Conn2.execute(strSql)
					If Not rsUsuario.Eof Then
						Do While Not rsUsuario.Eof
							%>
								var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
								comboBox.options[comboBox.options.length] = newOption;
							<%
							rsUsuario.movenext
						Loop
					Else
					%>
						var newOption = new Option('SIN USUARIO', '');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					CerrarSCG2()
					%>
					break;
				}

		<%
		  	rsGestCat.movenext
		  	Loop
		  	rsGestCat.close
		  	set rsGestCat=nothing
			CerrarSCG()
		%>
	}

}

function InicializaInforme()
{
		var comboBox = document.getElementById('CB_EJECUTIVO');
		comboBox.options.length = 0;
		var newOption = new Option('TODOS','');
		comboBox.options[comboBox.options.length] = newOption;
}

<%If sinCbUsario = "0" then%>
CargaUsuarios('<%=strCobranza%>','<%=strCodCliente%>');
<%End If%>

<%If strEjeAsig <> "" then%>
datos.CB_EJECUTIVO.value='<%=strEjeAsig%>';
<%End If%>
</script>


