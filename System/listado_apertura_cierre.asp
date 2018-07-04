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

	Response.CodePage=65001
	Response.charset ="utf-8"
	
	Usuario_session =Session("intCodUsuario")

	AbrirSCG()

	strFecha = request("TX_FECINICIO")

	if Trim(strFecha) = "" Then
		strFecha = TraeFechaActual(Conn)
	End If

	If Request("CB_CLIENTE") = "" then

	strCOD_CLIENTE = session("ses_codcli")

	Else

	strCOD_CLIENTE = Request("CB_CLIENTE")

	End If

	intCodEstadoCob=Trim(Request("CB_ESTADOCOB"))
	If Trim(intCodEstadoCob) = "" Then intCodEstadoCob = "0"
	strEjecutivo=Request("CB_EJECUTIVO")


	'Response.write "strCOD_CLIENTE=" & strCOD_CLIENTE

	'hoy=date

%>

	<title>CRM FACTORING</title>


	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo28 {color: #FFFFFF}
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

	$('#TX_FECINICIO').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

	$(document).tooltip();

})

function Refrescar()
{
	resp='no'
	datos.action = "listado_apertura_cierre.asp?resp="+ resp +"";
	datos.submit();
}

function enviar(){
			datos.action = "man_Export.asp?archivo=1&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&CB_ASIGNACION=" + document.datos.CB_ASIGNACION.value + "&CH_ACTIVO=" + document.datos.CH_ACTIVO.checked;
			datos.submit()
}

function Ingresa()
{
	with( document.datos )
	{
		action = "listado_apertura_cierre.asp";
		submit();
	}
}

function Reversar(cod_pago)
{
	with( document.datos )
	{

	if (confirm("¿ Está seguro de reversar el pago ? El pago se eliminará completamente y la deuda será reversada, volviendo a su estado original antes del pago."))
		{
			action = "reversar_pago.asp?cod_pago=" + cod_pago;
			submit();
		}
	else
		alert("Reverso del pago cancelado");
	}
}

function Modificar(cod_pago)
{
	with( document.datos )
	{
		action = "modif_caja_web2.asp?strOrigen=listado_apertura_cierre.asp&cod_pago=" + cod_pago;
		submit();
	}
}

function envia()
{
	resp='si'
	document.datos.action = "listado_apertura_cierre.asp?strBuscar=S&resp="+ resp +"";
	document.datos.submit();
}
function exportar()
{
	document.datos.action = "exp_Listado_proceso_Fact.asp";
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
</script>


</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">


<form name="datos" method="post">
<div class="titulo_informe">APERTURA CIERRE</div>
<br>
<table width="90%" height="500" border="0" align="center">
			
  <tr height="20">
    <td style="vertical-align: top;">
		<table width="100%" border="0" class="estilo_columnas">
		<thead>
			  <tr height="20">
				<td>CLIENTE</td>
				<td>EJECUTIVO</td>
				<td>ETAPA COBRANZA</td>
				<td>FECHA</td>
				<td>&nbsp;</td>
			  </tr>
		</thead>
			  <tr>

				<td>

				<SELECT NAME="CB_CLIENTE" id="CB_CLIENTE" onChange="envia();">

					<option value="0">TODOS</option>
					<%

						ssql="SELECT COD_CLIENTE,RAZON_SOCIAL, NOMBRE_FANTASIA FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
								<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCOD_CLIENTE)=Trim(rsTemp("COD_CLIENTE")) then response.Write("Selected") End If%>><%=rsTemp("NOMBRE_FANTASIA")%></option>
									<%
								rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing

					%>
				</SELECT>
				</td>
				<td>

					<SELECT NAME="CB_EJECUTIVO" id="CB_EJECUTIVO" onChange="envia();">

						<option value="">TODOS</option>
						<%

							strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
							strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO"

							If strCOD_CLIENTE <> "0" then
								strSql= strSql & " AND UC.COD_CLIENTE IN (" & strCOD_CLIENTE & ")"
							End If

							strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"

							If strCOD_CLIENTE <> "0" then
								strSql= strSql & " AND UC.COD_CLIENTE = " & strCOD_CLIENTE
							End If

							set rsTemp= Conn.execute(strSql)
							if not rsTemp.eof then
								do until rsTemp.eof%>
									<option value="<%=rsTemp("ID_USUARIO")%>"<%if Trim(strEjecutivo)=Trim(rsTemp("ID_USUARIO")) then response.Write("Selected") End If%>><%=rsTemp("LOGIN")%></option>
										<%
									rsTemp.movenext
								loop
							end if
							rsTemp.close
							set rsTemp=nothing

						%>
					</SELECT>
				</td>
				<td>
					<SELECT NAME="CB_ESTADOCOB" id="CB_ESTADOCOB" onChange="envia();">

						<option value="0">TODOS</option>
						<%

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

						%>
					</SELECT>
				</td>
				<td>	
					<input name="TX_FECINICIO" id="TX_FECINICIO" readonly="true" type="text" value="<%=strFecha%>" size="10" maxlength="10">
				</td>

				<td align = "right" >
				<input type="Button" name="Submit" class="fondo_boton_100" value="Ver" onClick="envia();">
				</td>
			  </tr>
		</table>
    </td>
   </tr>

	<%

	strSql = " SELECT COD_CLIENTE FROM CLIENTE WHERE COD_CLIENTE IN"
	strSql= strSql & " (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC "

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

	If strCOD_CLIENTE = "0" then
		strCOD_CLIENTE = mid(strTClienteUsu,3,len(strTClienteUsu))
	End If

	'Response.write "strTClienteUsu = " & strTClienteUsu

	'Response.write "strCOD_CLIENTE = " & strCOD_CLIENTE

	strSql = "proc_Apertura_Cierre_RPT '" & strCOD_CLIENTE & "','" & strFecha & "','" & strEjecutivo & "'," & intCodEstadoCob

	'Response.write "strSql = " & strSql
	'Response.write "strSql = " & strEstado

	'Response.End

	set rsDet=Conn.execute(strSql)
		
	if not rsDet.eof then
		
		strHoraMax = rsDet("HORA_CIERRE")

		strHoraMax = mid(strHoraMax,1,len(strHoraMax)-3)
		
%>

   <tr>
	<td style="vertical-align: top;">
	<table width="100%" border="0" class="intercalado" style="width:100%;">
		<thead>

	    <tr >
			<td class="subtitulo_informe" colspan = "2">> INFORME</td>
			<td class="subtitulo_informe" colspan = "11">&nbsp;</td>
			<td align = "center" colspan = "1" height = "20"><%=strHoraMax%></td>
	    </tr>
  
		<tr bordercolor="#999999">

			<td>&nbsp;</td>
			<td>CLIENTE</td>
			<td>ESTADO COBRANZA</td>
			<td>EJECUTIVO</td>
			<td>FECHA</td>
			<td>HORA CIE.</td>
			<td>RUT AP.</td>
			<td>DOC AP.</td>
			<td>MONTO AP.</td>

			<td>RUT CIE.</td>
			<td>DOC CIE.</td>
			<td>MONTO CIE.</td>
			<td>CAS. GEST.</td>
			<td>MONTO GEST.</td>

		</tr>
		</thead>
		<tbody>	
<%

			intReg = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>
			
			<tr>

			<%intMontoDif = ValNulo(rsDet("MONTO_APERTURA"),"N") - ValNulo(rsDet("MONTO_CIERRE"),"N") %>

			<td Width="15"><%=intReg%></td>
			<td><%=Mid(rsDet("COD_CLIENTE"),1,30)%></td>
			<td><%=Mid(rsDet("ESTADO_COBRANZA"),1,15)%></td>
			<td><%=Mid(rsDet("GESTIONADOR"),1,15)%></td>
			<td><%=Mid(rsDet("FECHA"),1,25)%></td>
			<td align="right"><%=rsDet("HORA_CIERRE")%></td>
			<td align="right"><%=FN(rsDet("RUT_APERTURA"),0)%></td>
			<td align="right"><%=FN(rsDet("DOCTO_APERTURA"),0)%></td>
			<td align="right"><%=FN(rsDet("MONTO_APERTURA"),0)%></td>

			<td align="right"><%=FN(rsDet("RUT_CIERRE"),0)%></td>
			<td align="right"><%=FN(rsDet("DOCTO_CIERRE"),0)%></td>
			<td align="right"><%=FN(rsDet("MONTO_CIERRE"),0)%></td>
			<td align="right"><%=FN(rsDet("RUT_APERTURA")-rsDet("RUT_CIERRE"),0)%></td>
			<td align="right"><%=FN(intMontoDif,0)%></td>

			</tr>
		
				<%

				intTotalRutAp = intTotalRutAp + ValNulo(rsDet("RUT_APERTURA"),"N")
				intTotalDoctoAp = intTotalDoctoAp + ValNulo(rsDet("DOCTO_APERTURA"),"N")
				intTotalMontoAp = intTotalMontoAp + ValNulo(rsDet("MONTO_APERTURA"),"N")
				intTotalRutCie = intTotalRutCie + ValNulo(rsDet("RUT_CIERRE"),"N")
				intTotalDoctoCie = intTotalDoctoCie + ValNulo(rsDet("DOCTO_CIERRE"),"N")
				intTotalMontoCie = intTotalMontoCie + ValNulo(rsDet("MONTO_CIERRE"),"N")
				intTotalDifCasos = intTotalRutAp - intTotalRutCie
				intTotalDifMonto = intTotalMontoAp - intTotalMontoCie

				rsDet.movenext
			loop

	%>
		</tbody>
		<thead>
			<tr  class="totales">
			
				<td colspan="2" >TOTALES</td>
				<td colspan="4" ALIGN="CENTER">&nbsp;</td>
				<td ALIGN="RIGHT"><%=FN(intTotalRutAp,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalDoctoAp,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalMontoAp,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalRutCie,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalDoctoCie,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalMontoCie,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalDifCasos,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalDifMonto,0)%></td>

			</tr>
		</thead>
	
	<%Else%>
	
		<tr class="estilo_columna_individual">
			<td Colspan = "1">&nbsp;</td>
		</tr>
			
		<tr >
			<td ALIGN="CENTER" Colspan = "1">NO HAY APERTURAS INGRESADAS SEGUN PARAMETROS DE BUSQUEDA</td>
		</tr>

		<tr class="estilo_columna_individual">
			<td Colspan = "1">&nbsp;</td>
		</tr>	
	
	<%end if%>

	</table>
	</td>
   </tr>
  </table>

</form>


</body>
</html>

