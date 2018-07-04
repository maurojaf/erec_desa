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

	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	Usuario_session = Session("intCodUsuario")

	AbrirSCG()

	termino = request("termino")
	inicio = request("inicio")

	resp = request("resp")
	if Trim(inicio) = "" Then
		strMesActual = Month(TraeFechaActual(Conn))
		strAnoActual = Cdbl(Year(TraeFechaActual(Conn)))

		If strMesActual = 1 Then strAnoActual = strAnoActual - 1
		If strMesActual = 1 Then strMesActual = 12
		strMesActual = strMesActual - 1

		if Len(strMesActual) = 1 Then strMesActual = "0" & strMesActual

		If Trim(inicio) = "" Then inicio = "01/" & strMesActual & "/" & strAnoActual

	End If


	if Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If



	strCliente = REQUEST("CB_CLIENTE")
	strTipo = REQUEST("CB_TIPO")
	dtmFechaProc = REQUEST("CB_FECHA")


	intCOD_CLIENTE = session("ses_codcli")

%>


<LINK href="../css/isk_style.css" type=text/css rel=stylesheet>
<title>CRM FACTORING</title>


<style type="text/css">
<!--
.Estilo13 {color: #FFFFFF}
.Estilo28 {color: #FFFFFF}
.Estilo27 {color: #FFFFFF}
-->


</style>


<script language="JavaScript" src="../javascripts/validaciones.js"></script>
<script src="../javascripts/SelCombox.js"></script>
<script src="../javascripts/OpenWindow.js"></script>




<script language="JavaScript " type="text/JavaScript">

function Refrescar()
{
	resp='no'
	datos.action = "listado_JIC.asp?resp="+ resp +"";
	datos.submit();
}


function envia()
{


if (datos.CB_TIPO.value=='0') {
	alert('Debe seleccionar un tipo de informe');
	}
else
	{
			resp='si'
			document.datos.action = "listado_JIC.asp?strBuscar=S&resp="+ resp +"";
			document.datos.submit();
	}


}

function imprimir()
{
	datos.action = "imprime_comprobantes.asp";
	datos.submit();
}


</script>

</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">


<form name="datos" method="post">
<div class="titulo_informe">INFORME ESPECIAL</div>	
<table width="90%" height="500" border="0" align="center">
<tr height="20">
    <td style="vertical-align: top;">
		<table width="100%" border="0" class="estilo_columnas">
		<thead>	
			  <tr height="20" >
				<td>TIPO</td>
				<td>FECHA 1</td>
				<td>FECHA 2</td>
				<td>&nbsp;</td>
			  </tr>
		</thead>			
			  <tr>
				<td>
					<SELECT NAME="CB_TIPO" id="CB_TIPO">
						<option value="0" <%If Trim(strTipo)="0" Then Response.write "SELECTED"%>>SELECCIONAR</option>
						<option value="1" <%If Trim(strTipo)="1" Then Response.write "SELECTED"%>>DOCUMENTOS</option>
						<option value="2" <%If Trim(strTipo)="2" Then Response.write "SELECTED"%>>MONTO</option>
					</SELECT>
				</td>

				<td>
					<input name="inicio" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
					<a href="javascript:showCal('Calendar7');"><img src="../imagenes/calendario.gif" border="0"></a>
				</td>
				<td>
					<input name="termino" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
				    <a href="javascript:showCal('Calendar6');"><img src="../imagenes/calendario.gif" border="0"></a>
				</td>

				<td align="CENTER">
					<input Name="SubmitButton" Value="Listar" Type="BUTTON" onClick="envia();">
				</td>
			  </tr>
		</table>
    </td>
   </tr>


   <tr>
	<td style="vertical-align: top;">

	<% If resp="si" then %>
	<table width="100%" border="0" bordercolor="#000000">

			<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

				<td>&nbsp;</td>
				<td>CLIENTE</td>
				<td>TOTAL</td>
				<td>NORM.</td>
				<td>DESAS.</td>
				<td>NO ASIG.</td>
				<td>ACTIVA</td>
				<td>ASIG.PERIODO</td>
				<td>NORM.PERIODO</td>
				<td>DESAS.PERIODO</td>
				<td>NO ASIGN.PERIODO</td>
				<td>SALDO ACTIVO</td>
				<td>VAR.PORC.</td>
		</tr>

	<%

		If Trim(strTipo) = "1" Then strTipoInf = "DOC"
		If Trim(strTipo) = "2" Then strTipoInf = "MONTO"

		strSql = "EXEC proc_listado_JIC '" & inicio & "','" & termino & "','DOC'"
		'Response.write "<br>strSql = " & strSql

		set rsDet=Conn.execute(strSql)

			intReg = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>
				<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">


					<td ALIGN="RIGHT"><%=intReg%></td>
					<td ALIGN="LEFT"><%=Mid(rsDet("CLIENTE"),1,40)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("TOTAL"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("NORMALIZADO"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("DESASIGNADO"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("NO_ASIGNABLE"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("ACTIVA"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("ASIGNACION_PERIODO"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("NORMALIZADO_PERIODO"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("DESASIGNACION_PERIODO"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("NO_ASIGNABLE_PERIODO"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("SALDO_ACTIVO"),0)%></td>
					<td ALIGN="RIGHT"><%=FN(rsDet("VAR_PORC"),0)%></td>

				</tr>
				<%
				rsDet.movenext
			loop
	%>

	</table>


	<br>




	<table width="100%" border="0" bordercolor="#000000">

				<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

					<td>&nbsp;</td>
					<td>CLIENTE</td>
					<td>TOTAL</td>
					<td>NORM.</td>
					<td>DESAS.</td>
					<td>NO ASIG.</td>
					<td>ACTIVA</td>
					<td>ASIG.PERIODO</td>
					<td>NORM.PERIODO</td>
					<td>DESAS.PERIODO</td>
					<td>NO ASIGN.PERIODO</td>
					<td>SALDO ACTIVO</td>
					<td>VAR.PORC.</td>
			</tr>

		<%

			If Trim(strTipo) = "1" Then strTipoInf = "DOC"
			If Trim(strTipo) = "2" Then strTipoInf = "MONTO"

			strSql = "EXEC proc_listado_JIC '" & inicio & "','" & termino & "','MONTO'"
			'Response.write "<br>strSql = " & strSql

			set rsDet=Conn.execute(strSql)

				intReg = 0
				intTotalTotal = 0
				do while not rsDet.eof
					intReg = intReg + 1
					intTotalTotal =  intTotalTotal + Cdbl(rsDet("TOTAL"))

					%>
					<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">


						<td ALIGN="RIGHT"><%=intReg%></td>
						<td ALIGN="LEFT"><%=Mid(rsDet("CLIENTE"),1,40)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("TOTAL"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("NORMALIZADO"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("DESASIGNADO"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("NO_ASIGNABLE"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("ACTIVA"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("ASIGNACION_PERIODO"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("NORMALIZADO_PERIODO"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("DESASIGNACION_PERIODO"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("NO_ASIGNABLE_PERIODO"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("SALDO_ACTIVO"),0)%></td>
						<td ALIGN="RIGHT"><%=FN(rsDet("VAR_PORC"),0)%></td>

					</tr>
					<%
					rsDet.movenext
				loop
		%>

	</table>


	<% End if %>
	</td>
   </tr>
  </table>

</form>


</body>
</html>

<div id="dhtmltooltip"></div>

<script type="text/javascript">
	CargaFechas(<%=strTipo%>,<%=strCliente%>);
	datos.CB_FECHA.value='<%=dtmFechaProc%>';
</script>


<script type="text/javascript">

/***********************************************
* Cool DHTML tooltip script- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/

var offsetxpoint=-60 //Customize x offset of tooltip
var offsetypoint=20 //Customize y offset of tooltip
var ie=document.all
var ns6=document.getElementById && !document.all
var enabletip=false
if (ie||ns6)
var tipobj=document.all? document.all["dhtmltooltip"] : document.getElementById? document.getElementById("dhtmltooltip") : ""

function ietruebody(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function ddrivetip(thetext, thecolor, thewidth){
if (ns6||ie){
if (typeof thewidth!="undefined") tipobj.style.width=thewidth+"px"
if (typeof thecolor!="undefined" && thecolor!="") tipobj.style.backgroundColor=thecolor
tipobj.innerHTML=thetext
enabletip=true
return false
}
}

function positiontip(e){
if (enabletip){
var curX=(ns6)?e.pageX : event.clientX+ietruebody().scrollLeft;
var curY=(ns6)?e.pageY : event.clientY+ietruebody().scrollTop;
//Find out how close the mouse is to the corner of the window
var rightedge=ie&&!window.opera? ietruebody().clientWidth-event.clientX-offsetxpoint : window.innerWidth-e.clientX-offsetxpoint-20
var bottomedge=ie&&!window.opera? ietruebody().clientHeight-event.clientY-offsetypoint : window.innerHeight-e.clientY-offsetypoint-20

var leftedge=(offsetxpoint<0)? offsetxpoint*(-1) : -1000

//if the horizontal distance isn't enough to accomodate the width of the context menu
if (rightedge<tipobj.offsetWidth)
//move the horizontal position of the menu to the left by it's width
tipobj.style.left=ie? ietruebody().scrollLeft+event.clientX-tipobj.offsetWidth+"px" : window.pageXOffset+e.clientX-tipobj.offsetWidth+"px"
else if (curX<leftedge)
tipobj.style.left="5px"
else
//position the horizontal position of the menu where the mouse is positioned
tipobj.style.left=curX+offsetxpoint+"px"

//same concept with the vertical position
if (bottomedge<tipobj.offsetHeight)
tipobj.style.top=ie? ietruebody().scrollTop+event.clientY-tipobj.offsetHeight-offsetypoint+"px" : window.pageYOffset+e.clientY-tipobj.offsetHeight-offsetypoint+"px"
else
tipobj.style.top=curY+offsetypoint+"px"
tipobj.style.visibility="visible"
}
}

function hideddrivetip(){
if (ns6||ie){
enabletip=false
tipobj.style.visibility="hidden"
tipobj.style.left="-1000px"
tipobj.style.backgroundColor=''
tipobj.style.width=''
}
}


document.onmousemove=positiontip

</script>
