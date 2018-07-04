<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->
<!--#include file="sesion_inicio.asp"-->
<!--#include file="../lib/asp/comunes/general/SoloNumeros.inc" -->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"


rut = request("strRUT_DEUDOR")
strFonoAgestionarO = request("strFonoAgestionar")

If session("permite_no_validar_fonos") = "N" Then
	If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then

	Else
		strNoValida = "disabled"
	End If
End If
%>

<html>
<head>

<meta http-equiv="Content-Type" content="text/html;charset=utf-8" /> 



<title>TELEFONOS DEL DEUDOR</title>
	<style type="text/css">
		<!--
		.Estilo35 {color: #333333}
		.Estilo36 {color: #FFFFFF}
		.Estilo37 {color: #000000}
		-->
	</style>

	<style type="text/css">
		 body {
		 scrollbar-arrow-color: white;
		 scrollbar-dark-shadow-color: #000080;
		 scrollbar-track-color: #0080C0;
		 scrollbar-face-color: #0080C0;
		 scrollbar-shadow-color: white;
		 scrollbar-highlight-color: white;
		 scrollbar-3d-light-color: a;
		 scrollbar-base-color:#ffeaff:
		 }
	 </style>

	<style type=text/css>

		#dhtmltooltip {
			position: absolute;
			width: 250px;
			border: 2px solid black;
			padding: 2px;
			background-color: lightyellow;
			visibility: hidden;
			z-index: 100;
			/*Remove below line to remove shadow. Below line should always appear last within this CSS*/
			filter: progid:DXImageTransform.Microsoft.Shadow(color=gray,direction=135);
		}

	</style>
</head>
<body>
	  <%
	  abrirscg()
	  	strSql="SELECT DIAS_ATENCION,HORA_DESDE, HORA_HASTA, ANEXO, [dbo].[fun_trae_estatus_telefono_solo] ('" & session("ses_codcli") & "', RUT_DEUDOR, cast(COD_AREA as varchar) + '-' + Telefono) as ANALISIS, ID_TELEFONO,COD_AREA,TELEFONO,CORRELATIVO,ESTADO,FECHA_INGRESO, ISNULL(TELEFONO_DAL,0) AS TELEFONO_DAL, DIAS_ATENCION FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR ='" & rut & "' AND ESTADO IN (0,1) ORDER BY (CASE WHEN ESTADO = 0 THEN 1 WHEN ESTADO = 1 THEN 0 ELSE ESTADO END), FECHA_INGRESO DESC"
		''Response.write "<br>strSql=" & strSql
		set rsTel=Conn.execute(strSql)
		if rsTel.eof then
		%>

		<table width="100%" border="0">
			<form action="" method="post" name="datos">
		  	<input name="rut" type="hidden" id="rut" value="<%=rut%>">

			<tr bordercolor="#FFFFFF" bgcolor="#d0cfd7" height=25>
			<td align="center" class="Estilo10"><b>No existen teléfonos válidos o sin auditar</b></td>
			<td align="center" bgcolor="#<%=session("COLTABBG2")%>">
				<a href="#" onClick="envia('NF');" onMouseover="ddrivetip('Nuevo Fono', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><img src="../imagenes/phone_add.png" border="0"></a>
			</td>
			<td align="center" bgcolor="#<%=session("COLTABBG2")%>">
				<a href="#" onClick="envia('NV');" onMouseover="ddrivetip('Ver No válidos', '#EFEFEF',25)"; onMouseout="hideddrivetip()"><img src="../imagenes/brick_delete.png" border="0"></a>
			</td>
			</tr>
			</form>
		</table>

		<%

		Else
	  %>
	  <table width="100%" border="0">
		<form action="" method="post" name="datos">
	  	<input name="rut" type="hidden" id="rut" value="<%=rut%>">

        <tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td align = "center">TIPO</td>
			<td align = "center">ÁREA</td>
			<td bgcolor="#<%=session("COLTABBG")%>">T&Eacute;LEFONO </td>
			<td align = "center">ANEXO</td>
			<td align = "center">DIAS ATENCION</td>
			<td colspan=2 align = "center">HORAS ATENCION</td>
			<td align="center">ESTADO</td>
			<td>
				<a href="#" onClick="envia('AF');" onMouseover="ddrivetip('Auditar Fonos', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><img src="../imagenes/phone.png" border="0"></a>
				<a href="#" onClick="envia('NF');" onMouseover="ddrivetip('Nuevo Fono', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><img src="../imagenes/phone_add.png" border="0"></a>
				<a href="#" onClick="envia('NV');" onMouseover="ddrivetip('Ver No validos', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><img src="../imagenes/phone_delete.png" border="0"></a>
			</td>
        </tr>
		<%
		sinauditar=0
		novalida=0
		valida=0
		intIdContacto = 1
		Do until rsTel.eof
			FECHA_REVISION=rsTel("FECHA_INGRESO")
			if isNULL(FECHA_REVISION) then
				FECHA_REVISION=""
			end if
			COD_AREA=rsTel("COD_AREA")
			Telefono=rsTel("Telefono")
			correlativo_deudor=rsTel("CORRELATIVO")
			strTelefonoDal=rsTel("TELEFONO_DAL")
			strFonoAgestionar = COD_AREA & "-" & Telefono
			srtAnexo = UCASE(rsTel("ANEXO"))
			Estado=rsTel("Estado")
			if estado="0" then
				strEstadoFono="SIN AUDITAR"
			elseif estado="1" then
				strEstadoFono="VALIDO"
			elseif estado="2" then
				strEstadoFono="NO VALIDO"
			end if

			strAnalisis=Trim(rsTel("ANALISIS"))
			strHoraDesde=Trim(rsTel("HORA_DESDE"))
			strHoraHasta=Trim(rsTel("HORA_HASTA"))
			strDiasAtencion=Trim(rsTel("DIAS_ATENCION"))

		%>

        <tr bordercolor="#FFFFFF">

		  <td>

		  <%
		  if COD_AREA="9" then
		  	response.Write("CELULAR")
		  Elseif COD_AREA="0" then
		  	response.Write("SIN ESPECIF.")
		  else
		  	response.Write("RED FIJA")
		  end if

		  If Trim(srtAnexo) <> "" Then
		  	srtAnexoMsg = srtAnexo
		  Else
		  	srtAnexoMsg = "Sin información"
		  End If

		  %>

          <td onMouseover="ddrivetip('<%=rsTel("ID_TELEFONO")%>', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><div align="CENTER"><%=COD_AREA%></div></td>
          <td ><div align="left">
             &nbsp;<a href="sip:<%=SoloNumeros(strTelefonoDal)%>"><%=Telefono%></a>
            </div>
          </td>

          	<td onMouseover="ddrivetip('<%=srtAnexoMsg%>', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><div align="CENTER"><input name="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=srtAnexo%>" size="30" maxlength="30"></td>

			<%

			strChequedLu = ""
			strChequedMa = ""
			strChequedMi = ""
			strChequedJu = ""
			strChequedVi = ""
			strChequedSa = ""

			If instr(strDiasAtencion,"LU") > 0 Then strChequedLu = "CHECKED"
			If instr(strDiasAtencion,"MA") > 0 Then strChequedMa = "CHECKED"
			If instr(strDiasAtencion,"MI") > 0 Then strChequedMi = "CHECKED"
			If instr(strDiasAtencion,"JU") > 0 Then strChequedJu = "CHECKED"
			If instr(strDiasAtencion,"VI") > 0 Then strChequedVi = "CHECKED"
			If instr(strDiasAtencion,"SA") > 0 Then strChequedSa = "CHECKED"
			%>
			<td>
			Lu
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" value="LU" <%=strChequedLu%>>
			Ma
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" value="MA" <%=strChequedMa%>>
			Mi
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" value="MI" <%=strChequedMi%>>
			Ju
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" value="JU" <%=strChequedJu%>>
			Vi
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" value="VI" <%=strChequedVi%>>
			Sa
			<INPUT TYPE=CHECKBOX NAME="CH_DIAS_<%=correlativo_deudor%>" value="SA" <%=strChequedSa%>>
            </td>

          	<td align = "center"><input name="TX_DESDE_<%=correlativo_deudor%>" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)"></td>
			<td><input name="TX_HASTA_<%=correlativo_deudor%>" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
			<td align="center">
				<div align="right"><span class="Estilo35">
				<input name="radiofon<%=correlativo_deudor%>" type="radio" value="1"
				<%if strEstadoFono="VALIDO" then
				Response.Write("checked")
				valida=valida+1
				end if%>>
				VA
				<input name="radiofon<%=correlativo_deudor%>" <%=strNoValida%> type="radio" value="2"
				<%if strEstadoFono="NO VALIDO" then
				Response.Write("checked")
				novalida=novalida+1
				end if%>>
				NV
				<input name="radiofon<%=correlativo_deudor%>" type="radio" value="0"
				<%if strEstadoFono="SIN AUDITAR" then
				Response.Write("checked")
				sinauditar=sinauditar+1
				end if%>>
				SA
				</span>
				</div>
			</td>
			<td align="center">
				<A HREF="modificar_contacto.asp?strOrigen=deudor_telefonos&strRut=<%=rut%>&intIdTelefono=<%=rsTel("ID_TELEFONO")%>"><img src="../imagenes/Agrega_contacto.png" border="0"></A>
			</td>
		</tr>
	<%
		intIdContacto = intIdContacto + 1
	rsTel.movenext
	Loop
	   %>
		<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>">
			<td bgcolor="#<%=session("COLTABBG")%>" colspan=2><span class="Estilo36">TOTAL</span></td>
			<td bgcolor="#<%=session("COLTABBG2")%>" colspan=2><span class="Estilo37">V&Aacute;LIDOS : <%=valida%></span></td>
			<td bgcolor="#<%=session("COLTABBG2")%>" colspan=1><span class="Estilo37">SIN AUDITAR : <%=sinauditar%></span></td>
			<td bgcolor="#<%=session("COLTABBG2")%>" colspan=2>&nbsp;</td>
			<td bgcolor="#<%=session("COLTABBG2")%>" colspan=1><span class="Estilo37">TOTAL TELÉFONOS : <%=(valida+novalida+sinauditar)%></span></td>
			<td bgcolor="#<%=session("COLTABBG2")%>">
				<a href="#" onClick="envia('AF');" onMouseover="ddrivetip('Auditar Fonos', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><img src="../imagenes/phone.png" border="0"></a>
				<a href="#" onClick="envia('NF');" onMouseover="ddrivetip('Nuevo Fono', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><img src="../imagenes/phone_add.png" border="0"></a>
				<a href="#" onClick="envia('NV');" onMouseover="ddrivetip('Ver No validos', '#EFEFEF',10)"; onMouseout="hideddrivetip()"><img src="../imagenes/phone_delete.png" border="0"></a>
			</td>
		</tr>
		</form>
      </table>

	  <%
		end if
		rsTel.close
		set rsTel=nothing
		cerrarscg()
	  %>



<script language="JavaScript" type="text/JavaScript">
function envia(strTipo) {

	if (strTipo == 'AF') {
		datos.action='audita_fon.asp?strOrigen=deudor_telefonos&strFonoAgestionar=<%=strFonoAgestionarO%>';
	}

	if (strTipo == 'NV') {
		datos.action='deudor_telefonos_nv.asp?strOrigen=deudor_telefonos&strRUT_DEUDOR=<%=rut%>';
	}

	if (strTipo == 'NF') {
		datos.action='nuevo_tel.asp?strOrigen=deudor_telefonos&rut=<%=rut%>';
	}
	datos.submit();

}

</script>

<script language="javascript" type="text/javascript">

function ValidaHora( ObjIng, strHora )
{
        var er_fh = /^(00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23)\:([0-5]0|[0-5][1-9])$/
        if( strHora == "" )
        {
                alert("Introduzca la hora.")
                return false
        }
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



</body>
</html>

<div id="dhtmltooltip"></div>

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


