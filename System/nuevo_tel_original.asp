<% @LCID = 1034 %>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"

rut = request("rut")
strOrigen = request("strOrigen")
%>


<% If strOrigen = "" Then %>
<!--#include file="sesion.asp"-->
<% Else %>
<!--#include file="sesion_inicio.asp"-->
<% End If %>


<meta http-equiv="Content-Type" content="text/html;charset=utf-8" /> 
<script src="http://code.jquery.com/jquery-1.5.js"></script>

<title>NUEVO TELEFONO</title>


<link href="../css/style.css" rel="stylesheet" type="text/css">


	<style type="text/css">
		 body {
		 scrollbar-arrow-color: white;
		 scrollbar-dark-shadow-color: #000080;
		 scrollbar-track-color: #0080C0;
		 scrollbar-face-color: #0080C0;
		 scrollbar-shadow-color: white;
		 scrollbar-highlight-color: white;
		 scrollbar-3d-light-color: a;
		 overflow: auto;
		 overflow-x: hidden;
		 scrollbar-base-color:#ffeaff:
		 }
	 </style>

<table width="100%" border="0">

<% If strOrigen = "" Then %>
	<tr>
		<TD width="100%" ALIGN=LEFT class="pasos2_i">
			<B>Nuevo Teléfono</B>
		</TD>
	</tr>
<% End If %>
  	<tr>

    	<td valign="top">
    	<table width="100%" border="0" bordercolor="#FFFFFF">
		<form name="datos" method="post">
		<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

        <td>CODIGO AREA </td>
        <td>TELEFONO </td>
        <td>ANEXO</td>
        <td>DIAS DE ATENCION</td>
        <td Colspan =5>HORA ATENCION</td>

		</tr>


		<input name="num_min" type="hidden" value="0">

        <td> <select name="COD_AREA" id="COD_AREA" onblur="num_min.value=asigna_minimo(COD_AREA,num_min)">>
         <%
         abrirscg()
         ssql="SELECT DISTINCT CODIGO_AREA FROM COMUNA WHERE ID_SADI<>0 UNION SELECT 9 AS CODIGO_AREA  ORDER BY CODIGO_AREA DESC"
	 	 set rsCOM= Conn.execute(ssql)
		 do until rsCOM.eof%>

				<option value="<%=rsCOM("codigo_area")%>" selected><%=rsCOM("codigo_area")%></option>

          <%
		  rsCOM.movenext
		  loop
		  rsCOM.close
		  set rsCOM=nothing
		  cerrarscg()
		  %>

				<option value="0" selected>--</option>
             </select>
         (CEL.9)
		</td>

		<td>
				<input name="numero" type="text" id="numero" size="11" maxlength="8" onKeyUp="numero.value=solonumero(numero)">
		&nbsp

		<td>
				<input name="TX_ANEXO" type="text" value="<%=strAnexo%>" size="35" maxlength="50">


			<%
			strChequedLu = ""
			strChequedMa = ""
			strChequedMi = ""
			strChequedJu = ""
			strChequedVi = ""
			strChequedSa = ""

			If instr(strDiasPago,"LU") > 0 Then strChequedLu = "CHECKED"
			If instr(strDiasPago,"MA") > 0 Then strChequedMa = "CHECKED"
			If instr(strDiasPago,"MI") > 0 Then strChequedMi = "CHECKED"
			If instr(strDiasPago,"JU") > 0 Then strChequedJu = "CHECKED"
			If instr(strDiasPago,"VI") > 0 Then strChequedVi = "CHECKED"
			If instr(strDiasPago,"SA") > 0 Then strChequedSa = "CHECKED"
			%>

		<td>
				Lu
				<INPUT TYPE=CHECKBOX NAME="CH_DIAS" value="LU" <%=strChequedLu%>>
				Ma
				<INPUT TYPE=CHECKBOX NAME="CH_DIAS" value="MA" <%=strChequedMa%>>
				Mi
				<INPUT TYPE=CHECKBOX NAME="CH_DIAS" value="MI" <%=strChequedMi%>>
				Ju
				<INPUT TYPE=CHECKBOX NAME="CH_DIAS" value="JU" <%=strChequedJu%>>
				Vi
				<INPUT TYPE=CHECKBOX NAME="CH_DIAS" value="VI" <%=strChequedVi%>>
				Sa
				<INPUT TYPE=CHECKBOX NAME="CH_DIAS" value="SA" <%=strChequedSa%>>

		</td>

		<td>
				<input name="TX_DESDE" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
				<input name="TX_HASTA" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">


		<td>
		              <input name="rut" type="hidden" id="rut" value="<%=rut%>">
        </td>


	</tr>



    </tr>

        <tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

		<td colspan=2>CONTACTO</td>

       	<td colspan=1>CARGO</td>

       	<% If TraeSiNo(session("perfil_emp")) <> "Si" Then %>

       	<td colspan=1>DEPARTAMENTO</td>

        <td colspan=5>FUENTE </td>

        <%Else%>

        <td colspan=6>DEPARTAMENTO</td>

       	<% End If%>



		</tr>

		<tr bordercolor="#FFFFFF">


		<td colspan=2><input name="TX_CONTACTO" type="text" id="TX_CONTACTO" size="35" maxlength="20"></td>

		<td colspan=1><input name="TX_CARGO" type="text" id="TX_CARGO" size="35" maxlength="20"></td>



		<% If TraeSiNo(session("perfil_emp")) <> "Si" Then %>

		<td colspan=1><input name="TX_DPTO" type="text" id="TX_DPTO" size="35" maxlength="20"></td>

		<td colspan=4><select name="CB_FUENTE" id="CB_FUENTE">>

			<%
			abrirscg()
			ssql="SELECT * FROM FUENTE_UBICABILIDAD ORDER BY COD_FUENTE"
			set rsFuente= Conn.execute(ssql)
			do until rsFuente.eof%>
				<option value="<%=rsFuente("NOM_FUENTE")%>" selected><%=rsFuente("NOM_FUENTE")%></option>
				<%
					rsFuente.movenext
					loop
					rsFuente.close
					set rsFuente=nothing
					cerrarscg()
				%>
				</select>
		</td>

			<% Else%>

		<td colspan=3><input name="TX_CARGO" type="text" id="TX_CARGO" size="35" maxlength="35"></td>

			<% End If%>

		<TD>

			<A HREF="#" onClick="envia();">
				<img ID=ImgSave src="../imagenes/save_as.png" border="0">
			</A>
			&nbsp;&nbsp;
			<A HREF="#" onClick="VolverDT();">
				<img src="../imagenes/arrow_left.png" border="0">
			</A>

		</TD>

		</tr>

		</form>
    </table>
    </td>
  </tr>
</table>

<script language="JavaScript" type="text/JavaScript">


///------x-x-x-x--x-x-x-x-x-x*x-x*x-x*x-x*x-x*x-x*x*-*-*-*

function asigna_minimo(campo, minimo1){
	if (campo.value!=0)	{
		if(campo.value==41 || campo.value==32 || campo.value==45 || campo.value==57 || campo.value==55 || campo.value==72 || campo.value==71 || campo.value==73 || campo.value==75){
			minimo1=7;
		}else if(campo.value.length==1 || campo.value==2){
			minimo1=8;
		}else {
			minimo1=6;
		}
	}else{minimo1=0}
	return(minimo1)
}



function valida_largo(campo, minimo){
//alert(datos.fono_aportado_area.value)
	//if (datos.fono_aportado_area.value!="0"){
		if(campo.value.length != minimo) {
			alert("Fono debe tener " + minimo + " digitos")
			campo.select()
			campo.focus()
			return(true)
		}
	//}
	return(false)
}

function solonumero(valor){
     //Compruebo si es un valor numérico
      if (isNaN(valor.value)) {
            //entonces (no es numero) devuelvo el valor cadena vacia
            valor.value=""
			return ""
      }else{
            //En caso contrario (Si era un número) devuelvo el valor
			valor.value
			return valor.value
      }
}

function checkvalidate(checks) {
    for (i = 0; lcheck = checks[i]; i++) {
        if (lcheck.checked) {
            return true;
        }
    }
    return false;
}

function envia(){
	var grupo = document.getElementById("datos").CH_DIAS;
	if(datos.numero.value==''){
		alert('Debe ingresar un numero');
	}else if (valida_largo(datos.numero, datos.num_min.value)){
	}else{
		//datos.ImgSave.disabled=true;
		//datos.action='scg_tel.asp?strOrigen=<%=strOrigen%>&strRUT_DEUDOR=<%=rut%>';

		window.Contenido.refrescaTelefonos()
		//datos.submit();
	}
}

function refrescaTelefonos(){

	alert("ok")
}
function VolverDT(){
	datos.action="deudor_telefonos.asp?strOrigen=<%=strOrigen%>&strRUT_DEUDOR=<%=rut%>";
	datos.submit();
}


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

