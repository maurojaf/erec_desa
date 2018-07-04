<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
    <link rel="stylesheet" href="../css/style_generales_sistema.css">

    <!--#include file="sesion.asp"-->
<%
	Response.CodePage 	=65001
	Response.charset	="utf-8"
%>

	<title>DIRECCIONES DEL DEUDOR</title>

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
<%
	rut = request("rut")
%>

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">


<style type="text/css">
	<!--
	.Estilo35 {color: #333333}
	.Estilo36 {color: #FFFFFF}
	.Estilo37 {color: #000000}
	-->

</style>
</head>
<body>
<form action="" method="post" name="datos">
<input name="rut" type="hidden" id="rut" value="<%=rut%>">
<DIV class="titulo_informe">DIRECCIÓN DEL DEUDOR</DIV>
<BR>
<table width="90%" border="0" align="center">
  <tr>
    <td valign="top" colspan="2">
	  <%

		abrirscg()
		ssql=""
		ssql="SELECT HORA_DESDE, HORA_HASTA,DIAS_PAGO, ID_DIRECCION,Calle,Numero,Comuna,CORRELATIVO,Resto,Estado,FECHA_INGRESO FROM DEUDOR_DIRECCION WHERE ESTADO IN (0,1) AND RUT_DEUDOR='"&rut&"' ORDER BY FECHA_INGRESO"
		'Response.write "ssql=" & ssql
		set rsDIR=Conn.execute(ssql)
		if not rsDIR.eof then
	  %>

	  <table width="100%" border="0" class="estilo_columnas">
	  	<thead>
        <tr >
          <td ALIGN = "CENTER"></td>
          <td ALIGN = "CENTER" Width = "200">DIRECCION</td>
          <td ALIGN = "CENTER">RESTO</td>
          <td ALIGN = "CENTER">DIAS PAGO</td>
          <td ALIGN = "CENTER">HORARIO PAGO</td>
          <td ALIGN = "CENTER">CONTACTO COBRO</td>
		  <td>&nbsp;</td>
          <td ALIGN = "CENTER">ESTADO</td>
        </tr>
    	</thead>
		<%
		sinauditar=0
		novalida=0
		valida=0
		do until rsDIR.eof
			intId = rsDIR("ID_DIRECCION")
			FECHA_REVISION=rsDIR("FECHA_INGRESO")
			if isNULL(FECHA_REVISION) then
				FECHA_REVISION=""
			end if
			calle_deudor=rsDIR("Calle")
			numero_deudor=rsDIR("Numero")
			comuna_deudor=rsDIR("Comuna")
			correlativo_deudor=rsDIR("CORRELATIVO")
			strResto=UCASE(rsDIR("RESTO"))
			If Trim(strResto) = "" Then
				strLabelResto = "Sin Información"
			Else
				strLabelResto = strResto
			End If
			Estado=rsDIR("Estado")
			if estado="0" then
				estado_direccion="SIN AUDITAR"
			elseif estado="1" then
				estado_direccion="VALIDA"
			elseif estado="2" then
				estado_direccion="NO VALIDA"
			end if


			strHoraDesde=Trim(rsDIR("HORA_DESDE"))
			strHoraHasta=Trim(rsDIR("HORA_HASTA"))
			strDiasPago=Trim(rsDIR("DIAS_PAGO"))

			strDireccion = calle_deudor & " " & numero_deudor & " " & comuna_deudor
			strDireccion = Trim(strDireccion)

			strDireccion_geo = replace(ucase(strDireccion),"CALLE","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"CALLE","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"POBLACION","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"POBLACIÓN","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PASAJE","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"AV.","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PJE.","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PSJE","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PGE","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIDA","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"CAYE","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"CALLLE","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIDAS","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIA","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"V.","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"AVDA","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PASAGE","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PARCELA","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PARC.","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PARCELAS","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PARSELA","")
			strDireccion_geo = replace(ucase(strDireccion_geo),"PARS.","")		

			strDireccion_geo = strDireccion_geo

		%>
		<input type="hidden" id="correlativo_deudor" name="correlativo_deudor" value="<%=correlativo_deudor%>">
        <tr>
			<td width="40" align="center">
				<img width="20" style="cursor:pointer;" onclick="bt_geolocalizacion('<%=trim(strDireccion_geo)%>')" height="20" src="../Imagenes/map.png" title="Consulta dirección mapa">
			</td>
          	<td><acronym title="<%=strDireccion%>"><%=Mid(strDireccion,1,35)%></acronym></td>
          	<td title="<%=strLabelResto%>"><div align="CENTER"><input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=strResto%>" size="15" maxlength="50"></td>

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
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="LU" <%=strChequedLu%>>
			Ma
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MA" <%=strChequedMa%>>
			Mi
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MI" <%=strChequedMi%>>
			Ju
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="JU" <%=strChequedJu%>>
			Vi
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="VI" <%=strChequedVi%>>
			Sa
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="SA" <%=strChequedSa%>>
            </td>


            <td>
             	<input name="TX_DESDE_<%=correlativo_deudor%>" id="TX_DESDE_<%=correlativo_deudor%>" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
				<input name="TX_HASTA_<%=correlativo_deudor%>" id="TX_HASTA_<%=correlativo_deudor%>" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
			</td>


		   <td>

			<select name="CB_CONTACTO" id="CB_CONTACTO"  onchange="this.style.width=130">
				<%
				strSql="SELECT ID_CONTACTO, CONTACTO  FROM DIRECCION_CONTACTO WHERE ID_DIRECCION = " & rsDIR("ID_DIRECCION") & " ORDER BY ID_CONTACTO DESC"
				''Response.write "strSql=" & strSql
				set rsTemp= Conn.execute(strSql)
				if not rsTemp.eof then
				Do until rsTemp.eof%>
					<option value="<%=rsTemp("ID_CONTACTO")%>" <%if Trim(strPrincipal) = "S" then response.Write("SELECTED") End If%>><%=UCASE(rsTemp("CONTACTO"))%></option>
					<%
					rsTemp.movenext
				Loop
				Else
				%>
				<option value="0">SIN CONTACTO</option>
				<%
				end if
				rsTemp.close
				set rsTemp=nothing
				%>
			</select>
		</td>

		  <td align="CENTER"><A HREF="modificar_contacto_dir.asp?strRut=<%=rut%>&intIdDireccion=<%=rsDIR("ID_DIRECCION")%>"><img src="../imagenes/Agrega_contacto.png" border="0"></A></td>


	      <td><div align="right"><span class="Estilo35">
              <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="1"
			  <%if estado_direccion="VALIDA" then
			   Response.Write("checked")
			   valida=valida+1
			   end if%>>
              VA
			  <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="2"
			  <%if estado_direccion="NO VALIDA" then
			  Response.Write("checked")
			  novalida=novalida+1
			  end if%>>
			  NV
              <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="0"
			  <%if estado_direccion="SIN AUDITAR" then
			  Response.Write("checked")
			  sinauditar=sinauditar+1
			  end if%>>
              SA
		    </span></div></td>
        </tr>
	<%
	rsDIR.movenext
	loop
	   %>


        <tr class="totales">
          <td colspan="2"><span class="">TOTALES :</span> V&Aacute;LIDAS : <%=valida%></span></td>
          <td colspan="3"><span class="">NO V&Aacute;LIDAS : <%=novalida%></span></td>
          <td ><span class="">SIN AUDITAR : <%=sinauditar%></span></td>
          <td colspan="1">&nbsp;</td>
          <td colspan="2"><span class="" COLSPAN=3>TOTAL DIRECCIONES : <%=(valida+novalida+sinauditar)%></span></td>
        </tr>

      </table>
	  <%

	  	else%>
	  		<div style="hegth:25px;" class="">SIN DIRECCIONES VÁLIDAS O SIN AUDITAR</div>

		<%end if
		rsDIR.close
		set rsDIR=nothing
		cerrarscg()

	  %>
	  </td>
  </tr>
</table>

<br>
<DIV class="titulo_informe">DIRECCIONES NO VALIDAS DEL DEUDOR</DIV>
<BR>
<table width="90%" border="0" align="center">
  <tr>
    <td valign="top" colspan="2">
	  <%

		abrirscg()
		ssql=""
		ssql="SELECT HORA_DESDE, HORA_HASTA,DIAS_PAGO, ID_DIRECCION,Calle,Numero,Comuna,CORRELATIVO,Resto,Estado,FECHA_INGRESO FROM DEUDOR_DIRECCION WHERE ESTADO = 2 AND RUT_DEUDOR='"&rut&"' ORDER BY FECHA_INGRESO"
		'Response.write "ssql=" & ssql
		set rsDIR=Conn.execute(ssql)
		if not rsDIR.eof then
	  %>
	  <table width="100%" border="0"class="estilo_columnas">
	  	<thead>
        <tr >
			<td ALIGN = "CENTER"></td>
			<td ALIGN = "CENTER" Width = "200">DIRECCION</td>
			<td ALIGN = "CENTER">RESTO</td>
			<td ALIGN = "CENTER">DIAS PAGO</td>
			<td ALIGN = "CENTER">HORARIO PAGO</td>
			<td ALIGN = "CENTER">CONTACTO COBRO</td>
			<td>&nbsp;</td>
			<td ALIGN = "CENTER">ESTADO</td>
        </tr>
    	</thead>
		<%
		sinauditar=0
		novalida=0
		valida=0
		do until rsDIR.eof
			intId = rsDIR("ID_DIRECCION")
			FECHA_REVISION=rsDIR("FECHA_INGRESO")
			if isNULL(FECHA_REVISION) then
				FECHA_REVISION=""
			end if
			calle_deudor=rsDIR("Calle")
			numero_deudor=rsDIR("Numero")
			comuna_deudor=rsDIR("Comuna")
			correlativo_deudor=rsDIR("CORRELATIVO")

			strResto=UCASE(rsDIR("RESTO"))
			If Trim(strResto) = "" Then
				strLabelResto = "Sin Información"
			Else
				strLabelResto = strResto
			End If


			Estado=rsDIR("Estado")
			if estado="0" then
				estado_direccion="SIN AUDITAR"
			elseif estado="1" then
				estado_direccion="VALIDA"
			elseif estado="2" then
				estado_direccion="NO VALIDA"
			end if


			strHoraDesde=Trim(rsDIR("HORA_DESDE"))
			strHoraHasta=Trim(rsDIR("HORA_HASTA"))
			strDiasPago=Trim(rsDIR("DIAS_PAGO"))

			strDireccion = calle_deudor & " " & numero_deudor & " " & comuna_deudor
			strDireccion = Trim(strDireccion)
		%>
<input type="hidden" id="correlativo_deudor" name="correlativo_deudor" value="<%=correlativo_deudor%>">
        <tr>
			<td>&nbsp;</td>
          	<td><acronym title="<%=strDireccion%>"><%=Mid(strDireccion,1,35)%></acronym></td>
          	<td title="<%=strLabelResto%>">
          	<div align="CENTER">
          	<input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=strResto%>" size="15" maxlength="50">
          	</td>

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
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="LU" <%=strChequedLu%>>
			Ma
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MA" <%=strChequedMa%>>
			Mi
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MI" <%=strChequedMi%>>
			Ju
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="JU" <%=strChequedJu%>>
			Vi
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="VI" <%=strChequedVi%>>
			Sa
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="SA" <%=strChequedSa%>>
            </td>

			<td>
				<input name="TX_DESDE_<%=correlativo_deudor%>" id="TX_DESDE_<%=correlativo_deudor%>" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
				<input name="TX_HASTA_<%=correlativo_deudor%>" id="TX_HASTA_<%=correlativo_deudor%>" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
			</td>

		   <td>

			<select name="CB_CONTACTO" id="CB_CONTACTO" onchange="this.style.width=130">
				<%
				strSql="SELECT ID_CONTACTO, CONTACTO  FROM DIRECCION_CONTACTO WHERE ID_DIRECCION = " & rsDIR("ID_DIRECCION") & " ORDER BY ID_CONTACTO DESC"
				''Response.write "strSql=" & strSql
				set rsTemp= Conn.execute(strSql)
				if not rsTemp.eof then
				Do until rsTemp.eof%>
					<option value="<%=rsTemp("ID_CONTACTO")%>" <%if Trim(strPrincipal) = "S" then response.Write("SELECTED") End If%>><%=UCASE(rsTemp("CONTACTO"))%></option>
					<%
					rsTemp.movenext
				Loop
				Else
				%>
				<option value="0">SIN CONTACTO</option>
				<%
				end if
				rsTemp.close
				set rsTemp=nothing
				%>
			</select>
		</td>

		  <td align="CENTER"><A HREF="modificar_contacto_dir.asp?strRut=<%=rut%>&intIdDireccion=<%=rsDIR("ID_DIRECCION")%>"><img src="../imagenes/Agrega_contacto.png" border="0"></A></td>


	      <td><div align="right"><span class="">
              <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="1"
			  <%if estado_direccion="VALIDA" then
			   Response.Write("checked")
			   valida=valida+1
			   end if%>>
              VA
			  <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="2"
			  <%if estado_direccion="NO VALIDA" then
			  Response.Write("checked")
			  novalida=novalida+1
			  end if%>>
			  NV
              <input name="radiodir<%=correlativo_deudor%>" id="radiodir<%=correlativo_deudor%>" type="radio" value="0"
			  <%if estado_direccion="SIN AUDITAR" then
			  Response.Write("checked")
			  sinauditar=sinauditar+1
			  end if%>>
              SA
		    </span></div></td>
        </tr>
	<%
	rsDIR.movenext
	loop
	   %>


        <tr class="totales" >
          <td colspan="2"><span class="">TOTALES :</span> V&Aacute;LIDAS : <%=valida%></span></td>
          <td colspan="3"><span class="">NO V&Aacute;LIDAS : <%=novalida%></span></td>
          <td ><span class="">SIN AUDITAR : <%=sinauditar%></span></td>
          <td colspan="1">&nbsp;</td>
          <td colspan="2"><span class="" COLSPAN=3>TOTAL DIRECCIONES : <%=(valida+novalida+sinauditar)%></span></td>
        </tr>

      </table>
	  <%

	  	else%>
	  		<div style="hegth:25px;" class="">SIN DIRECCIONES NO VÁLIDAS O SIN AUDITAR</div>
		<%end if
		rsDIR.close
		set rsDIR=nothing
		cerrarscg()
	%>
  </td>
   </tr>

    <tr bordercolor="#FFFFFF">
    	<td align="LEFT">
    	</td>
 		<td align="RIGHT">
&nbsp;&nbsp;&nbsp;<img ID=ImgSave src="../imagenes/save_as.png" border="0" style="cursor:pointer;" onClick="enviar();" alt="Guardar">&nbsp;&nbsp;&nbsp;<img src="../imagenes/arrow_left.png" border="0" style="cursor:pointer;" alt="Volver" onClick="location.href='principal.asp'"><!--<input name="Submit" type="button" class="Estilo8" onClick="enviar();" value="Guardar Cambios Realizados">-->
 		</td>
   </tr>
</table>






</form>

<script type="text/javascript">
	$(document).ready(function(){
		$(document).tooltip();
	})

function bt_geolocalizacion(direccion){
	window.open('geolocalizacion.asp?direccion='+encodeURIComponent(direccion),"DATOS1","width=610, height=610, scrollbars=no, menubar=no, location=no, resizable=yes")

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

function enviar(){

	var strFonoAgestionar 	=$('#strFonoAgestionar').val()
	var rut 			 	=$('#rut').val()

	$('input[name="correlativo_deudor"]').each(function(){

	 	var concat_anexo 		="#TX_ANEXO_"+$(this).val()
	 	var concat_radiomail	="input[id='radiodir"+$(this).val()+"']:checked"
	 	var concat_TX_DESDE 	="#TX_DESDE_"+$(this).val()
	 	var concat_TX_HASTA 	="#TX_HASTA_"+$(this).val()
	 	var concat_CH_DIAS 		="input[id='CH_DIAS_"+$(this).val()+"']:checked"
	 	var strDiasAtencion     =""

		$(concat_CH_DIAS).each(function () {
			strDiasAtencion =$(this).val()+","+strDiasAtencion
		})

		strDiasAtencion =strDiasAtencion.substring(0, strDiasAtencion.length-1)

	 	var strAnexo  			=$(concat_anexo).val()
	 	var estado_correlativo 	=$(concat_radiomail).val()
	 	var CORRELATIVO 		=$(this).val()
	 	var TX_HASTA 			=$(concat_TX_HASTA).val()
	 	var TX_DESDE 			=$(concat_TX_DESDE).val()
	 	


		var criterios ="alea="+Math.random()+"&strOrigen=deudor_direcciones&rut="+rut+"&strFonoAgestionar="+strFonoAgestionar+"&estado_correlativo="+encodeURIComponent(estado_correlativo)+"&strAnexo="+encodeURIComponent(strAnexo)+"&CORRELATIVO="+encodeURIComponent(CORRELATIVO)+"&accion_ajax=auditar_direccion&strDiasAtencion="+encodeURIComponent(strDiasAtencion)+"&TX_DESDE="+TX_DESDE+"&TX_HASTA="+TX_HASTA

	 	$('#carga_funcion_ajax').load('FuncionesAjax/audita_dir_ajax.asp', criterios, function(data){
	 		
	 	})
		

	});


	alert("¡Datos actualizados!")
	window.location.reload()
}
</script>

<div id="carga_funcion_ajax"></div>

</body>
</html>