<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
    <link rel="stylesheet" href="../css/style_generales_sistema.css">	

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<%
Response.CodePage=65001
Response.charset ="utf-8"

rut = request.QueryString("rut")
%>
<title>EMPRESA S.A.</title>
<style type="text/css">
<!--
.Estilo35 {color: #333333}
.Estilo36 {
	color: #<%=session("COLTABBG")%>;
	font-weight: bold;
}
.Estilo36 {color: #FFFFFF}
.Estilo37 {color: #000000}
-->
</style>

</head>
<body>
<form action="" method="post" name="datos">
<input name="rut" type="hidden" id="rut" value="<%=rut%>">
<DIV class="titulo_informe">EMAIL DEL DEUDOR</DIV>
<BR>
<table width="90%" border="0" align="center">  
  <tr>
    <td valign="top" background="">
	  <%
	    abrirscg()
		ssql=""
		ssql="SELECT ID_EMAIL,ANEXO,FECHA_INGRESO,EMAIL,CORRELATIVO,ESTADO,FECHA_REVISION FROM DEUDOR_EMAIL WHERE RUT_DEUDOR='" & rut& "' AND ESTADO IN (0,1) ORDER BY FECHA_INGRESO"
		''Response.write "<br>ssql=" & ssql
		set rsDIR=Conn.execute(ssql)
		if not rsDIR.eof then
	  %>

	  <table class="estilo_columnas" style="width:100%;">
	  	<thead>
        <tr>
        	<td ALIGN = "CENTER">ID</td>
			<td>&nbsp;</td>
			<td ALIGN="CENTER" width="200">EMAIL</td>
			<td ALIGN="CENTER">ANEXO</td>
			<td ALIGN="CENTER">F.INGRESO</td>
			<td ALIGN="CENTER">F.AUDITORIA</td>
			<td ALIGN="CENTER">CONTACTO</td>
			<td>&nbsp;</td>
			<td WIDTH= 125 ALIGN = "CENTER" colspan=2>ESTADO</td>
        </tr>
    	</thead>
		<%
		sinauditar=0
		novalida=0
		valida=0
		do until rsDIR.eof
			FECHA_INGRESO=rsDIR("FECHA_INGRESO")
			if isNULL(FECHA_INGRESO) then
				FECHA_INGRESO=""
			end if
			Email=rsDIR("Email")

			FECHA_REVISION=rsDIR("FECHA_REVISION")
			if isNULL(FECHA_REVISION) then
				FECHA_REVISION=""
			end if

			correlativo_deudor=rsDIR("CORRELATIVO")
			strEstado=Trim(rsDIR("Estado"))


			if strEstado="0" then
				estado_EMAIL="SIN AUDITAR"
			elseif strEstado="1" then
				estado_EMAIL="VALIDO"
			elseif strEstado="2" then
				estado_EMAIL="NO VALIDO"
			end if

			srtAnexo = UCASE(rsDIR("ANEXO"))

			If Trim(srtAnexo) <> "" Then
				srtAnexoMsg = srtAnexo
			Else
				srtAnexoMsg = "Sin información"
			End If
		%>
		<input type="hidden" id="correlativo_deudor" name="correlativo_deudor" value="<%=trim(correlativo_deudor)%>">
        <tr>
			<td>&nbsp;</td>
			<td align="CENTER"><A HREF="detalle_gestiones.asp?strFonoAgestionar=<%=strFonoAgestionar%>&strCategoria=3&rut=<%=rut%>&cliente=<%=session("ses_codcli")%>"><img src="../imagenes/gestionar.jpg" border="0"></A></td>
			<td><%=Email%></td>
			<td title="<%=srtAnexoMsg%>"><div align="CENTER">
				<input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=srtAnexo%>" size="30" maxlength="30">
			</td>
          	<td align="LEFT"><%=MID(Cstr(FECHA_INGRESO),1,10)%></td>

        <td><div align="LEFT"><%=MID(Cstr(FECHA_REVISION),1,10)%></div></td>

		<td>
			<select name="CB_CONTACTO_<%=intIdContacto%>" id="CB_CONTACTO_<%=intIdContacto%>" onchange="this.style.width=200">
			<%
			strSql="SELECT ID_CONTACTO, CONTACTO  FROM EMAIL_CONTACTO WHERE ID_EMAIL = " & rsDIR("ID_EMAIL") & " ORDER BY ID_CONTACTO DESC"
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

	      <td align="CENTER"><A HREF="modificar_contacto_email.asp?strRut=<%=rut%>&intIdEmail=<%=rsDIR("ID_EMAIL")%>"><img src="../imagenes/Agrega_contacto.png" border="0"></A></td>

          <td><div align="right"><span class="Estilo35">
              <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="1"
			  <%if Trim(estado_EMAIL)="VALIDO" then
			   Response.Write("checked")
			   valida=valida+1
			   end if%>>
              VA
			  <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="2"
			  <%if Trim(estado_EMAIL)="NO VALIDO" then
			  Response.Write("checked")
			  novalida=novalida+1
			  end if%>>
			  NV
              <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="0"
			  <%if Trim(estado_EMAIL)="SIN AUDITAR" then
			  Response.Write("checked")
			  sinauditar=sinauditar+1
			  end if%>>
              SA
		    </span></div>
           </td>
        </tr>
	<%
	rsDIR.movenext
	loop
	   %>
        <tr class="totales" >
          <td colspan="2"><span class="">TOTAL</span></td>
          <td colspan="1"><span class=""></span> V&Aacute;LIDOS : <%=valida%></span></td>
          <td colspan="2"><span class="">NO V&Aacute;LIDOS : <%=novalida%></span></td>
          <td colspan="2"><span class="">SIN AUDITAR : <%=sinauditar%></span></td>
          <td colspan="2"><span class="">TOTAL CORREOS : <%=(valida+novalida+sinauditar)%></span></td>
        </tr>
      </table>
	  <%
		else
		%>
			<div style="hegth:25px;" class="">SIN EMAIL VÁLIDOS O SIN AUDITAR</div>
		<%		  
		end if
		rsDIR.close
		set rsDIR=nothing
		cerrarscg()

	  %>
  </tr>
</table>

<br>
<DIV class="titulo_informe">EMAIL NO VÁLIDOS DEL DEUDOR</DIV>
<BR>
<table width="90%" border="0" align="center">
  <tr>
    <td valign="top" background="">
	  <%
	    abrirscg()
		ssql = "SELECT ID_EMAIL,ANEXO,FECHA_INGRESO,EMAIL,CORRELATIVO,ESTADO,FECHA_REVISION FROM DEUDOR_EMAIL WHERE RUT_DEUDOR='" & rut& "' AND ESTADO = 2 ORDER BY FECHA_INGRESO"
		''Response.write "<br>2ssql=" & ssql
		set rsDIR=Conn.execute(ssql)
		if not rsDIR.eof then
	  %>
	  <table width="100%" border="0" class="estilo_columnas">
	  	<thead>
        <tr>
        	<td ALIGN="CENTER">ID</td>
			<td>&nbsp;</td>
			<td width="200" ALIGN="CENTER">EMAIL</td>
			<td ALIGN="CENTER">ANEXO</td>
			<td ALIGN="CENTER">F.INGRESO</td>
			<td ALIGN="CENTER">F.AUDITORIA</td>
			<td ALIGN="CENTER">CONTACTO</td>
			<td>&nbsp;</td>
			<td ALIGN="CENTER" colspan=2>ESTADO</td>
        </tr>
    	</thead>
		<%
		sinauditar=0
		novalida=0
		valida=0
		do until rsDIR.eof
			FECHA_INGRESO=rsDIR("FECHA_INGRESO")
			if isNULL(FECHA_INGRESO) then
				FECHA_INGRESO=""
			end if
			Email=rsDIR("Email")

			FECHA_REVISION=rsDIR("FECHA_REVISION")
			if isNULL(FECHA_REVISION) then
				FECHA_REVISION=""
			end if

			correlativo_deudor=rsDIR("CORRELATIVO")
			strEstado=Trim(rsDIR("Estado"))


			if strEstado="0" then
				estado_EMAIL="SIN AUDITAR"
			elseif strEstado="1" then
				estado_EMAIL="VALIDO"
			elseif strEstado="2" then
				estado_EMAIL="NO VALIDO"
			end if

			srtAnexo = UCASE(rsDIR("ANEXO"))

			If Trim(srtAnexo) <> "" Then
				srtAnexoMsg = srtAnexo
			Else
				srtAnexoMsg = "Sin información"
			End If

			'REsponse.Write "strEstado=" & strEstado
			'REsponse.Write "estado_EMAIL=" & estado_EMAIL
		%>
		<input type="hidden" id="correlativo_deudor" name="correlativo_deudor" value="<%=trim(correlativo_deudor)%>">
        <tr >
			<td>&nbsp;</td>
			<td align="CENTER"><A HREF="detalle_gestiones.asp?strFonoAgestionar=<%=strFonoAgestionar%>&strCategoria=3&rut=<%=rut%>&cliente=<%=session("ses_codcli")%>"><img src="../imagenes/gestionar.jpg" border="0"></A></td>
			<td><%=Email%></td>
			<td title="<%=srtAnexoMsg%>"><div align="CENTER">
				<input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=srtAnexo%>" size="30" maxlength="30">
			</td>
          	<td align="LEFT"><%=MID(Cstr(FECHA_INGRESO),1,10)%></td>

        <td><div align="LEFT"><%=MID(Cstr(FECHA_REVISION),1,10)%></div></td>

		<td>
			<select name="CB_CONTACTO_<%=intIdContacto%>" id="CB_CONTACTO_<%=intIdContacto%>" onchange="this.style.width=200">
			<%
			strSql="SELECT ID_CONTACTO, CONTACTO  FROM EMAIL_CONTACTO WHERE ID_EMAIL = " & rsDIR("ID_EMAIL") & " ORDER BY ID_CONTACTO DESC"
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

	      <td align="CENTER"><A HREF="modificar_contacto_email.asp?strRut=<%=rut%>&intIdEmail=<%=rsDIR("ID_EMAIL")%>"><img src="../imagenes/Agrega_contacto.png" border="0"></A></td>

          <td><div align="right"><span class="Estilo35">
              <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="1"
			  <%if Trim(estado_EMAIL)="VALIDO" then
			   Response.Write("checked")
			   valida=valida+1
			   end if%>>
              VA
			  <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="2"
			  <%if Trim(estado_EMAIL)="NO VALIDO" then
			  Response.Write("checked")
			  novalida=novalida+1
			  end if%>>
			  NV
              <input name="radiomail<%=correlativo_deudor%>" id="radiomail<%=correlativo_deudor%>" type="radio" value="0"
			  <%if Trim(estado_EMAIL)="SIN AUDITAR" then
			  Response.Write("checked")
			  sinauditar=sinauditar+1
			  end if%>>
              SA
		    </span></div>
           </td>
        </tr>
	<%
	rsDIR.movenext
	loop
	   %>
        <tr class="totales">
          <td colspan="2"><span class="">TOTAL</span></td>
          <td colspan="1"><span class=""></span> V&Aacute;LIDOS : <%=valida%></span></td>
          <td colspan="2"><span class="">NO V&Aacute;LIDOS : <%=novalida%></span></td>
          <td colspan="2"><span class="">SIN AUDITAR : <%=sinauditar%></span></td>
          <td colspan="2"><span class="">TOTAL CORREOS : <%=(valida+novalida+sinauditar)%></span></td>
        </tr>
      </table>
	  <%
			else
			%>
				<div style="hegth:25px;" class="">SIN EMAIL NO VÁLIDOS O SIN AUDITAR</div>
			<%	

		end if
		rsDIR.close
		set rsDIR=nothing
		cerrarscg()

	  %>       

     </td>
 	</tr>
     <tr>
     <td  align="RIGHT">
        <span class="Estilo35">
		&nbsp;&nbsp;&nbsp;<img ID=ImgSave src="../imagenes/save_as.png" border="0" style="cursor:pointer;" onClick="envia();" alt="Guardar">&nbsp;&nbsp;&nbsp;<img src="../imagenes/arrow_left.png" border="0" style="cursor:pointer;" alt="Volver" onClick="location.href='principal.asp'">

      </span> 
  </td>
  </tr>
</table>
<div id="carga_funcion_ajax"></div>



</form>

</body>
</html>

<script type="text/javascript">
function envia(){

	var rut 			 	=$('#rut').val()

	$('input[name="correlativo_deudor"]').each(function(){

	 	var concat_anexo 		="#TX_ANEXO_"+$(this).val()
	 	var concat_radiomail	="input[id='radiomail"+$(this).val()+"']:checked"

	 	var strAnexo  			=$(concat_anexo).val()
	 	var estado_correlativo 	=$(concat_radiomail).val()
	 	var CORRELATIVO 		=$(this).val()


		var criterios ="alea="+Math.random()+"&strOrigen=deudor_email&rut="+rut+"&estado_correlativo="+estado_correlativo+"&strAnexo="+encodeURIComponent(strAnexo)+"&CORRELATIVO="+CORRELATIVO+"&accion_ajax=auditar_email"



	 	$('#carga_funcion_ajax').load('FuncionesAjax/audita_cor_ajax.asp', criterios, function(data){

	 		
	 	})

	});

	alert("¡Datos actualizados!")
	window.location.reload()
}

$(document).ready(function(){
	$(document).tooltip();
})
</script>
