<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
<%
Response.CodePage 	=65001
Response.charset	="utf-8"

rut = request("rut")
strOrigen = request("strOrigen")

%>

	<% If strOrigen = "" Then %>
	<!--#include file="sesion.asp"-->
	<% Else %>
	<!--#include file="sesion_inicio.asp"-->
	<% End If %>
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

</head>
<body>
<%
strRut 			= request("strRut")
intIdEmail 		= request("intIdEmail")
strGraba 		= request("strGraba")
strElimina 		= request("strElimina")
strContacto 	= request("TX_CONTACTO")
strApellido		= request("TX_APELLIDO")
strCargo 		= request("TX_CARGO")
strDpto 		= request("TX_DPTO")
intIdContacto 	= Request("intIdContacto")

AbrirSCG()

If strContacto <> "" and strCargo <> "" and strDpto <> "" Then
	strContactoCargo = strContacto &" /"& strCargo &" /"& strDpto

ElseIf strContacto <> "" and strCargo <> "" Then
	strContactoCargo = strContacto &" /"& strCargo

ElseIf strContacto <> "" and strDpto <> "" Then
	strContactoCargo = strContacto &" /"& strDpto

Else strContactoCargo = strContacto

End If


%>
<% If strOrigen = "" Then %>
<div class="titulo_informe">MODIFICAR CONTACTO EMAIL</div>
<br>
<% End If %>
<div id="carga_funcion">

<INPUT TYPE="hidden" NAME="intIdEmail"  id="intIdEmail" value="<%=intIdEmail%>">
<INPUT TYPE="hidden" NAME="strRut" 		id="strRut" 	value="<%=strRut%>">
<INPUT TYPE="hidden" NAME="strOrigen" 	id="strOrigen" 	value="<%=strOrigen%>">	
<INPUT TYPE="hidden" NAME="rut_" 		id="rut_" 		value="<%=strRut%>">

<table border="0" align="center" <%if strOrigen<>"" then%> style="width:100%;" <%else%> style="width:90%;" <%end if%>>
  <tr>
    <td style="vertical-align: top;" align="left"  width="480">
		<table width="100%" border="0" class="intercalado" style="width:100%;">
		<thead>
		  <tr >

			<td Colspan="1">CONTACTOS ASOCIADOS</td>
			<td colspan="1">FECHA INGRESO</td>
			<td colspan="2">USUARIO INGRESO</td>
			<td width = "30" >&nbsp;</td>
			
		  </tr>
		</thead>  
		<tbody>

		  <%
			strSql="SELECT UPPER(CONTACTO) AS CONTACTO, ID_CONTACTO, CONVERT(VARCHAR(10),FECHA_INGRESO,103) AS FECHA_INGRESO, USR_INGRESO FROM EMAIL_CONTACTO WHERE ID_EMAIL = " & intIdEmail
			''Response.write "strSql=" & strSql
			set rsTemp1= Conn.execute(strSql)

			if not rsTemp1.eof then
				Do until rsTemp1.eof%>

				<tr >

					<td Colspan="1"><%=rsTemp1("CONTACTO")%></td>
					<td Colspan="1"><%=rsTemp1("FECHA_INGRESO")%></td>
					<td Colspan="2"><%=rsTemp1("USR_INGRESO")%></td>
					<td Colspan="4" align="CENTER"><img src="../imagenes/eliminar.jpg" border="0" onclick="modifica_email_elimina('<%=strRut%>','<%=strOrigen%>','<%=rsTemp1("ID_CONTACTO")%>','<%=intIdEmail%>')"></td>
			  </tr>
					<%
					rsTemp1.movenext
				Loop
			ELSE
			%>
				<TR><TD COLSPAN="4">SIN CONTACTOS ASOCIADOS</TD></TR>
			<%				
			End If
			%>
		</tbody>	
		</table>
	</td>
	<td style="vertical-align: top;" align="left"  width="480">
		<table width="100%" border="0" class="estilo_columnas">
		<thead>
		<tr>
			<td align="left">NOMBRE</td>
			<td align="left">APELLIDO</td>
			<td align="left">CARGO</td>
			<td align="left">DEPARTAMENTO</td>
		  </tr>
		</thead>
		<tr>
			<td align="left"><input name="TX_CONTACTO" type="text" id="TX_CONTACTO" size="20" maxlength="20"></td>
			<td align="left"><input name="TX_APELLIDO" type="text" id="TX_APELLIDO" size="20" maxlength="20"></td>
			<td align="left"><input name="TX_CARGO" type="text" id="TX_CARGO" size="20" maxlength="20"></td>
			<td align="left"><input name="TX_DPTO" type="text" id="TX_DPTO" size="20" maxlength="20"></td>
		</tr>
		<tr >
			<td colspan="4" align="RIGHT">
				<A HREF="#" onClick="modifica_email_guarda('<%=strOrigen%>','<%=intIdEmail%>');">
					<img ID=ImgSave src="../imagenes/save_as.png" border="0">
				</A>
				&nbsp;&nbsp;
				<%if trim(strOrigen)="" then%>
					<A HREF="#" onClick="history.back();">
						<img ID=ImgVolver src="../imagenes/arrow_left.png" border="0">
					</A>
				<%else%>
					<A HREF="#" onClick="carga_funcion_email();">
						<img ID=ImgVolver src="../imagenes/arrow_left.png" border="0">
					</A>
				<%end if%>

			</td>
		</tr>
		</table>
    </td>
  </tr>
</table>
</div>

</body>
</html>

<%if trim(strOrigen)="" then%>


	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>

	<script type="text/javascript">



	function modifica_email_guarda(strOrigen, intIdEmail){


		var rut 			= $('#rut_').val()
		var TX_CONTACTO 	= $('#TX_CONTACTO').val()
		var TX_APELLIDO		= $('#TX_APELLIDO').val()
		var TX_CARGO 		= $('#TX_CARGO').val()
		var TX_DPTO 		= $('#TX_DPTO').val()

		if(TX_CONTACTO==''){
			alert('DEBE INGRESAR NOMBRE');
			return
		}
		
		if(TX_APELLIDO==''){
			alert('DEBE INGRESAR APELLIDO');
			return
		}
		
		if(TX_CARGO=='' && TX_DPTO==''){
			alert('DEBE INGRESAR AL MENOS UN CARGO O UN DEPARTAMENTO');
			return
		}

		var criterios ="alea="+Math.random()+"&strOrigen="+strOrigen+"&strRut="+rut+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&accion_ajax=guardar_mail&intIdEmail="+intIdEmail

		$('#carga_funcion').load('FuncionesAjax/modificar_email_ajax.asp', criterios, function(data){

		})

		if(strOrigen!=""){
			set_CB_CONTACTO_ASOCIADO_EMAIL(0)
			var criterios 	="alea="+Math.random()+"&rut="+rut+"&accion_ajax=actualiza_CB_EMAIL_GESTION" 
			$('#td_CB_EMAIL_GESTION').load('FuncionesAjax/deudor_email_ajax.asp', criterios, function(){})
		}
	
	}


	function modifica_email_elimina(strRut,strOrigen,intIdContacto,intIdEmail){
		var rut 			= $('#rut_').val()

		var criterios ="alea="+Math.random()+"&strOrigen="+strOrigen+"&strRut="+strRut+"&intIdContacto="+encodeURIComponent(intIdContacto)+"&intIdEmail="+encodeURIComponent(intIdEmail)+"&accion_ajax=elimina_mail"


		$('#carga_funcion').load('FuncionesAjax/modificar_email_ajax.asp', criterios, function(data){})

		if(strOrigen!=""){
			set_CB_CONTACTO_ASOCIADO_EMAIL(0)
			var criterios 	="alea="+Math.random()+"&rut="+rut+"&accion_ajax=actualiza_CB_EMAIL_GESTION" 
			$('#td_CB_EMAIL_GESTION').load('FuncionesAjax/deudor_email_ajax.asp', criterios, function(){})
		}
	}

	</script>



<%end if%>