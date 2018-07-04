<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CodePage = 65001
Response.charset="utf-8"

	strOrigen 		= request("strOrigen")
%>
<% If strOrigen = "" Then %>
	<!DOCTYPE html>
	<html lang="es">
	<head>		
	    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	    <meta charset="utf-8">
	    <link href="../css/normalize.css" rel="stylesheet">
	    <title>DIRECCIONES DEL DEUDOR</title>
<%end if%>

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
<%
strRut 			= request("strRut")
intIdTelefono 	= request("intIdTelefono")
strGraba 		= request("strGraba")
strElimina 		= request("strElimina")
strContacto 	= request("TX_CONTACTO")
strApellido		= request("TX_APELLIDO")
strCargo 		= request("TX_CARGO")
strDpto 		= request("TX_DPTO")


If strContacto <> "" and strApellido <> "" and strCargo <> "" and strDpto <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strCargo & " /" & strDpto
ElseIf strContacto <> "" and strApellido <> "" and strCargo <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strCargo
ElseIf strContacto <> "" and strApellido <> "" and strDpto <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strDpto
ElseIf strContacto <> "" and strApellido <> "" Then
	strContactoCargo = strContacto & " /"& strApellido
Else strContactoCargo = strContacto
END IF

'Response.Write "<br>strGraba=" & strGraba
'Response.Write "<br>intIdTelefono=" & intIdTelefono
''Response.Write "<br>strRut=" & strRut

AbrirSCG()

If Trim(strGraba) = "S" Then
		strSql = "INSERT INTO TELEFONO_CONTACTO (RUT_DEUDOR,ID_TELEFONO,CONTACTO,USR_INGRESO,FECHA_INGRESO)"
		strSql = strSql & " VALUES ('" & strRut & "'," & intIdTelefono & ",'" & UCASE(strContactoCargo) & "','" & session("session_login") & "',GETDATE())"
		Response.write "strSql=" & strSql &"<br>"
		set rsInsert = Conn.execute(strSql)

		If strOrigen = "" Then
			Response.Redirect "mas_telefonos.asp?rut=" + strRut
		Else
			Response.Redirect "deudor_telefonos.asp?strOrigen=" & strOrigen & "&strRUT_DEUDOR=" + strRut
		End If
End If

If Trim(strElimina) = "S" Then
	intIdContacto = Request("intIdContacto")
		strSql="DELETE FROM TELEFONO_CONTACTO WHERE ID_CONTACTO = " & intIdContacto
		''Response.write "strSql=" & strSql
		set rsInsert = Conn.execute(strSql)
End If

%>


<% If strOrigen = "" Then %>
	</head>
	<body>
<%end if%>

<input type="hidden" id="strOrigen" name="strOrigen" value="<%=trim(strOrigen)%>">

<% If strOrigen = "" Then %>
<div class="titulo_informe">MODIFICAR CONTACTO TELÉFONO</div>
<br>
<% End If %>
<div id="carga_funcion">
<table border="0" align="center" <%if strOrigen<>"" then%> style="width:100%;" <%else%> style="width:90%;" <%end if%>>
	<tr>
		<TD style="vertical-align: top;" align="left">
			<table width="100%" border="0" cellSpacing="0" cellPadding="0" class="intercalado" style="width:100%;">
				<thead>
				  <tr>
					<td>CONTACTOS ASOCIADOS</td>
					<td>FECHA ING.</td>
					<td colspan="2">USUARIO ING.</td>
					
				  </tr>
				</thead>
				<tbody>
				  <%
					strSql="SELECT UPPER(CONTACTO) AS CONTACTO, ID_CONTACTO, CONVERT(VARCHAR(10),FECHA_INGRESO,103) AS FECHA_INGRESO, USR_INGRESO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & intIdTelefono & " ORDER BY FECHA_INGRESO DESC"
					'Response.write "strSql=" & strSql &"<br>"
					set rsTemp1= Conn.execute(strSql)
					if not rsTemp1.eof then
						Do until rsTemp1.eof%>

						<tr >

							<td><%=rsTemp1("CONTACTO")%></td>
							<td><%=rsTemp1("FECHA_INGRESO")%></td>
							<td><%=rsTemp1("USR_INGRESO")%></td>
							<td align="CENTER"><img src="../imagenes/eliminar.jpg" border="0" onclick="modifica_contacto_elimina('<%=strRut%>','<%=trim(strOrigen)%>','<%=rsTemp1("ID_CONTACTO")%>','<%=intIdTelefono%>')"></td>

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

		</TD>
		<TD style="vertical-align: top;" align="left">
			<table width="100%" border="0" cellSpacing="0" cellPadding="0" class="estilo_columnas">
			<thead>	
				 <tr>
					<td align="left">NOMBRE</td>
					<td align="left">APELLIDO</td>
					<td align="left">CARGO</td>
					<td align="left">DEPARTAMENTO</td>
				  </tr>
			</thead>	  
			     <tr >
					<td align="left"><input name="TX_CONTACTO" type="text" id="TX_CONTACTO" size="20" maxlength="20"></td>
					<td align="left"><input name="TX_APELLIDO" type="text" id="TX_APELLIDO" size="20" maxlength="20"></td>
					<td align="left"><input name="TX_CARGO" type="text" id="TX_CARGO" size="20" maxlength="20"></td>
					<td align="left"><input name="TX_DPTO" type="text" id="TX_DPTO" size="20" maxlength="20"></td>
				</tr>
				 <tr >
					<td colspan="4" align="RIGHT">
						<A HREF="#" onClick="modifica_contacto_guarda('<%=intIdTelefono%>');">
							<img ID=ImgSave src="../imagenes/save_as.png" border="0">
						</A>
						&nbsp;&nbsp;
						<%if trim(strOrigen)="" then%>
							<A HREF="#" onClick="history.back();">
								<img ID=ImgVolver src="../imagenes/arrow_left.png" border="0">
							</A>
						<%else%>
							<A HREF="#" onClick="carga_funcion_telefono();">
								<img ID=ImgVolver src="../imagenes/arrow_left.png" border="0">
							</A>
						<%end if%>
					</td>

				</tr>
			</table>
		</TD>
	</tr>
</table>
</div>
<%if trim(strOrigen)="" then%>

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>

	<script type="text/javascript">

		function modifica_contacto_guarda(intIdTelefono){

		var rut 			= $('#rut_').val()
		var TX_CONTACTO 	= $('#TX_CONTACTO').val()
		var TX_APELLIDO		= $('#TX_APELLIDO').val()
		var TX_CARGO 		= $('#TX_CARGO').val()
		var TX_DPTO 		= $('#TX_DPTO').val()
		var strOrigen 		= $('#strOrigen').val()

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

		var criterios ="alea="+Math.random()+"&strRut="+rut+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)+"&accion_ajax=guardar_contacto&intIdTelefono="+intIdTelefono+"&strOrigen="+strOrigen

		$('#carga_funcion').load('FuncionesAjax/modificar_contacto_ajax.asp', criterios, function(data){

			var CB_FONO_CP_RUTA 	=$('#CB_FONO_CP_RUTA').val()
			var criterios 			="alea="+Math.random()+"&accion_ajax=actualiza_td_CB_CONTACTO_ASOCIADO_CP_RUTA&CB_FONO_CP_RUTA="+CB_FONO_CP_RUTA
			$('#td_CB_CONTACTO_ASOCIADO_CP_RUTA').load('FuncionesAjax/deudor_telefonos_ajax.asp', criterios, function(){
				var CB_FONO_GESTION 	=$('#CB_FONO_GESTION').val()
				set_CB_CONTACTO_ASOCIADO(CB_FONO_GESTION)
			})
		
		})

	}


	function modifica_contacto_elimina(strRut,strOrigen,intIdContacto,intIdTelefono)
	{
		var criterios ="alea="+Math.random()+"&strOrigen="+strOrigen+"&strRut="+strRut+"&intIdContacto="+intIdContacto+"&intIdTelefono="+intIdTelefono+"&accion_ajax=eliminar_contacto"

		$('#carga_funcion').load('FuncionesAjax/modificar_contacto_ajax.asp', criterios, function(data){

			var CB_FONO_CP_RUTA 	=$('#CB_FONO_CP_RUTA').val()
			var criterios 			="alea="+Math.random()+"&accion_ajax=actualiza_td_CB_CONTACTO_ASOCIADO_CP_RUTA&CB_FONO_CP_RUTA="+CB_FONO_CP_RUTA
			$('#td_CB_CONTACTO_ASOCIADO_CP_RUTA').load('FuncionesAjax/deudor_telefonos_ajax.asp', criterios, function(){

				var CB_FONO_GESTION 	=$('#CB_FONO_GESTION').val()
				set_CB_CONTACTO_ASOCIADO(CB_FONO_GESTION)

			})

		})
		
	}



	</script>


<%end if%>




<% If strOrigen = "" Then %>	
	</body>
	</html>
<%end if%>



