<% @LCID = 1034 %>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->

<%
Response.CodePage = 65001
Response.charset="utf-8"


strRut 			= request("strRut")
intIdDireccion 	= request("intIdDireccion")
strGraba 		= request("strGraba")
strElimina 		= request("strElimina")
strContacto 	= request("TX_CONTACTO")
strApellido		= request("TX_APELLIDO")
strCargo 		= request("TX_CARGO")
strDpto 		= request("TX_DPTO")
accion_ajax 	= request.querystring("accion_ajax")
strOrigen 		= request("strOrigen")

'Response.Write "<br>strGraba=" & strGraba
'Response.Write "<br>intIdDireccion=" & intIdDireccion
''Response.Write "<br>strRut=" & strRut

AbrirSCG()

If strContacto <> "" and strApellido <> "" and strCargo <> "" and strDpto <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strCargo & " /" & strDpto
ElseIf strContacto <> "" and strApellido <> "" and strCargo <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strCargo
ElseIf strContacto <> "" and strApellido <> "" and strDpto <> "" Then
	strContactoCargo = strContacto & " /" & strApellido & " /" & strDpto
ElseIf strContacto <> "" and strApellido <> "" Then
	strContactoCargo = strContacto & " /"& strApellido
Else strContactoCargo = strContacto

End If


If Trim(accion_ajax) = "agrega_direccion" Then


		strSql = "INSERT INTO DIRECCION_CONTACTO (RUT_DEUDOR,ID_DIRECCION,CONTACTO,USR_INGRESO,FECHA_INGRESO)"
		strSql = strSql & " VALUES ('" & strRut & "'," & intIdDireccion & ",'" & UCASE(strContactoCargo) & "','" & session("session_login") & "',GETDATE())"
		'Response.write "strSql=" & strSql
		Conn.execute(strSql)

		If strOrigen = "" Then
			'Response.Redirect "mas_direcciones.asp?rut=" + strRut
		Else
			'Response.Redirect "deudor_direcciones.asp?strOrigen=" & strOrigen & "&strRUT_DEUDOR=" + strRut
		End If

End If

If Trim(accion_ajax) = "elimina_direccion" Then

	intIdContacto = Request("intIdContacto")

	strSql="DELETE FROM DIRECCION_CONTACTO WHERE ID_CONTACTO = " & intIdContacto
	'Response.write "strSql=" & strSql
		'set rsInsert = Conn.execute(strSql)
	Conn.execute(strSql)

End If

%>


<INPUT TYPE="hidden" NAME="intIdDireccion" value="<%=intIdDireccion%>">
<INPUT TYPE="hidden" NAME="strRut" value="<%=strRut%>">
<INPUT TYPE="hidden" NAME="strOrigen" id="strOrigen" value="<%=strOrigen%>">



<div id="carga_funcion">
<table border="0" align="center" <%if strOrigen<>"" then%> style="width:100%;" <%else%> style="width:90%;" <%end if%>>
  <tr>
    <td width="480" style="vertical-align: top;" align="left">
    <table width="100%" border="0" class="intercalado" style="width:100%;">
    <thead>	
      <tr>

			<td Colspan="1">CONTACTOS ASOCIADOS</td>
			<td colspan="1">FECHA INGRESO</td>
			<td colspan="2">USUARIO INGRESO</td>
			<td width = "30" >&nbsp;</td>
       </tr>
    </thead>   
    <tbody>

      <%
		strSql="SELECT UPPER(CONTACTO) AS CONTACTO, ID_CONTACTO, CONVERT(VARCHAR(10),FECHA_INGRESO,103) AS FECHA_INGRESO, USR_INGRESO FROM DIRECCION_CONTACTO WHERE ID_DIRECCION = " & intIdDireccion
		''Response.write "strSql=" & strSql
		set rsTemp1= Conn.execute(strSql)
		if not rsTemp1.eof then
			Do until rsTemp1.eof%>

			<tr >

				<td Colspan="1"><%=rsTemp1("CONTACTO")%></td>
				<td Colspan="1"><%=rsTemp1("FECHA_INGRESO")%></td>
				<td Colspan="2"><%=rsTemp1("USR_INGRESO")%></td>
				<td align="CENTER">
					<img src="../imagenes/eliminar.jpg" border="0" style="cursor:pointer;" onclick="elimina_direccion('<%=rsTemp1("ID_CONTACTO")%>','<%=trim(intIdDireccion)%>','<%=strRut%>')"></td>
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
	<TD style="vertical-align: top;" align="left" width="480">

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
				<td align="RIGHT" colspan="4">
					<A HREF="#" onClick="agrega_contacto_direccion('<%=trim(intIdDireccion)%>','<%=strRut%>');">
						<img ID="ImgSave" src="../imagenes/save_as.png" border="0">
					</A>
					&nbsp;&nbsp;
					<%if trim(strOrigen)="" then%>
						<A HREF="#" onClick="history.back();">
							<img ID=ImgVolver src="../imagenes/arrow_left.png" border="0">
						</A>

					<%else%>
						<A HREF="#" onClick="carga_funcion_direccion();">
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
<%if trim(strOrigen)="" then%>


	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>

	<script type="text/javascript">
	function agrega_contacto_direccion(intIdDireccion,strRut){

		var TX_CONTACTO 	=$('#TX_CONTACTO').val()
		var TX_APELLIDO		=$('#TX_APELLIDO').val()
		var TX_CARGO 		=$('#TX_CARGO').val()
		var TX_DPTO 		=$('#TX_DPTO').val()			
		var strOrigen 		=$('#strOrigen').val()
		
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

		var criterios ="alea="+Math.random()+"&strRut="+strRut+"&intIdDireccion="+intIdDireccion+"&accion_ajax=agrega_direccion&TX_CONTACTO="+encodeURIComponent(TX_CONTACTO)+"&TX_APELLIDO="+encodeURIComponent(TX_APELLIDO)+"&strOrigen="+strOrigen+"&TX_CARGO="+encodeURIComponent(TX_CARGO)+"&TX_DPTO="+encodeURIComponent(TX_DPTO)

		$('#carga_funcion').load('FuncionesAjax/modificar_contacto_dir_ajax.asp', criterios, function(data){})	

		
	}

	function elimina_direccion(intIdContacto,intIdDireccion,strRut){
		var strOrigen 		=$('#strOrigen').val()
		
		var criterios ="alea="+Math.random()+"&intIdContacto="+intIdContacto+"&accion_ajax=elimina_direccion&intIdDireccion="+intIdDireccion+"&strRut="+strRut+"&strOrigen="+strOrigen

		$('#carga_funcion').load('FuncionesAjax/modificar_contacto_dir_ajax.asp', criterios, function(data){})	

	}
	</script>


<%end if%>

<%cerrarscg()%>
