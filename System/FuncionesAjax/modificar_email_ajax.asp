<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->


<%

Response.CodePage = 65001
Response.charset  ="utf-8"


strOrigen 		= request("strOrigen")
accion_ajax 	= request.queryString("accion_ajax")

strRut 			= request("strRut")
intIdEmail 		= request("intIdEmail")
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

End If

'Response.Write "<br>strGraba=" & strGraba
'Response.Write "<br>intIdTelefono=" & intIdTelefono
''Response.Write "<br>strRut=" & strRut

AbrirSCG()

If Trim(accion_ajax) = "guardar_mail" Then

	strSql = "INSERT INTO EMAIL_CONTACTO (RUT_DEUDOR,ID_EMAIL,CONTACTO,USR_INGRESO,FECHA_INGRESO)"
		strSql = strSql & " VALUES ('" & strRut & "'," & intIdEmail & ",'" & UCASE(strContactoCargo) & "','" & session("session_login") & "',GETDATE())"
		'Response.write "strSql=" & strSql
		'set rsInsert = Conn.execute(strSql)
		Conn.execute(strSql)


		If strOrigen = "" Then
			'Response.Redirect "mas_correos.asp?rut=" + strRut
		Else
			'Response.Redirect "deudor_email.asp?strOrigen=" & strOrigen & "&strRUT_DEUDOR=" + strRut
		End If
End If

If Trim(accion_ajax) = "elimina_mail" Then
	intIdContacto = Request("intIdContacto")
		strSql="DELETE FROM EMAIL_CONTACTO WHERE ID_CONTACTO = " & intIdContacto
		'Response.write "strSql=" & strSql
		'set rsInsert = Conn.execute(strSql)
		Conn.execute(strSql)
End If

%>


<table border="0" align="center" <%if strOrigen<>"" then%> style="width:100%;" <%else%> style="width:90%;" <%end if%>>

<INPUT TYPE="hidden" NAME="intIdEmail" value="<%=intIdEmail%>">
<INPUT TYPE="hidden" NAME="strRut" value="<%=strRut%>">

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

					<td Colspan=1><%=rsTemp1("CONTACTO")%></td>
					<td Colspan=1><%=rsTemp1("FECHA_INGRESO")%></td>
					<td Colspan=2><%=rsTemp1("USR_INGRESO")%></td>
					<td Colspan=4 align="CENTER"><img src="../imagenes/eliminar.jpg" border="0" onclick="modifica_email_elimina('<%=strRut%>','<%=strOrigen%>','<%=rsTemp1("ID_CONTACTO")%>','<%=intIdEmail%>')"></td>
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

		<tr bordercolor="#FFFFFF">
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

