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

End If

UCASE(strContactoCargo)

'Response.Write "<br>strGraba=" & strGraba
'Response.Write "<br>intIdTelefono=" & intIdTelefono
''Response.Write "<br>strRut=" & strRut

AbrirSCG()

If Trim(accion_ajax) = "guardar_contacto" Then

		strSql = "INSERT INTO TELEFONO_CONTACTO (RUT_DEUDOR,ID_TELEFONO,CONTACTO,USR_INGRESO,FECHA_INGRESO)"
		strSql = strSql & " VALUES ('" & strRut & "'," & intIdTelefono & ",'" & UCASE(strContactoCargo) & "','" & session("session_login") & "',GETDATE())"
		'Response.write "strSql=" & strSql &"<br>"
		'set rsInsert = Conn.execute(strSql)
		Conn.execute(strSql)
		If strOrigen = "" Then
			'Response.Redirect "mas_telefonos.asp?rut=" + strRut
		Else
			'Response.Redirect "deudor_telefonos.asp?strOrigen=" & strOrigen & "&strRUT_DEUDOR=" + strRut
		End If
End If

If Trim(accion_ajax) = "eliminar_contacto" Then
	intIdContacto = Request("intIdContacto")
		strSql="DELETE FROM TELEFONO_CONTACTO WHERE ID_CONTACTO = " & intIdContacto
		'Response.write "strSql=" & strSql
		'set rsInsert = Conn.execute(strSql)
		Conn.execute(strSql)
End If

%>



<input type="hidden" id="strOrigen" name="strOrigen" value="<%=trim(strOrigen)%>">

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
			    <tr>
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


