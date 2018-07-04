<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>


<!--#include file="../arch_utils.asp"-->
<%


	Response.CodePage = 65001
	Response.charset="utf-8"

	contentVar =request.queryString("contentVar")


	AbrirSCG1()
		'Response.write SetCB_CONTACTO_ASOCIADO(Conn1,request("contentVar"))

if trim(contentVar)<>0 then
	'function SetCB_CONTACTO_ASOCIADO(strConex, intIdTelefono)

		strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & contentVar
		strSql = strSql & " UNION"
		strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "


		set rsContacto = Conn1.execute(strSql)

		'response.write "strQuery == " & strSql
		%>
		<select name="CB_CONTACTO_ASOCIADO" id="CB_CONTACTO_ASOCIADO">
			<option value="0">SELECCIONE</option>
		<%
		Do While not rsContacto.eof
			strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))
		%>
			<option value="<%=Trim(rsContacto("ID_CONTACTO"))%>"><%=strContacto%></option>
		<%
			''strContacto = Trim(rsContacto("CONTACTO"))
			'value = value & Trim(rsContacto("ID_CONTACTO"))  & "*" & strContacto & "/"
			rsContacto.moveNext
		Loop
		rsContacto.close
		set rsContacto=nothing

		'SetCB_CONTACTO_ASOCIADO = value

	'end function

		CerrarSCG1()
%>
		</select>

<%else%>

	<select name="CB_CONTACTO_ASOCIADO" id="CB_CONTACTO_ASOCIADO"  onchange="this.style.width=260">
			<option value="0">SELECCIONE</option>
			</select>
<%end if%>




