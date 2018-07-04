<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>

<% ' Capa 1 ' %>
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc" -->
<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/asp/comunes/odbc/insertUpdate.inc"-->
<!--#include file="../lib/asp/comunes/odbc/ObtenerRecordset.inc"-->

<% ' Capa 2 ' %>
<!--#include file="../lib/asp/comunes/insert/Usuario.inc"-->
<!--#include file="../lib/asp/comunes/recordset/Usuario.inc"-->
<!--#include file="../lib/asp/comunes/general/funciones.inc"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"


AbrirSCG()

If Request("strFormMode") = "Nuevo" Then
	If Trim(request("ID_USUARIO")) <> "" Then
		recordset_Usuario Conn, srsRegistro, request("ID_USUARIO")
		If Not srsRegistro.EOF Then
			Response.Write "<P>Ya existe un registro con el código " &  request("ID_USUARIO")
			Response.Write "<P>Debe asignarle otro código si desea crear un registro nuevo"
			Response.Write "<FORM><INPUT VALUE=Volver TYPE=BUTTON onClick='javascript:history.back()'></FORM>"
			Response.End
		End If
	End If
End If
'Response.write "hola"
'Response.write "CB_BANCO" & request("CB_BANCO")
'Response.End

Set dicUsuario = CreateObject("Scripting.Dictionary")
dicUsuario.Add "ID_USUARIO", request("ID_USUARIO")
dicUsuario.Add "RUT_USUARIO", ValNulo(request("RUT_USUARIO"),"C")
dicUsuario.Add "ANEXO", ValNulo(request("ANEXO"),"C")
dicUsuario.Add "NOMBRES_USUARIO", ValNulo(request("NOMBRES_USUARIO"),"C")

dicUsuario.Add "APELLIDO_PATERNO", ValNulo(request("APELLIDO_PATERNO"),"C")
dicUsuario.Add "APELLIDO_MATERNO", ValNulo(request("APELLIDO_MATERNO"),"C")
dicUsuario.Add "FECHA_NACIMIENTO", ValNulo(request("FECHA_NACIMIENTO"),"C")
dicUsuario.Add "CORREO_ELECTRONICO", ValNulo(request("CORREO_ELECTRONICO"),"C")
dicUsuario.Add "TELEFONO_CONTACTO", ValNulo(request("TELEFONO_CONTACTO"),"C")

dicUsuario.Add "PERFIL", ValNulo(request("PERFIL"),"C")
dicUsuario.Add "LOGIN", ValNulo(request("LOGIN"),"C")
dicUsuario.Add "CLAVE", ValNulo(request("CLAVE"),"C")
dicUsuario.Add "PERFIL_ADM", ValNulo(request("PERFIL_ADM"),"N")
dicUsuario.Add "PERFIL_COB", ValNulo(request("PERFIL_COB"),"N")
dicUsuario.Add "PERFIL_SUP", ValNulo(request("PERFIL_SUP"),"N")
dicUsuario.Add "PERFIL_CAJA", ValNulo(request("PERFIL_CAJA"),"N")
dicUsuario.Add "PERFIL_PROC", ValNulo(request("PERFIL_PROC"),"N")
dicUsuario.Add "PERFIL_FULL", ValNulo(request("PERFIL_FULL"),"N")
dicUsuario.Add "PERFIL_EMP", ValNulo(request("PERFIL_EMP"),"N")

dicUsuario.Add "ACTIVO", ValNulo(request("ACTIVO"),"N")

dicUsuario.Add "OBSERVACIONES", ValNulo(request("OBSERVACIONES"),"C")


insert_Usuario Conn, dicUsuario


If Request("strFormMode") <> "Nuevo" Then

	strSql = "DELETE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & request("ID_USUARIO")
	set rsBorra= Conn.execute(strSql)

	strSql = "SELECT COD_CLIENTE FROM CLIENTE WHERE ACTIVO = 1"

	Response.write "strSql=" & strSql

	set rsEmpresa= Conn.execute(strSql)

	Do While not rsEmpresa.Eof
		strObjeto = "CH_CLIENTE_" & rsEmpresa("COD_CLIENTE")
		strValorObjeto = Request(strObjeto)

		'Response.write "<br>strObjeto=" & strObjeto
		'Response.write "<br>strValorObjeto=" & UCASE(strValorObjeto)

		If UCASE(strValorObjeto) = "ON" Then
			strSql = "INSERT INTO USUARIO_CLIENTE (ID_USUARIO, COD_CLIENTE) "
			strSql = strSql & "VALUES (" & request("ID_USUARIO") & "," & rsEmpresa("COD_CLIENTE") & ")"
			set rsInserta= Conn.execute(strSql)
		End if

		intCorr = intCorr + 1
		rsEmpresa.movenext
	Loop

End if

CerrarSCG()
Response.Redirect "man_Usuario.asp"
%>
