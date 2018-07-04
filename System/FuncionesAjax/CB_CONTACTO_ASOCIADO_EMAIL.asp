<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>


<!--#include file="../arch_utils.asp"-->
<%


	Response.CodePage = 65001
	Response.charset="utf-8"

	
	AbrirSCG1()
		Response.write SetCB_CONTACTO_ASOCIADO_EMAIL(Conn1,request("contentVar"))
	CerrarSCG1()

	function SetCB_CONTACTO_ASOCIADO_EMAIL(strConex, intIdEmail)

		strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & intIdEmail
		strSql = strSql & " UNION"
		strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "


		strSql="SELECT ID_CONTACTO,CONTACTO FROM EMAIL_CONTACTO WHERE ID_EMAIL = " & intIdEmail


		set rsContacto = strConex.execute(strSql)

		''response.write "strQuery == " & strSql

		Do While not rsContacto.eof
			strContacto = UCASE(Replace(Trim(rsContacto("CONTACTO")),"*"," "))

			''strContacto = Trim(rsContacto("CONTACTO"))
			value = value & Trim(rsContacto("ID_CONTACTO"))  & "*" & strContacto & " / "
			rsContacto.moveNext
		Loop
		rsContacto.close
		set rsContacto=nothing

		SetCB_CONTACTO_ASOCIADO_EMAIL = value
	end function
%>





