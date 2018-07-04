<!--#include file="../arch_utils.asp"-->
<%
	AbrirSCG1()
		Response.write SetCB_EMAIL_GESTION(Conn1,request("contentVar"))
	CerrarSCG1()

	function SetCB_EMAIL_GESTION(strConex, strRut)

		strSql = "SELECT ID_EMAIL, UPPER(EMAIL) AS EMAIL FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = '" & strRut & "' AND ESTADO <> 2"
		set rsEmail = strConex.execute(strSql)

		'response.write rsEmail.eof
		Do While not rsEmail.eof
			value = value & rsEmail("ID_EMAIL")  & "*" & rsEmail("EMAIL") & "/"
			rsEmail.moveNext
		Loop
		rsEmail.close
		set rsEmail=nothing

		SetCB_EMAIL_GESTION = value
	end function
%>

