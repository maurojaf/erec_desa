<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../arch_utils.asp"-->
<%


Response.CodePage = 65001
Response.charset="utf-8"



	AbrirSCG1()



	response.write SetCB_FONO_GESTION(Conn1,request("contentVar"))

	CerrarSCG1()

	function SetCB_FONO_GESTION(strConex, strRut)

		strQuery = "SELECT ID_TELEFONO, TELEFONO, COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & strRut & "' AND ESTADO <> 2"

		'response.write "strQuery=============" & strQuery
		'REsponse.End
		set rsTelefonos = strConex.execute(strQuery)

		'response.write rsTelefonos.eof
		Do While not rsTelefonos.eof
			value = value & rsTelefonos("ID_TELEFONO")  & "*" & rsTelefonos("COD_AREA") & "-" & rsTelefonos("TELEFONO") & "/"
			rsTelefonos.moveNext
		Loop
		rsTelefonos.close
		set rsTelefonos=nothing

		SetCB_FONO_GESTION = value
	end function
%>

