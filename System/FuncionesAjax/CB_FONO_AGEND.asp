<!--#include file="../arch_utils.asp"-->
<%
	AbrirSCG1()

	response.write SetCB_FONO_AGEND(Conn1,request("contentVar"))

	CerrarSCG1()

	function SetCB_FONO_AGEND(strConex, strRut)

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

		SetCB_FONO_AGEND = value
	end function
%>

