<!--#include file="../arch_utils.asp"-->
<%
	AbrirSCG1()
		Response.write SetCB_LUGAR_NORM2(Conn1,request("contentVar"))
	CerrarSCG1()

	function SetCB_LUGAR_NORM2(strConex, strRut)

		strSql = "SELECT 1 as TIPO, REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,'  ',' ') as LUGAR_PAGO , 0 AS ORDEN, id_direccion ID , 'DIRECCION' tipo FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = '" & TRIM(strRut) & "' AND ESTADO <> 2"
		strSql = strSql & " UNION"

		strSql = strSql & " SELECT 2 as TIPO, NOMBRE + ' ' + UBICACION as LUGAR_PAGO , ORDEN, ID_FORMA_RECAUDACION ID, 'FORMA_RECAUDACION' tipo FROM FORMA_RECAUDACION WHERE COD_CLIENTE = '" & TRIM(session("ses_codcli")) & "' ORDER BY ORDEN ASC"

		set rsDireccion = strConex.execute(strSql)

		''response.write "strQuery == " & strSql

		Do While not rsDireccion.eof
			'strDireccion = Replace(Replace(Trim(rsDireccion("LUGAR_PAGO")),"*"," "),"/"," ")
			strDireccion = Trim(rsDireccion("ID"))&"-"&TRIM(rsDireccion("TIPO"))
			value = value & strDireccion  & "*" & strDireccion & "/"
		
		rsDireccion.moveNext
		Loop

		rsDireccion.close
		set rsDireccion=nothing
		SetCB_LUGAR_NORM2 = value
	end function
%>

