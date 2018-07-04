<%
	sub TraeTelefonoDeudor(strConex, rut)
		strQuery = "SELECT ID_TELEFONO, TELEFONO, COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & rut & "' AND ESTADO <> 2"
		set rsTelefonos = strConex.execute(strQuery)
		Do While not rsTelefonos.eof
			strFonoCB = rsTelefonos("COD_AREA") & "-" & rsTelefonos("TELEFONO")
			strSel=""
			%>
			<OPTION value="<%=rsTelefonos("ID_TELEFONO")%>" <%=strSel%>><%=rsTelefonos("COD_AREA")%>-<%=rsTelefonos("TELEFONO")%></OPTION>
			<%
			rsTelefonos.moveNext
		Loop
		rsTelefonos.close
		set rsTelefonos=nothing
	end sub
%>
