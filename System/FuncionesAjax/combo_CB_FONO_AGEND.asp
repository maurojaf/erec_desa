<%
	AbrirSCG1()
		''response.write SetCB_FONO_AGEND(Conn1,request("contentVar"))
	CerrarSCG1()

	sub SetCB_FONO_AGEND(strConex, rut)
		strQuery = "SELECT IDTELEFONO, TELEFONO, CODAREA FROM DEUDOR_TELEFONO WHERE RUTDEUDOR = ' " & rut & "' AND ESTADO <> 2"
		vArrayE = split(ido,",")
		intTamvArrayE=ubound(vArrayE)
		set rsTelefonos = strConex.execute(strQuery)
		Do While not rsTelefonos.eof
			For indice = 0 to intTamvArrayE
				strSelected=""
				if Trim(rsTelefonos("IDTELEFONO")) = Trim(vArrayE(indice)) then
					strSelected = "SELECTED"
					exit for
				End If
			Next
			%>
			<OPTION VALUE="<%=Trim(rsTelefonos("IDTELEFONO"))%>" <%=strSelected%>><%=rsTelefonos("CODAREA") & "-" & rsTelefonos("TELEFONO")%></OPTION>
			<%
			rsTelefonos.moveNext
		Loop
		rsTelefonos.close
		set rsTelefonos=nothing
	end sub
%>