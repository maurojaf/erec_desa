<%
	SUMA1 = request.querystring("SUMA1")
	SUMA2 = request.querystring("SUMA2")
	
	if isnumeric(SUMA1) and isnumeric(SUMA2)  then
		response.write cint(SUMA1)+cint(SUMA2)
	else
		response.write "caracteres incorrectos"
	end if

%>