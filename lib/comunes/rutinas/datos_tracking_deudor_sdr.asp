<%
If Trim(hdd_cod_cliente) <> "" Then
	strNombreCliente = TraeNombreClienteSDR(conexionSDR,hdd_cod_cliente)
End if 
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td align="center">
<table border="0" cellspacing="1" cellpadding="1" class="SycFondoTableAdm">
<tr class="SycFondoTitTabAdm" height=25>
<td align="center" colspan=5><font class="TituloDatos">DETALLE DOCUMENTO <%=txt_NroDoc%>==> Tracking Distribución</font></td>
</tr>
	<tr class="SycFondoTitTabAdm" height=25>
		<td align="center"><font class="TituloDatos">Fecha Gestión</font></td>
    	<td align="center"><font class="TituloDatos">Remesa</font></td>
		<td align="center"><font class="TituloDatos">Cobrador</font></td>
		<td align="center"><font class="TituloDatos">Excusa</font></td>
		<td align="center"><font class="TituloDatos">Estado</font></td>
	    <!--td align="center"><font class="SycFontTituloTablaAdm">Observaciones</font></td-->
	</tr>
<%
If Trim(txt_NroDoc) <> "" and Trim(hdd_cod_cliente) <> "" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSql= ""
	strSql= "Select G.Fecha_Gestion As FGestion,C.CodRemesa as Remesa,C.CodCobrador as Cobrador,E.Descripcion as Excusa,A.DesAct as Actividad,G.Obs as OBS"
	strSql= strSql & " From Gestion G,Excusa E,Actividades A,Cuota C"
	strSql= strSql & " Where ltrim(replace(SubString(G.Obs,1, charindex('::',G.Obs)),':',''))='" & txt_NroDoc & "'"
	strSql= strSql & " And G.Excusa=E.Codigo and A.CodActividad=G.Actividad And C.NroDoc=ltrim(replace(SubString(Obs,1, charindex('::',obs)),':',''))"
	strSql= strSql & " Order By G.Fecha_Gestion"

	rs.open strSql, conexionSDR
	Do While not rs.eof
	dtmFechaGestion=rs("FGestion")
	strRemesa=rs("Remesa")
	strCobrador=rs("Cobrador")
	strExcusa=rs("Excusa")
	strActividad=rs("Actividad")
	strObs=rs("OBS")
%>
<tr class="SycFondoLineaTabAdm" height=20>
<td align="center"><font class="SycFontTextoTablaAdm12"><%=dtmFechaGestion%></font></td>
<td align="center"><font class="SycFontTextoTablaAdm12"><%=strRemesa%></font></td>
<td><font class="SycFontTextoTablaAdm12"><%=strCobrador%></font></td>
<td><font class="SycFontTextoTablaAdm12"><%=strExcusa%></font></td>
<td class="SycFontTextoTablaAdm"><font class="SycFontTextoTablaAdm12"><%=strActividad%></font></td>
<!--td><font class="SycFontTextoTablaAdm"><%=strObs%></font></td-->
</tr>
<tr class="SycFondoLineaTabAdm" height=16>
<td colspan=5 class="SycFontTextoTablaAdm"><font class="SycFontTituloTablaAdm">&nbsp&nbspOBS:&nbsp<%=strObs%></font></td>
</tr>
<%	
    rs.MoveNext
	Loop	  
End If
%>
</table>

