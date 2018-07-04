<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->
<!--#include file="../../lib/asp/comunes/general/SoloNumeros.inc" -->

<%

accion_ajax 		=request.querystring("accion_ajax")
rut					=request.querystring("rut")
intCodCliente=session("ses_codcli")

'Response.write "<br>accion_ajax=" & accion_ajax
'Response.write "<br>rut=" & rut


if trim(accion_ajax)="actualiza_CB_FONO_AGEND" then

rut					=request.querystring("rut")

%>
<SELECT NAME="CB_FONO_AGEND" id="CB_FONO_AGEND">
	<OPTION VALUE="0" >SELECCIONE</OPTION>
	<%if fono_con="0" or fono_con="" then%>
	<%
	AbrirSCG1()
	ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & rut & "' AND ESTADO <> 2"
	set rsFON=Conn1.execute(ssql_)
	Do until rsFON.eof
	
		strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
		strSel=""
		if strFonoCB = strFonoAsociado Then strSel = "SELECTED"	%>
		<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
	
	<%rsFON.movenext
	Loop
	rsFON.close
	set rsFON=nothing
	CerrarSCG1()
	%>
	<%else%>
		<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
	<%end if %>
</SELECT>
<%

elseif trim(accion_ajax)="actualiza_CB_FONO_GESTION" then
	rut		=request.querystring("rut")
%>

	<select name="CB_FONO_GESTION" id="CB_FONO_GESTION" onchange="set_CB_CONTACTO_ASOCIADO(this.value); return false;">
	<option value="0">SELECCIONE</option>
	<%if fono_con="0" or fono_con="" then%>
	  <%
		AbrirSCG1()
		ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & rut & "' AND ESTADO <> 2"
		set rsFON=Conn1.execute(ssql_)
		Do until rsFON.eof
			strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
			strSel=""
			if strFonoCB = strFonoAsociado Then strSel = "SELECTED"
			%>
			<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
			<%
				rsFON.movenext
		Loop
		rsFON.close
		set rsFON=nothing
		CerrarSCG1()
	 %>
	<%else%>
		<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
	<%end if %>
	</select>
<%

elseif trim(accion_ajax)="Carga_Fonos_AT" then
	
rut					=request.querystring("rut")
nombreDeudor		=request.querystring("nombreDeudor")
%>
</head>
<body>
<%


strFonoAgestionarO 	= request("strFonoAgestionar")
intUsaDiscador 	= request("intUsaDiscador")

strOrigen = request("strOrigen")

If session("permite_no_validar_fonos") = "N" Then
	If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then

	Else
		strNoValida = "disabled"
	End If
End If

abrirscg()

	strSql = "SELECT ISNULL(TIPO_SOFTPHONE,2) AS TIPO_SOFTPHONE "
	strSql = strSql & " FROM USUARIO "
	strSql = strSql & " WHERE ID_USUARIO = " & session("session_idusuario")

	set RsCli=Conn.execute(strSql)
	If not RsCli.eof then
		intTipoSoftPhone = RsCli("TIPO_SOFTPHONE")
	End if
	RsCli.close
	set RsCli=nothing

'Response.write "<br>rut=" & rut
'Response.write "<br>nombreDeudor=" & nombreDeudor
		
cerrarscg()
	
	
%>

		<table width="100%" border="0" bordercolor="#FFFFFF">
		  <tr>

			<td width="105" class="estilo_columna_individual">&nbsp;&nbsp;RUT DEUDOR</td>
			<td width="125" class="Estilo10" bgcolor="#C9DEF2">&nbsp;<%=rut%></td>			
				
			<td width="300"  class="estilo_columna_individual">&nbsp;&nbsp;NOMBRE DEUDOR</td>
			<td width="300"  class="Estilo10" bgcolor="#C9DEF2">&nbsp;<%=nombreDeudor%></td>
			
		  </tr>
		</table>
		
		
		<div class="subtitulo_informe">> FONOS VÁLIDOS Y SIN AUDITAR</div>
	  <%
	  abrirscg()
	  	strSql="SELECT DT.RUT_DEUDOR,IdTipoContacto, DIAS_ATENCION,HORA_DESDE, HORA_HASTA, ANEXO, [dbo].[fun_trae_estatus_telefono_solo] ('" & session("ses_codcli") & "', RUT_DEUDOR, ID_TELEFONO) as ANALISIS,"
		strSql = strSql & "ID_TELEFONO,COD_AREA,TELEFONO,CORRELATIVO,ESTADO,FECHA_INGRESO, ISNULL(TELEFONO_DAL,0) AS TELEFONO_DAL, DIAS_ATENCION, NOMBRE_ESTADO = ISNULL(EC.NOMBRE_ESTADO,'NO DEFINIDO'), FECHA_CONTACTABILIDAD = CONVERT(VARCHAR(10),FECHA_CONTACTABILIDAD,103) "
		strSql = strSql & "FROM DEUDOR_TELEFONO DT LEFT JOIN ESTADO_CONTACTABILIDAD EC ON DT.ID_ESTADO_CONTACTABILIDAD=EC.ID_ESTADO "
		strSql = strSql & "WHERE RUT_DEUDOR ='" & rut & "' AND ESTADO IN (1,0) "
		strSql = strSql & "ORDER BY ID_ESTADO_CONTACTABILIDAD ASC, FECHA_CONTACTABILIDAD DESC"
		
		'Response.write "<br>strSql=" & strSql
		
		set rsTel=Conn.execute(strSql)
		if rsTel.eof then
		%>
		
		<table width="100%" border="0">

			<tr bordercolor="#FFFFFF" bgcolor="#d0cfd7" height="25">
			<td align="center" class="Estilo10"><b>No existen teléfonos válidos o sin auditar</b></td>
			</tr>
			</form>
		</table>

	  <%
		Else
		%>
	  <table width="100%" border="0" class="intercalado" style="width:100%;">
	  	<thead>
        <tr >
			<td align = "center" width = "35">TIPO</td>
			<td align = "center" width = "70">ÁREA</td>
			<td align = "center" width = "100">TELEFONO </td>
			<td width = "25">&nbsp;</td>
			<td align = "center" width = "300">CONTACTABILIDAD</td>
			<td align = "center" width = "150">FECHA CONTACTO</td>
			<td align = "center" width = "150">ESTADO FONO</td>
			<td>&nbsp;</td>
        </tr>
    	</thead>
		<tr bordercolor="#FFFFFF">	
			<%
			Do until rsTel.eof

			intCodigoArea 				=rsTel("COD_AREA")
			Telefono 				=rsTel("Telefono")
			strFonoAgestionar 		= intCodigoArea & "-" & Telefono
			intFono 		= intCodigoArea & Telefono
			strTelefonoDal 			=rsTel("TELEFONO_DAL")
			strEstadoContactabilidad =Trim(rsTel("NOMBRE_ESTADO"))
			Estado 					=rsTel("Estado")
			strUltFecContacto					=rsTel("FECHA_CONTACTABILIDAD")
			
			  if intCodigoArea="9" then
				strTipoFono = "CELULAR"
			  Elseif intCodigoArea="0" then
				strTipoFono = "SIN ESPECIF."
			  else
				strTipoFono = "RED FIJA"
			  end if

			if estado="0" then
				strEstadoFono="SIN AUDITAR"
			elseif estado="1" then
				strEstadoFono="VALIDO"
			elseif estado="2" then
				strEstadoFono="NO VALIDO"
			end if
			
			'Response.write "<br>intFono=" & intFono
			
			%>	
			
			<td><div align="CENTER"><%=intCodigoArea%></div></td>	
			<td align="center"><%=strTipoFono%></td>
			
		<%If intTipoSoftPhone="1" then%>
			<td>
				<div align="center">
					<% j = 1 %>
					<a href="sip:<%=SoloNumeros(strTelefonoDal)%>" title="
					<% 	strLista = "SELECT CONTACTO FROM TELEFONO_CONTACTO WHERE RUT_DEUDOR = '"& RUT &"' AND ID_TELEFONO = '"& rsTel("ID_TELEFONO") &"' ORDER BY Fecha_ingreso DESC"
						set rsLista = Conn.execute(strLista)
						if not rsLista.Eof then
							Do While Not rsLista.Eof %>
								<% response.write(j) %> - <%=rsLista("CONTACTO") %>
						<% 	rsLista.movenext
							j = j + 1 
							Loop
							else
								response.write("No hay contactos ingresados.")
							end if %>
					"><%=Telefono%></a>
				</div>
			</td>
		<%ElseIf intTipoSoftPhone="2" Then%>
			<td>
				<div align="center">
					<% j = 1 %>
					<a href="callto://<%=SoloNumeros(strTelefonoDal)%>" title="
					<% 	strLista = "SELECT CONTACTO FROM TELEFONO_CONTACTO WHERE RUT_DEUDOR = '"& RUT &"' AND ID_TELEFONO = '"& rsTel("ID_TELEFONO") &"' ORDER BY Fecha_ingreso DESC"
						set rsLista = Conn.execute(strLista)
						if not rsLista.Eof then
							Do While Not rsLista.Eof %>
								<% response.write(j) %> - <%=rsLista("CONTACTO") %>
						<% 	rsLista.movenext
							j = j + 1 
							Loop
							else
								response.write("No hay contactos ingresados.")
							end if %>
					"><%=Telefono%></a>
				</div>
			</td>
		 <%End If%>
		 
			<td ALIGN="center">
				<a href="detalle_gestiones.asp?rut=<%=rut%>&cliente=<%=intCodCliente%>&strNuevaGestion=S&pagina_origen=agendamiento_tactico&fono_actual=<%=intFono%>" onclick="javascript:SetCustomer('<%=intCodCliente%>');">
					<img src="../imagenes/Contacto.azul.png" border="0">
				</a>
			</td>
				
			<td align="center"><%=strEstadoContactabilidad%></td>
			<td align="center"><%=strUltFecContacto%></td>
			<td align="center"><%=strEstadoFono%></td>

			<td align="center">
			
							<a href="javascript:ventanaGestionesFonos('gestiones_por_telefono.asp?intIdFono=<%=rsTel("ID_TELEFONO")%>&strFonoAgestionar=<%=strFonoAgestionar%>&strRutDeudor=<%=rut%>')">
							<img src="../imagenes/icon_gestiones.jpg" border="0">
						</a>
			</td>
			
		</tr>
			<%
			
			rsTel.movenext
			Loop		
			end if
			rsTel.close
			set rsTel=nothing
			cerrarscg()
		  %>
		<tr class="totales">
			<td colspan="8"><span class="" >&nbsp;</span></td>
			</td>
		</tr>
      </table>
	 
	<div class="subtitulo_informe">> FONOS INVÁLIDOS</div>
	 
	  <%
	  abrirscg()
	  	strSql="SELECT DT.RUT_DEUDOR,IdTipoContacto, DIAS_ATENCION,HORA_DESDE, HORA_HASTA, ANEXO, [dbo].[fun_trae_estatus_telefono_solo] ('" & session("ses_codcli") & "', RUT_DEUDOR, ID_TELEFONO) as ANALISIS,"
		strSql = strSql & "ID_TELEFONO,COD_AREA,TELEFONO,CORRELATIVO,ESTADO,FECHA_INGRESO, ISNULL(TELEFONO_DAL,0) AS TELEFONO_DAL, DIAS_ATENCION, NOMBRE_ESTADO = ISNULL(EC.NOMBRE_ESTADO,'NO DEFINIDO'), FECHA_CONTACTABILIDAD = CONVERT(VARCHAR(10),FECHA_CONTACTABILIDAD,103) "
		strSql = strSql & "FROM DEUDOR_TELEFONO DT LEFT JOIN ESTADO_CONTACTABILIDAD EC ON DT.ID_ESTADO_CONTACTABILIDAD=EC.ID_ESTADO "
		strSql = strSql & "WHERE RUT_DEUDOR ='" & rut & "' AND ESTADO IN (2) "
		strSql = strSql & "ORDER BY ID_ESTADO_CONTACTABILIDAD ASC, FECHA_CONTACTABILIDAD DESC"
		
		'Response.write "<br>intFono=" & intFono
		
		set rsTel=Conn.execute(strSql)
		if rsTel.eof then
		%>
		

		<table width="100%" border="0">

			<tr bordercolor="#FFFFFF" bgcolor="#d0cfd7" height="25">
			<td align="center" class="Estilo10"><b>No existen teléfonos inválidos</b></td>
			</tr>
			</form>
		</table>

	  <%
		Else
		%>
	  <table width="100%" border="0" class="intercalado" style="width:100%;">
	  	<thead>
        <tr >
			<td align = "center" width = "35">TIPO</td>
			<td align = "center" width = "70">ÁREA</td>
			<td align = "center" width = "100">TELEFONO </td>
			<td width = "25">&nbsp;</td>
			<td align = "center" width = "300">CONTACTABILIDAD</td>
			<td align = "center" width = "150">FECHA CONTACTO</td>
			<td align = "center" width = "150">ESTADO FONO</td>
			<td>&nbsp;</td>
        </tr>
    	</thead>
		<tr bordercolor="#FFFFFF">	
			<%
			Do until rsTel.eof
			
			Telefono 				=rsTel("Telefono")
			intCodigoArea 			=rsTel("COD_AREA")
			strFonoAgestionar 		= intCodigoArea & "-" & Telefono
			intFono 				= intCodigoArea & Telefono
			strTelefonoDal 			=rsTel("TELEFONO_DAL")
			
			strEstadoContactabilidad =Trim(rsTel("NOMBRE_ESTADO"))
			Estado 					=rsTel("Estado")
			strUltFecContacto					=rsTel("FECHA_CONTACTABILIDAD")
			
			  if intCodigoArea="9" then
				strTipoFono = "CELULAR"
			  Elseif intCodigoArea="0" then
				strTipoFono = "SIN ESPECIF."
			  else
				strTipoFono = "RED FIJA"
			  end if

			if estado="0" then
				strEstadoFono="SIN AUDITAR"
			elseif estado="1" then
				strEstadoFono="VALIDO"
			elseif estado="2" then
				strEstadoFono="NO VALIDO"
			end if
			
			%>	
			
			<td><div align="CENTER"><%=intCodigoArea%></div></td>	
			<td align="center"><%=strTipoFono%></td>
			
		<%If intTipoSoftPhone="1" then%>
			<td>
				<div align="center">
					<% j = 1 %>
					<a href="sip:<%=SoloNumeros(strTelefonoDal)%>" title="
					<% 	strLista = "SELECT CONTACTO FROM TELEFONO_CONTACTO WHERE RUT_DEUDOR = '"& RUT &"' AND ID_TELEFONO = '"& rsTel("ID_TELEFONO") &"' ORDER BY Fecha_ingreso DESC"
						set rsLista = Conn.execute(strLista)
						if not rsLista.Eof then
							Do While Not rsLista.Eof %>
								<% response.write(j) %> - <%=rsLista("CONTACTO") %>
						<% 	rsLista.movenext
							j = j + 1 
							Loop
							else
								response.write("No hay contactos ingresados.")
							end if %>
					"><%=Telefono%></a>
				</div>
			</td>
		<%ElseIf intTipoSoftPhone="2" Then%>
			<td>
				<div align="center">
					<% j = 1 %>
					<a href="callto://<%=SoloNumeros(strTelefonoDal)%>" title="
					<% 	strLista = "SELECT CONTACTO FROM TELEFONO_CONTACTO WHERE RUT_DEUDOR = '"& RUT &"' AND ID_TELEFONO = '"& rsTel("ID_TELEFONO") &"' ORDER BY Fecha_ingreso DESC"
						set rsLista = Conn.execute(strLista)
						if not rsLista.Eof then
							Do While Not rsLista.Eof %>
								<% response.write(j) %> - <%=rsLista("CONTACTO") %></br>
						<% 	rsLista.movenext
							j = j + 1 
							Loop
							else
								response.write("No hay contactos ingresados.")
							end if %>
					"><%=Telefono%></a>
				</div>
			</td>
		 <%End If%>

			<td ALIGN="center">
				<a href="detalle_gestiones.asp?rut=<%=rut%>&cliente=<%=intCodCliente%>&strNuevaGestion=S&pagina_origen=agendamiento_tactico&fono_actual=<%=intFono%>" onclick="javascript:SetCustomer('<%=intCodCliente%>');">
					<img src="../imagenes/Contacto.azul.png" border="0">
				</a>
			</td>
								
			<td align="center"><%=strEstadoContactabilidad%></td>
			<td align="center"><%=strUltFecContacto%></td>
			<td align="center"><%=strEstadoFono%></td>

			<td align="center">
			
							<a href="javascript:ventanaGestionesFonos('gestiones_por_telefono.asp?intIdFono=<%=rsTel("ID_TELEFONO")%>&strFonoAgestionar=<%=strFonoAgestionar%>&strRutDeudor=<%=rut%>')">
							<img src="../imagenes/icon_gestiones.jpg" border="0">
						</a>
			</td>
			
		</tr>
			<%
			
			rsTel.movenext
			Loop		
			end if
			rsTel.close
			set rsTel=nothing
			cerrarscg()
		  %>
		<tr class="totales">
			<td colspan="8"><span class="" >&nbsp;</span></td>
			</td>
		</tr>
      </table>
</body>
</html>
	
<%


elseif trim(accion_ajax)="actualiza_td_CB_FONO_CP_RUTA" then
	rut		=request.querystring("rut")

	
%>

	<select name="CB_FONO_CP_RUTA" id="CB_FONO_CP_RUTA" onchange="set_CB_CONTACTO_ASOCIADO_CP_RUTA(this.value); return false;">

	<option value="0">SELECCIONE</option>
	<%if fono_con="0" or fono_con="" then%>
	  <%
		AbrirSCG1()
		ssql_ = "SELECT ID_TELEFONO, TELEFONO,COD_AREA FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR='" & rut & "' AND ESTADO <> 2"
		set rsFON=Conn1.execute(ssql_)
		Do until rsFON.eof
			strFonoCB = rsFON("COD_AREA") & "-" & rsFON("Telefono")
			strSel=""
			if strFonoCB = strFonoAgend Then strSel = "SELECTED"
			%>
			<option value="<%=rsFON("ID_TELEFONO")%>" <%=strSel%>><%=rsFON("COD_AREA")%>-<%=rsFON("Telefono")%></option>
			<%
				rsFON.movenext
		Loop
		rsFON.close
		set rsFON=nothing
		CerrarSCG1()
	 %>
	<%else%>
		<option value="<%=fono_con%>"><%=area_con%>-<%=fono_con%></option>
	<%end if %>
	</select>

<%



elseif trim(accion_ajax)="actualiza_td_CB_CONTACTO_ASOCIADO_CP_RUTA" then
	rut				=request.querystring("rut")
	CB_FONO_CP_RUTA =request.querystring("CB_FONO_CP_RUTA")

	if trim(CB_FONO_CP_RUTA)<>0 then

		AbrirSCG1()
		strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & CB_FONO_CP_RUTA
		strSql = strSql & " UNION"
		strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "
		set rsContacto = Conn1.execute(strSql)
	%>

		<select name="CB_CONTACTO_ASOCIADO_CP_RUTA" id="CB_CONTACTO_ASOCIADO_CP_RUTA" onchange="this.style.width=260">
			<option value="0">SELECCIONE</option>
			<%do while not rsContacto.eof%>
				<option value="<%=trim(rsContacto("ID_CONTACTO"))%>"><%=trim(rsContacto("CONTACTO"))%></option>
			<%rsContacto.movenext
			loop%>
		</select>

	<%CerrarSCG1()
		'response.write strSql

	else
	%>
		<select name="CB_CONTACTO_ASOCIADO_CP_RUTA" id="CB_CONTACTO_ASOCIADO_CP_RUTA"  onchange="this.style.width=260">
			<option value="0">SELECCIONE</option>
		</select>
	<%
	end if

end if

%>

</head>

<script type="text/javascript">
	function ventanaGestionesFonos (URL){
		window.open(URL,"DATOS2","width=1300, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes")
	}
</script>