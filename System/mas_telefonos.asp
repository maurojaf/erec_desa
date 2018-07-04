<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/asp/comunes/general/SoloNumeros.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
    <link rel="stylesheet" href="../css/style_generales_sistema.css">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	rut = request.QueryString("rut")
	strRutSubCliente = request.QueryString("strRutSubCliente")

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

'Response.write "<br>intTipoSoftPhone=" & intTipoSoftPhone
		
cerrarscg()

%>
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

	<title>TELEFONOS DEL DEUDOR</title>

	<style type="text/css">
		<!--
		.Estilo35 {color: #333333}
		.Estilo36 {color: #FFFFFF}
		.Estilo37 {color: #000000}
		-->
	</style>

</head>
<body>
<form action="" method="post" name="datos">
<input name="rut" type="hidden" id="rut" value="<%=rut%>">
<DIV class="titulo_informe">TELÉFONOS DEL DEUDOR</DIV>
<BR>
<table width="90%" border="0" align="center">

    <tr>
    <td valign="top" colspan="2">
	  <%
		intIdContacto = 1
	  abrirscg()

	  	strSql="SELECT DIAS_ATENCION,HORA_DESDE, HORA_HASTA, ANEXO, [dbo].[fun_trae_estatus_telefono_solo] ('" & session("ses_codcli") & "', RUT_DEUDOR, ID_TELEFONO) as ANALISIS, ID_TELEFONO,COD_AREA,TELEFONO,CORRELATIVO,ESTADO,FECHA_INGRESO, ISNULL(TELEFONO_DAL,0) AS TELEFONO_DAL, DIAS_ATENCION FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR ='" & rut & "' AND ESTADO IN (0,1) ORDER BY (CASE WHEN ESTADO = 0 THEN 1 WHEN ESTADO = 1 THEN 0 ELSE ESTADO END), FECHA_INGRESO DESC"
		''Response.write "<br>strSql=" & strSql
		set rsTel=Conn.execute(strSql)
		if not rsTel.eof then
	  %>

	  <table width="100%" border="0" class="estilo_columnas">
	  	<thead>
        <tr bordercolor="#FFFFFF" >
          <td align = "center">LLAMAR</td>
          <td align = "center">TIPO</td>
          <td>&Aacute;REA </td>
          <td >T&Eacute;LEFONO </td>
          <td align = "center">ANEXO</td>
          <td align = "center">DIAS ATENCION</td>
          <td colspan=1 align = "center">HORAS ATENCION</td>
          <td align = "center">CONTACTO</td>
          <td>&nbsp;</td>
          <!--td>FEC.INGRESO </td-->
          <td align = "center" colspan=1>ESTADO</td>
        </tr>
    	</thead>
		<%
		sinauditar=0
		novalida=0
		valida=0

		Do until rsTel.eof
		ID_TELEFONO=rsTel("ID_TELEFONO")
			FECHA_REVISION=rsTel("FECHA_INGRESO")
			if isNULL(FECHA_REVISION) then
				FECHA_REVISION=""
			end if
			COD_AREA=rsTel("COD_AREA")
			Telefono=rsTel("Telefono")
			correlativo_deudor=rsTel("CORRELATIVO")
			strTelefonoDal=rsTel("TELEFONO_DAL")
			strFonoAgestionar = COD_AREA & "-" & Telefono
			srtAnexo = UCASE(rsTel("ANEXO"))
			Estado=rsTel("Estado")
			if estado="0" then
				strEstadoFono="SIN AUDITAR"
			elseif estado="1" then
				strEstadoFono="VALIDO"
			elseif estado="2" then
				strEstadoFono="NO VALIDO"
			end if

			strAnalisis=Trim(rsTel("ANALISIS"))
			strHoraDesde=Trim(rsTel("HORA_DESDE"))
			strHoraHasta=Trim(rsTel("HORA_HASTA"))
			strDiasAtencion=Trim(rsTel("DIAS_ATENCION"))

			strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD, ID_CUOTA FROM CUOTA "
			strSql = strSql & " WHERE ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) "
			strSql = strSql & " AND COD_CLIENTE = '" & session("ses_codcli") & "' AND RUT_DEUDOR = '" & rut & "'"
			strSql = strSql & " GROUP BY ID_CUOTA"
			''Response.write strSql


			AbrirScg1()
			set rsCiclo=Conn1.execute(strSql)
			strCondicion = 0
			Do While (not rsCiclo.eof and strCondicion = 0)

				strSql = "SELECT ID_CUOTA FROM GESTIONES G, GESTIONES_CUOTA GD "
				strSql = strSql & " WHERE GD.ID_GESTION = G.ID_GESTION"
				strSql = strSql & " AND G.ID_MEDIO_AGENDAMIENTO = '" & ID_TELEFONO & "'"
				strSql = strSql & " AND G.COD_CLIENTE = '" & session("ses_codcli") & "' AND G.RUT_DEUDOR = '" & rut & "'"
				strSql = strSql & " AND GD.ID_CUOTA = " & rsCiclo("ID_CUOTA")

				''Response.write "<BR> = " & strSql

				set rsAnaliza = Conn.execute(strSql)
				If rsAnaliza.Eof Then
					strCondicion = 1
				End If

				rsCiclo.movenext
			Loop

			CerrarScg1()

			''Response.Write "<br>" & strCondicion
			If strCondicion = 1 then
				intCond11 = 1
			Else
				intCond11 = 0
			End If





			strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD "
			strSql = strSql & " FROM CUOTA WHERE COD_CLIENTE = '" & session("ses_codcli") & "' AND RUT_DEUDOR='" & rut &"'"
			strSql = strSql & " AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) "

			set rsAnGest=Conn.execute(strSql)
			if not rsAnGest.eof then
				intCantDocActivos = rsAnGest("CANTIDAD")
			Else
				intCantDocActivos = 0
			End If


			If intCantDocActivos > 0 then
				intCond12 = 1
			Else
				intCond12 = 0
			End If

			If strEstadoFono="VALIDO" or strEstadoFono="SIN AUDITAR" Then
				intCond13 = 1
			Else
				intCond13 = 0
			End If



			strSql = "SELECT ISNULL(COUNT(*),0) AS CANTIDAD "
			strSql = strSql & " FROM CUOTA WHERE COD_CLIENTE = '" & session("ses_codcli") & "' AND RUT_DEUDOR='" & rut &"'"
			strSql = strSql & " AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) "
			strSql = strSql & " AND (ID_FONO_AGEND_ULT_GES = '" & ID_TELEFONO & "' OR ID_FONO_AGEND_ULT_GES IS NULL OR ID_FONO_AGEND_ULT_GES = '')"
			strSql = strSql & " AND (FECHA_AGEND_ULT_GES <= getdate() OR FECHA_AGEND_ULT_GES IS NULL)"
			''Response.write strSql
			set rsAnGest=Conn.execute(strSql)

			if not rsAnGest.eof then
				intCantDoc2 = rsAnGest("CANTIDAD")
			Else
				intCantDoc2 = 0
			End If


			if intCantDoc2 > 0 then
				intCond21 = 1
				intCantDoc2 = rsAnGest("CANTIDAD")
			Else
				intCond21 = 0
				intCantDoc2 = 0
			End If


			If strEstadoFono="VALIDO" or strEstadoFono="SIN AUDITAR" Then
				intCond22 = 1
			Else
				intCond22 = 0
			End If


			If intCond11 = 1 and intCond12 = 1 and intCond13= 1 Then
				strEstatusTodos = "GESTIONAR"
			ElseIf intCond21 = 1 and intCond22 = 1 Then
				strEstatusTodos = "GESTIONAR"
			Else
				strEstatusTodos = "NO GESTIONAR"
			End If


		%>
		<input type="hidden" name="correlativo_deudor" id ="correlativo_deudor" value="<%=trim(correlativo_deudor)%>">
        <tr >

			<% If strEstadoFono="NO VALIDO" Then %>
				<td>&nbsp;</td>
			<% Else %>
				<% If strEstatusTodos = "GESTIONAR" Then %>
					<td align="CENTER">
						<A HREF="#" onClick="EnvioGestiones('<%=strFonoAgestionar%>','<%=rut%>', CB_CONTACTO_<%=intIdContacto%>.options[CB_CONTACTO_<%=intIdContacto%>.selectedIndex].value)"><img src="../imagenes/gestionar.jpg" border="0"></A>
					</td>
				<% ElseIf strEstatusTodos = "NO GESTIONAR" Then%>
					<td>&nbsp;</td>
				<% Else%>
					<% If strAnalisis = "GESTIONAR" Then%>
						<td align="CENTER">
						<A HREF="#" onClick="EnvioGestiones('<%=strFonoAgestionar%>','<%=rut%>', CB_CONTACTO_<%=intIdContacto%>.options[CB_CONTACTO_<%=intIdContacto%>.selectedIndex].value)"><img src="../imagenes/gestionar.jpg" border="0"></A>
						</td>
					<% Else%>
						<td>&nbsp;</td>
					<% End If%>
				<% End If%>
			<% End If%>

		  <td>

		  <%
		  if COD_AREA="9" then
		  	response.Write("CELULAR")
		  Elseif COD_AREA="0" then
		  	response.Write("SIN ESPECIF.")
		  else
		  	response.Write("RED FIJA")
		  end if


		  If Trim(srtAnexo) <> "" Then
		  	srtAnexoMsg = srtAnexo
		  Else
		  	srtAnexoMsg = "Sin Información"
		  End If
		  %>


          <td title="<%=rsTel("ID_TELEFONO")%>"><div align="CENTER"><%=COD_AREA%></div></td>
		  
		 <%If intTipoSoftPhone="1" then%>
          <td ><div align="left">
             &nbsp;<a href="sip:<%=SoloNumeros(strTelefonoDal)%>"><%=Telefono%></a>
            </div>
          </td>
	     <%ElseIf intTipoSoftPhone="2" Then%>
          <td ><div align="left">
             &nbsp;<a href="CALLTO://<%=SoloNumeros(strTelefonoDal)%>"><%=Telefono%></a>
            </div>
          </td>
	    <%End If%>
		 

          	<td title="<%=srtAnexoMsg%>"><div align="CENTER"><input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=srtAnexo%>" size="12" maxlength="50"></td>

			<%

			strChequedLu = ""
			strChequedMa = ""
			strChequedMi = ""
			strChequedJu = ""
			strChequedVi = ""
			strChequedSa = ""

			If instr(strDiasAtencion,"LU") > 0 Then strChequedLu = "CHECKED"
			If instr(strDiasAtencion,"MA") > 0 Then strChequedMa = "CHECKED"
			If instr(strDiasAtencion,"MI") > 0 Then strChequedMi = "CHECKED"
			If instr(strDiasAtencion,"JU") > 0 Then strChequedJu = "CHECKED"
			If instr(strDiasAtencion,"VI") > 0 Then strChequedVi = "CHECKED"
			If instr(strDiasAtencion,"SA") > 0 Then strChequedSa = "CHECKED"
			%>
			<td>
			Lu
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="LU" <%=strChequedLu%>>
			Ma
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MA" <%=strChequedMa%>>
			Mi
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MI" <%=strChequedMi%>>
			Ju
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="JU" <%=strChequedJu%>>
			Vi
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="VI" <%=strChequedVi%>>
			Sa
			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="SA" <%=strChequedSa%>>
            </td>

          	<td aling = "center"><input name="TX_DESDE_<%=correlativo_deudor%>" id="TX_DESDE_<%=correlativo_deudor%>" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
				<input name="TX_HASTA_<%=correlativo_deudor%>" id="TX_HASTA_<%=correlativo_deudor%>" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)"></td>

		   <td>
			<select name="CB_CONTACTO_<%=intIdContacto%>" id="CB_CONTACTO_<%=intIdContacto%>" onchange="this.style.width=150">
				<option value="">SELECCIONE</option>
				<%
				strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & rsTel("ID_TELEFONO")
				strSql = strSql & " UNION"
				strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "

				''Response.write "strSql=" & strSql
				set rsTemp= Conn.execute(strSql)
				if not rsTemp.eof then
				Do until rsTemp.eof%>
					<option value="<%=rsTemp("ID_CONTACTO")%>" <%if Trim(strPrincipal) = "S" then response.Write("SELECTED") End If%>><%=UCASE(rsTemp("CONTACTO"))%></option>
					<%
					rsTemp.movenext
				Loop
				Else
				%>
				<option value="0">SIN CONTACTO</option>
				<%

			
				end if
				rsTemp.close
				set rsTemp=nothing

				%>
			</select>


		</td>

		<td><A HREF="modificar_contacto.asp?strRut=<%=rut%>&intIdTelefono=<%=rsTel("ID_TELEFONO")%>"><img src="../imagenes/Agrega_contacto.png" border="0"></A>

		</td>

          <td><div align="right">
          	  <input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" type="radio" value="1"
			  <%if strEstadoFono="VALIDO" then
			   Response.Write("checked")
			   valida=valida+1
			   end if%>>
              VA
              <input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" <%=strNoValida%> type="radio" value="2"
			  <%if strEstadoFono="NO VALIDO" then
			  Response.Write("checked")
			  novalida=novalida+1
			  end if%>>
			  NV
			  <input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" type="radio" value="0"
			  <%if strEstadoFono="SIN AUDITAR" then
			  Response.Write("checked")
			  sinauditar=sinauditar+1
			  end if%>>
              SA
		    </span></div>
		  </td>

        </tr>
	<%
		intIdContacto = intIdContacto + 1
	rsTel.movenext
	Loop
	   %>

        <tr class="totales">
          <td ><span class="">TOTAL</span></td>
          <td  colspan="2"><span class="">V&Aacute;LIDOS : <%=valida%></span></td>
          <td  colspan="2"><span class="">NO V&Aacute;LIDOS : <%=novalida%></span></td>
          <td  colspan="4"><span class="">SIN AUDITAR : <%=sinauditar%></span></td>
          <td  colspan="1"><span class="">TOTAL TELÉFONOS : <%=(valida+novalida+sinauditar)%></span></td>
        </tr>

      </table>
	  <%
		else
		%>
			<div style="hegth:25px;" class="">SIN TELÉFONOS VÁLIDOS O SIN AUDITAR</div>
		<%	

		end if
		rsTel.close
		set rsTel=nothing
		cerrarscg()
	  %>
    </td>
  </tr>

</table>

<br>
<DIV class="titulo_informe">TELÉFONOS NO VALIDOS DEL DEUDOR</DIV>
<BR>
 <table width="90%" border="0" align="center">
     <tr>
     <td valign="top" colspan="2">
 	  <%

 	  abrirscg()

 	  	strSql="SELECT DIAS_ATENCION,HORA_DESDE, HORA_HASTA, ANEXO, [dbo].[fun_trae_estatus_telefono_solo] ('" & session("ses_codcli") & "', RUT_DEUDOR, ID_TELEFONO) as ANALISIS, ID_TELEFONO,COD_AREA,TELEFONO,CORRELATIVO,ESTADO,FECHA_INGRESO, ISNULL(TELEFONO_DAL,0) AS TELEFONO_DAL, DIAS_ATENCION FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR ='" & rut & "' AND ESTADO = 2 ORDER BY (CASE WHEN ESTADO = 0 THEN 1 WHEN ESTADO = 1 THEN 0 ELSE ESTADO END), FECHA_INGRESO DESC"
 		''Response.write "<br>strSql=" & strSql
 		set rsTel=Conn.execute(strSql)
 		if not rsTel.eof then
 	  %>
 	  <table width="100%" border="0"class="estilo_columnas">
 	  	<thead>
         <tr >
           <td align = "center">LLAMAR</td>
           <td align = "center">TIPO</td>
           <td>&Aacute;REA </td>
           <td >T&Eacute;LEFONO </td>
           <td align = "center">ANEXO</td>
           <td align = "center">DIAS ATENCION</td>
           <td colspan=1 align = "center">HORAS ATENCION</td>
           <td align = "center">CONTACTO</td>
           <td>&nbsp;</td>
           <!--td>FEC.INGRESO </td-->
           <td align = "center" colspan=2>ESTADO</td>
         </tr>
     	</thead>
 		<%
 		sinauditar=0
 		novalida=0
 		valida=0
 		Do until rsTel.eof
 			FECHA_REVISION=rsTel("FECHA_INGRESO")
 			if isNULL(FECHA_REVISION) then
 				FECHA_REVISION=""
 			end if
 			COD_AREA=rsTel("COD_AREA")
 			Telefono=rsTel("Telefono")
 			correlativo_deudor=rsTel("CORRELATIVO")
 			strTelefonoDal=rsTel("TELEFONO_DAL")
 			strFonoAgestionar = COD_AREA & "-" & Telefono
 			srtAnexo = UCASE(rsTel("ANEXO"))
 			Estado=rsTel("Estado")
 			if estado="0" then
 				strEstadoFono="SIN AUDITAR"
 			elseif estado="1" then
 				strEstadoFono="VALIDO"
 			elseif estado="2" then
 				strEstadoFono="NO VALIDO"
 			end if

 			strAnalisis=Trim(rsTel("ANALISIS"))
 			strHoraDesde=Trim(rsTel("HORA_DESDE"))
 			strHoraHasta=Trim(rsTel("HORA_HASTA"))
 			strDiasAtencion=Trim(rsTel("DIAS_ATENCION"))

  		%>
		<input type="hidden" name="correlativo_deudor" id ="correlativo_deudor" value="<%=trim(correlativo_deudor)%>">
         <tr >

 			<% If strEstadoFono="NO VALIDO" Then %>
 				<td>&nbsp;</td>
 			<% Else %>
 				<% If strEstatusTodos = "GESTIONAR" Then %>
 					<td align="CENTER">
 						<A HREF="#" onClick="EnvioGestiones('<%=strFonoAgestionar%>','<%=rut%>', CB_CONTACTO_<%=intIdContacto%>.options[CB_CONTACTO_<%=intIdContacto%>.selectedIndex].value)"><img src="../imagenes/gestionar.jpg" border="0"></A>
 					</td>
 				<% ElseIf strEstatusTodos = "NO GESTIONAR" Then%>
 					<td>&nbsp;</td>
 				<% Else%>
 					<% If strAnalisis = "GESTIONAR" Then%>
 						<td align="CENTER">
 						<A HREF="#" onClick="EnvioGestiones('<%=strFonoAgestionar%>','<%=rut%>', CB_CONTACTO_<%=intIdContacto%>.options[CB_CONTACTO_<%=intIdContacto%>.selectedIndex].value)"><img src="../imagenes/gestionar.jpg" border="0"></A>
 						</td>
 					<% Else%>
 						<td>&nbsp;</td>
 					<% End If%>
 				<% End If%>
 			<% End If%>

 		  <td>

 		  <%
 		  if COD_AREA="9" then
 		  	response.Write("CELULAR")
 		  Elseif COD_AREA="0" then
 		  	response.Write("SIN ESPECIF.")
 		  else
 		  	response.Write("RED FIJA")
 		  end if

		  If Trim(srtAnexo) <> "" Then
		  	srtAnexoMsg = srtAnexo
		  Else
		  	srtAnexoMsg = "Sin Información"
		  End If
 		  %>


           <td title="<%=rsTel("ID_TELEFONO")%>"><div align="CENTER"><%=COD_AREA%></div></td>
		   
		   
		<%If intTipoSoftPhone="1" then%>		  
          <td ><div align="left">
             &nbsp;<a href="SIP:<%=SoloNumeros(strTelefonoDal)%>"><%=Telefono%></a>
            </div>
          </td>
		<%ElseIf intTipoSoftPhone="2" Then%>
          <td ><div align="left">
             &nbsp;<a href="CALLTO:<%=SoloNumeros(strTelefonoDal)%>"><%=Telefono%></a>
            </div>
          </td>		
		 <%End If%>		

           	<td title="<%=srtAnexoMsg%>"><div align="CENTER"><input name="TX_ANEXO_<%=correlativo_deudor%>" id="TX_ANEXO_<%=correlativo_deudor%>" type="text" value="<%=srtAnexo%>" size="12" maxlength="50"></td>

 			<%

 			strChequedLu = ""
 			strChequedMa = ""
 			strChequedMi = ""
 			strChequedJu = ""
 			strChequedVi = ""
 			strChequedSa = ""

 			If instr(strDiasAtencion,"LU") > 0 Then strChequedLu = "CHECKED"
 			If instr(strDiasAtencion,"MA") > 0 Then strChequedMa = "CHECKED"
 			If instr(strDiasAtencion,"MI") > 0 Then strChequedMi = "CHECKED"
 			If instr(strDiasAtencion,"JU") > 0 Then strChequedJu = "CHECKED"
 			If instr(strDiasAtencion,"VI") > 0 Then strChequedVi = "CHECKED"
 			If instr(strDiasAtencion,"SA") > 0 Then strChequedSa = "CHECKED"
 			%>
 			<td>
 			Lu
 			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="LU" <%=strChequedLu%>>
 			Ma
 			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MA" <%=strChequedMa%>>
 			Mi
 			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="MI" <%=strChequedMi%>>
 			Ju
 			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="JU" <%=strChequedJu%>>
 			Vi
 			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="VI" <%=strChequedVi%>>
 			Sa
 			<INPUT TYPE="CHECKBOX" NAME="CH_DIAS_<%=correlativo_deudor%>" id="CH_DIAS_<%=correlativo_deudor%>" value="SA" <%=strChequedSa%>>
             </td>

           	<td align = "center"><input name="TX_DESDE_<%=correlativo_deudor%>" id="TX_DESDE_<%=correlativo_deudor%>" type="text" value="<%=strHoraDesde%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)">
 				<input name="TX_HASTA_<%=correlativo_deudor%>" id="TX_HASTA_<%=correlativo_deudor%>" type="text" value="<%=strHoraHasta%>" size="4" maxlength="5" onChange="return ValidaHora(this,this.value)"></td>

 		   <td>
 			<select name="CB_CONTACTO_<%=intIdContacto%>" id="CB_CONTACTO_<%=intIdContacto%>" onchange="this.style.width=150">
 				<option value="">SELECCIONE</option>
 				<%
 				strSql=" SELECT 0 as ORDEN, ID_CONTACTO, CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & rsTel("ID_TELEFONO")
 				strSql = strSql & " UNION"
 				strSql = strSql & " SELECT ORDEN, ID_CONTACTO, UPPER(CONTACTO_BASE) AS CONTACTO FROM CONTACTO_BASE WHERE COD_CLIENTE = '" & session("ses_codcli") & "' ORDER BY ORDEN, ID_CONTACTO DESC "

 				''Response.write "strSql=" & strSql
 				set rsTemp= Conn.execute(strSql)
 				if not rsTemp.eof then
 				Do until rsTemp.eof%>
 					<option value="<%=rsTemp("ID_CONTACTO")%>" <%if Trim(strPrincipal) = "S" then response.Write("SELECTED") End If%>><%=UCASE(rsTemp("CONTACTO"))%></option>
 					<%
 					rsTemp.movenext
 				Loop
 				Else
 				%>
 				<option value="0">SIN CONTACTO</option>
 				<%
		
 				end if
 				rsTemp.close
 				set rsTemp=nothing

 				%>
 			</select>


 		</td>

		</td>

		<td><A HREF="modificar_contacto.asp?strRut=<%=rut%>&intIdTelefono=<%=rsTel("ID_TELEFONO")%>"><img src="../imagenes/Agrega_contacto.png" border="0"></A>

		</td>

           <td><div align="right"><span class="Estilo35">
           	  <input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" type="radio" value="1"
 			  <%if strEstadoFono="VALIDO" then
 			   Response.Write("checked")
 			   valida=valida+1
 			   end if%>>
               VA
               <input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" <%=strNoValida%> type="radio" value="2"
 			  <%if strEstadoFono="NO VALIDO" then
 			  Response.Write("checked")
 			  novalida=novalida+1
 			  end if%>>
 			  NV
 			  <input name="radiofon<%=correlativo_deudor%>" id="radiofon<%=correlativo_deudor%>" type="radio" value="0"
 			  <%if strEstadoFono="SIN AUDITAR" then
 			  Response.Write("checked")
 			  sinauditar=sinauditar+1
 			  end if%>>
               SA
 		    </span></div>
 		  </td>

         </tr>
 	<%
 		intIdContacto = intIdContacto + 1
 	rsTel.movenext
 	Loop
 	   %>
         <tr class="totales">
           <td ><span class="">TOTAL</span></td>
           <td colspan=2><span class="">V&Aacute;LIDOS : <%=valida%></span></td>
           <td colspan=2><span class="">NO V&Aacute;LIDOS : <%=novalida%></span></td>
           <td colspan=4><span class="">SIN AUDITAR : <%=sinauditar%></span></td>
           <td colspan=1><span class="">TOTAL TELÉFONOS : <%=(valida+novalida+sinauditar)%></span></td>
         </tr>

       </table>
 	  <%
		else
		%>
			<div style="hegth:25px;" class="">SIN TELÉFONOS VÁLIDOS O SIN AUDITAR</div>
		<% 	 	  
 		end if
 		rsTel.close
 		set rsTel=nothing
 		cerrarscg()
 	  %>
     </td>
   </tr>

    <tr >
 	<td align="LEFT">
	</td>
 	<td align="RIGHT">
 		&nbsp;&nbsp;&nbsp;<img ID=ImgSave src="../imagenes/save_as.png" border="0" style="cursor:pointer;" onClick="envia();" alt="Guardar">
 		&nbsp;&nbsp;&nbsp;<img src="../imagenes/arrow_left.png" border="0" style="cursor:pointer;" alt="Volver" onClick="location.href='principal.asp'">
 	</td>
 	</tr>
</table>

</form>

<script type="text/javascript">
	$(document).ready(function(){
		$(document).tooltip();
	})

function envia(){
	var strFonoAgestionar 	=$('#strFonoAgestionar').val()
	var rut 			 	=$('#rut').val()

	$('input[name="correlativo_deudor"]').each(function(){

	 	var concat_anexo 		="#TX_ANEXO_"+$(this).val()
	 	var concat_radiomail	="input[id='radiofon"+$(this).val()+"']:checked"
		var concat_TX_DESDE 	="#TX_DESDE_"+$(this).val()
	 	var concat_TX_HASTA 	="#TX_HASTA_"+$(this).val()
	 	var concat_CH_DIAS 		="input[id='CH_DIAS_"+$(this).val()+"']:checked"
	 	var strDiasAtencion     =""

		$(concat_CH_DIAS).each(function () {
			strDiasAtencion =$(this).val()+","+strDiasAtencion
		})

		strDiasAtencion =strDiasAtencion.substring(0, strDiasAtencion.length-1)

	 	var strAnexo  			=$(concat_anexo).val()
	 	var estado_correlativo 	=$(concat_radiomail).val()
	 	var CORRELATIVO 		=$(this).val()
		var TX_DESDE 			=$(concat_TX_DESDE).val()
	 	var TX_HASTA 			=$(concat_TX_HASTA).val()



		var criterios ="alea="+Math.random()+"&rut="+rut+"&estado_correlativo="+encodeURIComponent(estado_correlativo)+"&strAnexo="+encodeURIComponent(strAnexo)+"&CORRELATIVO="+encodeURIComponent(CORRELATIVO)+"&TX_DESDE="+encodeURIComponent(TX_DESDE)+"&TX_HASTA="+encodeURIComponent(TX_HASTA)+"&strDiasAtencion="+encodeURIComponent(strDiasAtencion)

	 	$('#carga_funcion_ajax').load('FuncionesAjax/audita_fon_ajax.asp', criterios, function(data){
	 		
	 		
	 	})
		


	});	
	
	alert("¡Datos actualizados!")
	window.location.reload()
		

}


function EnvioGestiones(strFonoAgestionar,rut,strContactoSel){


	datos.action='detalle_gestiones.asp?strMasTelefonos=S&strFonoAgestionar=' + strFonoAgestionar + '&strCategoria=2&rut=' + rut + '&cliente=<%=session("ses_codcli")%>&strContactoSel=' + strContactoSel
	datos.submit();
}


function ValidaHora( ObjIng, strHora )
{
	var er_fh = /^(00|01|02|03|04|05|06|07|08|09|10|11|12|13|14|15|16|17|18|19|20|21|22|23)\:([0-5]0|[0-5][1-9])$/
	if( strHora == "" )
	{
	        alert("Introduzca la hora.")
	        return false
	}
	if ( !(er_fh.test( strHora )) )
	{
	        alert("El dato en el campo hora no es válido.");
	        ObjIng.value = '';
	        ObjIng.focus();
	        return false
	}

	//alert("¡Campo de hora correcto!")
	return true
}

</script>

</form>

</body>
</html>

<div id="carga_funcion_ajax"></div>
