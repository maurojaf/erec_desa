<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/comunes/rutinas/funcionesBD.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">	
<%
		Response.CodePage=65001
		Response.charset ="utf-8"

		strFechaInicio = request("strFechaInicio")
		strFechaTermino = request("strFechaTermino")
		intCodTipoGes = request("intCodTipoGes")
		intCodUsuario=Request("strLogin")
		intTipoListado = request("CB_TIPO_LISTADO")
		strCodCliente = request("strCodCliente")
		strOrigen = request("strOrigen")
		strCobranza = Request("CB_COBRANZA")
		
		intIdFocoAT=Request("intIdFocoAT")
		intCodCampana = Request("intCodCampana")
		strTramoVenc = Request("strTramoVenc")
		strTramoMonto = Request("strTramoMonto")
		strSucursal = Request("strSucursal")
		
		'response.write "intIdFocoAT=" & intIdFocoAT
	
		intCodCategoria = request("intCodCategoria")
		intCodSubCategoria = request("intCodSubCategoria")
		intCodGestion = request("intCodGestion")

		if Trim(intTipoListado) = "" Then
			intTipoListado = 0
		End If

		AbrirSCG()

		if Trim(strFechaInicio) = "" Then
			strFechaInicio = TraeFechaActual(Conn)
		End If

		if Trim(strFechaTermino) = "" Then
			strFechaTermino = TraeFechaActual(Conn)

		Else strFechaTermino = strFechaTermino

		End If

		If strTramoVenc = "" Then strTramoVenc = 0 End If
		If strTramoMonto = "" Then strTramoMonto = 0 End If
		
		id_usuario      = session("session_idusuario")

        strSql="select EsInterno,EsExterno,PuedenEscucharMisGrabaciones,PuedoEscucharGrabaciones from usuario where id_usuario = " & id_usuario

        'response.write("intTipoListado--" & intTipoListado)

        set RsUser=Conn.execute(strSql)

    			    If not RsUser.eof then
	        Usr_EsInterno                       = RsUser("EsInterno")    
            Usr_EsExterno                       = RsUser("EsExterno")    
            Usr_PuedenEscucharMisGrabaciones    = RsUser("PuedenEscucharMisGrabaciones")    
            Usr_PuedoEscucharGrabaciones        = RsUser("PuedoEscucharGrabaciones")   

	    End if
	    RsUser.close
	    set RsUser=nothing	

		

		%>
		<title>Detalle Gestiones</title>
</head>
<body>
<form name="datos" method="post">
	<div class="titulo_informe">LISTADO GESTIONES</div>
	<br>

		<table width="90%" border="0" align="center" >
			<tr height="20">
				<td class="estilo_columna_individual" width="150">TIPO DE LISTADO</td>
				<td>

					<select name="CB_TIPO_LISTADO" onChange="envia();">
						<option value="0" <%if Trim(intTipoListado)="0" then response.Write("Selected") end if%>>CATEGORIZADAS</option>
						<option value="1" <%if Trim(intTipoListado)="1" then response.Write("Selected") end if%>>TOTALES</option>
					</select>
				</td>
				<td align="right"><input type="button" value="Volver" class="fondo_boton_100" onclick="javascript:history.back();"></td>				
			</tr>
		</table>

		<br>
		<table class="intercalado">
		
<%		If intTipoListado = 1 then %>

		<thead>	
		<tr>
			<td>&nbsp;</td>
			<td align="center">Foco</td>
			<td align="center">Fecha Ing.</td>
			<td align="center">Hora Ing.</td>
			<td align="center">Rut Deudor</td>
			<td>NOMBRE D.</td>
			<td align="center">Monto</td>
			<td align="center">Total Doc.</td>
			<td align="center">DMI</td>
			<td align="center">Gestión</td>
			<td align="center" >Medio</td>
			<td>Contacto</td>
			<td align="center">Ejecutivo</td>
			<td align="center">Fecha Comp.</td>
			<td align="center">Fecha Agend.</td>
			<td align="center">Hora Agend.</td>
			<td align="center">Días Agend.</td>
			<td align="center">OBS.</td>
			<td>&nbsp;</td>
		</tr>
		</thead>
		<tbody>
<%
			strSql = "Exec [uspInfListadoDetalleGestiones] '"&TRIM(strCodCliente)&"','" & intCodUsuario &"','"& intCodCategoria &"','"& intCodSubCategoria &"','"& intCodGestion &"','"&strFechaInicio&"','"&strFechaTermino&"',"&intIdFocoAT&","&intCodCampana&","&strTramoVenc&","&strTramoMonto&",'"&strSucursal&"'"

			'Response.write "strSql = " & strSql

			SET rsGES=Conn.execute(strSql)

			Do while not rsGES.eof

			strContacto = rsges("NOM_CONTACTO_GESTION")

			Obs=UCASE(Trim(rsges("OBSERVACIONES")))
			Obs=Trim(rsges("OBSERVACIONES"))

			If Obs="" then
				Obs="SIN INFORMACION ADICIONAL"
			End if

			intGestionGtel = rsges("PRIORIDAD_GTEL")
			intGestionGmail = rsges("PRIORIDAD_GMAIL")
			intGestionComunica = rsges("COMUNICA")

            Gst_EsInterno                       = rsges("EsInterno")    
            Gst_EsExterno                       = rsges("EsExterno")    
            Gst_PuedenEscucharMisGrabaciones    = rsges("PuedenEscucharMisGrabaciones")    
            Gst_PuedoEscucharGrabaciones        = rsges("PuedoEscucharGrabaciones")   
            Ubicabilidad                        = rsges("Ubicabilidad")
            StrAnexo                            = rsges("anexo")

			contador = contador + 1%>
				<tr>
					<td><%=contador%></font></td>
					<td class="DatosDeudorTexto" ><font class="TextoDatos"><%=rsges("NOMBRE_FOCO")%></font></td>
					<td class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("FECHA_INGRESO")%></font></td>
					<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("HORA_INGRESO")%></font></td>

					<td class="DatosDeudorTexto" ><font class="TextoDatos">
					<A HREF="detalle_gestiones.asp?rut=<%=rsges("RUT_DEUDOR")%>&cliente=<%=rsges("COD_CLIENTE")%>"><acronym title="Llevar a pantalla de Ingreso de Gestión"><%=rsges("RUT_DEUDOR")%></acronym></A></font></td>

					
					<td title="<%=rsges("NOMBRE_DEUDOR")%>">
						<%=Mid(rsges("NOMBRE_DEUDOR"),1,15)%></font></td>
						
					<td><%=FN(rsges("MONTO_GESTIONADO"),0)%></td>
					<td align="center"><%=FN(rsges("TOTAL_DOC"),0)%></td>
					<td align="center"><%=FN(rsges("DMING"),0)%></td>
					<td align= "left" class="Estilo4" title="<%=rsges("GESTION")%>">
						<%=rsges("DES3")%></font></td>
					<td
					  <%If intGestionGmail = 1 Then%>
					  align= "center" class="Estilo4" title="<%=rsges("EMAIL_ASOCIADO")%>">
					  <img src="../imagenes/Arroa.png" border="0">

					  <%ElseIf intGestionGtel = 1 Then%>
					  
					  <td align="center"><%=rsges("TELEFONO_ASOCIADO")%></td>

					  <%Else%>
					   >&nbsp;
					  <%End If%>
					</td>

					<td align="center"
						<%If intGestionComunica = 0 OR intGestionComunica = 2 AND intGestionGmail = 1 Then%>
						class="Estilo4" title="<%=strContacto%>">
						<img src="../imagenes/Contacto.rojo.png" border="0">

						<%ElseIf intGestionComunica = 0 OR intGestionComunica = 2 AND intGestionGtel = 1 Then%>
						class="Estilo4" title="<%=strContacto%>">
						<img src="../imagenes/Contacto.rojo.png" border="0">

						<%ElseIf intGestionComunica = 1 AND intGestionGtel = 1 Then%>
						class="Estilo4" title="<%=strContacto%>">
						<img src="../imagenes/Contacto.azul.png" border="0">

						<%ElseIf intGestionComunica = 1 AND intGestionGmail = 1 Then%>
						class="Estilo4" title="<%=strContacto%>">
						<img src="../imagenes/Contacto.azul.png" border="0">

						<%Else%>
						Width = "40">&nbsp;
						<%End If%>

					</td>

					<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("LOGIN")%></font></td>
					<td class="DatosDeudorTexto" align="center"><font class="TextoDatos"><%=rsges("FECHA_COMPROMISO")%></font></td>

					<td align="center"class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("FECHA_AGEND")%></font></td>
					<td align="center" class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("HORA_AGEND")%></font></td>
					<td align= "center" class="DatosDeudorTexto" ><font class="TextoDatos" align="center"><%= rsges("DA")%></font></td>

					<td align="center" title="<%=Obs%>">
						<img src="../imagenes/priorizar_normal.png" border="0">
					</td>

					<%if Usr_PuedoEscucharGrabaciones = "Verdadero" then         %>
                        <%if Gst_PuedenEscucharMisGrabaciones = "Verdadero" then  %>
                            <%if intCodCategoria=2 then  %>
                                    <td class="Estilo4">
                    	                <A HREF="#" onClick="TraerGrabacion('<%=rsges("TELEFONO_ASOCIADO")%>','<%=rsges("FECHA_INGRESO")%>','<%=rsges("HORA_INGRESO")%>','<%=rsges("ID_USUARIO")%>','<%=strAnexo%>')";>
							                <img src="../imagenes/sound.png" border="0">
						                </A>
					                </td>
                            <%else%>	
                                <td width="20">&nbsp;</td>
                            <% End if%>
		                <%else%>	
			                <td width="20">&nbsp;</td>
		                <% End if%>
    <%else%>	
			        <td width="20">&nbsp;</td>
                <%End if %>
            
				</tr>

				<%	rsGES.movenext
					Loop
				rsGES.close
				set rsGES=nothing

		CerrarSCG()

		End If
		%>
				<thead>	
		<tr>
			<td colspan=19>&nbsp;</td>
		</tr>
		</thead>
		
		</body>
		<%

		If intTipoListado = 0 then %>

		<thead>	
		<tr>
			<td>&nbsp;</td>
			<td align="center">Rut&nbsp;&nbsp;Deudor</td>
			<td align="center">Fecha Ing.</td>
			<td align="center">Hora Ing.</td>
			<td align="center" >Medio</td>
			<td>&nbsp;</td>
			<td align="center">Gestión</td>
			<td align="center">Ejecutivo</td>
			<td align="center">Fecha Comp.</td>
			<td align="center">Fecha Agend.</td>
			<td align="center">Hora Agend.</td>
			<td align="center">DA</td>
			<td align="center">OBS.</td>
			<td>&nbsp;</td>
		</tr>
		</thead>
		<tbody>
<%

			STRSQL = " SELECT CLIENTE1,PP.LOGIN, PP.RUT_DEUDOR,GESTIONES.ID_USUARIO,GTG.DESCRIPCION AS DES3,CLIENTE.DESCRIPCION,CONVERT(VARCHAR(10),GESTIONES.FECHA_INGRESO,103) AS FECHA_INGRESO,"
			
			
			STRSQL = STRSQL & " ISNULL(DD.TELEFONO_DAL,'&nbsp;') AS TELEFONO_ASOCIADO,ISNULL(DE.EMAIL,'&nbsp;') AS EMAIL_ASOCIADO,GTG.PRIORIDAD_GTEL,GTG.PRIORIDAD_GMAIL,ISNULL(GTG.COMUNICA,2) AS COMUNICA,"
			

			STRSQL = STRSQL & " CONVERT(VARCHAR(5),convert(datetime, GESTIONES.HORA_INGRESO), 108)  HORA_INGRESO,ISNULL(CONVERT(VARCHAR(10),GESTIONES.FECHA_COMPROMISO,103),'&nbsp;') AS FECHA_COMPROMISO,"
			STRSQL = STRSQL & " REPLACE(REPLACE(GESTIONES.OBSERVACIONES,CHAR(13),' '),CHAR(10),' ') AS OBSERVACIONES,(GTC.DESCRIPCION+'-'+GTSC.DESCRIPCION+'-'+GTG.DESCRIPCION) AS GESTION, "


			strSql= strSql & " case "
			strSql= strSql & " 	when GESTIONES.TIPO_MEDIO_GESTION = 1 then "
			strSql= strSql & " 		(SELECT UPPER(CONTACTO) AS CONTACTO FROM TELEFONO_CONTACTO WHERE ID_CONTACTO = GESTIONES.ID_CONTACTO_GESTION) "
			strSql= strSql & " 	when GESTIONES.TIPO_MEDIO_GESTION = 2 then "
			strSql= strSql & " 		(SELECT UPPER(CONTACTO) AS CONTACTO FROM EMAIL_CONTACTO WHERE ID_CONTACTO = GESTIONES.ID_CONTACTO_GESTION) "
			strSql= strSql & " 	when GESTIONES.TIPO_MEDIO_GESTION = 3 then "
			strSql= strSql & " 		(SELECT UPPER(CONTACTO) AS CONTACTO FROM DIRECCION_CONTACTO WHERE ID_CONTACTO = GESTIONES.ID_CONTACTO_GESTION)  "
			strSql= strSql & " else '' end NOM_CONTACTO_GESTION, "


			STRSQL = STRSQL & " CONVERT(VARCHAR(10),GESTIONES.FECHA_AGENDAMIENTO,103) AS FECHA_AGEND, (CASE WHEN HORA_AGENDAMIENTO = '' THEN '&nbsp;' ELSE ISNULL(GESTIONES.HORA_AGENDAMIENTO,'&nbsp;') END) AS HORA_AGEND,"
			STRSQL = STRSQL & " (CAST(GESTIONES.FECHA_AGENDAMIENTO-ISNULL(GESTIONES.FECHA_INGRESO,'01/01/1900') AS INT)) AS DA"
            STRSQL = STRSQL & " , isnull(GESTIONES.ANEXO_USUARIO,USUARIO.anexo) anexo ,USUARIO.EsInterno ,USUARIO.EsExterno "
            STRSQL = STRSQL & " ,USUARIO.PuedenEscucharMisGrabaciones,USUARIO.PuedoEscucharGrabaciones,gtg.Ubicabilidad"

			STRSQL = STRSQL & " FROM GESTIONES INNER JOIN (SELECT RUT_DEUDOR,MAX(ID_GESTION) AS ID_GESTION, MIN(TIPO_GESTION) AS TIPO_GESTION,CLIENTE1,LOGIN,ETAPA_COBRANZA"
			STRSQL = STRSQL & " FROM (SELECT GESTIONES.ID_USUARIO,GESTIONES.ID_GESTION,LOGIN,GESTIONES.RUT_DEUDOR,CLIENTE.COD_CLIENTE AS CLIENTE,ETAPA_COBRANZA,"

			STRSQL = STRSQL & " (CASE WHEN MIN(GESTIONES_TIPO_GESTION.CATEGORIZACION) IN (1,2)"
			STRSQL = STRSQL & " 	  THEN 1"
			STRSQL = STRSQL & " 	  WHEN MAX(GESTIONES_TIPO_GESTION.PRIORIDAD_GTIT)= 1"
			STRSQL = STRSQL & " 	  THEN 2"
			STRSQL = STRSQL & " 	  WHEN MIN(GESTIONES_TIPO_GESTION.CATEGORIZACION)= 8"
			STRSQL = STRSQL & " 	  THEN 3"
			STRSQL = STRSQL & " 	  WHEN MIN(GESTIONES_TIPO_GESTION.CATEGORIZACION)= 9"
			STRSQL = STRSQL & " 	  THEN 4"
			STRSQL = STRSQL & " 	  WHEN MIN(GESTIONES_TIPO_GESTION.CATEGORIZACION) = 10"
			STRSQL = STRSQL & " 	  THEN 5"
			STRSQL = STRSQL & " 	  WHEN MIN(GESTIONES_TIPO_GESTION.CATEGORIZACION) = 11"
			STRSQL = STRSQL & " 	  THEN 6"
			STRSQL = STRSQL & " 	  ELSE 9"
			STRSQL = STRSQL & " 	  END) AS TIPO_GESTION,"

			STRSQL = STRSQL & " CLIENTE.COD_CLIENTE AS CLIENTE1,ROW_NUMBER() OVER(PARTITION BY GESTIONES_CUOTA.ID_CUOTA "
			STRSQL = STRSQL & " ORDER BY GESTIONES.ID_GESTION DESC) AS CORR "
			STRSQL = STRSQL & " FROM GESTIONES INNER JOIN GESTIONES_TIPO_GESTION ON GESTIONES.COD_CATEGORIA = GESTIONES_TIPO_GESTION.COD_CATEGORIA AND "
			STRSQL = STRSQL & " GESTIONES.COD_SUB_CATEGORIA = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA AND GESTIONES.COD_GESTION = GESTIONES_TIPO_GESTION.COD_GESTION "
			STRSQL = STRSQL & " AND GESTIONES.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE "

			STRSQL = STRSQL & " INNER JOIN USUARIO ON GESTIONES.ID_USUARIO = USUARIO.ID_USUARIO"
			STRSQL = STRSQL & " INNER JOIN CLIENTE ON GESTIONES.COD_CLIENTE = CLIENTE.COD_CLIENTE "
			STRSQL = STRSQL & " INNER JOIN DEUDOR ON DEUDOR.RUT_DEUDOR = GESTIONES.RUT_DEUDOR AND DEUDOR.COD_CLIENTE = GESTIONES.COD_CLIENTE "
			STRSQL = STRSQL & " INNER JOIN GESTIONES_CUOTA ON GESTIONES.ID_GESTION = GESTIONES_CUOTA.ID_GESTION "

			strSql = strSql & " WHERE GESTIONES.FECHA_INGRESO BETWEEN CAST( '" & strFechaInicio & "' AS DATETIME) AND CAST( '" & strFechaTermino & "' AS DATETIME) AND ((PERFIL_COB = 1) OR PERFIL_BACK = 1) AND GESTIONES.COD_CLIENTE IN (" & strCodCliente & ")"

			If Trim(strCobranza) = "INTERNA" Then
				strSql = strSql & " AND GESTIONES.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCodCliente & "' AND CUSTODIO IS NOT NULL)"
			End if

			If Trim(strCobranza) = "EXTERNA" Then
				strSql = strSql & " AND GESTIONES.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCodCliente & "' AND CUSTODIO IS NULL)"
			End if

			strSql = strSql & " GROUP BY LOGIN,CLIENTE.COD_CLIENTE,ETAPA_COBRANZA,GESTIONES.RUT_DEUDOR,GESTIONES_CUOTA.ID_CUOTA,GESTIONES.ID_GESTION,GESTIONES.ID_USUARIO) AS GESTIONES "
			strSql = strSql & " WHERE CORR = 1"

			If intCodUsuario <> "" then
			strSql = strSql & " AND GESTIONES.ID_USUARIO IN (" & intCodUsuario & ")"
			End If


			strSql = strSql & " GROUP BY CLIENTE1,ETAPA_COBRANZA,GESTIONES.RUT_DEUDOR,LOGIN,GESTIONES.ID_USUARIO"

			If intCodTipoGes <> "" then
			strSql = strSql & " HAVING MIN(TIPO_GESTION) = " & intCodTipoGes
			End If

			strSql = strSql & " ) AS PP ON GESTIONES.ID_GESTION = PP.ID_GESTION"

			strSql = strSql & " INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON GTC.COD_CATEGORIA = GESTIONES.COD_CATEGORIA"
			strSql = strSql & " INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON GTSC.COD_CATEGORIA = GESTIONES.COD_CATEGORIA AND GTSC.COD_SUB_CATEGORIA = GESTIONES.COD_SUB_CATEGORIA"
			strSql = strSql & " INNER JOIN GESTIONES_TIPO_GESTION GTG ON GTG.COD_CATEGORIA = GESTIONES.COD_CATEGORIA AND GTG.COD_SUB_CATEGORIA = GESTIONES.COD_SUB_CATEGORIA AND GTG.COD_GESTION = GESTIONES.COD_GESTION AND GTG.COD_CLIENTE = GESTIONES.COD_CLIENTE"
			
			STRSQL = STRSQL & "	LEFT JOIN DEUDOR_TELEFONO DD ON DD.ID_TELEFONO = GESTIONES.ID_MEDIO_GESTION "			
			STRSQL = STRSQL & "	LEFT JOIN DEUDOR_EMAIL DE ON DE.ID_EMAIL = GESTIONES.ID_MEDIO_GESTION "

			strSql = strSql & " INNER JOIN CLIENTE ON PP.CLIENTE1 = CLIENTE.COD_CLIENTE"
            strSql = strSql & " INNER JOIN USUARIO ON GESTIONES.ID_USUARIO = USUARIO.ID_USUARIO "
			strSql = strSql & " ORDER BY CAST(GESTIONES.FECHA_INGRESO + ' ' + GESTIONES.HORA_INGRESO AS DATETIME) DESC"

			'Response.write "strSql = " & strSql

		AbrirSCG()

			SET rsGES=Conn.execute(strSql)

			Do while not rsGES.eof

					strContacto = rsges("NOM_CONTACTO_GESTION")


			Obs=UCASE(Trim(rsges("OBSERVACIONES")))
			Obs=Trim(rsges("OBSERVACIONES"))

            Gst_EsInterno                       = rsges("EsInterno")    
            Gst_EsExterno                       = rsges("EsExterno")    
            Gst_PuedenEscucharMisGrabaciones    = rsges("PuedenEscucharMisGrabaciones")    
            Gst_PuedoEscucharGrabaciones        = rsges("PuedoEscucharGrabaciones")   
            Ubicabilidad                        = rsges("Ubicabilidad")
            StrAnexo                            = rsges("anexo")

			If Obs="" then
				Obs="SIN INFORMACION ADICIONAL"
			End if

			intGestionGtel = rsges("PRIORIDAD_GTEL")
			intGestionGmail = rsges("PRIORIDAD_GMAIL")
			intGestionComunica = rsges("COMUNICA")
			contador = contador + 1%>
				<tr>
					<td><%=contador%></font></td>
					<td class="DatosDeudorTexto" ><font class="TextoDatos">
					<A HREF="principal.asp?TX_RUT=<%=rsges("RUT_DEUDOR")%>"><acronym title="Llevar a pantalla de Ingreso de Gestión"><%=rsges("RUT_DEUDOR")%></acronym></A>
					</font></td>
					<td class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("FECHA_INGRESO")%></font></td>
					<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("HORA_INGRESO")%></font></td>

					<td
					  <%If intGestionGmail = 1 Then%>
					  align= "center" class="Estilo4" title="<%=rsges("EMAIL_ASOCIADO")%>">
					  <img src="../imagenes/Arroa.png" border="0">

					  <%ElseIf intGestionGtel = 1 Then%>

					  align= "center" class="Estilo4" title="<%=rsges("TELEFONO_ASOCIADO")%>">
					  <img src="../imagenes/mod_telefono_va.png" border="0">

					  <%Else%>
					   >&nbsp;
					  <%End If%>
					</td>

					<td align="center"
						<%If intGestionComunica = 0 AND intGestionGmail = 1 Then%>
						class="Estilo4" title="<%=strContacto%>">
						<img src="../imagenes/Contacto.rojo.png" border="0">

						<%ElseIf intGestionComunica = 0 AND intGestionGtel = 1 Then%>
						class="Estilo4" title="<%=strContacto%>">
						<img src="../imagenes/Contacto.rojo.png" border="0">

						<%ElseIf intGestionComunica = 1 AND intGestionGtel = 1 Then%>
						class="Estilo4" title="<%=strContacto%>">
						<img src="../imagenes/Contacto.azul.png" border="0">

						<%ElseIf intGestionComunica = 1 AND intGestionGmail = 1 Then%>
						class="Estilo4" title="<%=strContacto%>">
						<img src="../imagenes/Contacto.azul.png" border="0">

						<%Else%>
						Width = "40">&nbsp;
						<%End If%>

					</td>

					<td align= "left" class="Estilo4" title="<%=rsges("GESTION")%>">
						<%=rsges("DES3")%></font></td>

					<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("LOGIN")%></font></td>
					<td class="DatosDeudorTexto"><font class="TextoDatos"><%=rsges("FECHA_COMPROMISO")%></font></td>

					<td class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("FECHA_AGEND")%></font></td>
					<td class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("HORA_AGEND")%></font></td>
					<td align= "center" class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("DA")%></font></td>

					<td align="center" title="<%=Obs%>">
						<img src="../imagenes/priorizar_normal.png" border="0">

					</td>

                    <%if Usr_PuedoEscucharGrabaciones = "Verdadero" then         %>
                        <%if Gst_PuedenEscucharMisGrabaciones = "Verdadero" then  %>
                            <%if Ubicabilidad = "1" or Ubicabilidad = "2" or Ubicabilidad = "3"   then  %>
                                    <td class="Estilo4">
                    	                <A HREF="#" onClick="TraerGrabacion('<%=rsges("TELEFONO_ASOCIADO")%>','<%=rsges("FECHA_INGRESO")%>','<%=rsges("HORA_INGRESO")%>','<%=rsges("ID_USUARIO")%>','<%=strAnexo%>')";>
							                <img src="../imagenes/sound.png" border="0"> 
						                </A>
					                </td>
                            <%else%>	
                                <td width="20">&nbsp;</td>
                            <% End if%>
		                <%else%>	
			                <td width="20">&nbsp;</td>
		                <% End if%>
                    <%else%>	
			        <td width="20">&nbsp;</td>
                <%End if %>

				</tr>

				<%	rsGES.movenext
					Loop
				rsGES.close
				set rsGES=nothing

		CerrarSCG()

		End If%>

		</table>

</form>
</body>
</html>

<script language="javascript">
$(document).ready(function(){
		$(document).tooltip();
	})
function envia(){
			datos.action = "Detalle_informe_gestiones_2.asp?intCodTipoGes=<%=intCodTipoGes%>&strLogin=<%=intCodUsuario%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&intCodCategoria=<%=intCodCategoria%>&intCodSubCategoria=<%=intCodSubCategoria%>&intCodGestion=<%=intCodGestion%>";
			datos.submit()
}

function TraerGrabacion(strTelefono, strFecIngreso, strHoraIngreso, intIdusuario, strAnexo) {
    URL = 'EscucharGrabacion.asp?strTelefono=' + strTelefono + '&strFecIngreso=' + strFecIngreso + '&strHoraIngreso=' + strHoraIngreso + '&intIdusuario=' + intIdusuario + '&strAnexo=' + strAnexo
	window.open(URL,"DATOS_GRABACION","width=470, height=230, scrollbars=no, menubar=no, location=no, resizable=yes")
}

</script>
