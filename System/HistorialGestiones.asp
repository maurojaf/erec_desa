<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="sesion_inicio.asp"-->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/lib.asp"-->

  	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<%
	Response.CodePage = 65001
	Response.charset="utf-8"

	strFiltro = request("CB_FILTRO")

	If Trim(strFiltro) = "" Then 
		strFiltro = "EFECTIVAS ACTIVAS"

	end if

	strRutDeudor 	=Request("strRUT_DEUDOR")
	strCodCliente 	=Request("strCOD_CLIENTE")
%>

</head>
<body>
<form name="FrmHistorial" method="post">
<div class="subtitulo_informe" style="float:left;">
	> HISTORIAL DE GESTIONES
</div>
<div style="float:right;">
	<span class="subtitulo_informe">> Filtro Gest. Telefonica</span>
	<select name="CB_FILTRO" id="CB_FILTRO" onChange="Refrescar();">
		<option value="TODAS" <%If strFiltro="TODAS" Then response.write "SELECTED"%> >TODAS</option>
		<option value="ACTIVAS" <%If strFiltro="ACTIVAS" Then response.write "SELECTED"%>>ACTIVAS</option>
		<option value="EFECTIVAS" <%If strFiltro="EFECTIVAS" Then response.write "SELECTED"%>>EFECTIVAS</option>
		<option value="EFECTIVAS ACTIVAS" <%If strFiltro="EFECTIVAS ACTIVAS" Then response.write "SELECTED"%>>EFECTIVAS ACTIVAS</option>
	</select>		
</div>	
<BR>

	  <%

		strSql="SELECT MAX( CAST((CONVERT(VARCHAR(10),G.FECHA_INGRESO,103)+' '+G.HORA_INGRESO) AS DATETIME)) AS MAX_FECHA_GES_TIT"
		strSql=strSql + " FROM GESTIONES G INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON G.COD_CATEGORIA = GTC.COD_CATEGORIA "
		strSql=strSql + " 				 INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA" 
		strSql=strSql + " 																   AND G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA "
		strSql=strSql + " 				 INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
		strSql=strSql + " 														   AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA"
		strSql=strSql + " 														   AND G.COD_GESTION = GTG.COD_GESTION"
		strSql=strSql + " 														   AND GTG.COD_CLIENTE = '" & strCodCliente & "'"
		strSql=strSql + " 				 INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION "
		strSql=strSql + " 				 INNER JOIN CUOTA C ON C.ID_CUOTA = GC.ID_CUOTA AND GC.ID_GESTION = G.ID_GESTION "
		strSql=strSql + " 										AND C.COD_CLIENTE = '" & strCodCliente & "'"
		strSql=strSql + " 				 INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"

		strSql=strSql + " WHERE G.COD_CLIENTE = '" & strCodCliente & "' AND G.RUT_DEUDOR = '" & strRutDeudor & "' AND ACTIVO=1 AND ISNULL(GTG.PRIORIDAD_GTIT,0)=1"
		
		AbrirSCG()


		
		set RsFec=Conn.execute(strSql)
		If not RsFec.eof then
			dtmMaxFecTitular = RsFec("MAX_FECHA_GES_TIT")
		End if
		RsFec.close
		set RsFec=nothing		
		
		CerrarSCG()
		
	  if Trim(strRutDeudor) <> "" then

	  AbrirSCG()
			
		strSql="SELECT TOP 25 PP.ID_GESTION,PP.DESCRIPCION AS DESCRIP,"
		strSql=strSql + " G.COD_CATEGORIA, G.COD_SUB_CATEGORIA, G.COD_GESTION , G.ID_GESTION, G.FECHA_INGRESO," 
		strSql=strSql + " CONVERT(VARCHAR(5),convert(datetime, G.HORA_INGRESO), 108) HORA_INGRESO,"
		strSql=strSql + " G.ID_USUARIO,CONVERT(VARCHAR(10),G.FECHA_COMPROMISO,103) AS FECHA_COMPROMISO, "
		strSql=strSql + " IsNull(DATEDIFF(DAY,GETDATE(),G.FECHA_COMPROMISO),0) AS DIFFDIAS,G.FECHA_AGENDAMIENTO, G.HORA_AGENDAMIENTO, "
		strSql=strSql + " REPLACE(REPLACE(G.OBSERVACIONES,char(13),' '),char(10),' ') as OBSERVACIONES, "
		strSql=strSql + " PRIORIDAD_GTEL,PRIORIDAD_GMAIL,GESTION_MODULOS,COMUNICA, G.TIPO_MEDIO_GESTION, "

		'MUESTRA TIPO MEDIO
		strSql=strSql + " case "
		strSql=strSql + " 	when G.TIPO_MEDIO_GESTION = 1 then "
		strSql=strSql + " 		(SELECT CONVERT(VARCHAR,COD_AREA)+'-'+CONVERT(VARCHAR,TELEFONO) "
		strSql=strSql + " 		FROM DEUDOR_TELEFONO DT "
		strSql=strSql + " 		WHERE DT.ID_TELEFONO=G.ID_MEDIO_GESTION) "				
		strSql=strSql + " 	when G.TIPO_MEDIO_GESTION = 2 then "
		strSql=strSql + " 		(SELECT UPPER(EMAIL) AS EMAIL "
		strSql=strSql + " 		FROM DEUDOR_EMAIL DE WHERE DE.ID_EMAIL=G.ID_MEDIO_GESTION) "		
		strSql=strSql + " 	when G.TIPO_MEDIO_GESTION = 3 then "
		strSql=strSql + " 		(SELECT REPLACE(CALLE + ' ' + NUMERO + ' ' + RESTO + ' ' + COMUNA,' ',' ') "
		strSql=strSql + " 		FROM DEUDOR_DIRECCION DD WHERE DD.ID_DIRECCION=G.ID_MEDIO_GESTION) "
		strSql=strSql + " else '' end NOM_MEDIO_GESTION, "

		strSql=strSql + " case "
		strSql=strSql + " 	when G.TIPO_MEDIO_GESTION = 1 then "
		strSql=strSql + " 		(SELECT CONTACTO "
		strSql=strSql + " 		FROM TELEFONO_CONTACTO TC "
		strSql=strSql + " 		WHERE TC.ID_CONTACTO= G.ID_CONTACTO_GESTION) "				
		strSql=strSql + " 	when G.TIPO_MEDIO_GESTION = 2 then "
		strSql=strSql + " 		(SELECT CONTACTO "
		strSql=strSql + " 		FROM EMAIL_CONTACTO TC "
		strSql=strSql + " 		WHERE TC.ID_CONTACTO= G.ID_CONTACTO_GESTION) "			
		strSql=strSql + " 	when G.TIPO_MEDIO_GESTION = 3 then "
		strSql=strSql + " 		(SELECT CONTACTO "
		strSql=strSql + " 		FROM DIRECCION_CONTACTO TC "
		strSql=strSql + " 		WHERE TC.ID_CONTACTO= G.ID_CONTACTO_GESTION) "	
		strSql=strSql + " else '' end NOM_CONTACTO_GESTION "



		strSql=strSql + " FROM"
		strSql=strSql + " (SELECT DISTINCT G.ID_GESTION,(GTC.DESCRIPCION+'-'+GTSC.DESCRIPCION+'-'+GTG.DESCRIPCION) AS DESCRIPCION,"
		strSql=strSql + " ISNULL(GTG.PRIORIDAD_GTEL,0) AS PRIORIDAD_GTEL,ISNULL(GTG.PRIORIDAD_GMAIL,0) AS PRIORIDAD_GMAIL,"
		strSql=strSql + " ISNULL(GTG.GESTION_MODULOS,0) AS GESTION_MODULOS,ISNULL(GTG.COMUNICA,0) AS COMUNICA"
		strSql=strSql + " FROM GESTIONES G INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON G.COD_CATEGORIA = GTC.COD_CATEGORIA" 
		strSql=strSql + " 				   INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA "
		strSql=strSql + " 																   AND G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA "
		
		strSql=strSql + " 				   INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
		strSql=strSql + " 														   AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA"
		strSql=strSql + " 														   AND G.COD_GESTION = GTG.COD_GESTION"
		strSql=strSql + " 														   AND GTG.COD_CLIENTE = '" & strCodCliente & "'"
		strSql=strSql + " 				   INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION "
		strSql=strSql + " 				   INNER JOIN CUOTA C ON C.ID_CUOTA = GC.ID_CUOTA AND GC.ID_GESTION = G.ID_GESTION "
		strSql=strSql + " 										AND C.COD_CLIENTE = '" & strCodCliente & "'"
		strSql=strSql + " 				   INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"

		strSql=strSql + " WHERE G.COD_CLIENTE = '" & strCodCliente & "' AND G.RUT_DEUDOR = '" & strRutDeudor & "'"

			if Trim(strFiltro) = "EFECTIVAS ACTIVAS" Then 'Todas las vigente menos las no comunica inferiores a la ultima fecha de gestion efectiva.
				strSql=strSql & " AND NOT (PRIORIDAD_GTEL = '1' AND PRIORIDAD_GEFE = '0' AND CAST((CONVERT(VARCHAR(10),G.FECHA_INGRESO,103)+' '+CONVERT(VARCHAR(5),convert(datetime, G.HORA_INGRESO), 108) ) AS DATETIME) < CAST('" & dtmMaxFecTitular & "' AS DATETIME))"
				strSql=strSql & " AND NOT (PRIORIDAD_GTEL = '1' AND PRIORIDAD_GEFE = '0' AND CAST((CONVERT(VARCHAR(10),G.FECHA_INGRESO,103)+' '+G.HORA_INGRESO) AS DATETIME) < CAST('" & dtmMaxFecTitular & "' AS DATETIME))"
				strSql=strSql & " AND ED.ACTIVO=1"
			End If

			if Trim(strFiltro) = "EFECTIVAS" Then 'Todas las gestiones menos las no comunica
				strSql=strSql & " AND NOT (PRIORIDAD_GTEL = '1' AND PRIORIDAD_GEFE = '0' AND CAST((CONVERT(VARCHAR(10),G.FECHA_INGRESO,103)+' '+CONVERT(VARCHAR(5),convert(datetime, G.HORA_INGRESO), 108) ) AS DATETIME) < CAST('" & dtmMaxFecTitular & "' AS DATETIME))"
				strSql=strSql & " AND NOT (PRIORIDAD_GTEL = '1' AND PRIORIDAD_GEFE = '0' AND CAST((CONVERT(VARCHAR(10),G.FECHA_INGRESO,103)+' '+CONVERT(VARCHAR(5),convert(datetime, G.HORA_INGRESO), 108) ) AS DATETIME) < CAST('" & dtmMaxFecTitular & "' AS DATETIME))"
			End If

			if Trim(strFiltro) = "ACTIVAS" Then 'Todas las gestiones menos las no comunica
				strSql=strSql & " AND ED.ACTIVO=1"
			End If
			
		End If
		
		strSql = strSql & " ) PP INNER JOIN GESTIONES G ON PP.ID_GESTION = G.ID_GESTION"
		strSql = strSql & " 	 LEFT JOIN USUARIO U ON G.ID_USUARIO = U.ID_USUARIO"

        strSql = strSql & " ORDER BY G.FECHA_INGRESO DESC , G.ID_GESTION DESC"
		
        'Response.write "<br>Sql=" & strSql
          'Response.End
		
	  set rsDET=Conn.execute(strSql)
	  if not rsDET.eof then
	  %>
     <table border="2" bordercolor="#FFFFFF" class="intercalado" style="width:100%;">
     	<thead>
        <tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
          <td class="Estilo4">&nbsp;</td>
          <td class="Estilo4">&nbsp;</td>
          <td width = "65" class="Estilo4">FECHA</td>
          <td class="Estilo4">HORA</td>
          <td class="Estilo4">GESTION</td>
          <td class="Estilo4">F.COMP.</td>
          <td width = "65" class="Estilo4">F.AGEND</td>
          <td class="Estilo4">H.AGEND</td>
          <td class="Estilo4">OBSERVACIONES</td>
          <td width = "55" class="Estilo4">MEDIO</td>
          <td class="Estilo4">&nbsp;</td>
          <td class="Estilo4">EJECUTIVO</td>
          <td class="Estilo4">&nbsp;</td>

			<% if ((TraeSiNo(session("perfil_sup"))="Si" or  TraeSiNo(session("perfil_adm"))="Si") and TraeSiNo(session("perfil_emp"))<>"Si") Then %>
          <td class="Estilo4">&nbsp;</td>

          	<% End If %>

          </tr>
      </thead>
      <tbody>
		<%
		intCorr = 0
		do until rsDET.eof
		Obs =UCASE(Trim(rsDET("OBSERVACIONES")))
		Obs =Trim(rsDET("OBSERVACIONES"))

		If Obs="" then
			Obs ="SIN INFORMACION ADICIONAL"
		End if

		strCodGestion = rsDET("COD_CATEGORIA") & rsDET("COD_SUB_CATEGORIA") & rsDET("COD_GESTION")

		strNomGestion 		= rsDET("DESCRIP")
		intGestionComunica 	= rsDET("COMUNICA")
		intGestionGtel 		= rsDET("PRIORIDAD_GTEL")
		intGestionGmail 	= rsDET("PRIORIDAD_GMAIL")
		strTipoGestion 		= rsDET("GESTION_MODULOS")

		AbrirSCG2()
			If Trim(rsDET("ID_USUARIO")) <> "" Then
				strLoginCobrador = TraeCampoId(Conn2, "LOGIN", rsDET("ID_USUARIO"), "USUARIO", "ID_USUARIO")
			End If

		CerrarSCG2()



		AbrirSCG1()
		strSql = "SELECT G.ID_CUOTA, E.GRUPO, ISNULL(G.CONFIRMACION_CP,'N') AS CONFIRMACION_CP, NRO_DOC, ID_ULT_GEST, ROW_NUMBER() OVER(PARTITION BY C.NRO_DOC ORDER BY ISNULL(NRO_CUOTA,0) ASC) AS SUMNRO_CUOTA, ISNULL(NRO_CUOTA,0) AS NRO_CUOTA "
		strSql = strSql & " FROM CUOTA C, GESTIONES_CUOTA G, ESTADO_DEUDA E"
		strSql = strSql & " WHERE C.ID_CUOTA = G.ID_CUOTA AND C.ESTADO_DEUDA = E.CODIGO"
		strSql = strSql & " AND G.ID_GESTION = " & Trim(rsDET("ID_GESTION"))
		strSql = strSql & " ORDER BY NRO_DOC ASC"

		SET rsNomGestion = Conn1.execute(strSql)

		strSigue 		= "1"
		strMuestraIcono = "N"
		strDocNoConf 	= 0
		strDocConf 		= 0

		strNroDoc 		= ""
		strNroDocPag 	= ""
		strNroDocRet 	= ""
		strNroDocNoAsig = ""
		
		dtmFecCompromiso = rsDet("FECHA_COMPROMISO")

		If Not rsNomGestion.Eof Then
			Do While Not rsNomGestion.Eof


			If rsNomGestion("SUMNRO_CUOTA") > "1" then

			   strnrocuota = "(" & rsNomGestion("NRO_CUOTA") & ")"

			Else

			   strnrocuota = ""

			End If

				If (Trim(rsNomGestion("GRUPO")) = "ACTIVOS") Then

					strNroDoc = strNroDoc & rsNomGestion("NRO_DOC") & " " & strnrocuota

						If strSigue = "1" Then

								if ((Trim(strTipoGestion) = "11" OR Trim(strTipoGestion) = "11" ) AND Trim(rsNomGestion("ID_ULT_GEST")) = Trim(rsDET("ID_GESTION"))) Then
									strMuestraIcono = "S"
									''strSigue = "0"
									strConfirmada = ""
									If rsNomGestion("CONFIRMACION_CP") = "S" Then
										strDocConf 	= strDocConf + 1
										strNroDoc 	= strNroDoc & "(C)"
									Else
										strDocNoConf = strDocNoConf + 1
									End If

								end if

						End If
					strNroDoc = strNroDoc & " - "


				End If

				If (Trim(rsNomGestion("GRUPO")) = "PAGADOS") Then
					strNroDocPag = strNroDocPag & rsNomGestion("NRO_DOC") & " " & strnrocuota & " - "
				End If

				If (Trim(rsNomGestion("GRUPO")) = "RETIROS") Then
					strNroDocRet = strNroDocRet & rsNomGestion("NRO_DOC") & " " & strnrocuota & " - "
				End If

				If (Trim(rsNomGestion("GRUPO")) = "NO ASIGNABLE") Then
					strNroDocNoAsig = strNroDocNoAsig & rsNomGestion("NRO_DOC") & " " & strnrocuota & " - "
				End If

				rsNomGestion.Movenext

			Loop

		End If
		CerrarSCG1()

		If rsDET("DIFFDIAS") < 0 Then
			strImgConfirmar ="icon_cruz_roja.jpg"

		Elseif strDocNoConf > 0 Then
			strImgConfirmar ="icon_amarillo.jpg"

		Else
			strImgConfirmar ="bt_confirmar.jpg"

		End If

		strsql="SELECT IsNull(ANEXO,'') as ANEXO FROM USUARIO WHERE ID_USUARIO = " & Trim(rsDET("ID_USUARIO"))
		set rsUsu = Conn.execute(strsql)
		If Not rsUsu.eof then
			strAnexo 		= Trim(rsUsu("ANEXO"))
		End if

		If Trim(strNroDoc) <> "" Then
			strNroDoc 		= Mid(strNroDoc,1,Len(strNroDoc)-2)
		End If

		If Trim(strNroDocPag) <> "" Then
			strNroDocPag 	= Mid(strNroDocPag,1,Len( strNroDocPag)-2)
		End If

		If Trim(strNroDocRet) <> "" Then
			strNroDocRet 	= Mid(strNroDocRet,1,Len( strNroDocRet)-2)
		End If

		If Trim(strNroDocNoAsig) <> "" Then
			strNroDocNoAsig = Mid(strNroDocNoAsig,1,Len( strNroDocNoAsig)-2)
		End If

		intCorr = intCorr + 1

		strTextoDocAct 		= ""
		strTextoDocPag 		= ""
		strTextoDocRet 		= ""
		strTextoDocNoAsig 	= ""
		strTextoDoc 		= ""

		If Trim(strNroDoc) <> "" Then
			strTextoDocAct 		= "Doc.Asociados : " & strNroDoc & "<BR>"
		End If

		If Trim(strNroDocPag) <> "" Then
			strTextoDocPag 		= "Doc.Cancelados : " & strNroDocPag & "<BR>"
		End If

		If Trim(strNroDocRet) <> "" Then
			strTextoDocRet 		= "Doc.Desasignados : " & strNroDocRet & "<BR>"
		End If

		If Trim(strNroDocNoAsig) <> "" Then
			strTextoDocNoAsig 	= "Doc.No Asignable : " & strNroDocNoAsig & "<BR>"
		End If

		strTextoDoc = strTextoDocAct & strTextoDocPag & strTextoDocRet & strTextoDocNoAsig

		%>
        <tr bordercolor="#FFFFFF" class="Estilo8">
          <td class="Estilo4" title="<%=rsDET("ID_GESTION")%>"><%=intCorr%></td>

		 <%If strMuestraIcono = "S" Then %>
          <td class="Estilo4" title="Confirmar / desconfirmar compromiso">
				<%If strImgConfirmar="icon_cruz_roja.jpg" Then %>
					<img src="../imagenes/<%=strImgConfirmar%>" border="0">
				<%Else%>
					<A HREF="#" onClick="ConfirmarCP(<%=rsDet("ID_GESTION")%>,'<%=dtmFecCompromiso%>','<%=strCodGestion%>')";><img src="../imagenes/<%=strImgConfirmar%>" border="0"></A>
				<%End If%>
			</td>
		<%Else%>
			<td>&nbsp;</td>
		<%End If%>

          <td class="Estilo4"><%=rsDET("FECHA_INGRESO")%></td>
          <td class="Estilo4"><%=rsDET("HORA_INGRESO")%></td>
          <td class="Estilo4"><%=strNomGestion%></td>
          <td class="Estilo4"><%=rsDET("FECHA_COMPROMISO")%></td>
          <td class="Estilo4"><%=rsDET("FECHA_AGENDAMIENTO")%></td>
          <td class="Estilo4"><%=rsDET("HORA_AGENDAMIENTO")%></td>
          <td class="Estilo4" title="<%=Obs%>"><%=Mid(Obs,1,50)%></td>
		  
		  	 
		  	  <%
		  	  if trim(rsDET("NOM_MEDIO_GESTION")) <> "" AND NOT ISNULL(rsDET("NOM_MEDIO_GESTION")) then
		  	  	NOM_MEDIO_GESTION =trim(rsDET("NOM_MEDIO_GESTION"))
		  	  else
		  	  	NOM_MEDIO_GESTION ="SIN MEDIO ASOCIADO"
		  	  end if

		  	  if trim(rsDET("NOM_CONTACTO_GESTION")) <> "" AND NOT ISNULL(rsDET("NOM_CONTACTO_GESTION")) then
		  	  	NOM_CONTACTO_GESTION =trim(rsDET("NOM_CONTACTO_GESTION"))
		  	  else
		  	  	NOM_CONTACTO_GESTION ="SIN CONTACTO ASOCIADO"
		  	  end if

		  	  
		  	  If rsDET("tipo_medio_gestion") = 2 Then%>
		  	  	<td WIDTH="80" align="center" class="Estilo4" title="<%=NOM_MEDIO_GESTION%>">
		  	  	<img src="../imagenes/Arroa.png" border="0">
		  	  	</td>

	          <%ElseIf rsDET("tipo_medio_gestion") = 1 Then%>
		  	  	<td WIDTH="80" class="Estilo4" align="center"><%=rsDET("NOM_MEDIO_GESTION")%></td>

	          <%ElseIf rsDET("tipo_medio_gestion") = 3 Then%>
		  	  	<td WIDTH="80" class="Estilo4" align="center">
		  	  		 <img src="../imagenes/mod_direccion_va.png" title="<%=NOM_MEDIO_GESTION%>">
		  	  	</td>

	          <%Else%>
		  	   	<td class="Estilo4">&nbsp;</td>
		  	  <%End If%>
		  
		 

		  	  <%If intGestionComunica = 0 AND intGestionGmail = 1 Then%>
		  	   <td class="Estilo4" align="center" title="<%=NOM_CONTACTO_GESTION%>">
			   <img src="../imagenes/Contacto.rojo.png" border="0">
			  </td>

			  <%ElseIf intGestionComunica = 0 AND intGestionGtel = 1 Then%>
			  <td class="Estilo4" align="center" title="<%=NOM_CONTACTO_GESTION%>">
			  <img src="../imagenes/Contacto.rojo.png" border="0"></td>

		  	  <%ElseIf intGestionComunica = 1 AND intGestionGtel = 1 Then%>
		  	  <td class="Estilo4" align="center" title="<%=NOM_CONTACTO_GESTION%>">
		  	  <img src="../imagenes/Contacto.azul.png" border="0"></td>

		  	  <%ElseIf intGestionComunica = 1 AND intGestionGmail = 1 Then%>
		  	  <td class="Estilo4" align="center" title="<%=NOM_CONTACTO_GESTION%>">
		  	  <img src="../imagenes/Contacto.azul.png" border="0"></td>

	          <%Else%>
		  	   <td>&nbsp;</td>
		  	  <%End If%>



          	<td class="Estilo4"><%=UCASE(strLoginCobrador)%></td>
			<td class="Estilo4" title="<%=strTextoDoc%>">
				<img src="../imagenes/carpeta1.png" border="0" onclick="trae_cuotas_por_gestion('<%=trim(rsDet("ID_GESTION"))%>')">
			</td>

			<% if ((TraeSiNo(session("perfil_sup"))="Si" or  TraeSiNo(session("perfil_adm"))="Si") and TraeSiNo(session("perfil_emp"))<>"Si") and intGestionGtel = 1 and strAnexo <> "" Then %>
			<td class="Estilo4">
				<A HREF="#" onClick="TraerGrabacion('<%=rsDET("NOM_MEDIO_GESTION")%>','<%=rsDET("FECHA_INGRESO")%>','<%=rsDET("HORA_INGRESO")%>','<%=rsDET("ID_USUARIO")%>')";>
					<img src="../imagenes/sound.png" border="0">
				</A>
			</td>
			<% End if %>

        </tr>

		 <%

		 response.Flush()
		 rsDET.movenext
		 Loop%>
		<%
	  	else
	  	%>	<table style="width:100%;" >
	  		<thead>
	  		<tr>
	  			<td style="background-color:#F2F2F2;"><br>
	  	<%
			  response.Write("MENSAJE : ")
			  response.Write("EL DEUDOR NO POSEE GESTIONES REGISTRADAS PARA EL CLIENTE ")
			  response.Write(nombre_cliente)
	  	%>
	  			</td>
	  		</tr>
	  		</thead>
	  		</table>
	  	<%		  
		end if

	  rsDET.close
	  set rsDET=nothing

	  CerrarSCG()
	  %>

</tbody>
</table>
<input type="hidden" name="strRUT_DEUDOR" id="strRUT_DEUDOR" value="<%=strRutDeudor%>">
</form>
</body>
</html>

<script type="text/javascript">
$(document).ready(function(){
	$(document).tooltip();

})

function trae_cuotas_por_gestion(ID_GESTION){
	parent.trae_cuotas_por_gestion(ID_GESTION)

}
function Refrescar(){
	FrmHistorial.action='HistorialGestiones.asp?strRUT_DEUDOR=<%=strRutDeudor%>&strCOD_CLIENTE=<%=strCodCliente%>';
	FrmHistorial.submit();
}

function TraerGrabacion (strTelefono,strFecIngreso,strHoraIngreso,intIdusuario){
	URL='EscucharGrabacion.asp?strTelefono=' + strTelefono + '&strFecIngreso=' + strFecIngreso + '&strHoraIngreso=' + strHoraIngreso + '&intIdusuario=' + intIdusuario
	window.open(URL,"DATOS_GRABACION","width=300, height=230, scrollbars=no, menubar=no, location=no, resizable=yes")
}

function ConfirmarCP(id_gestion, dtmFecCompGest, intCodGestConcat)
{
	FrmHistorial.action = "paso_a_confirmar_cp.asp?id_gestion=" + id_gestion + "&rut=<%=rut%>&cliente=<%=cliente%>&dtmFecCompGest=" + dtmFecCompGest + "&intCodGestConcat=" + intCodGestConcat ;
	FrmHistorial.submit();
}

</script>

