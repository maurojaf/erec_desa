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

	strFechaInicio 	= request("strFechaInicio")
	strFechaTermino = request("strFechaTermino")
	intCodTipoGes 	= request("intCodTipoGes")
	strEjeAsig 		= request("strEjeAsig")
	strCodCliente 	= request("strCodCliente")
	strCobranza 	= Request("strCobranza")

	AbrirSCG()

	if Trim(strFechaInicio) = "" Then
		strFechaInicio = TraeFechaActual(Conn)
	End If

	if Trim(strFechaTermino) = "" Then
		strFechaTermino = TraeFechaActual(Conn)

	Else strFechaTermino = strFechaTermino

	End If



	%>
	<title>Detalle Compromisos</title>
</head>
<body>
<form name="datos" method="post">
	<div class="titulo_informe">LISTADO GESTIONES</div>
		<table width="90%" border="0" cellspacing="0" cellpadding="0">
			<tr  class="Estilo20">
				<td align="right" ><input type="button"value="Volver" class="fondo_boton_100" onclick="javascript:history.back();"></td>
			</tr>
		</table>
		<br>

		<table class="intercalado" >
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

<%



				strSql = "SELECT G.COD_CLIENTE AS CLIENTE1,(GTC.DESCRIPCION+'-'+GTSC.DESCRIPCION+'-'+GTG.DESCRIPCION) AS GESTION,"
				strSql= strSql & " 	(CAST(G.FECHA_AGENDAMIENTO-ISNULL(G.FECHA_INGRESO,'01/01/1900') AS INT)) AS DA,GTG.DESCRIPCION AS DES3,ISNULL(CONVERT(VARCHAR(10),G.FECHA_AGENDAMIENTO,103),'&nbsp;') AS FECHA_AGEND,"
				strSql= strSql & " 	(CASE WHEN HORA_AGENDAMIENTO = '' THEN '&nbsp;' ELSE ISNULL(G.HORA_AGENDAMIENTO,'&nbsp;') END) AS HORA_AGEND, "

				strSql= strSql & " case "
				strSql= strSql & " 	when G.TIPO_MEDIO_GESTION = 1 then "
				strSql= strSql & " 		(SELECT UPPER(CONTACTO) AS CONTACTO FROM TELEFONO_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION) "
				strSql= strSql & " 	when G.TIPO_MEDIO_GESTION = 2 then "
				strSql= strSql & " 		(SELECT UPPER(CONTACTO) AS CONTACTO FROM EMAIL_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION) "
				strSql= strSql & " 	when G.TIPO_MEDIO_GESTION = 3 then "
				strSql= strSql & " 		(SELECT UPPER(CONTACTO) AS CONTACTO FROM DIRECCION_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION)  "
				strSql= strSql & " else '' end NOM_CONTACTO_GESTION ,REPLACE(REPLACE(G.OBSERVACIONES,CHAR(13),' '),CHAR(10),' ') AS OBSERVACIONES,PRIORIDAD_GTEL,PRIORIDAD_GMAIL,COMUNICA,"
				strSql= strSql & " 	G.RUT_DEUDOR,CONVERT(VARCHAR(10),G.FECHA_INGRESO,103) AS FECHA_INGRESO,				CONVERT(VARCHAR(5),convert(datetime, G.HORA_INGRESO), 108) HORA_INGRESO,"
				strSql= strSql & " 	ISNULL(DE.EMAIL,'&nbsp;') AS EMAIL_ASOCIADO,LOGIN,ISNULL(CONVERT(VARCHAR(10),G.FECHA_COMPROMISO,103),'&nbsp;') AS FECHA_COMPROMISO,"
				

				strSql= strSql & " 	ISNULL(DD.TELEFONO,'&nbsp;') AS TELEFONO_ASOCIADO ,G.ID_USUARIO"
				
				

				strSql= strSql & " 	FROM GESTIONES G	INNER JOIN DEUDOR D ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.COD_CLIENTE = D.COD_CLIENTE"
				strSql= strSql & " 						LEFT JOIN CAJA_FORMA_PAGO CFP ON G.FORMA_PAGO = CFP.ID_FORMA_PAGO"
				strSql= strSql & " 						LEFT JOIN USUARIO U ON G.ID_USUARIO = U.ID_USUARIO"
				strSql = strSql & " 					INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON GTC.COD_CATEGORIA = G.COD_CATEGORIA"
				strSql = strSql & " 					INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON GTSC.COD_CATEGORIA = G.COD_CATEGORIA AND GTSC.COD_SUB_CATEGORIA = G.COD_SUB_CATEGORIA"
				strSql = strSql & " 					INNER JOIN GESTIONES_TIPO_GESTION GTG ON GTG.COD_CATEGORIA = G.COD_CATEGORIA AND GTG.COD_SUB_CATEGORIA = G.COD_SUB_CATEGORIA AND GTG.COD_GESTION = G.COD_GESTION AND GTG.COD_CLIENTE = G.COD_CLIENTE"
				strSql= strSql & " 						INNER JOIN"
				strSql= strSql & " (SELECT G.ID_GESTION,"
				strSql= strSql & " (CASE WHEN COUNT(C.ID_CUOTA) = SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END))"
				strSql= strSql & " 	  THEN 'COMPROMISOS CUMPLIDOS'"
				strSql= strSql & " 	  WHEN SUM((CASE WHEN (ED.GRUPO='PAGADOS' AND G.ID_GESTION = C.ID_ULT_GEST_CP) THEN 1 ELSE 0 END)) > 0"
				strSql= strSql & " 	  THEN 'COMPROMISOS PARC. CUMPLIDOS'"
				strSql= strSql & " 	  ELSE 'COMPROMISOS NO CUMPLIDOS'"
 				strSql= strSql & " END) AS ESTADO_GESTION_CP"
				strSql= strSql & " FROM GESTIONES G		INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION"
				
				strSql= strSql & " 						INNER JOIN CUOTA C ON GC.ID_CUOTA = C.ID_CUOTA"
				strSql= strSql & " 						INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA AND"
				strSql= strSql & " G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA AND"
				strSql= strSql & " G.COD_GESTION = GTG.COD_GESTION AND"
				strSql= strSql & " 	G.COD_CLIENTE = GTG.COD_CLIENTE"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_CATEGORIA GTC ON G.COD_CATEGORIA = GTC.COD_CATEGORIA"
				strSql= strSql & " 						INNER JOIN GESTIONES_TIPO_SUBCATEGORIA GTSC ON G.COD_CATEGORIA = GTSC.COD_CATEGORIA AND "
				strSql= strSql & " G.COD_SUB_CATEGORIA = GTSC.COD_SUB_CATEGORIA"

				strSql= strSql & " WHERE GTG.CATEGORIZACION IN (1,2) AND G.COD_CLIENTE IN (" & strCodCliente &")"

				If strEjeAsig <> "" then

				strSql = strSql & " AND  G.ID_USUARIO = " & strEjeAsig

				End If

				If strFechaInicio <> "" then

				strSql = strSql & " AND  G.FECHA_COMPROMISO > = '" & strFechaInicio & " 00:00:00'"

				End If

				If strFechaTermino <> "" then

				strSql = strSql & " AND G.FECHA_COMPROMISO < = '" & strFechaTermino & " 23:59:59'"

				End If

				If Trim(strCobranza) = "INTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NOT NULL"
				End if

				If Trim(strCobranza) = "EXTERNA" Then
					strSql = strSql & " AND C.CUSTODIO IS NULL"
				End if
					
				strSql= strSql & " GROUP BY G.ID_GESTION) AS PP ON G.ID_GESTION=PP.ID_GESTION "

				strSql= strSql & " 	LEFT JOIN DEUDOR_TELEFONO DD ON DD.ID_TELEFONO = G.ID_MEDIO_GESTION "
				strSql= strSql & " 	LEFT JOIN DEUDOR_EMAIL DE ON DE.ID_EMAIL = G.ID_MEDIO_GESTION "

				If intCodTipoGes="1" then
				
				strSql= strSql & " WHERE PP.ESTADO_GESTION_CP = 'COMPROMISOS NO CUMPLIDOS'"
				
				End If

				If intCodTipoGes="2" then
				
				strSql= strSql & " WHERE PP.ESTADO_GESTION_CP = 'COMPROMISOS PARC. CUMPLIDOS'"
				
				End If

				If intCodTipoGes="3" then
				
				strSql= strSql & " WHERE PP.ESTADO_GESTION_CP = 'COMPROMISOS CUMPLIDOS'"
				
				End If
				
				'Response.write "<br>strSql=" & strSql



			set rsGES=Conn.execute(strSql)

			Do while not rsGES.eof


			strContacto = Trim(rsges("NOM_CONTACTO_GESTION"))




			Obs=UCASE(Trim(rsges("OBSERVACIONES")))
			Obs=Trim(rsges("OBSERVACIONES"))

			If Obs="" then
				Obs="SIN INFORMACION ADICIONAL"
			End if

			intGestionGtel = rsges("PRIORIDAD_GTEL")
			intGestionGmail = rsges("PRIORIDAD_GMAIL")
			intGestionComunica = rsges("COMUNICA")
			contador = contador + 1%>
				<tbody>
				<tr>
					<td><%=contador%></font></td>
					<td class="DatosDeudorTexto" ><font class="TextoDatos">
					<A HREF="principal.asp?TX_RUT=<%=rsges("RUT_DEUDOR")%>"><acronym title="Llevar a pantalla de selección"><%=rsges("RUT_DEUDOR")%></acronym></A>
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

					<td class="Estilo4">
						<A HREF="#" onClick="TraerGrabacion('<%=rsges("TELEFONO_ASOCIADO")%>','<%=rsges("FECHA_INGRESO")%>','<%=rsges("HORA_INGRESO")%>','<%=rsges("ID_USUARIO")%>')";>
							<img src="../imagenes/sound.png" border="0">
						</A>
					</td>

				</tr>
				</tbody>

				<%	rsGES.movenext
					Loop
				rsGES.close
				set rsGES=nothing

		CerrarSCG()%>

		</table>

</form>
</body>
</html>
<script type="text/javascript">
$(document).ready(function(){
		$(document).tooltip();
	})
</script>
<script language="javascript">

function envia(){
			datos.action = "Detalle_informe_gestiones_2.asp?intCodTipoGes=<%=intCodTipoGes%>&strLogin=<%=intCodUsuario%>&strFechaInicio=<%=strFechaInicio%>&strFechaTermino=<%=strFechaTermino%>&strCodCliente=<%=strCodCliente%>&intCodCategoria=<%=intCodCategoria%>&intCodSubCategoria=<%=intCodSubCategoria%>&intCodGestion=<%=intCodGestion%>";
			datos.submit()
}

function TraerGrabacion (strTelefono,strFecIngreso,strHoraIngreso,intIdusuario){
	URL='EscucharGrabacion.asp?strTelefono=' + strTelefono + '&strFecIngreso=' + strFecIngreso + '&strHoraIngreso=' + strHoraIngreso + '&intIdusuario=' + intIdusuario
	window.open(URL,"DATOS_GRABACION","width=300, height=230, scrollbars=no, menubar=no, location=no, resizable=yes")
}

</script>
