<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

<!--#include file="sesion.asp"-->
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/lib.asp"-->

<!--#include file="../lib/comunes/rutinas/funciones.inc"-->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc"-->
<!--#include file="../lib/asp/comunes/general/Minimo.inc"-->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<link rel="stylesheet" href="../css/style_generales_sistema.css">
<%

Response.CodePage=65001
Response.charset ="utf-8"

COD_CLIENTE 	= Request.Querystring("COD_CLIENTE")
strNombres 		= Request("TX_NOMBRES")
strRut 			= Request("TX_RUT")
strDesde 		= Request("TX_DESDE")
strHasta 		= Request("TX_HASTA")
intCOD_REMESA 	= Request("CB_REMESA")
intCodCampana 	= Request("CB_CAMPANA")
strCodCliente 	= session("ses_codcli")
strEjeAsig  	= Request("CB_EJECUTIVO")
strTipoInf 		= Request("tipo_busqueda")
strRubro 		= Request("CB_RUBRO")
intEstadoCob 	= Request("CB_TIPOCOB")
intTipoDoc 		= Request("CB_TIPODOC")
tipo_busqueda 	= Request("tipo_busqueda")

if trim(COD_CLIENTE)<>"" then
	strCodCliente = COD_CLIENTE
end if

if tipo_busqueda = "" then
	tipo_busqueda = "0"
end if

'Response.write "<br>intEstadoCob=" & intEstadoCob

'--Calcula Objetos relacionados al tipo de cobranza (Interna, Externa), CB_COBRANZA Y CB_EJECUTIVO--'

strCobranza = Request("CB_COBRANZA")

abrirscg()

		strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
		strSql = strSql & " FROM CLIENTE CL"
		strSql = strSql & " WHERE CL.COD_CLIENTE = '" & strCodCliente & "'"

		set RsCli=Conn.execute(strSql)
		If not RsCli.eof then
			intUsaCobInterna = RsCli("USA_COB_INTERNA")
		End if
		RsCli.close
		set RsCli=nothing

cerrarscg()

intVerCobExt = "1"
intVerEjecutivos = "1"

If TraeSiNo(session("perfil_emp")) = "Si" and strCobranza = "" and intUsaCobInterna = "1" Then

	strCobranza="INTERNA"

ElseIf TraeSiNo(session("perfil_emp")) = "No" and strCobranza = "" then

	strCobranza="EXTERNA"

End If

If TraeSiNo(session("perfil_emp")) = "Si" Then

	intVerEjecutivos="0"

End If

If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

	sinCbUsario="0"
    'strEjeAsig = "0"

End If

'---Fin codigo tipo de cobranza---'

If Trim(Request("strBuscar")) = "S" Then
	session("FtroCA_Ejecutivo") = strEjeAsig
	session("FtroCA_Campana") = intCodCampana
	session("FtroCA_Asignacion") = intCOD_REMESA
	session("FtroCA_TipoCartera") = strTipoInf
	session("FtroCA_Rubro") = strRubro
	session("FtroCA_EstadoCob") = intEstadoCob
	session("FtroCB_Cobranza") = strCobranza


End If

If Trim(Request("strBuscar")) = "N" or Trim(Request("strLimpiar")) = "S" Then

	''Response.write "<br>strBuscar=" & Request("strBuscar")
	session("FtroCA_Ejecutivo") = ""
	session("FtroCA_Campana") = ""
	session("FtroCA_Asignacion") = ""
	session("FtroCA_TipoCartera") = ""
	session("FtroCA_Rubro") = ""
	session("FtroCA_EstadoCob") = ""
	session("FtroCB_Cobranza") = ""
End If

If Trim(Request("strLimpiar")) = "S" Then
	 strEjeAsig = ""
	 intCodCampana = ""
	 intCOD_REMESA = ""
	 strTipoInf = ""
	 strRubro = ""
	 intEstadoCob = ""
	 strCobranza = ""

End If

'Response.write "<br>FtroCA_EstadoCob=" & session("FtroCA_EstadoCob")
'Response.write "<br>strEjeAsig=" & strEjeAsig

''Response.write "<br>FtroCA_TipoCartera=" & session("FtroCA_TipoCartera")
'Response.write "<br>intUsaCobInterna=" & intUsaCobInterna
'response.write("sdakjaskdhaks -->>" & sinCbUsario) 
'response.end

If strTipoInf = "" Then strTipoInf = session("FtroCA_TipoCartera")
If strEjeAsig = "" and not (TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si") Then 
    strEjeAsig = session("session_idusuario") 'session("FtroCA_Ejecutivo")
end if 
If intCodCampana = "" Then intCodCampana = session("FtroCA_Campana")
If intCOD_REMESA = "" Then intCOD_REMESA = session("FtroCA_Asignacion")
If strRubro = "" Then strRubro = session("FtroCA_Rubro")
If intEstadoCob = "" Then intEstadoCob = session("FtroCA_EstadoCob")
If strCobranza = "" Then strCobranza = session("FtroCB_Cobranza")

'MODIFICAR AQUI PARA CAMBIAR EL Nº DE REGISTRO POR PAGINA
TamPagina=100

'Leemos qué página mostrar. La primera vez será la inicial
if Request.Querystring("pagina")="" then
	PaginaActual=1
else
	PaginaActual=CInt(Request.Querystring("pagina"))
end if


%>
<title>CARTERA ASIGNADA</title>

</head>
<%strTitulo="MI CARTERA"%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">
<div class="titulo_informe">CARTERA ASIGNADA</div>
	<table width="90%" align="CENTER" class="estilo_columnas">
			<thead>
				<tr>
				  <td>COBRANZA</td>
				  <td width="">TIPO COBRANZA</td>
				  <td width="">RUBRO</td>
				  <td width="">TIPO DOC</td>
				  <td width="">CAMPAÑA</td>

				  <% If sinCbUsario = "0" Then %>
					<td>EJECUTIVO</td>
				  <% End If %>

				  <td width="">ESTADO CARTERA</td>
			</tr>
			</thead>
				<td>
					<select name="CB_COBRANZA" style="width:130px;" <%If sinCbUsario = "0" then%> onChange="CargaUsuarios(CB_COBRANZA.value);" <%End If%> >

						<%If Trim(intVerCobExt) = "1" and Trim(intUsaCobInterna) = "1" Then%>
							<option value="0" <%If Trim(strCobranza) ="" Then Response.write "SELECTED"%>>TODOS</option>
						<%End If%>

						<%If Trim(intUsaCobInterna) = "1" Then%>
							<option value="INTERNA" <%If Trim(strCobranza) ="INTERNA" Then Response.write "SELECTED"%>>INTERNA</option>
						<%End If%>

						<%If Trim(intVerCobExt) = "1" Then%>
							<option value="EXTERNA" <%If Trim(strCobranza) ="EXTERNA" Then Response.write "SELECTED"%>>EXTERNA</option>
						<%End If%>

					</select>
				</td>

				<td>
					<select name="CB_TIPOCOB" style="width:140px;">
						<option value="">TODOS</option>
						<%
						abrirscg()
						ssql="SELECT COD_ESTADO_COBRANZA, NOM_ESTADO_COBRANZA FROM ESTADO_COBRANZA ORDER BY 1"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("COD_ESTADO_COBRANZA")%>"<%if Trim(intEstadoCob)=Trim(rsTemp("COD_ESTADO_COBRANZA")) then response.Write("Selected") End If%>><%=rsTemp("NOM_ESTADO_COBRANZA")%></option>
							<%
							rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing
						cerrarscg()
						%>
					</select>
				</td>
				<td>
					<select name="CB_RUBRO" style="width:140px;">
					<option value="" <%if Trim(strRubro)="" then response.Write("Selected") end if%>>SELECCIONE</option>
						<%
						abrirscg()
						ssql="SELECT DISTINCT ISNULL(ADIC_2,'OTRO') AS ADIC_2 FROM DEUDOR  WHERE COD_CLIENTE = '" & strCodCliente & "' ORDER BY ADIC_2"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("ADIC_2")%>"<%if strRubro=rsTemp("ADIC_2") then response.Write("Selected") End If%>><%=rsTemp("ADIC_2")%></option>
							<%
							rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing
						cerrarscg()
						%>
					</select>
				</td>
				<td>
					<select name="CB_TIPODOC" style="width:140px;">
						<option value="">TODOS</option>
						<%
						abrirscg()
						strSql="SELECT DISTINCT COD_TIPO_DOCUMENTO, NOM_TIPO_DOCUMENTO"
						strSql=strSql & " FROM CUOTA LEFT JOIN TIPO_DOCUMENTO ON TIPO_DOCUMENTO = COD_TIPO_DOCUMENTO"
						strSql=strSql & " WHERE CUOTA.COD_CLIENTE = '" & strCodCliente & "'"
						strSql=strSql & " ORDER BY NOM_TIPO_DOCUMENTO ASC"
	
						set rsTemp= Conn.execute(strSql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("COD_TIPO_DOCUMENTO")%>"<%if Trim(intTipoDoc)=Trim(rsTemp("COD_TIPO_DOCUMENTO")) then response.Write("Selected") End If%>><%=rsTemp("NOM_TIPO_DOCUMENTO")%></option>
							<%
							rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing
						cerrarscg()
						%>
					</select>
				</td>

				<td>
					<select name="CB_CAMPANA" style="width:140px;">
						<option value="">TODAS</option>
						<%
						AbrirSCG()
							strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE = '" & strCodCliente & "'"
							set rsCampana=Conn.execute(strSql)
							Do While not rsCampana.eof
								If Trim(intCodCampana)=Trim(rsCampana("ID_CAMPANA")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsCampana("ID_CAMPANA")%>" <%=strSelCam%>> <%=rsCampana("ID_CAMPANA") & " - " & rsCampana("NOMBRE")%></option>
								<%
								rsCampana.movenext
							Loop
							rsCampana.close
							set rsCampana=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>

			<% If sinCbUsario="0" Then %>
				<td>
					<select name="CB_EJECUTIVO" id="CB_EJECUTIVO" style="width:140px;">
					</select>
				</td>
			<% End If %>

			<td>
				<select name="tipo_busqueda" style="width:140px;">
						<option value="0" <%If Trim(tipo_busqueda) ="0" Then Response.write "SELECTED"%>>TODOS</option>
						<option value="1" <%If Trim(tipo_busqueda) ="1" Then Response.write "SELECTED"%>>GESTIONABLES</option>
						<option value="1" <%If Trim(tipo_busqueda) ="2" Then Response.write "SELECTED"%>>INUBICABLES</option>
						<option value="3" <%If Trim(tipo_busqueda) ="3" Then Response.write "SELECTED"%>>GESTIONADOS</option>
						<option value="4" <%If Trim(tipo_busqueda) ="4" Then Response.write "SELECTED"%>>PENDIENTES</option>
						<option value="5" <%If Trim(tipo_busqueda) ="5" Then Response.write "SELECTED"%>>GESTIÓN POSITIVA</option>
						<option value="6" <%If Trim(tipo_busqueda) ="6" Then Response.write "SELECTED"%>>GESTIÓN NEGATIVA</option>
						<option value="7" <%If Trim(tipo_busqueda) ="7" Then Response.write "SELECTED"%>>TITULAR</option>
						<option value="8" <%If Trim(tipo_busqueda) ="8" Then Response.write "SELECTED"%>>TERCERO</option>
	
				</select>
			</td>

			</tr>
	</table>

	<table width="90%" align="CENTER" border="0" class="estilo_columnas">
		<thead>
		<tr >
			<td width="">NOMBRE O RAZON SOCIAL</td>
			<td width="">RUT</td>
			<td width="">MONTO DESDE</td>
			<td width="">MONTO HASTA</td>
			<td>&nbsp;</td>
		</tr>
		</thead>
		<tr >
			<td><input name="TX_NOMBRES" type="text" value="" size="50" maxlength="570"></td>
			<td><input name="TX_RUT" type="text" value="" size="12" maxlength="12"></td>
			<td><input name="TX_DESDE" type="text" value="" size="12" maxlength="12"></td>
			<td><input name="TX_HASTA" type="text" value="" size="12" maxlength="12"></td>

			<td align="right">
				<input name="Limpiar" type="button" class="fondo_boton_100" value="Limpiar"  onClick="limpiar();">
				<input name="Buscar" type="button" class="fondo_boton_100" value="Buscar"  onClick="buscar();">
			</td>
		</tr>
	</table>

	<br>
	<table width="90%" align="CENTER">
		<tr class="Estilo13">
			<td width="60%" align="center"><%=strMensaje%></td>
		</tr>
	</table>


				<%
					AbrirSCG()

					strSql = " SELECT * "
					strSql = strSql & " FROM "
					strSql = strSql & " 		( "
					strSql = strSql & " 		SELECT PP.RUT_DEUDOR,MIN(PP.PRIORIDAD_CUOTA) AS PRIORIDAD_CUOTA,  "
					strSql = strSql & " 		(	SELECT SUM(SALDO)  "
					strSql = strSql & " 			FROM CUOTA C "
					strSql = strSql & " 			INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO "
					strSql = strSql & " 			WHERE C.RUT_DEUDOR=PP. RUT_DEUDOR AND C.COD_CLIENTE=PP.COD_CLIENTE AND ED.ACTIVO=1 "
					strSql = strSql & " 		) AS SALDO_RUT, "
					strSql = strSql & " 		( "
					strSql = strSql & " 			SELECT COUNT(ID_CUOTA) "
					strSql = strSql & " 			FROM CUOTA C "
					strSql = strSql & " 			INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO "
					strSql = strSql & " 			WHERE C.RUT_DEUDOR=PP. RUT_DEUDOR AND C.COD_CLIENTE=PP.COD_CLIENTE AND ED.ACTIVO=1 "
					strSql = strSql & " 		) AS TOTAL_DOC,  "
					strSql = strSql & " 		PP.NOMBRE_DEUDOR, PP.RUBRO, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_TEL_VA >0 THEN 1 ELSE 0 END) AS TEL_VA, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_TEL_SA >0 THEN 1 ELSE 0 END) AS TEL_SA, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_TEL_NV >0 THEN 1 ELSE 0 END) AS TEL_NV, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_EMAIL_VA >0 THEN 1 ELSE 0 END) AS EMAIL_VA, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_EMAIL_SA >0 THEN 1 ELSE 0 END) AS EMAIL_SA, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_EMAIL_NV >0 THEN 1 ELSE 0 END) AS EMAIL_NV, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_DIR_VA >0 THEN 1 ELSE 0 END) AS DIR_VA, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_DIR_SA >0 THEN 1 ELSE 0 END) AS DIR_SA, "
					strSql = strSql & " 		(CASE WHEN PP.TOTAL_DIR_NV >0 THEN 1 ELSE 0 END) AS DIR_NV, "
					strSql = strSql & " 		SUM(PP.GEST_GENERAL) AS TT_GEST_GENERAL, "
					strSql = strSql & " 		SUM(PP.GEST_TEL) AS TT_GEST_TEL, "
					strSql = strSql & " 		SUM(PP.GEST_MAIL) AS TT_GEST_MAIL, "
					strSql = strSql & " 		SUM(PP.GEST_DIR) AS TT_GDIR, "
					strSql = strSql & " 		SUM(PP.GEST_EFE) AS TT_GEFE, "
					strSql = strSql & "  		SUM(PP.GEST_TIT) AS TT_GTIT "
					strSql = strSql & " 		FROM "
					strSql = strSql & " 		( "
					strSql = strSql & " 			SELECT D.RUT_DEUDOR,D.COD_CLIENTE,D.NOMBRE_DEUDOR, ISNULL(D.ADIC_2,'OTRO') RUBRO,MIN(ISNULL(C.PRIORIDAD_CUOTA,99)) AS PRIORIDAD_CUOTA, "
					strSql = strSql & " 			(CASE WHEN G.ID_GESTION IS NOT NULL THEN 1 ELSE 0 END) AS GEST_GENERAL, ISNULL((GTG.PRIORIDAD_GTEL),0) AS GEST_TEL, "
					strSql = strSql & " 			ISNULL((GTG.PRIORIDAD_GMAIL),0) AS GEST_MAIL, ISNULL((GTG.PRIORIDAD_GDIR),0) AS GEST_DIR, "
					strSql = strSql & " 			ISNULL((GTG.PRIORIDAD_GEFE),0) AS GEST_EFE, ISNULL((GTG.PRIORIDAD_GTIT),0) AS GEST_TIT, "
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_TEL_VA, "
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_TEL_SA, "
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_TEL_NV,  "
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_EMAIL_VA, "
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_EMAIL_SA, " 
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_EMAIL WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_EMAIL_NV, " 
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 1) AS TOTAL_DIR_VA, " 
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 0) AS TOTAL_DIR_SA, " 
					strSql = strSql & " 			(SELECT COUNT(RUT_DEUDOR) FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR = D.RUT_DEUDOR AND ESTADO = 2) AS TOTAL_DIR_NV "

					strSql = strSql & " 			FROM CUOTA C "
					strSql = strSql & " 			INNER JOIN DEUDOR D ON C.RUT_DEUDOR = D.RUT_DEUDOR AND C.COD_CLIENTE = D.COD_CLIENTE "
					strSql = strSql & " 			INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA = ED.CODIGO "
					strSql = strSql & " 			LEFT JOIN GESTIONES_CUOTA GC ON C.ID_CUOTA = GC.ID_CUOTA "
					strSql = strSql & " 			LEFT JOIN GESTIONES G ON GC.ID_GESTION = G.ID_GESTION "
					strSql = strSql & " 			LEFT JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
					strSql = strSql & " 			AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA AND G.COD_GESTION = GTG.COD_GESTION AND G.COD_CLIENTE = GTG.COD_CLIENTE " 
					strSql = strSql & " 			WHERE ED.ACTIVO=1 "
					strSql = strSql & " 			AND D.COD_CLIENTE = " & TRIM(strCodCliente)

					if trim(strCobranza)="INTERNA" THEN
						strSql = strSql & " AND D.CUSTODIO IS not NULL "

					elseif trim(strCobranza)="EXTERNA" THEN
						strSql = strSql & " AND D.CUSTODIO IS NULL "
					end if
	
					IF TRIM(intEstadoCob)<>"" THEN
						strSql = strSql & " AND D.ETAPA_COBRANZA = '"&TRIM(intEstadoCob)&"'"

					END IF
	

					if trim(strRubro)<>"" then
						strSql = strSql & " AND ISNULL(D.ADIC_2,'OTRO') = '"&trim(strRubro)&"' " 
					end if

					if trim(intTipoDoc)<>"" then
						strSql = strSql & " AND C.TIPO_DOCUMENTO ='"&trim(intTipoDoc)&"' " 
					end if


					IF TRIM(intCodCampana)<>"" THEN
						strSql = strSql & " AND d.ID_CAMPANA = '"&TRIM(intCodCampana)&"'"

					END IF


					IF TRIM(strEjeAsig)<>""  THEN
						strSql = strSql & " AND ISNULL(C.USUARIO_ASIG,0) = '"&TRIM(strEjeAsig)&"'"
					END IF


					strSql = strSql & " 			GROUP BY D.RUT_DEUDOR,D.NOMBRE_DEUDOR, D.ADIC_2,D.COD_CLIENTE,G.ID_GESTION,GTG.PRIORIDAD_GTEL,GTG.PRIORIDAD_GMAIL, "
					strSql = strSql & " 			GTG.PRIORIDAD_GDIR,GTG.PRIORIDAD_GEFE,GTG.PRIORIDAD_GTIT"
					strSql = strSql & " 		) AS PP "

					strSql = strSql & " 		GROUP BY PP.COD_CLIENTE,PP.RUT_DEUDOR,PP.NOMBRE_DEUDOR, PP.RUBRO,PP.TOTAL_TEL_VA,TOTAL_TEL_SA,TOTAL_TEL_NV,TOTAL_EMAIL_VA, "
					strSql = strSql & " TOTAL_EMAIL_SA,TOTAL_EMAIL_NV,TOTAL_DIR_VA,TOTAL_DIR_SA,TOTAL_DIR_NV "
					strSql = strSql & " ) AS PP2 "
					
					If Trim(tipo_busqueda) <> "0" Then 'FILTRADO
					
					strSql = strSql & " WHERE "
					
					END IF

					If Trim(tipo_busqueda) ="1" Then 'GESTIONABLES

						strSql = strSql & " (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR PP2.DIR_SA=1) "
					END IF

					If Trim(tipo_busqueda) ="2" Then 'INUBICABLES

						strSql = strSql & "  (PP2.TEL_VA=0 AND PP2.TEL_SA=0 AND PP2.EMAIL_VA=0 AND PP2.EMAIL_SA=0 AND PP2.DIR_VA=0 AND PP2.DIR_SA=0) "
					END IF


					If Trim(tipo_busqueda) ="3" Then 'GESTIONADOS

						strSql = strSql & " (PP2.TT_GEST_GENERAL>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) "
					END IF

					If Trim(tipo_busqueda) ="4" Then 'PENDIENTES

						strSql = strSql & " (PP2.TT_GEST_GENERAL=0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR PP2.DIR_SA=1))"
					END IF


					If Trim(tipo_busqueda) ="5" Then 'GESTIÓN POSITIVA

						strSql = strSql & " (PP2.TT_GEFE>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) "
					END IF

					If Trim(tipo_busqueda) ="6" Then 'GESTIÓN NEGATIVA

						strSql = strSql & " (PP2.TT_GEST_GENERAL>0 AND PP2.TT_GEFE=0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR PP2.DIR_SA=1))"
					END IF


					If Trim(tipo_busqueda) ="7" Then 'TITULAR

						strSql = strSql & " (PP2.TT_GTIT>0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) "
					END IF

					If Trim(tipo_busqueda) ="8" Then 'TERCERO

						strSql = strSql & " (PP2.TT_GEFE>0 AND PP2.TT_GTIT=0 AND (PP2.TEL_VA=1 OR PP2.TEL_SA=1 OR PP2.EMAIL_VA=1 OR PP2.EMAIL_SA=1 OR PP2.DIR_VA=1 OR PP2.DIR_SA=1)) "
					END IF

					'Response.write "strSql=" & strSql
					
					set rsCuota=Server.CreateObject("ADODB.Recordset")
					rsCuota.Open strSql, Conn, 1, 2
					
					intTotalSaldo = 0
					intTotalDoc= 0
					intTotalRut = 0

					rsCuota.PageSize=TamPagina
					rsCuota.CacheSize=TamPagina
					PaginasTotales=rsCuota.PageCount
					''Response.write "PaginaActual=" & PaginasTotales

					'Compruebo que la pagina actual está en el rango
					if PaginaActual < 1 then
						PaginaActual = 1
					end if
					if PaginaActual > PaginasTotales then
						PaginaActual = PaginasTotales
					end if

					'Por si la consulta no devuelve registros!
					if PaginasTotales=0 then
						strMensaje = "No se encontraron resultados"
					else
						rsCuota.AbsolutePage=PaginaActual
					End If


					sintPagina = PaginaActual
					sintTotalPaginas = PaginasTotales
					%>


					  <table style="width:90%" align="CENTER" class="intercalado">
					  	<thead>
						<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
							<td width = "80">RUT DEUDOR</td>
							<td>NOMBRE O RAZON SOCIAL</td>
							<td>RUBRO</td>
							<td>SALDO</td>
							<td width="40">DOC.</td>
							<td>PRIOR.</td>

							<% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
							<td width="70">&nbsp;</td>
							<% End If%>

							<td>&nbsp</td>
						</tr>

						<TR>
							<TD COLSPAN=8>
								<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="100%">
									<TR BGCOLOR="#F3F3F3">
										<TD WIDTH="20%" ALIGN=left>
											<%if PaginaActual > 1 then %>
											<INPUT TYPE=BUTTON NAME="Retroceder" VALUE="  &lt;  " onClick="IrPagina( 'Retroceder')">
											<% end if %>
										</TD>
										<TD WIDTH="60%" ALIGN=center>
											<FONT FACE="verdana, Sans-Serif" Size=1 COLOR="#FF0000"><b>Página <%= sintPagina %> de <%= sintTotalPaginas %></b></FONT>
										</TD>
										<TD WIDTH="20%" ALIGN=right>
											<%if PaginaActual < PaginasTotales then%>
											<INPUT TYPE=BUTTON NAME="Avanzar" VALUE="  &gt;  " onClick="IrPagina( 'Avanzar')">
											<% end if %>
										</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>

						<TR>
							<TD COLSPAN=8 ALIGN="CENTER">
								<%=strMensaje%>
							</TD>
						</TR>
						</thead>
						<tbody>
					<%

						If Not rsCuota.eof Then
							totalventa=0
							Do while not rsCuota.eof and CuantosRegistros < TamPagina

								'strSqlG = "SELECT * FROM GESTIONES WHERE COD_CLIENTE = " & strCodCliente
								'set rsGestion=Conn.execute(strSqlG)
								'If rsGestion.eof Then
								'	strbgcolor="#F6F6F6"
								'Else
								'	strbgcolor="#F6F6CF"
								'End If

								'Response.write "<br>DIFERENCIA=" & dtmDif
								'Response.write "<br>dtmFecUG=" & dtmFecUG
								'Response.write "<br>FECHA_AGENDAMIENTO=" & rsFecha("FECHA_AGENDAMIENTO")
								'Response.write "<br>FECHA_INGRESO=" & rsFecha("FECHA_INGRESO")

								''rESPONSE.WRITE "valor_moneda=" & session("valor_moneda")

								intValorSaldo = Round(session("valor_moneda") * ValNulo(rsCuota("SALDO_RUT"),"N"),0)

								intDoc = rsCuota("TOTAL_DOC")

								intTotalDoc = intTotalDoc + intDoc
								intTotalSaldo = intTotalSaldo + intValorSaldo
								intTotalRut = intTotalRut + 1


								'dtmFecAgend= rsCuota("FEC_VENC_IA")

								%>
									<tr bgcolor="<%=strbgcolor%>" class="Estilo8">

										<td ALIGN="left"><%=rsCuota("RUT_DEUDOR")%></td>
										<td><%=rsCuota("NOMBRE_DEUDOR")%></td>
										<td><%=rsCuota("RUBRO")%></td>
										<td ALIGN="right"><%=FN(intValorSaldo,0)%></td>

										<td ALIGN="center"><%=intDoc%></td>

										<td ALIGN="center"><b><%=rsCuota("PRIORIDAD_CUOTA")%></b></td>

										<% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
										<td>
											<A HREF="asigna_masiva.asp?TA_RUT=<%=rsCuota("RUT_DEUDOR")%>&CB_COBRANZA=<%=strCobranza%>">
												<acronym title="Asigna Deudor">Asignar</acronym>
											</A>
										</td>
										<% End If%>

										<td>
											<A HREF="principal.asp?TX_RUT=<%=rsCuota("RUT_DEUDOR")%>">
												<acronym title="Llevar a pantalla de selección">Seleccionar</acronym>
											</A>
										</td>
									</tr>
								<%
								rESPONSE.flush()
								CuantosRegistros=CuantosRegistros+1
								rsCuota.movenext
							Loop
						End If
					rsCuota.close
					set rsCuota=NOTHING
					%>
					<tr>
						<td COLSPAN="8">

							<table BORDER="0" CELLSPACING="0" CELLPADDING=0 WIDTH="100%">
								<tr BGCOLOR="#F3F3F3">
									<td WIDTH="20%" ALIGN=left>
										<%if PaginaActual > 1 then %>
										<INPUT TYPE=BUTTON NAME="Retroceder" VALUE="  &lt;  " onClick="IrPagina( 'Retroceder')">
										<% end if %>
									</td>
									<td WIDTH="60%" ALIGN=center>
										<FONT FACE="verdana, Sans-Serif" Size=1 COLOR="#FF0000"><b>Página <%= sintPagina %> de <%= sintTotalPaginas %></b></FONT>
									</td>
									<td WIDTH="20%" ALIGN=right>
										<%if PaginaActual < PaginasTotales then%>
										<INPUT TYPE=BUTTON NAME="Avanzar" VALUE="  &gt;  " onClick="IrPagina( 'Avanzar')">
										<% end if %>
									</td>
								</tr>
							</table>

						</TD>
					</tr>
					<tr class="totales">
						<td ><b>Totales</b></td>
						<td align="right"  colspan=3>$ <%=FN(intTotalSaldo,0)%></td>
						<td align="right"  colspan=1><%=FN(intTotalDoc,0)%></td>
						<td align="center" colspan=5>Total Rut : <%=intTotalRut%> </td>
					</tr>
			</tbody>		
			</table>

</form>
</body>
<script language="JavaScript1.2">

function buscar(){
	datos.Buscar.disabled = true;
	datos.action='cartera_asignada.asp?strBuscar=S';
	datos.submit();

}

function limpiar(){
	datos.action='cartera_asignada.asp?strLimpiar=S';
	datos.submit();

}

function IrPagina( sintAccion ) {

	datos.Buscar.disabled = true;
	if (sintAccion == 'Retroceder') {
    	self.location.href = 'cartera_asignada.asp?pagina=<%=PaginaActual - 1%>&TX_NOMBRES=<%=strNombres%>&CB_REMESA=<%=intCOD_REMESA%>&CB_RUBRO=<%=strRubro%>&CB_EJECUTIVO=<%=strEjeAsig%>&CB_CAMPANA=<%=intCodCampana%>&tipo_busqueda=<%=strTipoInf%>'
    }
    if (sintAccion == 'Avanzar') {
	    self.location.href = 'cartera_asignada.asp?pagina=<%=PaginaActual + 1%>&TX_NOMBRES=<%=strNombres%>&CB_REMESA=<%=intCOD_REMESA%>&CB_RUBRO=<%=strRubro%>&CB_EJECUTIVO=<%=strEjeAsig%>&CB_CAMPANA=<%=intCodCampana%>&tipo_busqueda=<%=strTipoInf%>'
    }

}

function CargaUsuarios(subCat)
{
	//alert(subCat);

	var comboBox = document.getElementById('CB_EJECUTIVO');
	comboBox.options.length = 0;

		if (subCat=='INTERNA') {
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('SIN ASIGNACION', '0');
			comboBox.options[comboBox.options.length] = newOption;
			<%

			AbrirSCG2()

			strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = '" & strCodCliente & "'"

			strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
			strSql = strSql & " AND U.PERFIL_EMP=1"

			'Response.write "<br>strSql=" & strSql

			set rsUsuario=Conn2.execute(strSql)
			If Not rsUsuario.Eof Then
				Do While Not rsUsuario.Eof
					%>
						var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					rsUsuario.movenext
				Loop
			Else
			%>
				var newOption = new Option('SIN USUARIO', '');
				comboBox.options[comboBox.options.length] = newOption;
			<%
			End if
			CerrarSCG2()
			%>
		}

		else if ((subCat=='EXTERNA') && (<%=intVerEjecutivos%>=='1')) {
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('SIN ASIGNACION', '0');
			comboBox.options[comboBox.options.length] = newOption;
			<%

			AbrirSCG2()

			strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = '" & strCodCliente & "'"

			strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
			strSql = strSql & " AND U.PERFIL_EMP=0"

			'Response.write "<br>strSql=" & strSql

			set rsUsuario=Conn2.execute(strSql)
			If Not rsUsuario.Eof Then
				Do While Not rsUsuario.Eof
					%>
						var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					rsUsuario.movenext
				Loop
			Else
			%>
				var newOption = new Option('SIN USUARIO', '');
				comboBox.options[comboBox.options.length] = newOption;
			<%
			End if
			CerrarSCG2()
			%>
		}
		else if ((subCat=='EXTERNA') && (<%=intVerEjecutivos%>=='0')) {
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;

		}
		else {
			var newOption = new Option('TODOS', '');
			comboBox.options[comboBox.options.length] = newOption;
			var newOption = new Option('SIN ASIGNACION', '0');
			comboBox.options[comboBox.options.length] = newOption;
			<%

			AbrirSCG2()

			strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
			strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE = '" & strCodCliente & "'"

			strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"

			If intVerEjecutivos = "0" then
			strSql = strSql & " AND U.PERFIL_EMP=1"
			end If

			set rsUsuario=Conn2.execute(strSql)
			If Not rsUsuario.Eof Then
				Do While Not rsUsuario.Eof
					%>
						var newOption = new Option('<%=rsUsuario("LOGIN")%>', '<%=rsUsuario("ID_USUARIO")%>');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					rsUsuario.movenext
				Loop
			Else
			%>
				var newOption = new Option('SIN USUARIO', '');
				comboBox.options[comboBox.options.length] = newOption;
			<%
			End if
			CerrarSCG2()
			%>
		}

}

function InicializaInforme()
{
		var comboBox = document.getElementById('CB_EJECUTIVO');
		comboBox.options.length = 0;
		var newOption = new Option('TODOS','');
		comboBox.options[comboBox.options.length] = newOption;
}

<%If sinCbUsario = "0" then%>
CargaUsuarios('<%=strCobranza%>');
<%End If%>

<%If strEjeAsig <> "" then%>
datos.CB_EJECUTIVO.value='<%=strEjeAsig%>';
<%End If%>
</script>