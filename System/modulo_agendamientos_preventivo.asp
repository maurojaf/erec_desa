<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib2.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/Minimo.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->

<script language="JavaScript">
	function ventanaSecundaria (URL){
		window.open(URL,"DETALLE","width=1000, height=300, scrollbars=YES, menubar=no, location=no, resizable=yes")
	}
</script>
<%

Dim PaginaActual ' en qué pagina estamos
Dim PaginasTotales ' cuántas páginas tenemos
Dim TamPagina ' cuantos registros por pagina
Dim CuantosRegistros ' para imprimir solo el nº de registro por pagina que

strNombres= Request("TX_NOMBRES")
strRut = Request("TX_RUT")
intCodCampana = Request("CB_CAMPANA")
strCOD_CLIENTE=session("ses_codcli")
''strCOD_CLIENTE=Request("CB_CLIENTE")
strEjeAsig = Request("CB_EJECUTIVO")
strTipoInf = Request("CB_TIPOCARTERA")
intEstadoCob = Request("CB_TIPOCOB")


intGestionPrinc = Request("CB_TIPOGESTION_PRINC")
intGestion = Request("CB_TIPOGESTION")
strProridad = Request("CB_PRIORIDAD")
dtmInicio = Request("TX_INICIO")
dtmTermino = Request("TX_TERMINO")
intTipoDoc = Request("CB_TIPODOC")


If trim(strProridad) = "" Then strProridad = "TODAS"
''Response.write "strProridad=" & strProridad

If Trim(intEstadoCob) = "" Then
	If Trim(session("tipo_cliente")) = "JUDICIAL" Then
		intEstadoCob = 4
	Else
		intEstadoCob = 2
	End If
End If

If Trim(strTipoInf) = "" Then strTipoInf = "GESTIONABLES"
''If Trim(strCOD_CLIENTE) = "" Then strCOD_CLIENTE = "1000"


''Response.write "strBuscar=" & Request("strBuscar")
If Trim(Request("strBuscar")) = "S" Then
	session("Ftro_Ejecutivo") = strEjeAsig
	session("Ftro_Campana") = intCodCampana
	session("Ftro_DtmInicio") = dtmInicio
	session("Ftro_DtmTermino") = dtmTermino
	session("Ftro_TipoCartera") = strTipoInf
	session("Ftro_TipoGPpal") = intGestionPrinc
	session("Ftro_TipoGTel") = intGestion
	session("Ftro_Cliente") = strCOD_CLIENTE
	session("Ftro_TipoDoc") = intTipoDoc
	session("Ftro_NivelAtrasp") = strProridad
End If

''Response.write "strLimpiar=" & Trim(Request("strLimpiar"))

If Trim(Request("strBuscar")) = "N" or Trim(Request("strLimpiar")) = "S" Then
	session("Ftro_Ejecutivo") = ""
	session("Ftro_Campana") = ""
	session("Ftro_DtmInicio") = ""
	session("Ftro_DtmTermino") = ""
	session("Ftro_TipoCartera") = ""
	session("Ftro_TipoGPpal") = ""
	session("Ftro_TipoGTel") = ""
	session("Ftro_Cliente") = ""
	session("Ftro_TipoDoc") = ""
	session("Ftro_NivelAtrasp") = "TODAS"
End If

If Trim(Request("strLimpiar")) = "S" Then
	strEjeAsig = "0"
	intCodCampana = ""
	dtmInicio = ""
	dtmTermino = ""
	intGestionPrinc = ""
	intGestion = ""
	strCOD_CLIENTE = ""
	intTipoDoc = ""
	strProridad=""

End If

If strEjeAsig <> "0" Then strEjeAsig = session("Ftro_Ejecutivo")
If intCodCampana = "" Then intCodCampana = session("Ftro_Campana")
If dtmInicio = "" Then dtmInicio = session("Ftro_DtmInicio")
If dtmTermino = "" Then dtmTermino = session("Ftro_DtmTermino")
If intGestionPrinc = "" Then intGestionPrinc = session("Ftro_TipoGPpal")
If intGestion = "" Then intGestion = session("Ftro_TipoGTel")
If strCOD_CLIENTE = "" Then strCOD_CLIENTE = session("Ftro_Cliente")
If intTipoDoc = "" Then intTipoDoc = session("Ftro_TipoDoc")
If strProridad = "" Then strProridad = session("Ftro_NivelAtrasp")

'Response.write "strEjeAsig=" & strEjeAsig

'MODIFICAR AQUI PARA CAMBIAR EL Nº DE REGISTRO POR PAGINA
TamPagina=100

'Leemos qué página mostrar. La primera vez será la inicial
if Request.Querystring("pagina")="" then
	PaginaActual=1
else
	PaginaActual=CInt(Request.Querystring("pagina"))
end if


%>
<title>MODULO AGENDAMIENTO PREVENTIVO</title>

<%strTitulo="MI CARTERA"%>

<script language='javascript' src="../javascripts/popcalendar.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">

<table width="100%" border="1" bordercolor="#FFFFFF">
	<tr>
		<TD height="20" ALIGN=LEFT class="pasos2_i">
			<B>MODULO DE AGENDAMIENTOS PREVENTIVO</B>
		</TD>
	</tr>
</table>


	<table width="1000" align="CENTER" border="0" bordercolor="#FFFFFF">
			<tr bordercolor="#999999"  bgcolor="#FFFFFF" class="Estilo13">
				<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				  <td width="200">MANDANTE</td>
				  <td width="220">TIPO COBRANZA</td>
				  <td width="150">TIPO DOC</td>
				  <td width="150">CAMPAÑA</td>
				  <td width="150">PRIORIDAD</td>


				  <% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
				  	<td align="center" width="107">EJECUTIVO</td>
				  <% Else %><td width="115">&nbsp;</td>
				  <% End If %>
				</tr>

				<td>
					<select name="CB_CLIENTE">
						<option value="">TODOS</option>
						<%
						abrirscg()
						strSql = "SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE "
						strSql = strSql & " WHERE 1=1 "
						If Trim(strCOD_CLIENTE) <> "" Then
							strSql = strSql & " AND COD_CLIENTE = '" & strCOD_CLIENTE & "'"
						End If
						strSql = strSql & " AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") ORDER BY RAZON_SOCIAL"
						set rsTemp= Conn.execute(strSql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCOD_CLIENTE)=Trim(rsTemp("COD_CLIENTE")) then response.Write("Selected") End If%>><%=rsTemp("RAZON_SOCIAL")%></option>
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
					<select name="CB_TIPOCOB">
						<option value="">TODOS</option-->
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
					<select name="CB_TIPODOC">
						<option value="">TODOS</option-->
						<%
						abrirscg()

						strSql="SELECT DISTINCT COD_TIPO_DOCUMENTO, NOM_TIPO_DOCUMENTO"
						strSql=strSql & " FROM CUOTA LEFT JOIN TIPO_DOCUMENTO ON TIPO_DOCUMENTO = COD_TIPO_DOCUMENTO"
						strSql=strSql & " WHERE CUOTA.COD_CLIENTE = '" & strCOD_CLIENTE & "'"
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
					<select name="CB_CAMPANA">
						<option value="">TODAS</option>
						<%
						AbrirSCG()
							strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
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

				<td>
				<select name="CB_PRIORIDAD">
					<option value="TODAS" <%if Trim(strProridad)="TODAS" then response.Write("Selected") end if%>>TODAS</option>
					<option value="1" <%if Trim(strProridad)="1" then response.Write("Selected") end if%>>1</option>
					<option value="2" <%if Trim(strProridad)="2" then response.Write("Selected") end if%>>2</option>
					<option value="2.1" <%if Trim(strProridad)="2.1" then response.Write("Selected") end if%>>2,1</option>
					<option value="2.2" <%if Trim(strProridad)="2.2" then response.Write("Selected") end if%>>2,2</option>
					<option value="3" <%if Trim(strProridad)="3" then response.Write("Selected") end if%>>3</option>
					<option value="4" <%if Trim(strProridad)="4" then response.Write("Selected") end if%>>4</option>
					<option value="5" <%if Trim(strProridad)="5" then response.Write("Selected") end if%>>5</option>
					<option value="6" <%if Trim(strProridad)="6" then response.Write("Selected") end if%>>6</option>
					<option value="7" <%if Trim(strProridad)="7" then response.Write("Selected") end if%>>7</option>
					<option value="8" <%if Trim(strProridad)="8" then response.Write("Selected") end if%>>8</option>
					<option value="9" <%if Trim(strProridad)="9" then response.Write("Selected") end if%>>9</option>
					<option value="10" <%if Trim(strProridad)="10" then response.Write("Selected") end if%>>10</option>
					<option value="99" <%if Trim(strProridad)="99" then response.Write("Selected") end if%>>99</option>
					<option value="100" <%if Trim(strProridad)="100" then response.Write("Selected") end if%>>100</option>
				</select>
				</td>

				<% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
				<td>
					<select name="CB_EJECUTIVO" id="CB_EJECUTIVO">
						<option value="0" <%if Trim(strEjeAsig)="0" then response.Write("Selected") end if%>>SELECCIONE</option>
						<%
						AbrirScg()

						strSql="SELECT USUARIO.ID_USUARIO,USUARIO.LOGIN "
						strSql=strSql & " FROM USUARIO INNER JOIN USUARIO_CLIENTE ON USUARIO.ID_USUARIO = USUARIO_CLIENTE.ID_USUARIO "
						strSql=strSql & " AND USUARIO_CLIENTE.COD_CLIENTE = '" & strCOD_CLIENTE & "' WHERE ACTIVO = 1 AND (PERFIL_COB = 1)"

						strSql=strSql & " ORDER BY PERFIL_COB, PERFIL_SUP, PERFIL_ADM, LOGIN"

						''Response.write "strSql=" & strSql
						If trim(intGrupo) <> "" and trim(intGrupo) <> "0" Then
							strSql = strSql & " and grupo = '" & intGrupo & "'"
						End if
						set rsEjecutivo=Conn.execute(strSql)
						if not rsEjecutivo.eof then
							do until rsEjecutivo.eof
							%>
							<option value="<%=rsEjecutivo("ID_USUARIO")%>" <%if Trim(strEjeAsig)=Trim(rsEjecutivo("ID_USUARIO")) then response.Write("selected") end if%>><%=ucase(rsEjecutivo("LOGIN"))%></option>
							<%rsEjecutivo.movenext
							loop
						end if
						rsEjecutivo.close
						set rsEjecutivo=nothing
						CerrarScg()
						%>
					</select>
				</td>
				<% End If%>

			</tr>
	</table>

	<table width="1000" align="CENTER" border="0" bordercolor="#FFFFFF">
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="423">NOMBRE O RAZON SOCIAL</td>
			<td width="150">RUT</td>
			<td width="150">FEC.INICIO</td>
			<td width="150">FEC.TERMINO</td>
			<td>&nbsp;</td>
		</tr>
		<tr bgcolor="#f6f6f6" class="Estilo8">
			<td><input name="TX_NOMBRES" type="text" value="" size="57" maxlength="77"></td>
			<td><input name="TX_RUT" type="text" value="" size="12" maxlength="12"></td>
			<td>
				<input name="TX_INICIO" type="text" id="TX_INICIO" value="<%=dtmInicio%>" size="10">
				<a href="#" onClick="popUpCalendar(this, datos.TX_INICIO, 'dd/mm/yyyy');"><img src="../Imagenes/calendario.gif" border="0">
			</td>
			<td>
				<input name="TX_TERMINO" type="text" id="TX_TERMINO" value="<%=dtmTermino%>" size="10">
				<a href="#" onClick="popUpCalendar(this, datos.TX_TERMINO, 'dd/mm/yyyy');"><img src="../Imagenes/calendario.gif" border="0">
			</td>

			<td align="right" width="107"><input name="Buscar" type="button" value="Buscar"  onClick="buscar();"></td>

		</tr>
	</table>

	<table width="1000" align="CENTER" border="0" bordercolor="#FFFFFF">
			<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td Colspan =4>FILTRO GESTION</td>
			</tr>
			<tr bgcolor="#f6f6f6" class="Estilo8">
				<td>
					<select name="CB_TIPOGESTION">
						<option value="" <%if Trim(intGestion) = ""  Then Response.write "SELECTED" %>>TODAS</option>
						<option value="SIN GESTION" <%if Trim(intGestion) = "SIN GESTION" Then Response.write "SELECTED" %>>SIN GESTION DOC</option>
						<option value="SIN GESTION EFECTIVA" <%if Trim(intGestion) = "SIN GESTION EFECTIVA" Then Response.write "SELECTED" %>>SIN GESTION EFECTIVA DOC</option>
						<option value="SIN GESTION CASO" <%if Trim(intGestion) = "SIN GESTION CASO" Then Response.write "SELECTED" %>>SIN GESTION CASO</option>
						<option value="SIN GESTION TEL CASO" <%if Trim(intGestion) = "SIN GESTION TEL CASO" Then Response.write "SELECTED" %>>SIN GESTION TELEFONICA CASO</option>
						<option value="SIN GESTION MAIL CASO" <%if Trim(intGestion) = "SIN GESTION MAIL CASO" Then Response.write "SELECTED" %>>SIN GESTION MAIL CASO</option>
						<%
						abrirscg()
							strSql = "SELECT * FROM GESTIONES_TIPO_GESTION WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
							strSql = strSql & " AND ISNULL(GESTIONES_TIPO_GESTION.VER_AGEND,1) = 1"

							set rsGest = Conn.execute(strSql)
							''strCodComPago = ""

							Do While not rsGest.eof

								strSql = "SELECT DESCRIPCION FROM GESTIONES_TIPO_CATEGORIA WHERE COD_CATEGORIA = " & rsGest("COD_CATEGORIA")
								set rsTemp = Conn.execute(strSql)
								If Not rsTemp.Eof Then
									strNomCategoria = rsTemp("DESCRIPCION")
								End if

								strSql = "SELECT DESCRIPCION FROM GESTIONES_TIPO_SUBCATEGORIA WHERE COD_CATEGORIA = " & rsGest("COD_CATEGORIA") & " AND COD_SUB_CATEGORIA = " & rsGest("COD_SUB_CATEGORIA")
								set rsTemp = Conn.execute(strSql)
								If Not rsTemp.Eof Then
									strNomSubCategoria = rsTemp("DESCRIPCION")
								End if

								strNombreGestion = rsGest("DESCRIPCION")
								strGestionTotal = strNomCategoria & "-" & strNomSubCategoria & "-" & strNombreGestion
								'strGestionTotal = strNomSubCategoria & "-" & strNombreGestion
								strCodigo = rsGest("COD_CATEGORIA") & "*" & rsGest("COD_SUB_CATEGORIA") & "*" & rsGest("COD_GESTION")

								if strCodigo = Trim(intGestion) Then strGestSel="SELECTED" Else strGestSel=""
							%>
								<option value="<%=Trim(strCodigo)%>" <%=strGestSel%>><%=strGestionTotal%></option>

							<%
								rsGest.movenext
							Loop

						cerrarscg()
						%>
					</select>
				</td>

				<td align="right" Width = "105"><input name="Limpiar" type="button" value="Limpiar"  onClick="limpiar();"></td>
			</tr>
	</table>

	<table width="1000" align="CENTER">
		<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="60%" align="center"><%=strMensaje%></td>
		</tr>
	</table>

				<%
					AbrirSCG()

					strSql = "SELECT IsNull(datediff(minute,DEUDOR.FECHA_CONF,IsNull(DEUDOR.FECHA_UG_TITULAR,'01/01/1900')),0) as DIFMINUTOS,"
					strSql = strSql & " MAX(DATEDIFF(DAY,FECHA_VENC,GETDATE())) AS DIAVENC,"
					strSql = strSql & " DEUDOR.OBSERVACIONES_CONF,"
					strSql = strSql & " DEUDOR.FECHA_CONF,"
					strSql = strSql & " DEUDOR.USUARIO_CONF,"
					strSql = strSql & " ISNULL(DEUDOR.RESP_EMAIL,0) AS RESP_EMAIL,"
					strSql = strSql & " MAX(GESTIONES_TIPO_CATEGORIA.DESCRIPCION+'-'+GESTIONES_TIPO_SUBCATEGORIA.DESCRIPCION+'-'+GESTIONES_TIPO_GESTION.DESCRIPCION) AS NOM_GEST,"

					strSql = strSql & " CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),103) AS FEC_AGEND,"
					strSql = strSql & " MIN((ISNULL(CUOTA.FECHA_AGEND_ULT_GES,GETDATE()+300) + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))) AS FEC_AGEND2,"
					strSql = strSql & " (CASE WHEN CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),108) = '00:00:00'"
					strSql = strSql & "		  THEN ''"
					strSql = strSql & " 	  WHEN SUBSTRING(CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),108),5,1)= ':' "
					strSql = strSql & " 	  THEN SUBSTRING(CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),108),1,4)"
					strSql = strSql & " 	  ELSE SUBSTRING(CONVERT(VARCHAR(10),MIN((FECHA_AGEND_ULT_GES + convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108))),108),1,5)"
					strSql = strSql & " END) AS HORA_AGEND,"


					strSql = strSql & " MAX(CUOTA.COD_ULT_GEST) AS UGT,"
					strSql = strSql & " [dbo].[fun_ubicabilidad_telefono_email] (DEUDOR.RUT_DEUDOR) AS ESTATUS_TEL,"
					strSql = strSql & " DEUDOR.RUT_DEUDOR,"
					strSql = strSql & " NOMBRE_DEUDOR,"
					strSql = strSql & " SUM(SALDO) as SALDO,"
					strSql = strSql & " COUNT(CUOTA.ID_CUOTA) as DOC,"
					strSql = strSql & " CLIENTE.COD_CLIENTE as COD_CLIENTE,"
					strSql = strSql & " MIN(ISNULL(CUOTA.PRIORIDAD_CUOTA,11)) AS PRIORIDAD_CUOTA"


					strSql = strSql & " FROM DEUDOR INNER JOIN CUOTA 					  ON DEUDOR.RUT_DEUDOR = CUOTA.RUT_DEUDOR AND DEUDOR.COD_CLIENTE = CUOTA.COD_CLIENTE"
					strSql = strSql & " 			INNER JOIN CLIENTE					  ON DEUDOR.COD_CLIENTE = CLIENTE.COD_CLIENTE"
					strSql = strSql & " 			LEFT JOIN GESTIONES_TIPO_CATEGORIA 	  ON SUBSTRING(CUOTA.COD_ULT_GEST,1,1) = GESTIONES_TIPO_CATEGORIA.COD_CATEGORIA"
					strSql = strSql & " 			LEFT JOIN GESTIONES_TIPO_SUBCATEGORIA ON SUBSTRING(CUOTA.COD_ULT_GEST,1,1) = GESTIONES_TIPO_SUBCATEGORIA.COD_CATEGORIA"
					strSql = strSql & " 													 AND SUBSTRING(CUOTA.COD_ULT_GEST,3,1) = GESTIONES_TIPO_SUBCATEGORIA.COD_SUB_CATEGORIA"
					strSql = strSql & " 			LEFT JOIN GESTIONES_TIPO_GESTION 	  ON SUBSTRING(CUOTA.COD_ULT_GEST,1,1) = GESTIONES_TIPO_GESTION.COD_CATEGORIA"
					strSql = strSql & " 													 AND SUBSTRING(CUOTA.COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
					strSql = strSql & " 											   		 AND SUBSTRING(CUOTA.COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
					strSql = strSql & " 					 								 AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"

					strSql = strSql & " WHERE CUOTA.COD_CLIENTE = '" & strCOD_CLIENTE & "'"

					strSql = strSql & " AND (DEUDOR.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_TELEFONO WHERE ESTADO IN (0,1))"
					strSql = strSql & " OR DEUDOR.RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM DEUDOR_EMAIL WHERE ESTADO IN (0,1)))"

					strSql = strSql & " AND (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0"
					strSql = strSql & " AND ISNULL(GESTIONES_TIPO_GESTION.VER_AGEND,1) = 1)"

					strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (SELECT ESTADO_DEUDA.CODIGO FROM ESTADO_DEUDA WHERE ESTADO_DEUDA.ACTIVO = 1)"

					strSql = strSql & " AND CUOTA.RUT_DEUDOR IN (SELECT RUT_DEUDOR"
					strSql = strSql & " FROM CUOTA LEFT JOIN GESTIONES_TIPO_GESTION 	  ON SUBSTRING(CUOTA.COD_ULT_GEST,1,1) = GESTIONES_TIPO_GESTION.COD_CATEGORIA"
					strSql = strSql & " 													 AND SUBSTRING(CUOTA.COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
					strSql = strSql & " 													 AND SUBSTRING(CUOTA.COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
					strSql = strSql & " 													 AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
					strSql = strSql & " WHERE CUOTA.COD_CLIENTE = 1100 AND ISNULL(GESTIONES_TIPO_GESTION.VER_AGEND,1) = 1"
					strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (SELECT ESTADO_DEUDA.CODIGO FROM ESTADO_DEUDA WHERE ESTADO_DEUDA.ACTIVO = 1)"
					strSql = strSql & " GROUP BY RUT_DEUDOR"
					strSql = strSql & " HAVING MAX((CAST((CAST(convert(varchar(10), getdate(),103) AS DATETIME)-FECHA_VENC) AS INT)))<-5)"

					If trim(strEjeAsig) <> "0" AND trim(strEjeAsig) <> "" Then
						strSql = strSql & " AND CUOTA.USUARIO_ASIG = " & strEjeAsig
					End if


					strParametro = "0"

					If Trim(strNombres) <> "" Then
						strSql = strSql & " AND NOMBRE_DEUDOR  LIKE '%" & strNombres & "%'"
						strParametro = "1"
					End if

					If Trim(strRut) <> "" Then
						strSql = strSql & " AND CUOTA.RUT_DEUDOR  LIKE '" & strRut & "%'"
						strParametro = "1"
					End if

					If Trim(intEstadoCob) <> "0" and Trim(intEstadoCob) <> "" Then
						strSql = strSql & " AND DEUDOR.ETAPA_COBRANZA = " & intEstadoCob
						strParametro = "1"
					End if

					If Trim(intTipoDoc) <> "0" and Trim(intTipoDoc) <> "" Then
						strSql = strSql & " AND CUOTA.TIPO_DOCUMENTO = '" & intTipoDoc & "'"
						strParametro = "1"
					End if

					If Trim(intCodCampana) <> "0" and Trim(intCodCampana) <> "" Then
						strSql = strSql & " AND CUOTA.RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
						strParametro = "1"
					End if

					If Trim(intGestion) = "SIN GESTION TEL CASO" Then
						strSql = strSql & " AND [dbo].[fun_trae_estatus_gestion] (DEUDOR.COD_CLIENTE,DEUDOR.RUT_DEUDOR,'TELEFONICA') = 0"

					ElseIf Trim(intGestion) = "SIN GESTION CASO" Then
						strSql = strSql & " AND [dbo].[fun_trae_estatus_gestion] (DEUDOR.COD_CLIENTE,DEUDOR.RUT_DEUDOR,'GENERAL') = 0"

					ElseIf Trim(intGestion) = "SIN GESTION MAIL CASO" Then
						strSql = strSql & " AND [dbo].[fun_trae_estatus_gestion] (DEUDOR.COD_CLIENTE,DEUDOR.RUT_DEUDOR,'MAIL') = 0"

					ElseIf Trim(intGestion) = "SIN GESTION" Then
						strSql = strSql & " AND (ID_ULT_GEST_GENERAL IS NULL OR ID_ULT_GEST_GENERAL=0)"

					ElseIf Trim(intGestion) = "SIN GESTION EFECTIVA" Then
						strSql = strSql & " AND (ID_ULT_GEST_EFE IS NULL OR ID_ULT_GEST_EFE=0)"

					ElseIf Trim(intGestion) <> "" Then
						strSql = strSql & " AND CUOTA.COD_ULT_GEST= '" & intGestion & "'"
					End If

					If Trim(strProridad) <> "TODAS" Then
						strSql = strSql & " AND CUOTA.PRIORIDAD_CUOTA  = '" & strProridad & "'"
					End If

					If Trim(dtmInicio) <> "" Then
						strSql = strSql & " AND FECHA_AGEND_ULT_GES >= '" & dtmInicio & " 00:00:00'"
					End If

					If Trim(dtmTermino) <> "" Then
						strSql = strSql & " AND FECHA_AGEND_ULT_GES <= '" & dtmTermino & " 23:58:59'"
					End If

					strSql = strSql & " GROUP BY DEUDOR.FECHA_UG_TITULAR,DEUDOR.OBSERVACIONES_CONF, DEUDOR.FECHA_CONF, DEUDOR.USUARIO_CONF,RESP_EMAIL, DEUDOR.RUT_DEUDOR,"
					strSql = strSql & " 		 NOMBRE_DEUDOR, CLIENTE.COD_CLIENTE "

					strSql = strSql & " ORDER BY CLIENTE.COD_CLIENTE, PRIORIDAD_CUOTA ASC, DIAVENC DESC, FEC_AGEND2 ASC, SUM(SALDO) DESC"



					''RESPONSE.WRITE "strSql=" & strSql
					'RESPONSE.eND

					set rsCuota=Server.CreateObject("ADODB.Recordset")
					rsCuota.Open strSql, Conn, 1, 2
					intTotalSaldo = 0
					intTotalRut = 0

					' Defino el tamaño de las páginas
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
						strMensaje = "No hay Casos Agendados Para Gestionar"
					else
						rsCuota.AbsolutePage=PaginaActual
					End If

					sintPagina = PaginaActual
					sintTotalPaginas = PaginasTotales

					%>


					  <table width="1000" align="CENTER">
						<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
							<td width="10">CONT.</td>
							<td align="center">RUT</td>
							<td width="350">NOMBRE O RAZON SOCIAL</td>
							<td align="center">DOC.</td>
							<td align="center">SALDO</td>
							<td width="80" align="center">ULT.GESTION</td>
							<td align="center">F.AGEND.</td>
							<td align="center">H.AGEND.</td>
							<!--td>&nbsp;</td-->
							<td align="center">PRIOR.</td>
							<td>&nbsp;</td>
							<td>&nbsp</td>
						</tr>
						<TR>
							<TD COLSPAN=12>
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
					<%
						'Response.write "valor_moneda=" & session("valor_moneda")
						'Response.write "SALDO=" & rsCuota("SALDO")

						'session("valor_moneda") = 22000
						'session("valor_moneda") = 1

						If Not rsCuota.eof Then
							totalventa=0

							Do while not rsCuota.eof and CuantosRegistros < TamPagina


								strNomGestion = rsCuota("NOM_GEST")
								strCodGestion = rsCuota("UGT")
								dtmFecAgend = rsCuota("FEC_AGEND")
								dtmHoraAgend = rsCuota("HORA_AGEND")


								intMinDif = rsCuota("DIFMINUTOS")
								strObsConf = rsCuota("OBSERVACIONES_CONF")
								strFechaConf = rsCuota("FECHA_CONF")
								strUsuarioConf = rsCuota("USUARIO_CONF")
								strRespEmail = rsCuota("RESP_EMAIL")
								strTextoConf=""
								If Trim(strFechaConf) <> "" and Trim(strUsuarioConf) <> "" then
									strTextoConf = "Fecha : " & strFechaConf & " , Usuario : " & strUsuarioConf & ", Obs : "
								End If

								intValorSaldo = Round(session("valor_moneda") * ValNulo(rsCuota("SALDO"),"N"),0)
								intTotalSaldo = intTotalSaldo + intValorSaldo
								intValorDoc = rsCuota("DOC")
								intTotalDoc = intTotalDoc + intValorDoc
								intTotalRut = intTotalRut + 1


								if Trim(rsCuota("ESTATUS_TEL")) = "CONTACTADO" Then
									strContactado = "tel_contactado.jpg"
								Else
									strContactado = "tel_nocontactado.jpg"
								End If


								%>
									<tr bgcolor="<%=strbgcolor%>" class="Estilo8">

										<td ALIGN="center"><img src="../imagenes/<%=strContactado%>" border="0"></td>
										<td ALIGN="center"><%=rsCuota("RUT_DEUDOR")%></td>
										<td><%=Mid(rsCuota("NOMBRE_DEUDOR"),1,30)%></td>
										<td ALIGN="right"><%=FN(intValorDoc,0)%></td>
										<td ALIGN="right"><%=FN(intValorSaldo,0)%></td>

										<td ALIGN="center">
										<acronym title="<%=strNomGestion%>">
										<%=strCodGestion%>
										</acronym>
										</td>
										<td ALIGN="center"><%=dtmFecAgend%></td>
										<td ALIGN="center"><%=dtmHoraAgend%></td>
										<td ALIGN="center"><b><%=rsCuota("PRIORIDAD_CUOTA")%></b></td>
										<td ALIGN="center">
										<% If Trim(strObsConf) <> "" and strRespEmail = 0 Then%>
											<acronym title="<%=strTextoConf & " " & strObsConf%>">
											<% If (Cdbl(intMinDif) >= 0 and strRespEmail = 0) then %>
												<img src="../imagenes/priorizar_normal.png" border="0" height=15 onMouseover="ddrivetip('<%=strTextoConf & " " & strObsConf%>', '#EFEFEF',300)"; onMouseout="hideddrivetip()">
											<% Else %>
												<img src="../imagenes/priorizar_urgente.png" border="0" height=15 onMouseover="ddrivetip('<%=strTextoConf & " " & strObsConf%>', '#EFEFEF',300)"; onMouseout="hideddrivetip()">
											<% End If %>
											</acronym>
										<% End If %>

										</td>
										<!--td ALIGN="right"><img src="../imagenes/<%=strColorG%>" border="0"></td-->
										<td>
											<A HREF="principal.asp?TX_RUT=<%=rsCuota("RUT_DEUDOR")%>">
												<acronym title="Llevar a pantalla de selección">Seleccionar</acronym>
											</A>
										</td>
									</tr>
								<%
								CuantosRegistros=CuantosRegistros+1
								rsCuota.movenext
							Loop
						End If
					rsCuota.close
					set rsCuota=NOTHING
					%>
					<TR>
						<TD COLSPAN=12>
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
					<tr>
						<td Colspan = "2" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">Totales</td>
						<td Colspan = "2" bgcolor="#<%=session("COLTABBG2")%>" span class="Estilo28" align="right"  colspan=2>Documentos Agendados :<%=FN(intTotalDoc,0)%> </td>
						<td Colspan = "3" bgcolor="#<%=session("COLTABBG2")%>" span class="Estilo28" align="right"  colspan=2>Saldo Agendados : $<%=FN(intTotalSaldo,0)%></td>
						<td bgcolor="#<%=session("COLTABBG2")%>" span class="Estilo28" align="center" colspan=7>Total Rut : <%=intTotalRut%> </td>
					</tr>
			</table>


</form>
</body>
</html>
<script language="JavaScript1.2">

function buscar(){
	datos.action='modulo_agendamientos_preventivo.asp?strBuscar=S';
	datos.submit();

}

function limpiar(){
	datos.action='modulo_agendamientos_preventivo.asp?strLimpiar=S';
	datos.submit();

}

function IrPagina( sintAccion ) {
	if (sintAccion == 'Retroceder') {
    	self.location.href = 'modulo_agendamientos_preventivo.asp?pagina=<%=PaginaActual - 1%>&TX_NOMBRES=<%=strNombres%>&CB_REMESA=<%=intCodRemesa%>&CB_CLIENTE=<%=strCOD_CLIENTE%>&CB_EJECUTIVO=<%=strEjeAsig%>&CB_CAMPANA=<%=intCodCampana%>&CB_TIPOCARTERA=<%=strTipoInf%>&TX_INICIO=<%=dtmInicio%>&TX_TERMINO=<%=dtmTermino%>&CB_TIPOGESTION=<%=intGestion%>&CB_PRIORIDAD=<%=strProridad%>&CB_TIPOGESTION_PRINC=<%=intGestionPrinc%>'
    }
    if (sintAccion == 'Avanzar') {
	    self.location.href = 'modulo_agendamientos_preventivo.asp?pagina=<%=PaginaActual + 1%>&TX_NOMBRES=<%=strNombres%>&CB_REMESA=<%=intCodRemesa%>&CB_CLIENTE=<%=strCOD_CLIENTE%>&CB_EJECUTIVO=<%=strEjeAsig%>&CB_CAMPANA=<%=intCodCampana%>&CB_TIPOCARTERA=<%=strTipoInf%>&TX_INICIO=<%=dtmInicio%>&TX_TERMINO=<%=dtmTermino%>&CB_TIPOGESTION=<%=intGestion%>&CB_PRIORIDAD=<%=strProridad%>&CB_TIPOGESTION_PRINC=<%=intGestionPrinc%>'
    }

}

</script>