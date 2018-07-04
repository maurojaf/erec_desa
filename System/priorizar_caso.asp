<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta charset="utf-8">
<link href="../css/normalize.css" rel="stylesheet">

<%
	strOrigen = request("strOrigen")
	
	If strOrigen = "1" Then 
	
	strOrigen = "2" %>
		<!--#include file="sesion_inicio.asp"-->
	<% Else %>
		<!--#include file="sesion.asp"-->
	<% End If %>

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">


<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	strRutDeudor = request("rut")
	strCodCliente = request("strCodCliente")
	id_gestion = request("id_gestion")
	intCodGestConcat = request("intCodGestConcat")
	dtmFecCompGest = request("dtmFecCompGest")
	intSolResp = request("cmb_SolResp")
	

	If intSolResp = "" Then intSolResp = 0 End If

	strRutDeudor = request("strRut")
	intIdTipoSol = request("cmb_TipoSol")
	intIdTipoReclamo = request("cmb_TipoReclamo")
	
	If intIdTipoReclamo = "" Then intIdTipoReclamo = 1 End If
	
	strObservaciones=Mid(Replace(request("OBSERVACIONES"),";"," "),1,599)
	strChTodos = Request("CH_TODOS")

	If UCASE(Request("CH_TODOS")) = "ON" Then
		strChTodos="CHECKED"
	End if

	strGraba = request("strGraba")

	intUsuario=session("session_idusuario")

	AbrirSCG1()
	strSql="SELECT FORMULA_HONORARIOS, FORMULA_INTERESES, RAZON_SOCIAL, USA_SUBCLIENTE, USA_INTERESES, USA_HONORARIOS, USA_PROTESTOS, ISNULL(RETIRO_SABADO,0) AS RETIRO_SABADO, [dbo].[fun_ubicabilidad_telefono] ('" & strRutDeudor & "') as UBIC_FONO, [dbo].[fun_ubicabilidad_email] ('" & strRutDeudor & "') as UBIC_EMAIL, [dbo].[fun_ubicabilidad_direccion] ('" & strRutDeudor & "') as UBIC_DIRECCION  FROM CLIENTE WHERE COD_CLIENTE='" & strCodCliente & "'"

	'Response.write "strSql=" & strSql
	set rsCLI=Conn1.execute(strSql)
	if not rsCLI.eof then
		strNomFormHon = ValNulo(rsCLI("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsCLI("FORMULA_INTERESES"),"C")

		strUsaSubCliente = rsCLI("USA_SUBCLIENTE")
		strUsaInteres = rsCLI("USA_INTERESES")
		strUsaHonorarios = rsCLI("USA_HONORARIOS")
		strUsaProtestos = rsCLI("USA_PROTESTOS")


		nombre_cliente=rsCLI("RAZON_SOCIAL")
		intRetiroSabado=Cint(rsCLI("RETIRO_SABADO"))
		strMsjRetiroSabado = ""
		If Trim(intRetiroSabado) = "1" Then
			strMsjRetiroSabado = "sabados,"
		End if

		strUbicFono =rsCLI("UBIC_FONO")
		strUbicEmail =rsCLI("UBIC_EMAIL")
		strUbicDireccion =rsCLI("UBIC_DIRECCION")
	end if
	rsCLI.close
	set rsCLI=nothing
	CerrarSCG1()

	AbrirSCG1()
		strSql="SELECT IsNull(COUNT(*),0) AS CANT FROM CUOTA WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE = '" & strCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
		set RsDeudor=Conn1.execute(strSql)
		If not RsDeudor.eof then
			intCantidadDocAct = RsDeudor("CANT")
		Else
			intCantidadDocAct = 0
		End if
		RsDeudor.close
		set RsDeudor=nothing


		If intCantidadDocAct = 0 Then
		%>
		<SCRIPT>
			alert('El deudor que intenta priorizar no posee actualmente deuda activa, si necesita dar PRIORIDAD para su cobranza favor pongase en contacto con supervisor o Backoffice llacruz.');
			location.href="principal.asp?strRut=<%=strRutDeudor%>";		
		</SCRIPT>

		<%

		End If

	CerrarSCG1()

	AbrirSCG()

	If Trim(request("strGraba")) = "SI" Then

		AbrirSCG()

		strSql = "SELECT USUARIO_ASIG FROM DEUDOR"
		strSql = strSql & " WHERE COD_CLIENTE =  '" & strCodCliente & "' AND RUT_DEUDOR =  '" & strRutDeudor & "'"

		'Response.write "strSql=" & strSql
		set rsInf=Conn.execute(strSql)

		intUsuarioAsig = rsInf("USUARIO_ASIG")

		CerrarSCG()

		AbrirSCG()
		strSql = "INSERT INTO PRIORIZACION (COD_CLIENTE,RUT_DEUDOR, FECHA_PRIORIZACION, ID_TIPO_SOLICITUD, ID_TIPO_RECLAMO, SOLICITA_RESPUESTA, ID_USUARIO_PRIORIZACION, OBSERVACION_PRIORIZACION,USUARIO_ASIG)"
		strSql = strSql & " VALUES ( '" & strCodCliente & "', '" & strRutDeudor & "', GETDATE(), " & intIdTipoSol & ", " & intIdTipoReclamo & ", " & intSolResp & ", '" & intUsuario & "', '" & UCASE(strObservaciones) & "','" & intUsuarioAsig & "')"
		'Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)
		CerrarSCG()

		AbrirSCG()
		strSql = "SELECT MAX(ID_PRIORIZACION) AS ID_PRIORIZACION FROM PRIORIZACION WHERE RUT_DEUDOR  = '" & strRutDeudor & "' AND COD_CLIENTE = '" & strCodCliente & "'"
		set rsPrioridad = Conn.execute(strSql )
		If Not rsPrioridad.eof Then
			intIdPriorizacion= rsPrioridad("ID_PRIORIZACION")
		End If
		CerrarSCG()

		AbrirSCG()
		strSql = "SELECT ID_CUOTA FROM CUOTA WHERE RUT_DEUDOR='" & strRutDeudor & "' AND COD_CLIENTE='" & strCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) "
		set rsTemp= Conn.execute(strSql)

		Do until rsTemp.eof

			strObjeto = "CH_" & rsTemp("ID_CUOTA")
			If UCASE(Request(strObjeto)) = "ON" Then

			AbrirSCG1()
				strSql = "INSERT INTO PRIORIZACIONES_CUOTA (ID_PRIORIZACION, ID_CUOTA, ESTADO_PRIORIZACION, FECHA_ESTADO,ID_USUARIO_ESTADO)"
				strSql = strSql&"VALUES (" & intIdPriorizacion & "," & rsTemp("ID_CUOTA") & ",0,GETDATE(),'" & intUsuario & "')"
				Conn1.execute(strSql)
			CerrarSCG1()

			AbrirSCG1()
				strSql = "UPDATE CUOTA"
				strSql = strSql & " SET PRIORIDAD_CUOTA = 2.1"
				strSql = strSql & " WHERE ID_CUOTA = " & rsTemp("ID_CUOTA")
				Conn1.execute(strSql)
			CerrarSCG1()
			
			End If

			rsTemp.movenext
		Loop
		rsTemp.close
		set rsTemp=nothing
		CerrarSCG()
		
		If strOrigen="" Then

		Response.redirect "principal.asp?strRut=" & strRutDeudor
		
		End If

	End If


%>
	<title>Empresa</title>
	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo27 {color: #FFFFFF}
	.Estilo1 {
		color: #FF0000;
		font-weight: bold;
		font-family: Arial, Helvetica, sans-serif;
	-->

	}
	</style>
</head>	
<body>
<form name="datos" method="post">
<INPUT TYPE="hidden" NAME="intOrigen" value="<%=intOrigen%>">
<%

	AbrirSCG1()
	ssql="SELECT NOMBRE_DEUDOR,RUT_DEUDOR, FECHA_PRORROGA FROM DEUDOR WHERE RUT_DEUDOR='"& strRutDeudor &"' AND COD_CLIENTE = '" & strCodCliente & "'"
	set rsDEU=Conn1.execute(ssql)
	if not rsDEU.eof then
		strNombre_deudor = rsDEU("NOMBRE_DEUDOR")
		strRutDeudor = rsDEU("RUT_DEUDOR")
	else
		strNombre_deudor = "SIN NOMBRE"
	end if
	rsDEU.close
	set rsDEU=nothing
	CerrarSCG1()

%>

<% If Trim(strRutDeudor) <> "" then %>
<div class="titulo_informe">Módulo Ingreso de Priorizaciones <%=nombre_cliente%></div>
<br>
	<table width="90%" border="0" ALIGN="CENTER">
		<tr>
			<td valign="top">


			<table width="100%" border="0" bordercolor="#FFFFFF">
		      <tr>
		        <td height="16%" width="120" class="estilo_columna_individual">&nbsp;&nbsp;RUT DEUDOR</td>
				<td bgcolor="#<%=session("COLTABBG2")%>">
					<A HREF="principal.asp?TX_RUT=<%=strRutDeudor%>">
						<acronym title="Llevar a pantalla principal"><%=strRutDeudor%></acronym>
					</A>
				</td>

		        <td width="17%" class="estilo_columna_individual">&nbsp;&nbsp;NOMBRE DEUDOR</td>
		        <td width="50%" class="Estilo10" bgcolor="#<%=session("COLTABBG2")%>">&nbsp <%=strNombre_deudor%></td>

		      </tr>
		    </table>

			<table class="intercalado" style="width:100%;">
				<thead>
				<tr class="Estilo34">
					<td colspan=15 align="LEFT">
					<a href="#" onClick= "marcar_boxes(true);">Marcar todos</a>&nbsp;&nbsp;&nbsp;
					<a href="#" onClick="desmarcar_boxes(true);">Desmarcar todos</a>
					</td>
				</tr>

				<tr class="Estilo34">
					<td>&nbsp;</td>

					<%If Trim(strUsaSubCliente)="1" Then%>
					<td>RUT CLIENTE</td>
					<td>NOMBRE CLIENTE</td>
					<%End If%>

					<td>NºDOC</td>
					<td>CUOTA</td>
					<td>FEC.VENC.</td>
					<td>ANT.</td>
					<td>TIPO DOC.</td>
					<td align="center" width="70">CAPITAL</td>
					<%If Trim(strUsaInteres)="1" Then%>
					<td align="center" width="70">INTERES</td>
					<%End If%>
					<%If Trim(strUsaProtestos)="1" Then%>
					<td align="center" width="70">PROTESTOS</td>
					<%End If%>
					<%If Trim(strUsaHonorarios)="1" Then%>
					<td align="center" width="70">HONORARIOS</td>
					<%End If%>
					<td align="center" width="70">ABONO</td>
					<td align="center" width="70">SALDO</td>
					<td>FECHA AGEND.</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				</thead>
				<tbody>

				<%
				AbrirSCG()

				strSql = "SELECT dbo." & strNomFormInt & "(CUOTA.ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(CUOTA.ID_CUOTA) as HONORARIOS,"
				strSql = strSql & "	ISNULL(FACTURA_RECEPCIONADA,2) AS FACTURA_RECEPCIONADA, COD_ULT_GEST, NRO_CUOTA,"
				strSql = strSql & "	ISNULL(NOTIFICACION_RECEPCIONADA,2) AS NOTIFICACION_RECEPCIONADA,"
				strSql = strSql & "	VALOR_CUOTA, CUOTA.ID_CUOTA,RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, NRO_DOC, SALDO, NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO,"
				strSql = strSql & "	GASTOS_PROTESTOS, CUENTA, FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,"
				strSql = strSql & "	CUSTODIO, CUOTA.FECHA_ESTADO, FECHA_CREACION,"
				strSql = strSql & "	ID_ULT_GEST,ISNULL((SUBSTRING(CONVERT(VARCHAR(11),CAST(CONVERT(VARCHAR(10),CUOTA.FECHA_AGEND_ULT_GES,103) AS DATETIME),6),1,7) + '/ ' + (CASE WHEN CUOTA.HORA_AGEND_ULT_GES = '' THEN '00:00' ELSE CUOTA.HORA_AGEND_ULT_GES END)),'SIN AGEND') AS FECHA_AGEND_ULT_GES,"
				strSql = strSql & "	ISNULL(PRC.ESTADO_PRIORIZACION,1) AS ESTADO_PRIORIZACION"

				strSql = strSql & " FROM CUOTA LEFT JOIN GESTIONES_CUOTA GC ON CUOTA.ID_CUOTA = GC.ID_CUOTA AND GC.ID_GESTION = CUOTA.ID_ULT_GEST_GENERAL"
				strSql = strSql & "			   LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
				strSql = strSql & "			   LEFT JOIN GESTIONES_TIPO_GESTION ON     SUBSTRING(CUOTA.COD_ULT_GEST,1,1) = GESTIONES_TIPO_GESTION.COD_CATEGORIA"
				strSql = strSql & "													   AND SUBSTRING(CUOTA.COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
				strSql = strSql & "											  		   AND SUBSTRING(CUOTA.COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
				strSql = strSql & "				    				  		        	AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
				strSql = strSql & " 		   LEFT JOIN PRIORIZACIONES_CUOTA PRC ON CUOTA.ID_CUOTA = PRC.ID_CUOTA AND ESTADO_PRIORIZACION = '0'"

				strSql = strSql & " WHERE RUT_DEUDOR='" & strRutDeudor & "' AND CUOTA.COD_CLIENTE='" & strCodCliente & "' AND  ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) "

				strSql = strSql & " ORDER BY RUT_SUBCLIENTE, FECHA_VENC DESC"

				'Response.write "strSql=" & strSql
				'Response.End
				set rsTemp= Conn.execute(strSql)

				intTasaMensual = 2/100
				intTasaDiaria = intTasaMensual/30
				intCorrelativo = 1
				strArrID_CUOTA=""
				intTotSelSaldo= 0
				intTotSelIntereses= 0
				intTotSelProtestos= 0
				intTotSelHonorarios= 0
				strDetCuota="mas_datos_adicionales.asp"

				strArrConcepto = ""
				strArrID_CUOTA = ""

				Do until rsTemp.eof

						intSaldo =  rsTemp("SALDO")
						intValorCuota =  rsTemp("VALOR_CUOTA")
						intAbono = intValorCuota - intSaldo
						strNroDoc = rsTemp("NRO_DOC")
						strNroCuota = rsTemp("NRO_CUOTA")
						strFechaVenc = rsTemp("FECHA_VENC")
						strTipoDoc = rsTemp("TIPO_DOCUMENTO")
						intAntiguedad = ValNulo(rsTemp("ANTIGUEDAD"),"N")
						intEstadoPrio = rsTemp("ESTADO_PRIORIZACION")

						If intEstadoPrio ="0" then
						strDisabled = "disabled"
						Else
						strDisabled = ""
						End If

						intIntereses = rsTemp("INTERESES")
						intHonorarios = rsTemp("HONORARIOS")

						'Response.write "intHonorarios=" & intHonorarios

						intProtestos = ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")
						intTotDoc= intSaldo+intIntereses+intProtestos+intHonorarios

						intTotSelSaldo = intTotSelSaldo+intSaldo
						intTotSelAbono = intTotSelAbono+intAbono
						intTotSelValorCuota = intTotSelValorCuota+intValorCuota

						intTotSelIntereses= intTotSelIntereses+intIntereses
						intTotSelProtestos= intTotSelProtestos+intProtestos
						intTotSelHonorarios= intTotSelHonorarios+intHonorarios
						intTotSelDoc = intTotSelDoc+intTotDoc

						strArrConcepto = strArrConcepto & ";" & "CH_" & rsTemp("ID_CUOTA")
						strArrID_CUOTA = strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")

						%>
						<tr class="">
						<input name="TX_SALDO_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intTotDoc%>">
						<input name="TX_CAPITAL_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intValorCuota%>">
						<input name="TX_HONORARIOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intHonorarios%>">
						<input name="TX_INTERESES_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intIntereses%>">
						<input name="TX_PROTESTOS_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intProtestos%>">

						<TD><INPUT TYPE="checkbox" NAME="CH_<%=rsTemp("ID_CUOTA")%>" <%=strDisabled%> <%if trim(intEstadoPrio)>0 then response.write " checked " end if%> onClick="suma_capital(this,TX_CAPITAL_<%=rsTemp("ID_CUOTA")%>.value,TX_INTERESES_<%=rsTemp("ID_CUOTA")%>.value,TX_HONORARIOS_<%=rsTemp("ID_CUOTA")%>.value,TX_PROTESTOS_<%=rsTemp("ID_CUOTA")%>.value,TX_SALDO_<%=rsTemp("ID_CUOTA")%>.value);"></TD>

						<%If Trim(strUsaSubCliente)="1" Then%>

						<td><%=rsTemp("RUT_SUBCLIENTE")%></td>

						<td class="Estilo4" title="<%=rsTemp("NOMBRE_SUBCLIENTE")%>">
						<%=Mid(rsTemp("NOMBRE_SUBCLIENTE"),1,30)%></td>

						<%End If%>

						<td><%=strNroDoc%></td>
						<td><%=strNroCuota%></td>
						<td><%=strFechaVenc%></td>
						<td><%=intAntiguedad%></td>
						<td><%=strTipoDoc%></td>
						<td ALIGN="RIGHT"><%=FN(intValorCuota,0)%></td>
						<%If Trim(strUsaInteres)="1" Then%>
						<td ALIGN="RIGHT"><%=FN(intIntereses,0)%></td>
						<%End If%>
						<%If Trim(strUsaProtestos)="1" Then%>
						<td ALIGN="RIGHT"><%=FN(intProtestos,0)%></td>
						<%End If%>
						<%If Trim(strUsaHonorarios)="1" Then%>
						<td ALIGN="RIGHT"><%=FN(intHonorarios,0)%></td>
						<%End If%>

						<td ALIGN="RIGHT"><%=FN(intAbono,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intTotDoc,0)%></td>
						<td ALIGN="RIGHT"><%=rsTemp("FECHA_AGEND_ULT_GES")%></td>
						<td ALIGN="CENTER">
						<a href="javascript:ventanaGestionesPorDoc('gestiones_por_documento.asp?intID_CUOTA=<%=trim(rsTemp("ID_CUOTA"))%>&cliente=<%=strCodCliente%>&strNroDoc=<%=trim(rsTemp("NRO_DOC"))%>')">
						<img src="../imagenes/icon_gestiones.jpg" border="0">
						</a>
						</td>
						<td align="center">

							<%
							strImagenGest1=""

							If intEstadoPrio ="0" Then
								strImagenGest1 = "GestionarRoj.png"
							Else
								strImagenGest1 = ""
							End If
							%>

							<% If strImagenGest1 <> "" Then %>

							<img src="../imagenes/<%=strImagenGest1%>" border="0">

							<% Else %>

							&nbsp;

							<% End If %>

						</td>
						</tr>
						<%

					rsTemp.movenext
				intCorrelativo = intCorrelativo + 1
				loop

				vArrConcepto = split(strArrConcepto,";")
				vArrID_CUOTA = split(strArrID_CUOTA,";")

				intTamvConcepto = ubound(vArrConcepto)

				rsTemp.close
				set rsTemp=nothing
				CerrarSCG()

				strArrID_CUOTA = Mid(strArrID_CUOTA,2,len(strArrID_CUOTA))
		%>
			</tbody>
			<thead class="totales">
			<tr class="">
				<%If Trim(strUsaSubCliente)="1" Then%>
				<td colspan = "2">&nbsp;</td>
				<%End If%>
				<td colspan = "6">Totales Seleccionados:</td>
				<td ALIGN="RIGHT">
					<INPUT TYPE="TEXT" NAME="TX_CAPITAL" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)">
				</td>


				<% If Trim(strUsaInteres)="1" Then%>
					<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_INTERESES" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
				<% Else%>
					<INPUT TYPE="hidden" NAME="TX_INTERESES">
				<% End If%>

				<% If Trim(strUsaProtestos)="1" Then%>
					<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_PROTESTOS" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
				<% Else%>
					<INPUT TYPE="hidden" NAME="TX_PROTESTOS">
				<% End If%>

				<% If Trim(strUsaHonorarios)="1" Then%>
					<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_HONORARIOS" DISABLED style="text-align:right;width:90" size="10" onkeyup="format(this)" onchange="format(this)"></td>
				<% Else%>
					<INPUT TYPE="hidden" NAME="TX_HONORARIOS">
				<% End If%>

				<td>&nbsp;</td>
				<td ALIGN="RIGHT"><INPUT TYPE="TEXT" NAME="TX_SALDO" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			</thead>
			<INPUT TYPE="hidden" NAME="strArrID_CUOTA" VALUE="<%=strArrID_CUOTA%>">

			</table>
			<% Response.Flush %>

			<table width="100%" border="0" class="estilo_columnas">
				<thead>
				<tr >
					<td width="33%">TIPO SOLICITUD</td>
					<td width="34%">TIPO RECLAMO</td>
					<td width="33%">URGENTE</td>
				</tr>
				</thead>
				<tr >

					<td><select name="cmb_TipoSol" id="cmb_TipoSol"onchange="this.style.width=300">
				  		<option value="0">SELECCIONE</option>
			<%
					  AbrirSCG1()

						strSql = "SELECT ID_TIPO_SOLICITUD, (CODIGO + '-' + NOM_TIPO_SOLICITUD) AS NOM_TIPO_SOLICITUD FROM TIPO_SOLICITUD_PRIORIZACION AS TSP"
						strSql = strSql & " WHERE TSP.COD_CLIENTE = '" & strCodCliente & "'"
						strSql = strSql & "ORDER BY ORDEN"

						'Response.write "strSql=" & strSql

						set rsTipoPrio=Conn1.execute(strSql)

						Do While not rsTipoPrio.eof
				%>
						<option value="<%=rsTipoPrio("ID_TIPO_SOLICITUD")%>"><%=rsTipoPrio("NOM_TIPO_SOLICITUD")%></option>
				<%
						rsTipoPrio.movenext
						Loop
						rsTipoPrio.close
						set rsTipoPrio=nothing

					  CerrarSCG1()
			%>

					</select>
					</td>

					<td><select name="cmb_TipoReclamo" id="cmb_TipoReclamo" onchange="this.style.width=300">
			<%
					  AbrirSCG1()

						strSql = "SELECT ID_TIPO_RECLAMO, (CODIGO + '-' + NOM_TIPO_RECLAMO) AS NOM_TIPO_RECLAMO FROM TIPO_RECLAMO_PRIORIZACION AS TRP"
						strSql = strSql & " WHERE TRP.COD_CLIENTE = '" & strCodCliente & "'"
						strSql = strSql & "ORDER BY ORDEN"

						'Response.write "strSql=" & strSql

						set rsTipoReclamo=Conn1.execute(strSql)

						Do While not rsTipoReclamo.eof
				%>
						<option value="<%=rsTipoReclamo("ID_TIPO_RECLAMO")%>"><%=rsTipoReclamo("NOM_TIPO_RECLAMO")%></option>
				<%

						rsTipoReclamo.movenext
						Loop
						rsTipoReclamo.close
						set rsTipoReclamo=nothing

					  CerrarSCG1()
			%>
					</select>
					</td>

					<td>
					<select name="cmb_SolResp">
						<option value="1" <%if Trim(intSolResp)="1" then response.Write("Selected") end if%>>SI</option>
						<option value="0" <%if Trim(intSolResp)="0" then response.Write("Selected") end if%>>NO</option>
					</select>
					</td>

				</tr>

			</table>

			<table width="100%" border="0">
				<tr >
					<td colspan="5" class="estilo_columna_individual">OBSERVACIONES (Max. 600 Caract.)</td>
				</tr>
				<tr>
				   <td colspan="4" align="LEFT">
				  <TEXTAREA NAME="OBSERVACIONES" ROWS="4" COLS="87"></TEXTAREA>
				  </td>
				   <td align="CENTER">
				  <input name="ingresar" class="fondo_boton_100" type="button" onClick="envia();" value="Priorizar documentos">
					</td>
				</tr>

			</table>

			<div class="titulo_informe">Historial de Priorizaciones</div>

			<table border="0" class="intercalado" style="width:100%;">
			<thead>
			<%
					AbrirSCG()

						strSql = "SELECT PR.ID_PRIORIZACION AS ID_PRIORIZACION, CONVERT(VARCHAR(10),FECHA_PRIORIZACION,103) AS FECHA, SUBSTRING(CONVERT(VARCHAR(38),FECHA_PRIORIZACION,121),12,5) AS HORA,"
						strSql = strSql & " TSP.NOM_TIPO_SOLICITUD AS NOM_TIPO_SOLICITUD, TRP.NOM_TIPO_RECLAMO AS NOM_TIPO_RECLAMO, OBSERVACION_PRIORIZACION,"
						strSql = strSql & " U1.LOGIN AS USUARIO_PRIORIZADOR, (CASE WHEN PR.SOLICITA_RESPUESTA = 1 THEN 'SI' ELSE 'NO' END) AS SOL_RESP, U2.LOGIN AS USUARIO_ASIG"


						strSql = strSql & " FROM PRIORIZACION PR INNER JOIN TIPO_SOLICITUD_PRIORIZACION TSP ON TSP.ID_TIPO_SOLICITUD = PR.ID_TIPO_SOLICITUD"
						strSql = strSql & " 				  	 INNER JOIN TIPO_RECLAMO_PRIORIZACION TRP ON TRP.ID_TIPO_RECLAMO = PR.ID_TIPO_RECLAMO"
						strSql = strSql & " 				  	 INNER JOIN USUARIO U1 ON U1.ID_USUARIO = PR.ID_USUARIO_PRIORIZACION"
						strSql = strSql & " 				  	 INNER JOIN DEUDOR ON PR.COD_CLIENTE = DEUDOR.COD_CLIENTE AND PR.RUT_DEUDOR = DEUDOR.RUT_DEUDOR"
						strSql = strSql & " 				  	 LEFT JOIN USUARIO U2 ON DEUDOR.USUARIO_ASIG = U2.ID_USUARIO"

						strSql = strSql & " WHERE PR.COD_CLIENTE = '" & strCodCliente & "' AND PR.RUT_DEUDOR = '" & strRutDeudor & "'"

						strSql = strSql & " ORDER BY ID_PRIORIZACION DESC,FECHA_PRIORIZACION DESC"

						set rsPriorizacion=Conn.execute(strSql)

						If Not rsPriorizacion.Eof Then

						%>
							<tr>
								<td width = "65" class="Estilo4">FECHA</td>
								<td class="Estilo4">HORA</td>
								<td class="Estilo4">TIPO SOLICITUD</td>
								<td class="Estilo4">TIPO RECLAMO</td>
								<td class="Estilo4">SOL. RESP.</td>
								<td width = "100" class="Estilo4">OBSERVACION</td>
								<td class="Estilo4">SOLICITANTE</td>
								<td class="Estilo4">COBRADOR</td>
								<td class="Estilo4">&nbsp;</td>
								<td class="Estilo4">&nbsp;</td>
							</tr>
					</thead>
					<tbody>
						<%

							Do While Not rsPriorizacion.Eof

							AbrirSCG1()
							strSql = "SELECT C.ID_CUOTA, NRO_DOC, C.ID_ULT_GEST, E.GRUPO, ROW_NUMBER() OVER(PARTITION BY C.NRO_DOC ORDER BY ISNULL(NRO_CUOTA,0) ASC) AS SUMNRO_CUOTA,C.NRO_CUOTA, PRC.ESTADO_PRIORIZACION AS ESTADO_PRIORIZACION"

							strSql = strSql & " FROM CUOTA C LEFT JOIN PRIORIZACIONES_CUOTA PRC ON C.ID_CUOTA = PRC.ID_CUOTA"
							strSql = strSql & " 			 INNER JOIN ESTADO_DEUDA E ON C.ESTADO_DEUDA = E.CODIGO"

							strSql = strSql & " WHERE PRC.ID_PRIORIZACION = " & Trim(rsPriorizacion("ID_PRIORIZACION"))
							strSql = strSql & " ORDER BY C.NRO_DOC ASC"

							'Response.write "strSql=" &strSql

							SET rsNomGestion = Conn1.execute(strSql)

							strNroDoc = ""
							strNroDocPag = ""
							strNroDocRet = ""
							strNroDocNoAsig = ""
							intEstadoPrio = 0

							If Not rsNomGestion.Eof Then
								Do While Not rsNomGestion.Eof

									If rsNomGestion("ESTADO_PRIORIZACION") = "0" then

									   intEstadoPrio = intEstadoPrio + 1
									Else
									   intEstadoPrio = intEstadoPrio
									End If


									If rsNomGestion("ESTADO_PRIORIZACION") ="1" then

									   strConfirm = "(C)"

									Else

									   strConfirm = ""

									End If

									If rsNomGestion("SUMNRO_CUOTA") > "1" then

									   strnrocuota = "(" & rsNomGestion("NRO_CUOTA") & ")"
									Else
									   strnrocuota = ""
									End If

									If (Trim(rsNomGestion("GRUPO")) = "ACTIVOS") Then

										strNroDoc = strNroDoc & rsNomGestion("NRO_DOC") & " " & strnrocuota & " " & strConfirm & " - "
										strNroDoc = strNroDoc

									End If

									If (Trim(rsNomGestion("GRUPO")) = "PAGADOS") Then
										strNroDocPag = strNroDocPag & rsNomGestion("NRO_DOC") & " " & strnrocuota & " " & strConfirm & " - "
									End If

									If (Trim(rsNomGestion("GRUPO")) = "RETIROS") Then
										strNroDocRet = strNroDocRet & rsNomGestion("NRO_DOC") & " " & strnrocuota & " " & strConfirm & " - "
									End If

									If (Trim(rsNomGestion("GRUPO")) = "NO ASIGNABLE") Then
										strNroDocNoAsig = strNroDocNoAsig & rsNomGestion("NRO_DOC") & " " & strnrocuota & " " & strConfirm & " - "
									End If

								rsNomGestion.Movenext
								Loop

							End If

							CerrarSCG1()


							Obs=REPLACE(Trim(rsPriorizacion("OBSERVACION_PRIORIZACION")),CHR(13)," ")

							If Obs="" then
								Obs="SIN INFORMACION ADICIONAL"
							End if


							If Trim(strNroDoc) <> "" Then
								strNroDoc = Mid(strNroDoc,1,Len(strNroDoc)-2)
							End If

							If Trim(strNroDocPag) <> "" Then
								strNroDocPag = Mid(strNroDocPag,1,Len( strNroDocPag)-2)
							End If

							If Trim(strNroDocRet) <> "" Then
								strNroDocRet = Mid(strNroDocRet,1,Len( strNroDocRet)-2)
							End If

							If Trim(strNroDocNoAsig) <> "" Then
								strNroDocNoAsig = Mid(strNroDocNoAsig,1,Len( strNroDocNoAsig)-2)
							End If

							intCorr = intCorr + 1

							strTextoDocAct = ""
							strTextoDocPag = ""
							strTextoDocRet = ""
							strTextoDocNoAsig = ""
							strTextoDoc = ""

							If Trim(strNroDoc) <> "" Then
								strTextoDocAct = "Doc.Asociados : " & strNroDoc & "<BR>"
							End If

							If Trim(strNroDocPag) <> "" Then
								strTextoDocPag = "Doc.Cancelados : " & strNroDocPag & "<BR>"
							End If

							If Trim(strNroDocRet) <> "" Then
								strTextoDocRet = "Doc.Desasignados : " & strNroDocRet & "<BR>"
							End If

							If Trim(strNroDocNoAsig) <> "" Then
								strTextoDocNoAsig = "Doc.No Asignable : " & strNroDocNoAsig & "<BR>"
							End If

							strTextoDoc = strTextoDocAct & strTextoDocPag & strTextoDocRet & strTextoDocNoAsig

							%>

							<tr>
								<td><%=rsPriorizacion("FECHA")%></td>
								<td><%=rsPriorizacion("HORA")%></td>
								<td><%=rsPriorizacion("NOM_TIPO_SOLICITUD")%></td>
								<td><%=rsPriorizacion("NOM_TIPO_RECLAMO")%></td>
								<td><%=rsPriorizacion("SOL_RESP")%></td>

								<td class="Estilo4" title="<%=Obs%>"><%=Mid(Obs,1,50)%></td>

								<td><%=rsPriorizacion("USUARIO_PRIORIZADOR")%></td>
								<td><%=rsPriorizacion("USUARIO_ASIG")%></td>
								<td class="Estilo4" title="<%=strTextoDoc%>">
									<img src="../imagenes/carpeta1.png" border="0">
								</td>

								<% If intEstadoPrio > "0" then %>

								<td class="Estilo4" title="Confirmar Priorización">
										<A HREF="#" onClick="ConfirmarPrio(<%=rsPriorizacion("ID_PRIORIZACION")%>,<%=strRutDeudor%>)";><img src="../imagenes/GestionarRoj.png" border="0"></A>

								<% Else%>

								<td class="Estilo4" title="Ver información documentos confirmados">
										<A HREF="#" onClick="ConfirmarPrio(<%=rsPriorizacion("ID_PRIORIZACION")%>,<%=strRutDeudor%>)";><img src="../imagenes/bt_confirmar.jpg" border="0"></A>

								<%End If%>

								</td>

							</tr>

							<%
								rsPriorizacion.movenext
							Loop

							%>
							</tbody>
							<%
							rsPriorizacion.close
							set rsPriorizacion=nothing
					CerrarSCG()

						Else

							%>
							<thead>
							<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
								<td>&nbsp;</td>
							</tr>

							<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
								<td height="30" Align="CENTER" Colspan = "12">DEUDOR NO CUENTA CON HISTORIAL DE PRIORIZACIONES</td>
							</tr>

							<tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
								<td>&nbsp;</td>
							</tr>
							</thead>
							<%
						End If
			%>

			</table>

			</td>
		</tr>

	</table>

<% End If %>

<INPUT TYPE="hidden" NAME="strGraba" value="">

</form>
</body>
</html>
<script type="text/javascript">
	$(document).ready(function(){
		$(document).tooltip();
	})

function envia() {
	//alert(validaCheckbox());
	if(datos.cmb_TipoSol.value=='0')
	{
		alert('Debe seleccionar el tipo de Solicitud');
	}
	else if(validaCheckbox()==false)
	{
		alert('Debe seleccionar al menos 1 documento');
	}

	else
	{
			datos.strGraba.value='SI';
			datos.submit();
			alert('Los documentos seleccionados fueron priorizados!');
	}

}


function validaCheckbox(){
	booResult = false;
	<% For i=1 TO intTamvConcepto %>
			if (document.forms[0].<%=vArrConcepto(i)%>.checked == true)
				booResult = true;
	<% Next %>

	return booResult;
}

function marcar_boxes(){

	datos.TX_CAPITAL.value = 0;
	datos.TX_INTERESES.value = 0;
	datos.TX_PROTESTOS.value = 0;
	datos.TX_HONORARIOS.value = 0;
	datos.TX_SALDO.value = 0;

	desmarcar_boxes()
	<% For i=1 TO intTamvConcepto %>
			if (document.forms[0].<%=vArrConcepto(i)%>.disabled == false) {
			document.forms[0].<%=vArrConcepto(i)%>.checked=true;
			suma_capital(document.forms[0].<%=vArrConcepto(i)%>,document.forms[0].TX_CAPITAL_<%=vArrID_CUOTA(i)%>.value,document.forms[0].TX_INTERESES_<%=vArrID_CUOTA(i)%>.value,document.forms[0].TX_HONORARIOS_<%=vArrID_CUOTA(i)%>.value,document.forms[0].TX_PROTESTOS_<%=vArrID_CUOTA(i)%>.value,document.forms[0].TX_SALDO_<%=vArrID_CUOTA(i)%>.value);
			}
	<% Next %>
}
function desmarcar_boxes(){
		datos.TX_CAPITAL.value = 0;
		datos.TX_INTERESES.value = 0;
		datos.TX_PROTESTOS.value = 0;
		datos.TX_HONORARIOS.value = 0;
		datos.TX_SALDO.value = 0;

		<% For i=1 TO intTamvConcepto %>
			document.forms[0].<%=vArrConcepto(i)%>.checked=false;
		<% Next %>
}

function suma_capital(objeto , intValorSaldoCapital, intValorIntereses, intValorHonorarios, intValorProtestos, intValorSaldo){
		//alert(objeto.checked);

		if (datos.TX_CAPITAL.value == '') datos.TX_CAPITAL.value = 0;
		if (datos.TX_INTERESES.value == '') datos.TX_INTERESES.value = 0;
		if (datos.TX_HONORARIOS.value == '') datos.TX_HONORARIOS.value = 0;
		if (datos.TX_PROTESTOS.value == '') datos.TX_PROTESTOS.value = 0;
		if (datos.TX_SALDO.value == '') datos.TX_SALDO.value = 0;

		if (objeto.checked == true) {
			datos.TX_CAPITAL.value = eval(datos.TX_CAPITAL.value) + eval(intValorSaldoCapital);
			datos.TX_INTERESES.value = eval(datos.TX_INTERESES.value) + eval(intValorIntereses);
			datos.TX_HONORARIOS.value = eval(datos.TX_HONORARIOS.value) + eval(intValorHonorarios);
			datos.TX_PROTESTOS.value = eval(datos.TX_PROTESTOS.value) + eval(intValorProtestos);
			datos.TX_SALDO.value = eval(datos.TX_SALDO.value) + eval(intValorSaldo);
		}
		else
		{
			datos.TX_CAPITAL.value = eval(datos.TX_CAPITAL.value) - eval(intValorSaldoCapital);
			datos.TX_INTERESES.value = eval(datos.TX_INTERESES.value) - eval(intValorIntereses);
			datos.TX_HONORARIOS.value = eval(datos.TX_HONORARIOS.value) - eval(intValorHonorarios);
			datos.TX_PROTESTOS.value = eval(datos.TX_PROTESTOS.value) - eval(intValorProtestos);
			datos.TX_SALDO.value = eval(datos.TX_SALDO.value) - eval(intValorSaldo);
		}
	}
function ventanaGestionesPorDoc (URL){
	window.open(URL,"DATOS","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ventanaMas (URL){
	window.open(URL,"DATOS","width=400, height=500, scrollbars=no, menubar=no, location=no, resizable=yes")
}

function ventanaGestionesPorDoc (URL){
	window.open(URL,"DATOS","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ConfirmarPrio(id_gestion,strRutDeudor)
{
	datos.action = "Confirmar_Priorizacion.asp?strOrigen=<%=strOrigen%>&strRutDeudor=<%=strRutDeudor%>&id_gestion=" + id_gestion ;
	datos.submit();

}


</script>
















