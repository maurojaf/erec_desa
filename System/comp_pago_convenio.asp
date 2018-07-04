<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->

<%
Response.CodePage=65001
Response.charset ="utf-8"

intNroComp=request("intNroComp")
strImprime=request("strImprime")

intCodRemesa = request("CB_REMESA")
intCodUsuario = request("CB_COBRADOR")

intCliente=session("ses_codcli")

If Trim(intCliente) = "" Then intCliente = "1000"

	AbrirSCG()
		strSql="SELECT * FROM PARAMETROS"
		set rsParam = Conn.execute(strSql)
		If not rsParam.eof then
			strNomLogo = Trim(rsParam("NOMBRE_LOGO_TOP_IZQ"))
			strNomSistema = Trim(rsParam("NOMBRE_SISTEMA"))
			strNomEmpresa = Trim(rsParam("NOMBRE_EMPRESA"))
			strDirEmpresa = Trim(rsParam("DIRECCION_EMPRESA"))
			strTelEmpresa = Trim(rsParam("TELEFONOS_EMPRESA"))
			strSitioWebEmpresa = Trim(rsParam("SITIO_WEB_EMPRESA"))
		End if
	CerrarSCG()

%>
	<style type="text/css">
	<!--
	.Estilo37 {color: #FFFFFF}
	.Estilo370 {color: #000000}
	.Estilo372 {
		color: #000000;
		font-size: 14px;
	}

	-->
	</style>
</head>
<body>
<table width="600" align="CENTER" border="0">

<tr>
    <td valign="top">


		<%

		If Trim(intCliente) <> "" then
		abrirscg()

		strSql = "SELECT ID_PAGO, COMP_INGRESO, USR_INGRESO, TIPO_PAGO, COD_CLIENTE, NRO_BOLETA, RUT_DEUDOR , MONTO_CAPITAL, INTERES_PLAZO, GASTOS_JUDICIALES, INDEM_COMP, MONTO_EMP, CONVERT(VARCHAR(10),FECHA_PAGO,103) as FECHA_PAGO, ID_CONVENIO, GASTOS_OTROS, GASTOS_ADMINISTRATIVOS "
		strSql = strSql & " FROM CAJA_WEB_EMP WHERE COMP_INGRESO = " & intNroComp

		set rsCaja=Conn.execute(strSql)
		If Not rsCaja.eof Then

			intIdPago = rsCaja("ID_PAGO")
			intCliente = rsCaja("COD_CLIENTE")
			intUsrIngreso = rsCaja("USR_INGRESO")
			strTipoPago = rsCaja("TIPO_PAGO")
			strBoleta = rsCaja("NRO_BOLETA")

			intIdConvenio = ValNulo(rsCaja("ID_CONVENIO"),"N")


			strRut = rsCaja("RUT_DEUDOR")

			strRutSD = FormatNumber(Mid(strRut,1,len(strRut)-2),0)
			strRutCD = Right(strRut,2)
			strRutForm = strRutSD & strRutCD

			strFechaPago = rsCaja("FECHA_PAGO")
			strMostrarRut = strRut
			intMontoCapital = ValNulo(rsCaja("MONTO_CAPITAL"),"N")
			intIntereses = ValNulo(rsCaja("INTERES_PLAZO"),"N")
			intGastosJudiciales = ValNulo(rsCaja("GASTOS_JUDICIALES"),"N")
			intIndemnizacion = ValNulo(rsCaja("INDEM_COMP"),"N")
			intHonorarios = ValNulo(rsCaja("MONTO_EMP"),"N")

			intGastosOperacionales = ValNulo(rsCaja("GASTOS_OTROS"),"N")
			intGastosAdministrativos = ValNulo(rsCaja("GASTOS_ADMINISTRATIVOS"),"N")
			intTotalPago = intMontoCapital + intIntereses + intGastosJudiciales + intIndemnizacion + intHonorarios + intGastosOperacionales + intGastosAdministrativos



			strSql = "SELECT FOLIO,SEDE FROM CONVENIO_ENC WHERE ID_CONVENIO = " & intIdConvenio
			set rsConv=Conn.execute(strSql)
			if not rsConv.eof then
				intFolio = rsConv("FOLIO")
				strSede = rsConv("SEDE")
			end if
			rsConv.close
			set rsConv=nothing

			ssql="SELECT NOMBRE_DEUDOR FROM DEUDOR WHERE RUT_DEUDOR='" & strRut & "' AND COD_CLIENTE = '" & intCliente & "'"
			set RsDeudor=Conn.execute(ssql)
			if not RsDeudor.eof then
				strNombreDeudor = RsDeudor("NOMBRE_DEUDOR")
			end if
			RsDeudor.close
			set RsDeudor=nothing

			ssql=""
			ssql="SELECT TOP 1 Calle,Numero,Comuna,Resto,CORRELATIVO,Estado FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR='"&strRut&"' and ESTADO<>'2' ORDER BY CORRELATIVO DESC"
			set rsDIR=Conn.execute(ssql)
			if not rsDIR.eof then
				calle_deudor=rsDIR("Calle")
				numero_deudor=rsDIR("Numero")
				comuna_deudor=rsDIR("Comuna")
				resto_deudor=rsDIR("Resto")
				strDirDeudor = calle_deudor & " " & numero_deudor & " " & resto_deudor & " " & comuna_deudor
			end if
			rsDIR.close
			set rsDIR=nothing

			ssql=""
			ssql="SELECT TOP 1 COD_AREA,TELEFONO,CORRELATIVO,ESTADO FROM DEUDOR_TELEFONO WHERE  RUT_DEUDOR='"&strRut&"' and ESTADO<>'2' ORDER BY CORRELATIVO DESC"
			set rsFON=Conn.execute(ssql)
			if not rsFON.eof then
				codarea_deudor = rsFON("COD_AREA")
				Telefono_deudor = rsFON("Telefono")
				strFono = codarea_deudor & "-" & Telefono_deudor
			end if
			rsFON.close
			set rsFON=nothing


			strSql="SELECT TOP 1 SUCURSAL, ADIC_5, CUOTA.NRO_CLIENTE_DEUDOR AS NRO_CLIENTE_DEUDOR FROM CUOTA WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strRut & "'"
			''strSql = strSql & " AND NRO_DOC IN (SELECT NRO_DOC FROM CAJA_WEB_EMP_DETALLE  WHERE ID_PAGO = " & intIdPago & ") AND SUCURSAL IS NOT NULL"
			''rESPONSE.WRITE strSql
			set rsSuc=Conn.execute(strSql)
			if not rsSuc.eof then
				strSucursal = rsSuc("SUCURSAL")
				strInterlocutor = rsSuc("NRO_CLIENTE_DEUDOR")
			end if
			rsSuc.close
			set rsSuc=nothing


			ssql=""
			ssql="SELECT TOP 1 RUT_DEUDOR,CORRELATIVO,FECHA_INGRESO,EMAIL,ESTADO FROM DEUDOR_EMAIL WHERE  RUT_DEUDOR='"&strRut&"' and ESTADO<>'2' ORDER BY CORRELATIVO DESC"
			set rsMAIL=Conn.execute(ssql)
			if not rsMAIL.eof then
				strEmail = rsMAIL("EMAIL")
			end if
			rsMAIL.close
			set rsMAIL=nothing

			strNomCliente = TraeCampoId(Conn, "DESCRIPCION", intCliente, "CLIENTE", "COD_CLIENTE")
			strRutCliente = TraeCampoId(Conn, "RUT", intCliente, "CLIENTE", "COD_CLIENTE")


			If Trim(strRutCliente) <> "" Then
				strRutCSD = FormatNumber(Mid(strRutCliente,1,len(strRutCliente)-2),0)
				strRutCCD = Right(strRutCliente,2)
				strRutCli = strRutCSD & strRutCCD
				strRutCli = "RUT: " & strRutCli
			End If
			strNomCliente = strRutCli  & " " & strNomCliente

			strSql="SELECT B.RUT, B.RAZON_SOCIAL FROM SEDE A, CONVENIO_CORRELATIVO B WHERE A.RUT = B.RUT AND A.COD_CLIENTE = B.COD_CLIENTE AND B.COD_CLIENTE = '" & intCliente & "' AND A.SEDE = '" & strSede & "'"
			'Response.write "strSql=" & strSql
			set rsSede=Conn.execute(strSql)
			if not rsSede.eof then
				strRutCli = rsSede("RUT")
				If Trim(strRutCli) <> "" Then
					strRutCSD = FormatNumber(Mid(strRutCli,1,len(strRutCli)-2),0)
					strRutCCD = Right(strRutCli,2)
					strRutCli = strRutCSD & strRutCCD
					strRutCli = "RUT: " &strRutCli
				End If
				strRSocialCli = rsSede("RAZON_SOCIAL")
				strDescCli = strRutCli & " " & strRSocialCli
			end if
			rsSede.close
			set rsSede=nothing
			''Response.write "strDescCli=" & strDescCli
			If Trim(strDescCli) <> "" Then
				strNomCliente = strDescCli
			End if




			strUsrIngreso = TraeCampoId(Conn, "NOMBRES_USUARIO", intUsrIngreso, "USUARIO", "ID_USUARIO") & " " & TraeCampoId(Conn, "APELLIDO_PATERNO", intUsrIngreso, "USUARIO", "ID_USUARIO")
			strDescTipoPago = TraeCampoId2(Conn, "DESC_TIPO_PAGO", strTipoPago, "CAJA_TIPO_PAGO", "ID_TIPO_PAGO")

			strFormaPago=""
			strSql = "SELECT DISTINCT IsNull(FORMA_PAGO,'') AS FORMA_PAGO FROM CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO = " & intIdPago
			set rsFPago=Conn.execute(strSql)
			Do While Not rsFPago.eof
				strFormaPago = strFormaPago & " - " & TRIM(TraeCampoId2(Conn, "DESC_FORMA_PAGO", rsFPago("FORMA_PAGO"), "CAJA_FORMA_PAGO", "ID_FORMA_PAGO"))
				rsFPago.movenext
			Loop
			rsFPago.close
			set rsFPago=nothing

			If Trim(strFormaPago) <> "" Then strFormaPago = Mid(strFormaPago, 3, LEN(strFormaPago))


			'strSql = "SELECT DISTINCT IsNull(COD_REMESA,0) AS COD_REMESA FROM CUOTA WHERE RUT_DEUDOR = '" & strRut & "' AND COD_CLIENTE = '" & intCliente & "'"
			'strSql = strSql & " AND NRO_DOC IN (SELECT NRO_DOC FROM CONVENIO_CUOTA WHERE ID_CONVENIO = " & intIdConvenio & ")"

			strSql = "SELECT DISTINCT IsNull(COD_REMESA,0) AS COD_REMESA FROM CONVENIO_CUOTA WHERE ID_CONVENIO = " & intIdConvenio
			'REsponse.write "strSql=" & strSql
			'REsponse.End

			set rsAsign=Conn.execute(strSql)
			Do While Not rsAsign.eof
				strAsignacion = strAsignacion & " - " & TRIM(rsAsign("COD_REMESA"))
				rsAsign.movenext
			Loop
			rsAsign.close
			set rsAsign=nothing
			If Trim(strAsignacion) <> "" Then strAsignacion = Mid(strAsignacion, 3, LEN(strAsignacion))

			%>

		<table width="600" border="0">
			<tr>
				<td><img src="../imagenes/reintegra.jpg" width="154" height="50"></td>
				<td><span class="Estilo370"><B>COMPROBANTE DE PAGO DE CONVENIO</B></td>
				<td width="154">
					<table border="0">
						<tr>
							<td><span class="Estilo370"><B>NRO.OPERACION :</B></td>
							<td align="RIGHT"><span class="Estilo370"><B><%=rsCaja("COMP_INGRESO")%></B></td>
						</tr>
						<tr>
							<td><span class="Estilo370">NRO.PAGARÉ :</td>
							<td align="RIGHT"><span class="Estilo370"><%=intFolio%></td>
						</tr>
						<tr>
							<td><span class="Estilo370">Nro.Comprobante :</td>
							<td align="RIGHT"><span class="Estilo370"><%=rsCaja("COMP_INGRESO")%></td>
						</tr>
						<tr>
							<td><span class="Estilo370">Boleta :</td>
							<td align="RIGHT"><span class="Estilo370"><%=strBoleta%></td>
						</tr>
						<tr>
							<td><span class="Estilo370">Fecha :</td>
							<td align="RIGHT"><span class="Estilo370"><%=rsCaja("FECHA_PAGO")%></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>

		<table width="600" border="0">
				<tr>
					<td colspan=4>&nbsp</td>
				</tr>
				<tr>
					<td colspan=4><span class="Estilo370"><b>Datos Deudor</b></td>
				</tr>
				<tr>
					<td width="100"><span class="Estilo370">Nombre :</td>
					<td width="300"><span class="Estilo370"><%=strNombreDeudor%></td>
					<td width="100"><span class="Estilo370">Rut :</td>
					<td width="100"><span class="Estilo370"><%=strRutForm%></td>
				</tr>
				<tr>
					<td><span class="Estilo370">Direccion :</td>
					<td><span class="Estilo370"><%=strDirDeudor%></td>
					<td><span class="Estilo370">Interlocutor :</td>
					<td><span class="Estilo370"><%=strInterlocutor%></td>
				</tr>
				<tr>
					<td><span class="Estilo370">Telefono celular :</td>
					<td><span class="Estilo370"><%=strFonoCelular%></td>
					<td><span class="Estilo370">Email :</td>
					<td><span class="Estilo370"><%=strEmail%></td>
				</tr>
			</table>



		<table width="600" border="0">
			<tr>
				<td colspan=4>&nbsp</td>
			</tr>
			<tr>
				<td colspan=4><span class="Estilo370"><b>Datos Deuda</b></td>
			</tr>
			<tr>
				<td width="100"><span class="Estilo370">Cedente :</td>
				<td width="500"><span class="Estilo372"><B><%=strNomCliente%></B></td>
			</tr>
		</table>

		<table width="600" border="0">
			<tr>
				<td>&nbsp</td>
			</tr>
			<tr>
				<td colspan=4><span class="Estilo370"><b>Detalle Pago</b></td>
			</tr>
			<tr>
				<td width="100"><span class="Estilo370">Forma de Pago :</td>
				<td width="300"><span class="Estilo370"><%=strFormaPago%></td>
				<td width="100"><span class="Estilo370">Tipo de Pago :</td>
				<td width="100"><span class="Estilo370"><%=strDescTipoPago%>&nbsp;(<%=intIdConvenio%>)</td>
			</tr>
			<tr>
				<td>&nbsp</td>
			</tr>
		</table>


<table width="600" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
	<tr>
		<td><span class="Estilo370">TIPO</span></td>
		<td><span class="Estilo370">BANCO</span></td>
		<td><span class="Estilo370">CTA CTE NRO.</span></td>
		<td><span class="Estilo370">CHEQUE</span></td>
		<td><span class="Estilo370">MONTO</span></td>
		<td><span class="Estilo370">FECHA.</span></td>
	</tr>

	<%
		AbrirSCG1()
			if 1=2 Then
				strSql = "SELECT SUM(MONTO) AS MONTO , TIPO_PAGO, DIVIDE FROM CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO = " & rsCaja("ID_PAGO") & " AND DIVIDE = '1' GROUP BY TIPO_PAGO, DIVIDE"
				set rsTemp=Conn.execute(strSql)
				Do While Not rsTemp.eof
					If Trim(rsTemp("TIPO_PAGO")) = "1" Then 'Cliente
						intValorHonorarios = ValNulo(rsTemp("MONTO"),"N")
					End If
					If Trim(rsTemp("TIPO_PAGO")) = "0" Then 'Empresa
						intValorRemesar = ValNulo(rsTemp("MONTO"),"N")
					End If
					rsTemp.movenext
				Loop
			End if
		CerrarSCG1()



			AbrirSCG1()

				strSql = "SELECT SUM(MONTO) AS MONTO , TIPO_PAGO, DIVIDE FROM CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO = " & rsCaja("ID_PAGO") & " GROUP BY TIPO_PAGO, DIVIDE"
				'Response.write "strSql=" & strSql
				set rsTemp=Conn.execute(strSql)
				Do While Not rsTemp.eof

					If Trim(rsTemp("TIPO_PAGO")) = "1" and Trim(rsTemp("DIVIDE")) = "1" Then 'Cliente
						intValorHonorarios = ValNulo(rsTemp("MONTO"),"N")
					End If

					If Trim(rsTemp("TIPO_PAGO")) = "0" and Trim(rsTemp("DIVIDE")) = "2" Then 'Empresa
						intAbonoInicial = ValNulo(rsTemp("MONTO"),"N")
					End If

					If Trim(rsTemp("TIPO_PAGO")) = "0" and Trim(rsTemp("DIVIDE")) = "0" Then 'Empresa
						intSaldoCliente = ValNulo(rsTemp("MONTO"),"N")
					End If

					rsTemp.movenext
				Loop
				CerrarSCG1()

					IntTotalCP = intValorHonorarios + intAbonoInicial + intSaldoCliente
					IntTotalRecaudacion = intValorHonorarios + intAbonoInicial









		strSql = "SELECT ISNULL(NOMBRE_B,'') AS NOMBRE_B, ISNULL(NRO_CTA_CTE,'') AS NRO_CTA_CTE, ISNULL(NRO_DEPOSITO,'') AS NRO_DEPOSITO, ISNULL(NRO_CHEQUE,'') AS NRO_CHEQUE, MONTO, ISNULL(VENCIMIENTO,'') AS VENCIMIENTO, FORMA_PAGO FROM CAJA_WEB_EMP_DOC_PAGO, BANCOS WHERE BANCOS.CODIGO =* CAJA_WEB_EMP_DOC_PAGO.COD_BANCO AND ID_PAGO = " & rsCaja("ID_PAGO") & " ORDER BY FORMA_PAGO, VENCIMIENTO, NRO_CTA_CTE, NRO_CHEQUE"
		''rESPONSE.WRITE strSql
		set rsDocPago=Conn.execute(strSql)
		If Not rsDocPago.eof Then
			strBanco = 0
			intSumaCapital = 0

		Do While not rsDocPago.Eof
			strBanco = rsDocPago("NOMBRE_B")
			strTipoPag=Trim(rsDocPago("FORMA_PAGO"))

			If Trim(rsDocPago("FORMA_PAGO")) = "EF" Then strBanco  = "PAGO EN EFECTIVO"
			If Trim(rsDocPago("FORMA_PAGO")) = "CU" Then strBanco  = "CUOTA"
			If Trim(rsDocPago("FORMA_PAGO")) = "AB" Then strBanco  = "ABONO"

			If Trim(rsDocPago("FORMA_PAGO")) = "DP" Then
				strNroCheque = rsDocPago("NRO_DEPOSITO")
			Else
				strNroCheque = rsDocPago("NRO_CHEQUE")
			End If


			strCtaCte = rsDocPago("NRO_CTA_CTE")
			If trim(strCtaCte) = "" Then strCtaCte = "&nbsp;"
			strNroCheque = rsDocPago("NRO_CHEQUE")
			If trim(strNroCheque) = "" Then strNroCheque = "&nbsp;"
			strMonto = rsDocPago("MONTO")
			strVencimiento = Saca1900(rsDocPago("VENCIMIENTO"))
			If trim(strVencimiento) = "" Then strVencimiento = "&nbsp;"

			If TrIM(strMonto) <> "0" Then
	%>

		<tr>
			<td><%=strTipoPag%></td>
			<td><%=strBanco%></td>
			<td><%=strCtaCte%></td>
			<td><%=strNroCheque%></td>
			<td ALIGN="RIGHT"><%=FN(strMonto,0)%></td>
			<td><%=strVencimiento%></td>
		</tr>

	<%
			End If
			rsDocPago.movenext
		Loop
		End If

		If Trim(strTipoPago) = "CO" Then
			strGlosaTotal = "TOTAL PAGO"
		Else
			strGlosaTotal = "TOTAL PAGO"
		End if
	%>
</table>

<table width="600">
	<tr>
		<td>&nbsp</td>
	</tr>
</table>



<%


				strSql = "SELECT COUNT(*) as CANTC FROM CONVENIO_DET WHERE CUOTA <> 0 AND ID_CONVENIO = " & intIdConvenio
				set rsDetConv=Conn.execute(strSql)
				If Not rsDetConv.Eof Then
					intNroCuotas = rsDetConv("CANTC")
				Else
					intNroCuotas = 0
				End If

				strSql = "SELECT * FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO = " & rsCaja("ID_PAGO")
				set rsDetCaja=Conn.execute(strSql)
				If Not rsDetCaja.eof Then
					intNroDoc = 0
					intSumaCapital = 0
					strFacturas=""
					strPie=""
					strCuotas=""
					intPagoPie = 0
					intPagoCuota = 0
					strDetalleCuotas = ""
					Do While not rsDetCaja.Eof

						If Trim(rsDetCaja("NRO_DOC")) = "0" Then
							strPie = rsDetCaja("NRO_DOC")
							intPagoPie = intPagoPie + rsDetCaja("CAPITAL")
						End If
						If Trim(rsDetCaja("NRO_DOC")) <> "0" Then
							strCuotas = strCuotas & ", " & rsDetCaja("NRO_DOC")
							intPagoCuota = intPagoCuota + rsDetCaja("CAPITAL")
						End If

						strSql="SELECT CUOTA, FECHA_PAGO FROM CONVENIO_DET WHERE CUOTA <> 0 AND ID_CONVENIO = " & intIdConvenio & " AND CUOTA = " & rsDetCaja("NRO_DOC")
						set RsCuota=Conn.execute(strSql)
						If not RsCuota.eof then
							strNroCuota = RsCuota("CUOTA")
							strFechaPagoCuota = RsCuota("FECHA_PAGO")
						Else
							strNroCuota = ""
							strFechaPagoCuota = ""

						End if
						RsCuota.close
						set RsCuota=nothing
						strNombreDeudor=""
						strMostrarRut=""

						If Trim(strNroCuota) <> "" Then
							strDetalleCuotas = strDetalleCuotas & "<BR>Cuota " & strNroCuota & ", Vencimiento = " & strFechaPagoCuota
						End if

					rsDetCaja.movenext
					intNroDoc = intNroDoc + 1
					Loop
				End If
				strCuotas=Mid(strCuotas,2,len(strCuotas))

				strCajaNro="Santiago"

				If trim(strPie) <> "" Then
					strPie = "Pie del convenio / pagaré"
				End if
				If trim(strCuotas) <> "" Then
					strCuotas = "Pago de las siguientes Cuotas :" & strCuotas & " de " & intNroCuotas
				End if
			%>


<table width="600">
	<tr>
		<td>&nbsp</td>
	</tr>
</table>

<table width="600"  border="0">
	<tr>
		<td>
			<table width="300" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
					<tr>
						<td colspan=2><span class="Estilo370"><b>Detalle Deuda</b></td>
					</tr>
					<tr>
						<td><span class="Estilo370" width="125">PIE</span></td>
						<td ALIGN="RIGHT" width="125"><%=FN(intAbonoInicial,0)%></td>
					</tr>
					<tr>
						<td><span class="Estilo370">CUOTAS</span></td>
						<td ALIGN="RIGHT"><%=FN(intPagoCuota,0)%></td>
					</tr>
					<tr>
						<td><span class="Estilo370">INTERESES</span></td>
						<td ALIGN="RIGHT"><%=FN(intIntereses,0)%></td>
					</tr>
					<tr>
						<td><span class="Estilo370">GASTOS JUDICIALES</span></td>
						<td ALIGN="RIGHT"><%=FN(intGastosJudiciales,0)%></td>
					</tr>
					<tr>
						<td><span class="Estilo370">INDEMNIZACION</span></td>
						<td ALIGN="RIGHT"><%=FN(intIndemnizacion,0)%></td>
					</tr>
					<tr>
						<td><span class="Estilo370">GASTOS DE COBRANZAS</span></td>
						<td ALIGN="RIGHT"><%=FN(intValorHonorarios,0)%></td>
					</tr>
					<tr>
						<td><span class="Estilo370">GASTOS OPERACIONALES</span></td>
						<td ALIGN="RIGHT"><%=FN(intGastosOperacionales,0)%></td>
					</tr>
					<tr>
						<td><span class="Estilo370">GASTOS ADMINISTRATIVOS</span></td>
						<td ALIGN="RIGHT"><%=FN(intGastosAdministrativos,0)%></td>
					</tr>
					<tr>
						<td><span class="Estilo370"><%=strGlosaTotal%></span></td>
						<td ALIGN="RIGHT"><%=FN(intTotalPago,0)%></td>
				</tr>
			</table>

		</td>

		<!--td style="vertical-align: bottom;" align="center">
			Información uso exclusivo de <%=strNomEmpresa%>
			<table width="250" border="1" bordercolor = "#808080" cellSpacing=0 cellPadding=1>
				<tr class="Estilo371">
					<td ALIGN="CENTER">Boleta N.</td>
					<td ALIGN="CENTER">Fecha</td>
					<td ALIGN="CENTER">Honorarios</td>
				</tr>
				<tr class="Estilo371">
					<td ALIGN="CENTER"><%=strBoleta%></td>
					<td ALIGN="CENTER"><%=strFechaPago%></td>
					<td ALIGN="RIGHT"><%=FN(intValorHonorarios,0)%></td>
				</tr>
				<tr class="Estilo371">
					<td colspan=2>Valor a remesar</td>
					<td ALIGN="RIGHT"><%=FN(intValorRemesar,0)%></td>
				</tr>
				<tr class="Estilo371">
					<td colspan=2>Totales</td>
					<td ALIGN="RIGHT"><%=FN(intValorHonorarios + intValorRemesar,0)%></td>
				</tr>
			</table>
		</td-->

		<td style="vertical-align: top" align="center">

			<table width="250" border="1" bordercolor = "#808080" cellSpacing=0 cellPadding=1>
				<tr class="Estilo371">
					<td>Descripción</td>
					<td>Nº Boleta</td>
					<td>Monto</td>
				</tr>
				<tr class="Estilo371">
					<td ALIGN="LEFT">Honorarios</td>
					<td ALIGN="LEFT"><%=strBoleta%></td>
					<td ALIGN="RIGHT"><%=FN(intValorHonorarios,0)%></td>
				</tr>
				<tr class="Estilo371">
					<td ALIGN="LEFT" colspan = "2">Abono Inicial</td>
					<td ALIGN="RIGHT"><%=FN(intAbonoInicial,0)%></td>
				</tr>
				<tr class="Estilo371">
					<td colspan = "2" ALIGN="LEFT">Total Recaudación</td>
					<td ALIGN="RIGHT"><%=FN(IntTotalRecaudacion,0)%></td>
				</tr>
				<tr class="Estilo371">
					<td ALIGN="LEFT" colspan = "2">Saldo Cliente</td>
					<td ALIGN="RIGHT"><%=FN(intSaldoCliente,0)%></td>
				</tr>
				<tr class="Estilo371">
					<td colspan = "2" ALIGN="LEFT">Total Comprobante de Pago</td>
					<td ALIGN="RIGHT"><%=FN(IntTotalCP,0)%></td>
				</tr>
			</table>
			Información uso exclusivo de <%=strNomEmpresa%>
		</td>


	</tr>
</table>
<table width="600">
	<tr>
		<td>&nbsp</td>
	</tr>
</table>

<table width="600" border="0">
	<tr>
		<TD VALIGN = "TOP">

			<TABLE VALIGN = "TOP" width="400" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
				<tr>
					<td><span class="Estilo370"><b>Detalle del Pago</b></td>
				</tr>
				<tr>
					<td><span class="Estilo370" width="125"><%=strPie&""&strCuotas&" "&strDetalleCuotas%>Nº<%=intFolio%></td>
				</tr>
			</table>

		</td>
		<td VALIGN = "TOP">

			<TABLE width="200" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
				<tr>
					<td width="100"><span class="Estilo370"><b>Caja:</b></td>
					<td width="100"><span class="Estilo370" width="125"><%=strCajaNro%></td>
				</tr>
				<tr>
					<td><span class="Estilo370"><b>Ejecutivo :</b></td>
					<td><span class="Estilo370" width="125"><%=strUsrIngreso%></td>
				</tr>
			</table>



		</td>
	</tr>
</table>


<%If Trim(strTipoPago) = "CO" Then%>
<br>
<br>
<!--TABLE WIDTH="600" border="0">
	<TR>
		<TD VALIGN = "TOP"><b>
			En caso de incumpliento o simple atraso en el pago de cualquiera de las
			cuotas  establecidas, LLACRUZ.  y/o nuestro  Mandante  quedan  facultadas  para  continuar el
			ejercicio  de las acciones legales de  cobro, devengandose como interes, el maximo convencional
			estipulado por la Ley.
			</b>
		</TD>
	</TR>
</TABLE-->
<br>
<%End If%>




<table width="600">
	<tr>
		<td>&nbsp</td>
	</tr>
</table>

<table width="600" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
		<tr>
		<td class="Estilo370" ALIGN="CENTER">
		<%=strNomEmpresa%>&nbsp;&nbsp;Dirección : <%=strDirEmpresa%>&nbsp;&nbsp;Telefono : 2<%=strTelEmpresa%>
		<BR>			<%=strSitioWebEmpresa%>

		</td>
	</tr>
</table>



			<%

		End If
		rsCaja.close
		set rsCaja=nothing
		''Response.End
		%>


		<%	cerrarscg()
		end if %>

	</td>
	</tr>
	</table>


	</td>



</table>
</body>
</html>