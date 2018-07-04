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

intNroComp=request("intNroComp")
strImprime=request("strImprime")

intCodRemesa = request("CB_REMESA")
intCodUsuario = request("CB_COBRADOR")

intCliente=session("ses_codcli")

If Trim(intCliente) = "" Then intCliente = "1000"

%>
	<style type="text/css">
	<!--
	.Estilo37 {color: #FFFFFF}
	.Estilo370 {color: #000000}
	.Estilo371 {color: #808080}
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

		strSql = "SELECT ID_CONVENIO,ID_PAGO, COMP_INGRESO, USR_INGRESO, TIPO_PAGO, COD_CLIENTE, NRO_BOLETA, RUT_DEUDOR , MONTO_CAPITAL, INTERES_PLAZO, GASTOS_JUDICIALES, INDEM_COMP, MONTO_EMP, GASTOS_ADMINISTRATIVOS, GASTOS_OTROS, CONVERT(VARCHAR(10),FECHA_PAGO,103) as FECHA_PAGO"
		strSql = strSql & " FROM CAJA_WEB_EMP WHERE COMP_INGRESO = " & intNroComp

		set rsCaja=Conn.execute(strSql)
		If Not rsCaja.eof Then

			intIdPago = rsCaja("ID_PAGO")
			intCliente = rsCaja("COD_CLIENTE")
			intUsrIngreso = rsCaja("USR_INGRESO")
			strTipoPago = rsCaja("TIPO_PAGO")
			strBoleta = rsCaja("NRO_BOLETA")
            
			strRut = rsCaja("RUT_DEUDOR")
            intIdConvenio = ValNulo(rsCaja("ID_CONVENIO"),"N")

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

			ssql="SELECT NOMBRE_DEUDOR FROM DEUDOR WHERE RUT_DEUDOR='" & strRut & "' AND COD_CLIENTE = '" & intCliente & "'"
			set RsDeudor=Conn.execute(ssql)
			if not RsDeudor.eof then
				strNombreDeudor = RsDeudor("NOMBRE_DEUDOR")
			end if
			RsDeudor.close
			set RsDeudor=nothing

			ssql=""
			ssql="SELECT TOP 1 Calle,Numero,Comuna,Resto,CORRELATIVO,Estado FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR='"& strRut &"' and ESTADO<>'2' ORDER BY CORRELATIVO DESC"
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


			strSql="SELECT TOP 1 SUCURSAL, ADIC_5, CUOTA.NRO_CLIENTE_DEUDOR AS NRO_CLIENTE_DEUDOR FROM CUOTA WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strRut & "'"
			
			if strTipoPago <> "CO" then
				strSql = strSql & " AND NRO_DOC IN (SELECT NRO_DOC FROM CAJA_WEB_EMP_DETALLE  WHERE ID_PAGO = " & intIdPago & ")"
			end if
			
			strSql = strSql & " AND SUCURSAL IS NOT NULL"
			set rsSuc=Conn.execute(strSql)
			if not rsSuc.eof then
				strSucursal = rsSuc("SUCURSAL")
				strInterlocutor = rsSuc("NRO_CLIENTE_DEUDOR")
			end if
			rsSuc.close
			set rsSuc=nothing


			strSql="SELECT B.RUT, B.RAZON_SOCIAL FROM SEDE A, CONVENIO_CORRELATIVO B WHERE A.RUT = B.RUT AND A.COD_CLIENTE = B.COD_CLIENTE AND B.COD_CLIENTE = '" & intCliente & "' AND A.SEDE = '" & strSucursal & "'"
			''Response.write "strSql=" & strSql
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
				strDescCli = strRutCli & "   /   " & strRSocialCli
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

			strSql = "SELECT DISTINCT IsNull(COD_REMESA,0) AS COD_REMESA FROM CUOTA WHERE RUT_DEUDOR = '" & strRut & "' AND COD_CLIENTE = '" & intCliente & "'"
			strSql = strSql & " AND NRO_DOC IN (SELECT NRO_DOC FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO = " & intIdPago & ")"
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
				<td><img src="../Imagenes/Logos/llacruz/llacruz_logo_1.jpg" width="154" height="50"></td>
				<td><span class="Estilo370">
                <%if strTipoPago <> "CO" then %>
                <B>COMPROBANTE DE PAGO</B></td>
                <%else%>
                <B>COMPROBANTE DE PAGO DE CONVENIO</B>
                <%end if %>
				<td width="154">
					<table border="0">
						<tr>
							<td><span class="Estilo370"><B>N.OPERACION :</B></td>
							<td align="RIGHT"><span class="Estilo370"><B><%=rsCaja("COMP_INGRESO")%></B></td>
						</tr>
						<tr>
							<td><span class="Estilo370">N.COMPROB.:</td>
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
				<td colspan=8>&nbsp</td>
			</tr>
			<tr>
				<td colspan=8><span class="Estilo370"><b>Datos Deuda</b></td>
			</tr>
			<tr>
				<td width="100"><span class="Estilo370">Cedente :</td>
				<td width="400"><span class="Estilo372"><B><%=strNomCliente%></B></td>
				<td width="70">&nbsp;</td>
				<td width="30">&nbsp;</td>
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
				<td width="100"><span class="Estilo370"><%=strDescTipoPago%></td>
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
		<td><span class="Estilo370">CHEQUE/DEP</span></td>
		<td><span class="Estilo370">FECHA.</span></td>
		<td><span class="Estilo370">MONTO</span></td>
	</tr>

	<%
    
		AbrirSCG1()

		strSql = "SELECT SUM(MONTO) AS MONTO , TIPO_PAGO, DIVIDE FROM CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO = " & rsCaja("ID_PAGO") & " GROUP BY TIPO_PAGO, DIVIDE"
		'
        'Response.write "strSql=" & strSql
        'Response.end
		set rsTemp=Conn.execute(strSql)
		Do While Not rsTemp.eof

			If Trim(rsTemp("TIPO_PAGO")) = "1" and Trim(rsTemp("DIVIDE")) = "1" Then 'Cliente
				intValorHonorarios = ValNulo(rsTemp("MONTO"),"N")
			End If

			If Trim(rsTemp("TIPO_PAGO")) = "0" and Trim(rsTemp("DIVIDE")) = "2" Then 'Empresa
				intAbonoInicial = ValNulo(rsTemp("MONTO"),"N")
			End If

			If Trim(rsTemp("TIPO_PAGO")) = "0" and Trim(rsTemp("DIVIDE")) = "3" Then 'Empresa
				intSaldoCliente = ValNulo(rsTemp("MONTO"),"N")
			End If

			rsTemp.movenext
		Loop
		CerrarSCG1()

			IntTotalCP = intValorHonorarios + intAbonoInicial + intSaldoCliente + intGastosAdministrativos
			IntTotalRecaudacion = intValorHonorarios + intAbonoInicial + intGastosAdministrativos + intSaldoCliente

		strSql = "SELECT NOMBRE_B , IsNull(NRO_CTA_CTE,'') as NRO_CTA_CTE, IsNull(NRO_CHEQUE,'') as NRO_CHEQUE, IsNull(NRO_DEPOSITO,'') as NRO_DEPOSITO, SUM(MONTO) as MONTO,  IsNull(VENCIMIENTO,'') as VENCIMIENTO , FORMA_PAGO "
		strSql = strSql & " FROM CAJA_WEB_EMP_DOC_PAGO, BANCOS WHERE BANCOS.CODIGO = IsNull(CAJA_WEB_EMP_DOC_PAGO.COD_BANCO,0) AND ID_PAGO = " & rsCaja("ID_PAGO")
		strSql = strSql & " GROUP BY NOMBRE_B , NRO_CTA_CTE, NRO_CHEQUE, NRO_DEPOSITO, VENCIMIENTO , FORMA_PAGO ORDER BY VENCIMIENTO ASC, NRO_CTA_CTE, NRO_CHEQUE"

		set rsDocPago=Conn.execute(strSql)
		If Not rsDocPago.eof Then
			strBanco = 0
			intSumaCapital = 0
			intTotalMonto = 0

		Do While not rsDocPago.Eof
			strBanco = rsDocPago("NOMBRE_B")
			'Response.write "<br>FORMA_PAGO = " & Trim(rsDocPago("FORMA_PAGO"))
			'Response.write "<br>strBanco = " & strBanco

			If Trim(rsDocPago("FORMA_PAGO")) = "EF" THEN
				strTipoPag="EFECTIVO"
			ElseIf Trim(rsDocPago("FORMA_PAGO")) = "DP" THEN
				strTipoPag="DEPOSITO"
			ElseIf Trim(rsDocPago("FORMA_PAGO")) = "TR" THEN
				strTipoPag="TRANSFERENCIA"
			ElseIf Trim(rsDocPago("FORMA_PAGO")) = "CD" THEN
				strTipoPag="CHEQUE AL DIA"
			ElseIf Trim(rsDocPago("FORMA_PAGO")) = "CF" THEN
				strTipoPag="CHEQUE A FECHA"
			ElseIf Trim(rsDocPago("FORMA_PAGO")) = "VV" THEN
				strTipoPag="VALE VISTA"
			ElseIf Trim(rsDocPago("FORMA_PAGO")) = "EF" THEN
				strTipoPag="EFECTIVO"
			ELSE
				strTipoPag = Trim(rsDocPago("FORMA_PAGO"))
			END IF


			If Trim(rsDocPago("FORMA_PAGO")) = "DP" Then
				strNroCheque = rsDocPago("NRO_DEPOSITO")
			Else
				strNroCheque = rsDocPago("NRO_CHEQUE")
			End If

			'Response.write "<br>strBanco = " & strBanco


			strCtaCte = rsDocPago("NRO_CTA_CTE")
			If trim(strCtaCte) = "" Then strCtaCte = "&nbsp;"

			If trim(strNroCheque) = "" Then strNroCheque = "&nbsp;"
			strMonto = rsDocPago("MONTO")

			strVencimiento = Saca1900(rsDocPago("VENCIMIENTO"))

			If trim(strVencimiento) = "" Then strVencimiento = "&nbsp;"


			strTotalMonto = strTotalMonto + strMonto
	%>

		<tr>
			<td><%=strTipoPag%></td>
			<td><%=strBanco%></td>
			<td><%=strCtaCte%></td>
			<td><%=strNroCheque%></td>
			<td><%=strVencimiento%></td>
			<td ALIGN="RIGHT"><%=FN(strMonto,0)%></td>
		</tr>

	<%
			rsDocPago.movenext
		Loop
		End If %>

		<tr>
			<td Colspan = "1">TOTAL DEUDA</td>
			<td Colspan = "4">&nbsp</td>
			<td ALIGN="RIGHT"><%=FN(strTotalMonto,0)%></td>
		</tr>

	<%	If Trim(strTipoPago) = "CO" Then
			strGlosaTotal = "TOTAL DEUDA"
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



<%if strTipoPago <> "CO" then %>
<table width="600"  border="0">
	<tr>
		<td>
			<table width="300" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
				<tr>
					<td colspan=2><span class="Estilo370"><b>Detalle Deuda</b></td>
				</tr>
				<tr>
					<td><span class="Estilo370" width="125">CAPITAL</span></td>
					<td ALIGN="RIGHT" width="125"><%=FN(intMontoCapital,0)%></td>
				</tr>
				<tr>
					<td><span class="Estilo370">INDEMNIZACION</span></td>
					<td ALIGN="RIGHT"><%=FN(intIndemnizacion,0)%></td>
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
					<td><span class="Estilo370">GASTOS DE COBRANZA</span></td>
					<td ALIGN="RIGHT"><%=FN(intHonorarios,0)%></td>
				</tr>
				<tr>
					<td><span class="Estilo370">GASTOS PROTESTOS</span></td>
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
					<td colspan = "2" ALIGN="LEFT">Gastos Administrativos</td>
					<td ALIGN="RIGHT"><%=FN(intGastosAdministrativos,0)%></td>
				</tr>
				<tr class="Estilo371">
					<td colspan = "2" ALIGN="LEFT">Subtotal Llacruz</td>
					<td ALIGN="RIGHT"><%=FN(intValorHonorarios+intGastosAdministrativos,0)%></td>
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
	<tr>
		<td colspan="2">
<%
				strSql = "SELECT * FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO = " & rsCaja("ID_PAGO")
				'''Response.write "<br>strSql=" & strSql
				set rsDetCaja=Conn.execute(strSql)
				If Not rsDetCaja.eof Then
					intNroDoc = 0
					intSumaCapital = 0
					strFacturas=""
					strCuotas=""
					strGlosaUMayor=""
					Do While not rsDetCaja.Eof

						strFacturas = strFacturas & ", " & rsDetCaja("NRO_DOC")
						intCapital = rsDetCaja("CAPITAL")
						strSql="SELECT CUENTA, FECHA_VENC, ADIC_5, C.NRO_CLIENTE_DEUDOR AS NRO_CLIENTE_DEUDOR, ISNULL(NRO_CUOTA,'ND') AS NRO_CUOTA, NOM_TIPO_DOCUMENTO FROM CUOTA C, TIPO_DOCUMENTO T WHERE C.TIPO_DOCUMENTO = T.COD_TIPO_DOCUMENTO AND C.RUT_DEUDOR = '" & strRut & "' AND COD_CLIENTE = '" & intCliente & "' AND ID_CUOTA = " & rsDetCaja("ID_CUOTA")
						''Response.write "<br>strSql=" & strSql
						''Response.End

						set RsCuota=Conn.execute(strSql)
						If not RsCuota.eof then
							strInterLoc = RsCuota("NRO_CLIENTE_DEUDOR")
							strTipoDoc = RsCuota("NOM_TIPO_DOCUMENTO")
							strCuenta = RsCuota("CUENTA")
							strFechaVenc = RsCuota("FECHA_VENC")
							strCuotas = RsCuota("NRO_CUOTA")

							strGlosaUMayor = strGlosaUMayor & strTipoDoc & " : " & rsDetCaja("NRO_DOC")& " - F.VENC: " & strFechaVenc & " - CAPITAL PAGADO :" & FN(intCapital,0) & "<br>"
						End if
						RsCuota.close
						set RsCuota=nothing
			%>

			<%
					strNombreDeudor=""
					strMostrarRut=""
					rsDetCaja.movenext
					intNroDoc = intNroDoc + 1
					Loop
				End If
				strFacturas=Mid(strFacturas,2,len(strFacturas))

				strCajaNro="Santiago"
			%>

<table width="600" border="0">
	<tr>
		<TD VALIGN = "TOP">
			<TABLE VALIGN = "TOP" width="400" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
				<tr>
					<td><span class="Estilo370"><b>Documento(s) pagado(s)</b></td>
				</tr>
				<tr>
					<td><span class="Estilo370" width="125"><%=strGlosaUMayor%></td>
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


		</td>

	</tr>
</table>

<%else %>

  <%


				strSql = "SELECT COUNT(*) as CANTC FROM CONVENIO_DET WHERE CUOTA <> 0 AND ID_CONVENIO = " & intIdConvenio
				set rsDetConv=Conn.execute(strSql)
				If Not rsDetConv.Eof Then
					intNroCuotas = rsDetConv("CANTC")
				Else
					intNroCuotas = 0
				End If

				strSql = "SELECT * FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO = " & rsCaja("ID_PAGO")
                'response.Write("asdasd" & strSql)
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
			
               strSql = "SELECT FOLIO,SEDE, TOTAL_CONVENIO, PIE FROM CONVENIO_ENC WHERE ID_CONVENIO = " & intIdConvenio
                set rsConv=Conn.execute(strSql)
			    if not rsConv.eof then
				    intFolio = rsConv("FOLIO")
                    strSucursal = rsConv("SEDE")
					intTotalConvenio = rsConv("TOTAL_CONVENIO")
					intPie = rsConv("PIE")
					
					intSaldoCliente = intTotalConvenio - intPie
			    end if
			    rsConv.close

                dim strGlosaUMayor 
                strGlosaUMayor ="<br>" 

                strSql="SELECT  C.NRO_DOC, C.FECHA_VENC, NOM_TIPO_DOCUMENTO "
                strSql= strSql & " FROM CUOTA C, TIPO_DOCUMENTO T ,CONVENIO_CUOTA CC"
                strSql= strSql & " WHERE C.COD_CLIENTE = '" & intCliente & "'  AND C.ID_CUOTA = CC.ID_CUOTA  "
                strSql= strSql & " AND C.TIPO_DOCUMENTO = T.COD_TIPO_DOCUMENTO "
                strSql= strSql & " AND C.RUT_DEUDOR = '" & strRut & "' "
                strSql= strSql & " AND CC.ID_CONVENIO ="   & intIdConvenio
			
                  set RsCuota2=Conn.execute(strSql)

            '        response.Write(strSql)
         		If not RsCuota2.eof then
                    Do While not RsCuota2.Eof

				   			strTipoDoc = RsCuota2("NOM_TIPO_DOCUMENTO")
                            strFechaVenc = RsCuota2("FECHA_VENC")
							strNroDoc = RsCuota2("NRO_DOC")
                            strGlosaUMayor = strGlosaUMayor & strTipoDoc & " : " & strNroDoc & " - F.VENC: " & strFechaVenc & " <br>"
					RsCuota2.movenext
                    Loop

                 End if
			%>

<table width="600">
	<tr height = "50">
		<td>
<table width="600"  border="0">
	<tr>
		<td>
			<table width="300" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
					<tr>
						<td colspan=2><span class="Estilo370"><b>Detalle Deuda</b></td>
					</tr>
					<tr>
						<td><span class="Estilo370" width="125">PIE</span></td>
						<td ALIGN="RIGHT" width="125"><%=FN(intPagoPie,0)%></td>
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
					<td colspan = "2" ALIGN="LEFT">Gastos Administrativos</td>
					<td ALIGN="RIGHT"><%=FN(intGastosAdministrativos,0)%></td>
				</tr>
				<tr class="Estilo371">
					<td colspan = "2" ALIGN="LEFT">Subtotal Llacruz</td>
					<td ALIGN="RIGHT"><%=FN(intValorHonorarios+intGastosAdministrativos,0)%></td>
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
        </td>
	</tr>
	<tr height = "50">
		<td>
      
     

<table width="600" border="0">
	<tr>
		<TD VALIGN = "TOP">

			<TABLE VALIGN = "TOP" width="400" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
				<tr>
					<td><span class="Estilo370"><b>Detalle del Pago</b></td>
				</tr>
				<tr>
                	<td><span class="Estilo370" width="125"><%=strPie&" "&strCuotas&" "&strDetalleCuotas  %> Nº<%=" <b>" & intFolio & "</b>"%></td>
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
    <tr>
    <td >
    <TABLE VALIGN = "TOP" width="400" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
				<tr><td>
                <b>Documentos Normalizados Convenio </b><%=strGlosaUMayor%>
                </td>
        </tr>
	</table>
    </td>
    
    </tr>
</table>


        </td>
	</tr>
</table>

<%end if %>


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