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

		strSql = "SELECT ID_PAGO, COMP_INGRESO, USR_INGRESO, TIPO_PAGO, COD_CLIENTE, NRO_BOLETA, RUT_DEUDOR , MONTO_CAPITAL, INTERES_PLAZO, GASTOS_JUDICIALES, INDEM_COMP, MONTO_EMP, GASTOS_ADMINISTRATIVOS, GASTOS_OTROS, CONVERT(VARCHAR(10),FECHA_PAGO,103) as FECHA_PAGO "
		strSql = strSql & " FROM CAJA_WEB_EMP WHERE COMP_INGRESO = " & intNroComp

		set rsCaja=Conn.execute(strSql)
		If Not rsCaja.eof Then

			intIdPago = rsCaja("ID_PAGO")
			intCliente = rsCaja("COD_CLIENTE")
			intUsrIngreso = rsCaja("USR_INGRESO")
			strTipoPago = rsCaja("TIPO_PAGO")
			strBoleta = rsCaja("NRO_BOLETA")

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
				strRutCSD = strRutCliente
				strRutCCD = Right(strRutCliente,2)
				strRutCli = strRutCSD & strRutCCD
				strRutCli = "RUT: " & strRutCli
			End If
			strNomCliente = strRutCli  & " " & strNomCliente


			strSql="SELECT TOP 1 SUCURSAL, ADIC_5 FROM CUOTA WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strRut & "'"
			strSql = strSql & " AND NRO_DOC IN (SELECT NRO_DOC FROM CAJA_WEB_EMP_DETALLE  WHERE ID_PAGO = " & intIdPago & ") AND SUCURSAL IS NOT NULL"
			set rsSuc=Conn.execute(strSql)
			if not rsSuc.eof then
				strSucursal = rsSuc("SUCURSAL")
				strInterlocutor = rsSuc("ADIC_5")
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
					strRutCli = strRutCli
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


	<table width="600" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>

	<%

		AbrirSCG1()
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
		CerrarSCG1()




		strFechaBoleta = rsCaja("FECHA_PAGO")
		strNombreCliente = strNomCliente
		strCP = rsCaja("COMP_INGRESO")
		strRUT_DEUDOR = strRut

		strRutCSD = FormatNumber(Mid(strRUT_DEUDOR,1,len(strRUT_DEUDOR)-2),0)
		strRutCCD = Right(strRUT_DEUDOR,2)
		strRUT_DEUDOR = strRutCSD & strRutCCD
		strNombreDeudor = strNombreDeudor
		strMontoBoleta = intValorHonorarios


	%>
</table>

<table width="800">
	<tr height="20">
		<td width="800" colspan="2">&nbsp;</td>
	</tr>
	<tr height="165">
		<td width="400">&nbsp;</td>

		<td width="400">
			<table>
				<tr><td class="Estilo372">&nbsp;&nbsp;&nbsp;&nbsp;FECHA&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;<%=strFechaBoleta%></td></tr>
				<tr><td class="Estilo372">&nbsp;&nbsp;&nbsp;&nbsp;RUT&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;<%=strRutCli%></td></tr>
				<tr><td class="Estilo372">&nbsp;&nbsp;&nbsp;&nbsp;NOMBRE&nbsp;&nbsp;&nbsp;:&nbsp;<%=strRSocialCli%></td></tr>
				<tr><td class="Estilo372">&nbsp;&nbsp;&nbsp;&nbsp;CP &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;<%=strCP%></td></tr>
			</table>

		</td>
	</tr>
	<tr height="45">
		<td width="800" colspan="2">&nbsp;</td>
	</tr>
	<tr>
			<td width="800" colspan=2>
				<table border=0>
					<tr height="20">
						<td width="650" class="Estilo372">RUT &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : <%=strRUT_DEUDOR%></td>
						<td width="150">&nbsp;</td></tr>
					<tr height="20">
						<td class="Estilo372">NOMBRE : <%=strNombreDeudor%></td><td>&nbsp;</td>
					</tr>
					<tr height="208">
					<td>&nbsp;</td>
					<td style="vertical-align: bottom;"  class="Estilo372">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=FN(strMontoBoleta,0)%></td>
					</tr>

					<tr>
					<td ALIGN="left">
					<!--input name="imp" type="button" onClick="window.print();" value="Imprimir Ficha"-->
					</td>
					</tr>

				</table>
			</td>
	</tr>
</table>






			<%

	end if
		end if %>
</body>
</html>