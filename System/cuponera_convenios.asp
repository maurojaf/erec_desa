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

	intNroConvenio=request("intNroConvenio")
	strImprime=request("strImprime")

	intCodRemesa = request("CB_REMESA")
	intCodUsuario = request("CB_COBRADOR")

	intCliente=session("ses_codcli")

	If Trim(intCliente) = "" Then intCliente = "1000"

	%>

	<style type="text/css">
		H1.SaltoDePagina {PAGE-BREAK-AFTER: always}
		.transpa {
		background-color: transparent;
		border: 1px solid #FFFFFF;
		text-align:center
		}
		<!--
		.Estilo37 {color: #FFFFFF}
		.Estilo370 {color: #000000}
		-->

	</style>


		<%

		If Trim(intCliente) <> "" then
		abrirscg()

		strSql = "SELECT RUT_DEUDOR, COD_CLIENTE "
		strSql = strSql & " FROM CONVENIO_ENC WHERE ID_CONVENIO = " & intNroConvenio

		set rsCaja=Conn.execute(strSql)
		If Not rsCaja.eof Then

			intCliente = rsCaja("COD_CLIENTE")
			strRut = rsCaja("RUT_DEUDOR")

			ssql="SELECT NOMBRE_DEUDOR FROM DEUDOR WHERE RUT_DEUDOR='" & strRut & "' AND COD_CLIENTE = '" & intCliente & "'"
			set RsDeudor=Conn.execute(ssql)
			if not RsDeudor.eof then
				strNombreDeudor = RsDeudor("NOMBRE_DEUDOR")
			end if
			RsDeudor.close
			set RsDeudor=nothing




			strSql = "SELECT COUNT(*) as CANTC FROM CONVENIO_DET WHERE CUOTA <> 0 AND ID_CONVENIO = " & intNroConvenio
			set rsDetConv=Conn.execute(strSql)
			If Not rsDetConv.Eof Then
				intNroCuotas = rsDetConv("CANTC")
			Else
				intNroCuotas = 0
			End If

			%>

</head>
<body>

<table width="720" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
	<%
	For intNroCuota = 1 to intNroCuotas


	strSql="SELECT CUOTA, FECHA_PAGO, TOTAL_CUOTA FROM CONVENIO_DET WHERE CUOTA <> 0 AND ID_CONVENIO = " & intNroConvenio & " AND CUOTA = " & intNroCuota
	set RsCuota=Conn.execute(strSql)
	If not RsCuota.eof then
		strFechaPagoCuota = RsCuota("FECHA_PAGO")
		intTotalCuota = RsCuota("TOTAL_CUOTA")
	Else
		strFechaPagoCuota = ""
		intTotalCuota = ""
	End if
	RsCuota.close
	set RsCuota=nothing

	%>
	<tr height="285">
		<td valign="TOP">
			<table width="360" border="0" height="60">
			<tr>
				<td><img src="../imagenes/reintegra.jpg" width="77" height="25"></td>
				<td><span class="Estilo370"><B>Cuponera de <%=session("NOMBRE_CONV_PAGARE")%></B></td>
				<td>copia Cliente</td>
			</tr>
			</table>
			<table width="360" border="0">
			<tr>
				<td width="260" colspan=2><span class="Estilo370"><b>Nombre del Cliente</b></td>
				<td width="100"><span class="Estilo370">Cuota <%=intNroCuota%> de <%=intNroCuotas%></td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370"><%=strNombreDeudor%></td>
				<td width="100"><span class="Estilo370">&nbsp;</td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370"><b>R.U.T.</b></td>
				<td width="100"><span class="Estilo370"><b>Fecha Venc.</b></td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370"><%=strRut%></td>
				<td width="100"><span class="Estilo370"><%=strFechaPagoCuota%></td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370"><b>Nro. <%=session("NOMBRE_CONV_PAGARE")%></b></td>
				<td width="100">&nbsp;</td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370"><%=intNroConvenio%></td>
				<td width="100">&nbsp;</td>
			</tr>
			<tr>
				<td width="260" colspan=2><b>Monto Cuota</b></td>
				<td width="100">&nbsp;</td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370">$&nbsp;<%=FN(intTotalCuota,0)%></td>
				<td width="100">&nbsp;</td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370"><b>Intereses</b></td>
				<td width="100">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370">$&nbsp;____________</td>
				<td width="100">&nbsp;</td>
			</tr>
			<tr>
				<td width="260" colspan=2><span class="Estilo370"><b>Total a pagar</b></td>
				<td width="100">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;___&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;___</td>
			</tr>
			<tr height="30">
				<td width="260" colspan=2><span class="Estilo370">$&nbsp;____________</td>
				<td width="100">Efectivo Cheque</td>
			</tr>
			</table>
		</td>
		<td valign="TOP">
			<table width="360" border="0" height="60">
				<tr>
					<td><img src="../imagenes/reintegra.jpg" width="77" height="25"></td>
					<td><span class="Estilo370"><B>Cuponera de <%=session("NOMBRE_CONV_PAGARE")%></B></td>
					<td>copia Empresa</td>
				</tr>
				</table>
				<table width="360" border="0">
				<tr>
					<td width="260" colspan=2><span class="Estilo370"><b>Nombre del Cliente</b></td>
					<td width="100"><span class="Estilo370">Cuota <%=intNroCuota%> de <%=intNroCuotas%></td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370"><%=strNombreDeudor%></td>
					<td width="100"><span class="Estilo370">&nbsp;</td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370"><b>R.U.T.</b></td>
					<td width="100"><span class="Estilo370"><b>Fecha Venc.</b></td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370"><%=strRut%></td>
					<td width="100"><span class="Estilo370"><%=strFechaPagoCuota%></td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370"><b>Nro. <%=session("NOMBRE_CONV_PAGARE")%></b></td>
					<td width="100">&nbsp;</td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370"><%=intNroConvenio%></td>
					<td width="100">&nbsp;</td>
				</tr>
				<tr>
					<td width="260" colspan=2><b>Monto Cuota</b></td>
					<td width="100">&nbsp;</td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370">$&nbsp;<%=FN(intTotalCuota,0)%></td>
					<td width="100">&nbsp;</td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370"><b>Intereses</b></td>
					<td width="100">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370">$&nbsp;____________</td>
					<td width="100">&nbsp;</td>
				</tr>
				<tr>
					<td width="260" colspan=2><span class="Estilo370"><b>Total a pagar</b></td>
					<td width="100">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;___&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;___</td>
				</tr>
				<tr height="30">
					<td width="260" colspan=2><span class="Estilo370">$&nbsp;____________</td>
					<td width="100">Efectivo   Cheque</td>
				</tr>
			</table>
		</td>


	</tr>

	<%	If intNroCuota mod 3 = 0 Then	%>
	</table>
	<H1 class=SaltoDePagina></H1>
	<table width="720" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
	<%	End If

	Next
	%>
</table>






			<%

		End If
		rsCaja.close
		set rsCaja=nothing
		''Response.End
		%>


		<%	cerrarscg()
		end if %>

</body>
</html>