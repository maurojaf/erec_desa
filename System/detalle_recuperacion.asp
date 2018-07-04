<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/comunes/rutinas/funcionesBD.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
<%

Response.CodePage=65001
Response.charset ="utf-8"

hdd_cod_cliente = request("cmb_cliente")
txt_FechaIni = request("txt_FechaIni")
txt_FechaFin = request("txt_FechaFin")

strCobranza = request("strCobranza")
intCliente = request("intCliente")
intOrigen = request("intOrigen")
intCodRemesa = request("intCodRemesa")
intCodUsuario = request("intCodUsuario")
intFecha = request("intFecha")
intFechaIni = request("intFechaIni")
intFechaFin = request("intFechaFin")

If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

	sinCbUsario="0"

End If

abrirscg()

	strSql = "SELECT ID_CUOTA, NRO_CUOTA, DATENAME(WEEKDAY,FECHA_ESTADO) AS DIA, (CUOTA.RUT_DEUDOR) AS RUT_DEUDOR, ISNULL(CUOTA.NRO_CUOTA,0) AS NRO_CUOTA, ESTADO_DEUDA.DESCRIPCION, IsNull(FECHA_VENC,'01/01/1900') AS FECHA_VENC, IsNull(datediff(d,FECHA_VENC,FECHA_ESTADO),0) AS ANTIGUEDAD,"
	strSql = strSql & " NRO_DOC AS NUMDOC,IsNull(VALOR_CUOTA,0) AS MONTO,IsNull(SALDO,0) AS SALDO,IsNull(CUOTA.USUARIO_ASIG,0) AS USUARIO_ASIG, NRO_CUOTA, IsNull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS,"
	strSql = strSql & " SUCURSAL , ESTADO_DEUDA, COD_REMESA, CUENTA, NRO_DOC, TIPO_DOCUMENTO, CONVERT(VARCHAR(10),FECHA_ESTADO,103) AS FECHA_ESTADO, IsNull((CUOTA.ADIC_1), 0) AS ADIC_1, IsNull((CUOTA.ADIC_2), 0) AS ADIC_2, IsNull((CUOTA.ADIC_3), 0) AS ADIC_3,"
	strSql = strSql & " IsNull(NRO_CLIENTE_DEUDOR, 0) AS NRO_CLIENTE_DEUDOR, ISNULL(NRO_CLIENTE_DOC,'&nbsp;') as NRO_CLIENTE_DOC,(VALOR_CUOTA) AS VALOR_CAPITAL, (VALOR_CUOTA - SALDO)AS MONTO_PAGADO, (DEUDOR.NOMBRE_DEUDOR) AS NOMBRE_DEUDOR, "
	strSql = strSql & " IsNull(CUOTA.CUSTODIO,'LLACRUZ') AS CUSTODIO, NOM_TIPO_DOCUMENTO, (CUOTA.USUARIO_ASIG) AS USUARIO_ASIG "

	strSql = strSql & " FROM CUOTA	   INNER JOIN DEUDOR ON DEUDOR.COD_CLIENTE = CUOTA.COD_CLIENTE AND CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR"
	strSql = strSql & " 			   INNER JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
	strSql = strSql & " 			   INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"

	strSql = strSql & " WHERE CUOTA.COD_CLIENTE = '" & intCliente & "'"

	If Trim(intCodRemesa) <> "" Then
		strSql = strSql & " AND COD_REMESA = " & intCodRemesa
	End If
	strSql = strSql & " AND SALDO = 0 "

	If Trim(intFecha) <> "" Then
		strSql = strSql & " AND CONVERT(VARCHAR(10),FECHA_ESTADO,103) =  CAST('"&intFecha & "'  AS DATETIME) "
	Else
		strSql = strSql & " AND FECHA_ESTADO >= CAST('"&intFechaIni&"' AS DATETIME) AND FECHA_ESTADO <=  CAST('"& intFechaFin&"' AS DATETIME) "
	End If

	strSql = strSql & " AND ESTADO_DEUDA IN "

	If Trim(intOrigen) = "T" Then
		strSql = strSql & " (3,4,7,8,10,11)"
	End if
	If Trim(intOrigen) = "E" Then
		strSql = strSql & " (4,8,10,11)"
	End if
	If Trim(intOrigen) = "C" Then
		strSql = strSql & " (3,7)"
	End if

	If sinCbUsario = "" Then
	 strSql = strSql & " AND CUOTA.USUARIO_ASIG = " & session("session_idusuario")
	End If
	
	If Trim(intCodUsuario) <> "" Then
		strSql = strSql & " AND CUOTA.USUARIO_ASIG = " & intCodUsuario
	End If

	'Response.write strSql
	'Response.End

	set rsDetPago= Conn.execute(strSql)
	rsDetPago.close
	set rsDetPago=nothing


SET rsDetPago=Conn.execute(strSql)
%>

	<title>Detalle Pagos</title>
	<style type="text/css">
	<!--
	.Estilo37 {color: #FFFFFF}
	-->
	</style>

</head>	
<body>
<form name="Free" method="post">
<div class="titulo_informe">DETALLE DE RECUPERACIÓN</div>
<BR>
<center>

<table width="90%" border="0" cellspacing="0" cellpadding="0" class="Estilo13">
<tr  class="Estilo20">
<td width="100%" align="RIGHT"><input type="button" value="Volver" class="fondo_boton_100" onclick="javascript:history.back();"></td>

</tr>
</table>

		<table width="100%" class="intercalado">
		<thead>
		<tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

					<td align="center">&nbsp;</td>
					<td align="center">RUT DEUDOR</td>
					<td align="center">INTER.</TD>
					<td align="center">NOM. DEUDOR</td>
					<td align="center">TIPO DOC.</td>
					<td align="center">NRO. DOC.</td>
					<td align="center">CUOTA</td>
					<td align="center">COD. SAP</td>
					<td align="center">MONTO CAPITAL</td>
					<td align="center">MONTO PAGADO</td>
					<td align="center">F.VENCIM.</td>
					<td align="center">ESTADO</td>
					<td align="center">F.ESTADO</td>
		</tr>
		</thead>
		<tbody>
		<%

		intCont = 0
		Do while not rsDetPago.eof

			intCont = intCont + 1

			If Trim(rsDetPago("USUARIO_ASIG")) <> "" Then
				strEjecutivo = TraeCampoId(Conn, "NOMBRES_USUARIO", Trim(rsDetPago("USUARIO_ASIG")), "USUARIO", "ID_USUARIO") & "-" & TraeCampoId(Conn, "APELLIDO_PATERNO", Trim(rsDetPago("USUARIO_ASIG")), "USUARIO", "ID_USUARIO")
			End if

			strOrigen = rsDetPago("ESTADO_DEUDA")
			'strEjecutivo = rsDetPago("USUARIO_ASIG")

			If Trim(rsDetPago("ESTADO_DEUDA")) <> "" Then
				strOrigen = TraeCampoId(Conn, "DESCRIPCION", Trim(rsDetPago("ESTADO_DEUDA")), "ESTADO_DEUDA", "CODIGO")
			End if


			If Trim(rsDetPago("RUT_DEUDOR")) <> "" Then
				strNombreDeudor = TraeCampoId2(Conn, "NOMBRE_DEUDOR", Trim(rsDetPago("RUT_DEUDOR")), "DEUDOR", "RUT_DEUDOR")
			End if

			'RESPONSE.write "USUARIO_ASIG=" & Trim(rsDetPago("USUARIO_ASIG"))
			'RESPONSE.END

		%>

			<tr>
				<td><div align="right"><%=intCont%></div></td>
				<td align="center">
					<A HREF="principal.asp?TX_RUT=<%=rsDetPago("RUT_DEUDOR")%>">
						<acronym title="Llevar a pantalla principal"><%=rsDetPago("RUT_DEUDOR")%></acronym>
					</A>
				</td>
				<td><div align="right"><%=rsDetPago("NRO_CLIENTE_DEUDOR")%></div></td>
				<td ALIGN="left" title="<%=rsDetPago("NOMBRE_DEUDOR")%>">
				<%=Mid(rsDetPago("NOMBRE_DEUDOR"),1,15)%>
				<td><div align="right"><%=rsDetPago("NOM_TIPO_DOCUMENTO")%></div></td>
				<td><div align="right"><%=rsDetPago("NRO_DOC")%></div></td>
				<td><div align="right"><%=rsDetPago("NRO_CUOTA")%></div></td>
				<td><div align="right"><%=rsDetPago("NRO_CLIENTE_DOC")%></div></td>
				<td><div align="right"><%=rsDetPago("VALOR_CAPITAL")%></div></td>
				<td><div align="right"><%=rsDetPago("MONTO_PAGADO")%></div></td>
				<td><div align="right"><%=rsDetPago("FECHA_VENC")%></div></td>

				<td ALIGN="left" title="<%=rsDetPago("DESCRIPCION")%>">
				<%=Mid(rsDetPago("DESCRIPCION"),1,10)%>

				<td><div align="right"><%=rsDetPago("FECHA_ESTADO")%></div></td>
			</tr>

<%			rsDetPago.movenext
			Loop
			rsDetPago.close
		set rsDetPago=nothing
		cerrarscg()%>
		</tbody>
		</table>

		<input type="Hidden" name="cmb_cliente">
</form>
</body>
</html>


<script type="text/javascript">
	$(document).ready(function(){
		$(document).tooltip();
	})
</script>