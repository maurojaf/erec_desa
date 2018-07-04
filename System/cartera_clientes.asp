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
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->

<script language="JavaScript">
function ventanaSecundaria (URL){
window.open(URL,"DETALLE","width=800, height=300, scrollbars=YES, menubar=no, location=no, resizable=yes")
}
</script>
<%

Response.CodePage=65001
Response.charset ="utf-8"
	
strRut= Request("TX_RUT")
strApellidos= Request("TX_APELLIDOS")
strNombres= Request("TX_NOMBRES")
strSolicitud = Request("TX_SOLICITUD")
strArea = Request("TX_AREA")
strTelefono = Request("TX_TELEFONO")

intCodRemesa = Request("intRemesa")

intEstadoDeuda = Request("intEstadoDeuda")

intGestion = Request("CB_GESTION")

''strCOD_CLIENTE = Request("CB_CLIENTE")
strCOD_CLIENTE=Request("strCliente")

strEjeAsig = Request("CB_EJECUTIVO")

If Trim(strCOD_CLIENTE) = "" Then strCOD_CLIENTE = "1000"

%>
<title>CARTERA ASIGNADA</title>


</head>
<%strTitulo="MI CARTERA"%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">

<table width="100%" border="1" bordercolor="#FFFFFF">
	<tr>
		<TD height="20" ALIGN=LEFT class="pasos2_i">
			<B>CARTERA ASIGNADA</B>
		</TD>
	</tr>
</table>


		<table width="900" align="CENTER">
		  <tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="100%" align="center">RESULTADO BUSQUEDA</td>
		  </tr>
		  </table>
		  <table width="900" align="CENTER">
			<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td>RUT</td>
				<td>NOMBRE O RAZON SOCIAL</td>
				<td>MONTO ORIGINAL</td>
				<td>SALDO</td>
				<td>PAG-ABON-RET</td>
				<td>ESTADO DEUDA</td>
				<td>&nbsp</td>
				<td>&nbsp</td>
			</tr>
			<%
				AbrirSCG()

				strSql = "SELECT TOP 2000 CUOTA.RUT_DEUDOR, NOMBRE_DEUDOR, ESTADO_DEUDA, SUM(VALOR_CUOTA) as VALOR_CUOTA, SUM(SALDO) as SALDO, CUOTA.COD_CLIENTE as COD_CLIENTE "
				strSql = strSql & " FROM DEUDOR , CUOTA	WHERE DEUDOR.RUT_DEUDOR = CUOTA.RUT_DEUDOR "


				If Trim(intEstadoDeuda) = "1" Then 'Saldo Activo
					strSql = strSql & " AND CUOTA.SALDO > 0 "
				End If

				If Trim(intEstadoDeuda) = "2" Then 'Retiros
					strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (2,6)"
				End If

				If Trim(intEstadoDeuda) = "10" Then 'Convenios
					strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (10,11)"
				End If

				If Trim(intEstadoDeuda) = "3" Then 'Pagos
					strSql = strSql & " AND CUOTA.ESTADO_DEUDA IN (3,4,7,8)"
				End If

				strSql = strSql & " AND CUOTA.COD_CLIENTE = '" & strCOD_CLIENTE & "' AND DEUDOR.COD_CLIENTE = CUOTA.COD_CLIENTE "

				'If trim(strEjeAsig) = "" Then
				'	strSql = strSql & " AND USUARIO_ASIG = " & session("session_idusuario")
				'Else
				'	strSql = strSql & " AND USUARIO_ASIG = " & strEjeAsig
				'End if

				strParametro = "0"

				'If Trim(strRut) <> "" Then
				'	strSql = strSql & " AND RUT = '" & strRut & "'"
				'	strParametro = "1"
				'End if

				If Trim(strNombres) <> "" Then
					strSql = strSql & " AND NOMBRE_DEUDOR  LIKE '%" & strNombres & "%'"
					strParametro = "1"
				End if

				If Trim(intCodRemesa) <> "0" and Trim(intCodRemesa) <> "" Then
					strSql = strSql & " AND CUOTA.COD_REMESA = " & intCodRemesa
					strParametro = "1"
				End if


				strSql = strSql & " GROUP BY CUOTA.RUT_DEUDOR, NOMBRE_DEUDOR, ESTADO_DEUDA, CUOTA.COD_CLIENTE "
				strSql = strSql & " ORDER BY SUM(SALDO) DESC"

				'Response.write(strSql)
				'Response.end
				set rsCuota=Conn.execute(strSql)
				intTotalSaldo = 0
				intTotalRut = 0
				intMAPR = 0
				If Not rsCuota.eof Then
					Do while not rsCuota.eof

						intValorSaldo = Round(session("valor_moneda") * ValNulo(rsCuota("SALDO"),"N"),0)
						intValorCuota = Round(session("valor_moneda") * ValNulo(rsCuota("VALOR_CUOTA"),"N"),0)


						intTotalSaldo = intTotalSaldo + intValorSaldo
						intTotalDOrig = intTotalDOrig + intValorCuota
						intMAPR = intValorCuota - intValorSaldo

						intTotalMAPR = intTotalMAPR + intMAPR
						intTotalRut = intTotalRut + 1

						%>
							<tr bgcolor="<%=strbgcolor%>" class="Estilo8">
								<td ALIGN="right"><%=rsCuota("RUT_DEUDOR")%></td>
								<td><%=Mid(rsCuota("NOMBRE_DEUDOR"),1,50)%></td>
								<td ALIGN="right"><%=FN(intValorCuota,0)%></td>
								<td ALIGN="right"><%=FN(intValorSaldo,0)%></td>
								<td ALIGN="right"><%=FN(intMAPR,0)%></td>

								<td ALIGN="center"><%=Mid(TraeCampoId(Conn, "DESCRIPCION", rsCuota("ESTADO_DEUDA"), "ESTADO_DEUDA", "CODIGO"),1,50)%></td>
								<td>
									<A HREF="principal.asp?rut=<%=rsCuota("RUT_DEUDOR")%>">
										<acronym title="Llevar a pantalla de selección">Seleccionar</acronym>
									</A>
								</td>
								<td>
									<% If TraeSiNo(session("perfil_adm")) = "Si" Then %>
									<A HREF="asigna_manual.asp?strCOD_CLIENTE=<%=rsCuota("COD_CLIENTE")%>&strRUT_DEUDOR=<%=rsCuota("RUT_DEUDOR")%>">
										<acronym title="Asigna Deudor">Asignar</acronym>
									</A>
									<% End If%>
								</td>
							</tr>
						<%
						rsCuota.movenext
					Loop
				End If
				rsCuota.close
				set rsCuota=NOTHING
				%>
				<tr>
					<td bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" colspan=2>Totales</td>
					<td bgcolor="#<%=session("COLTABBG2")%>" span class="Estilo28" align="right">$ <%=FN(intTotalDOrig,0)%></td>
					<td bgcolor="#<%=session("COLTABBG2")%>" span class="Estilo28" align="right">$ <%=FN(intTotalSaldo,0)%></td>
					<td bgcolor="#<%=session("COLTABBG2")%>" span class="Estilo28" align="right">$ <%=FN(intTotalMAPR,0)%></td>
					<td bgcolor="#<%=session("COLTABBG2")%>" span class="Estilo28" align="center" colspan=3>Total Rut : <%=intTotalRut%> </td>
				</tr>
		</table>


</form>
</body>
</html>
<script language="JavaScript1.2">

function buscar(){
	//if ((datos.TX_NOMBRES.value.length == 0) && (datos.TX_APELLIDOS.value.length==0) && (datos.TX_AREA.value.length==0) && (datos.TX_TELEFONO.value.length==0) && (datos.TX_SOLICITUD.value.length==0) && (datos.TX_RUT.value.length==0)){
	//	alert('Debe ingresar algun parametro de búsqueda');
	//	return(false);
	//}else if ((datos.TX_AREA.value.length==0) && (datos.TX_TELEFONO.value.length != 0)){
	//	alert('Si ingreso telefono debe ingresar codigo de area');
	//	return(false);
	//}else{
		datos.action='cartera_clientes.asp';
		datos.submit();
	//}
}

</script>