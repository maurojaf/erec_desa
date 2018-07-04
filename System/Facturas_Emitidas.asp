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
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<!--#include file="../lib/lib.asp"-->

	<!--#include file="../lib/comunes/rutinas/chkFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/sondigitos.inc"-->
	<!--#include file="../lib/comunes/rutinas/formatoFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/validarFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	Usuario_session =Session("intCodUsuario")

	AbrirSCG()

	termino = request("termino")
	inicio = request("inicio")

	'resp = request("resp")
	if Trim(inicio) = "" Then
		'inicio = TraeFechaMesActual(Conn,-1)
		'inicio = "01" & Mid(inicio,3,10)
		'inicio = TraeFechaActual(Conn)

		strMesActual = Month(TraeFechaActual(Conn))
		strAnoActual = Cdbl(Year(TraeFechaActual(Conn)))

		''strDiaActual = Cdbl(day(TraeFechaActual(Conn)))




		If strMesActual = 1 Then strAnoActual = strAnoActual - 1
		If strMesActual = 1 Then strMesActual = 12
		strMesActual = strMesActual - 1

		if Len(strMesActual) = 1 Then strMesActual = "0" & strMesActual

		If Trim(inicio) = "" Then inicio = "01/" & strMesActual & "/" & strAnoActual

	End If


	if Trim(termino) = "0" Then
		termino = TraeFechaActual(Conn)
	End If

	strCliente = REQUEST("CB_CLIENTE")
	strEstado = REQUEST("CB_ESTADO")
	strTipobus = REQUEST("CB_TIPOBUS")


	'Response.write "strTipobus=" & strTipobus

	'hoy=date

	intCOD_CLIENTE = session("ses_codcli")

%>

	<title>CRM FACTORINg</title>


	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo28 {color: #FFFFFF}
	.Estilo27 {color: #FFFFFF}
	-->
	</style>

	
	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>
	<script src="../javascripts/SelCombox.js"></script>
	<script src="../javascripts/OpenWindow.js"></script>


	<script language="JavaScript " type="text/JavaScript">

	function Refrescar()
	{
		resp='no'
		datos.action = "Facturas_Emitidas.asp?resp="+ resp +"";
		datos.submit();
	}

	function enviar(){
				datos.action = "man_Export.asp?archivo=1&CB_CLIENTE=" + document.datos.CB_CLIENTE.value + "&CB_TIPOPROCESO=" + document.datos.CB_TIPOPROCESO.value + "&CB_ASIGNACION=" + document.datos.CB_ASIGNACION.value + "&CH_ACTIVO=" + document.datos.CH_ACTIVO.checked;
				datos.submit()
	}

	function Ingresa()
	{
		with( document.datos )
		{
			action = "Facturas_Emitidas.asp";
			submit();
		}
	}

	function Reversar(cod_pago)
	{
		with( document.datos )
		{
			//alert("Opción deshabilitada");


		if (confirm("¿ Está seguro de reversar el pago ? El pago se eliminará completamente y la deuda será reversada, volviendo a su estado original antes del pago."))
			{
				action = "reversar_pago.asp?cod_pago=" + cod_pago;
				submit();
			}
		else
			alert("Reverso del pago cancelado");
		}
	}

	function Modificar(cod_pago)
	{
		with( document.datos )
		{
			action = "modif_caja_web2.asp?strOrigen=Facturas_Emitidas.asp&cod_pago=" + cod_pago;
			submit();
		}
	}

	function envia()
	{
		//datos.TX_RUT.value='';
		//datos.TX_PAGO.value='';
		resp='si'
		document.datos.action = "Facturas_Emitidas.asp?strBuscar=S&resp="+ resp +"";
		document.datos.submit();
	}

	function imprimir()
	{
		datos.action = "imprime_comprobantes.asp";
		datos.submit();
	}


	function envia_excel(URL){
		window.open(URL,"INFORMACION","width=200, height=200, scrollbars=yes, menubar=yes, location=yes, resizable=yes")
	}
	</script>


</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<form name="datos" method="post">
<div class="titulo_informe">LISTADO DOCUMENTOS FACTURADOS</div>	
<br>
<table width="90%" height="500" border="0" align="center">
  <tr height="20">
    <td style="vertical-align: top;">
		<table width="100%" border="0" class="estilo_columnas">
			<thead>
			  <tr height="20">
				<td>CLIENTE</td>
				<td>TIPO BUSQUEDA</td>
				<td>FECHA/NRO FACTURA</td>
				<td></td>

				<td align="CENTER">EXPORTAR</td>

			  </tr>
			</thead>
			  <tr>

				<td>

				<SELECT NAME="CB_CLIENTE" id="CB_CLIENTE">

					<option value="0">SELECCIONAR</option>
					<%

						ssql="SELECT COD_CLIENTE,RAZON_SOCIAL, NOMBRE_FANTASIA FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
								<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCliente)=Trim(rsTemp("COD_CLIENTE")) then response.Write("Selected") End If%>><%=rsTemp("NOMBRE_FANTASIA")%></option>
									<%
								rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing

					%>
				</SELECT>


				</td>

				<td>
					<SELECT NAME="CB_ESTADO" id="CB_ESTADO" onChange="CargaFechas(this.value,CB_CLIENTE.value);">
						<option value="0" <%If Trim(strEstado)="0" Then Response.write "SELECTED"%>>SELECCIONAR</option>
						<option value="1" <%If Trim(strEstado)="1" Then Response.write "SELECTED"%>>BUSQUEDA FECHA</option>
						<option value="2" <%If Trim(strEstado)="2" Then Response.write "SELECTED"%>>BUSQUEDA FACTURA</option>
						<!--<option value="3" <%If Trim(strEstado)="3" Then Response.write "SELECTED"%>>DOCUMENTOS CON NC</option>-->
					</SELECT>
				</td>

				<td>
					<SELECT NAME="CB_TIPOBUS" id="CB_TIPOBUS">
					</SELECT>
				</td>

				<td align = "CENTER" >
				<input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();">
				</td>

				<td align="CENTER">
					<input Name="SubmitButton" class="fondo_boton_100" Value="Exportar" Type="BUTTON" onClick="enviar();">
				</td>
			  </tr>
		</table>
    </td>
   </tr>


   <tr>
	<td style="vertical-align: top;">
	<table width="100%" border="0" class="intercalado" style="width:100%;">
		<thead>
		<tr>

		<%If strEstado = 0 Then%>
			<td colspan = "5" ALIGN="CENTER" >BUSQUEDA POR FECHA O NUMERO FACTURA</td>

		<%ElseIf strEstado = 1 OR strEstado = 2 Then%>

			<td>&nbsp;</td>
			<td>NUMERO.</td>
			<td>FECHA FACT</td>
			<td>MONTO VISACION</td>
			<td>MONTO FACTURADO</td>
			<td>DIFERENCIA</td>
			<td>DOC. FACTURADOS</td>
			<td>OBSERVACION</td>
			<td>EMISOR</td>
			<td>FECHA EMISION</td>

		<%End If%>

		</tr>
		</thead>
		<tbody>
	<%


	resp="si"
	If resp="si" then

	strSql = "SELECT 	NUMERO_FACTURA,FECHA_FACTURA,MONTO_TOTAL_FACTURA,TDOC_FACT,OBSERVACION_FACTURA,MONTO_ENVIADO_VISAR, USUARIO.LOGIN as EMISOR, FECHA_EMISION,(MONTO_TOTAL_FACTURA-MONTO_ENVIADO_VISAR) AS DIFERENCIA"
	strSql = strSql & "	FROM FACTURACION_CLIENTES   INNER JOIN CLIENTE ON FACTURACION_CLIENTES.COD_CLIENTE = CLIENTE.COD_CLIENTE"
	strSql = strSql & "								LEFT JOIN USUARIO ON FACTURACION_CLIENTES.USUARIO_ESTADO_FACTura = USUARIO.ID_USUARIO"


	strSql = strSql & " WHERE 		ESTADO_FACTURA = '3' AND FACTURACION_CLIENTES.COD_CLIENTE =  '" & strCliente & "'"


	if Trim(strEstado) = "1" Then

	strSql = strSql & " AND CAST( '" & strTipobus & "' AS DATETIME) = FECHA_FACTURA"

	End If

	if Trim(strEstado) = "2"  Then

	strSql = strSql & " AND NUMERO_FACTURA =  '" & strTipobus & "'"

	End If


	End if

	if strSql <> "" then
		set rsDet=Conn.execute(strSql)


		if not rsDet.eof then
			intReg = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>
				<tr >

		<%If strEstado = 1 or strEstado = 2 Then%>

			<td><%=intReg%></td>
			<td><%=Mid(rsDet("NUMERO_FACTURA"),1,30)%></td>
			<td><%=Mid(rsDet("FECHA_FACTURA"),1,18)%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("MONTO_ENVIADO_VISAR"),0)%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("MONTO_TOTAL_FACTURA"),0)%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("DIFERENCIA"),0)%></td>
			<td ALIGN="RIGHT"><%=Mid(rsDet("TDOC_FACT"),1,15)%></td>
			<td><%=Mid(rsDet("OBSERVACION_FACTURA"),1,25)%></td>
			<td ALIGN="RIGHT"><%=rsDet("EMISOR")%></td>
			<td><%=Mid(rsDet("FECHA_EMISION"),1,18)%></td>

		<%ElseIf strEstado = 2 Then%>

			<td><%=intReg%></td>
			<td><%=Mid(rsDet("DESCCLI"),1,30)%></td>
			<td><%=Mid(rsDet("ID_CUOTA"),1,15)%></td>
			<td><%=Mid(rsDet("RUT_DEUDOR"),1,15)%></td>
			<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("VALOR_CUOTA"),0)%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("MONTO_VISACION"),0)%></td>
			<td><%=Mid(rsDet("FECHA_ENVIO_VISAR"),1,28)%></td>
			<td><%=Mid(rsDet("estado_factura"),1,28)%></td>
			<td><%=Mid(rsDet("USUARIO"),1,30)%></td>

		<%ElseIf strEstado = 3 Then%>

			<td><%=intReg%></td>
			<td><%=Mid(rsDet("DESCCLI"),1,30)%></td>
			<td><%=Mid(rsDet("ID_CUOTA"),1,15)%></td>
			<td><%=Mid(rsDet("RUT_DEUDOR"),1,15)%></td>
			<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("VALOR_CUOTA"),0)%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("MONTO_FACTURACION"),0)%></td>
			<td><%=Mid(rsDet("FECHA_ENVIO_FACTURAR"),1,28)%></td>
			<td><%=Mid(rsDet("estado_factura"),1,28)%></td>
			<td><%=Mid(rsDet("USUARIO"),1,30)%></td>

		<%End If%>

				</tr>
				<%
				rsDet.movenext
			loop
		end if
	end if

	If resp="si" and (strEstado = 1 OR strEstado = 2) then

	strSql = "SELECT  SUM(MONTO_TOTAL_FACTURA) as tot1,SUM(MONTO_ENVIADO_VISAR) as tot2, SUM(TDOC_FACT) AS con1"
	strSql = strSql & "	FROM FACTURACION_CLIENTES "
	strSql = strSql & " GROUP BY COD_CLIENTE,ESTADO_FACTURA"

	strSql = strSql & " HAVING 		COD_CLIENTE = '" & strCliente & "' AND ESTADO_FACTURA = '3'"


	''Response.write "strSql = " & strSql


		if strSql <> "" then
		set rsTot1=Conn.execute(strSql)

		end if

	End if

	%>
		</tbody>
		<thead class="totales">
			<tr >

		<%If strEstado = 1 or strEstado = 2 Then%>

			<td colspan="3">TOTALES</td>
			<td colspan="1" ALIGN="RIGHT">$ <%=FN(rsTot1("tot2"),0)%></td>
			<td colspan="1" ALIGN="RIGHT">$ <%=FN(rsTot1("tot1"),0)%></td>
			<td colspan="1" ALIGN="RIGHT"></td>
			<td colspan="1" ALIGN="RIGHT"><%=FN(rsTot1("CON1"),0)%></td>
			<td colspan="3" ALIGN="RIGHT"></td>


		<%ElseIf strEstado = 3 Then%>

			<td colspan="2" >TOTALES</td>
			<td colspan="1" ALIGN="RIGHT">TOTAL DOCUMENTOS</td>
			<td colspan="1" ALIGN="RIGHT"><%=FN(rsTot1("con1"),0)%></td>
			<td colspan="3" ALIGN="RIGHT">MONTO FACTURADO</td>
			<td colspan="1" ALIGN="RIGHT">$ <%=FN(rsTot1("Tot1"),0)%></td>
			<td colspan="2" ALIGN="RIGHT"></td>


		<%End If%>

				</tr>
		</thead>
	</table>
	</td>
   </tr>
  </table>

</form>


</body>
</html>

<div id="dhtmltooltip"></div>

<script type="text/javascript">

function CargaFechas(subCat,cat)
{
	var comboBox = document.getElementById('CB_TIPOBUS');
	switch (cat)
	{
		<%
		  AbrirSCG()
			strSql="SELECT COD_CLIENTE FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE"
			set rsGestCat=Conn.execute(strSql)
			Do While not rsGestCat.eof
		%>
		case '<%=rsGestCat("COD_CLIENTE")%>':


			comboBox.options.length = 0;
				if (subCat=='1') {
					var newOption = new Option('SELECCIONE', '01/01/1900');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql = "SELECT DISTINCT FECHA_FACTURACION "
					strSql = strSql & " FROM CUOTA WHERE ESTADO_FACTURA IS NOT NULL"
					strSql = strSql & " AND COD_CLIENTE = '" & rsGestCat("COD_CLIENTE") & "' AND ESTADO_FACTURA = '3'"
					strSql = strSql & " ORDER BY FECHA_FACTURACION DESC"
					'Response.write "<br>strSql=" & strSql
					set rsGestion=Conn2.execute(strSql)
					If Not rsGestion.Eof Then
						Do While Not rsGestion.Eof
							%>
								var newOption = new Option('<%=rsGestion("FECHA_FACTURACION")%>', '<%=rsGestion("FECHA_FACTURACION")%>');comboBox.options[comboBox.options.length] = newOption;
							<%
							rsGestion.movenext
						Loop
					Else
					%>
						var newOption = new Option('SIN FECHA FACTURACION', '0');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					CerrarSCG2()
					%>
					break;
				}
				if (subCat=='2') {
					var newOption = new Option('SELECCIONE', '01/01/1900');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql = "SELECT DISTINCT NUMERO_FACTURA"
					strSql = strSql & " FROM CUOTA"
					strSql = strSql & " WHERE COD_CLIENTE = '" & rsGestCat("COD_CLIENTE") & "' AND ESTADO_FACTURA IS NOT NULL AND ESTADO_FACTURA = '3'"
					'Response.write "<br>strSql=" & strSql
					set rsGestion=Conn2.execute(strSql)
					If Not rsGestion.Eof Then
						Do While Not rsGestion.Eof
							%>
								var newOption = new Option('<%=rsGestion("NUMERO_FACTURA")%>', '<%=rsGestion("NUMERO_FACTURA")%>');comboBox.options[comboBox.options.length] = newOption;
							<%
							rsGestion.movenext
						Loop
					Else
					%>
						var newOption = new Option('SIN FECHA FACTURACION', '0');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					CerrarSCG2()
					%>
					break;
				}

		<%
		  	rsGestCat.movenext
		  	Loop
		  	rsGestCat.close
		  	set rsGestCat=nothing
			CerrarSCG()
		%>
	}
}


</script>



<script type="text/javascript">
	CargaFechas(document.datos.CB_ESTADO.value,document.datos.CB_CLIENTE.value);
	datos.CB_TIPOBUS.value='<%=strTipobus%>';
</script>

