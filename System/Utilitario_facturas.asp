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
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<!--#include file="../lib/lib.asp"-->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<script type="text/javascript">
$(document).ready(function(){

	$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

})
</script>
<script language="JavaScript">
function ventanaSecundaria (URL){
	window.open(URL,"DETALLE","width=200, height=200, scrollbars=no, menubar=no, location=no, resizable=no")
}
</script>

<%
Response.CodePage=65001
Response.charset ="utf-8"

inicio= request("inicio")
termino= request("termino")


strCliente = REQUEST("CB_CLIENTE")
strTipoCarga = request("CB_TIPO")
strNroFact = request("TX_NRO_FACT")
dtmFecFact = request("inicio")
strObsFact3 = request("TX_OBS_FACT3")


abrirscg()
	If Trim(inicio) = "" Then
		inicio = TraeFechaActual(Conn)
		inicio = "01/" & Mid(TraeFechaActual(Conn),4,10)
	End If

	If Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If

strGraba = request("strGraba")

  If Trim(strGraba) = "S" Then


			If Trim(strTipoCarga) = "1" Then

			AbrirSCG1()

			strSql = "UPDATE CUOTA SET FECHA_ESTADO_FACT = getdate(), USUARIO_ESTADO_FACT = " & session("session_idusuario") & ", ESTADO_FACTURA = '5', OBS_PROCESO_FACT = 'FACTURA'+' '+'" & strNroFact & "'+' '+'ANULADA' FROM CUOTA WHERE NUMERO_FACTURA = '" & strNroFact & "' AND CONVERT(VARCHAR(10),FECHA_FACTURACION,103) = CONVERT(VARCHAR(10),'" & dtmFecFact & "',103) AND ESTADO_FACTURA = '3' AND COD_CLIENTE = '" &strCliente& "'"

			abrirscg()
			set rsUpdate = Conn.execute(strSql)
			cerrarscg()

			strSql2 = "UPDATE FACTURACION_CLIENTES SET ESTADO_FACTURA = '5', USUARIO_ESTADO_FACTURA = " & session("session_idusuario") & ", OBSERVACION_FACTURA = '" & strObsFact3 & "', FECHA_ESTADO_FACTURA = getdate() FROM FACTURACION_CLIENTES WHERE NUMERO_FACTURA = '" & strNroFact & "' AND CONVERT(VARCHAR(10),FECHA_FACTURA,103) = CONVERT(VARCHAR(10),'" & dtmFecFact & "',103) AND ESTADO_FACTURA = '3' AND COD_CLIENTE = '" &strCliente& "'"

			'Response.write "<br>strSql2=" & strSql2

			abrirscg()
			set rsUpdate2 = Conn.execute(strSql2)
			cerrarscg()

			End if

	%>
	<script>
		alert('Proceso realizado correctamente');
	</script>
	<%


  End if


If Trim(intCodUsuario) = "" Then intCodUsuario = session("session_idusuario")

''Response.write "intCliente=" & intCliente

If Trim(intCliente) = "" Then intCliente = session("ses_codcli")
%>
<title>UTILITARIO</title>

<style type="text/css">
<!--
.Estilo37 {color: #FFFFFF}
-->
</style>

</head>
<body>
	<div class="titulo_informe">UTILITARIO FACTURACIÓN</div>
	<br>
<table width="90%" align="CENTER" border="0">

   <tr>
    <td valign="top">
	<BR>
	<FORM name="datos" method="post">
	<table width="100%" border= "0" class="estilo_columnas">
		<thead>
		<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td aling = "left" colspan = "2" width="39%">TIPO PROCESO</td>
			<td aling = "left" colspan = "2" width="21%">CLIENTE</td>
			<td aling = "left" colspan = "1" width="21%">NUMERO FACTURA</td>
			<td aling = "left" colspan = "1" width="25%">FECHA FACTURA</td>
		</tr>
		</thead>
		<tr>
			<td Colspan= "2">
				<select name="CB_TIPO">
					<option value="0" <%If strTipoCarga = "0" then response.write "SELECTED"%>>SELECCIONAR</option>
					<option value="1" <%If strTipoCarga = "1" then response.write "SELECTED"%>>ANULAR FACTURA</option>
				</select>
			</td>

			<td Colspan= "2">

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

			<td Colspan= "1">
				<input name="TX_NRO_FACT" type="text" value="<%=strNroFact%>" size="10" maxlength="12">
			</td>

			<td Colspan= "1">
				<input name="inicio" readonly="true" type="text" id="inicio" value="<%=dtmFecFact%>" size="8" maxlength="10">
					<!--<a href="javascript:showCal('Calendar7');"><img src="../Imagenes/calendario.gif" border="0"></a>-->
			</td>

		</tr>
	</table>


	<table width="100%" border= "0" class="estilo_columnas" >
		<thead>
		<tr >
			<td aling = "left" colspan = "1" width="20%">OBSERVACION FACTURA</td>
		</tr>
		</thead>
		<tr>
			<td Colspan= "1">
				<input name="TX_OBS_FACT3" type="text" value="<%=strObsFact3%>" size="45" maxlength="45">
			</td>
		</tr>
	</table>


<table width="100%" border="0" bordercolor="#FFFFFF" ALIGN="LEFT">

	<TR>
		<TD colspan="5" ALIGN="RIGHT">
			<INPUT TYPE="BUTTON" class="fondo_boton_100" value="Procesar" name="B1" onClick="envia('G');return false;">
		</TD>
	</TR>
</table>
</form>

</table>
</body>
</html>

<script language="JavaScript1.2">
function envia(intTipo)	{
		if ((document.forms[0].CB_TIPO.value == 'ELIMINAR') || (document.forms[0].CB_TIPO.value == 'ELIMINAR_ID')){
			if (confirm("¿ Está seguro de eliminar los documentos ingresados ? Este proceso es IRREVERSIBLE"))
			{
				if (confirm("¿ Está REALMENTE seguro de eliminar los documentos ingresados ? Este proceso es COMPLETAMENTE IRREVERSIBLE"))
				{
					if (intTipo=='G'){
								document.forms[0].action='Utilitario_facturas.asp?strGraba=S';
							}else{
								document.forms[0].action='Utilitario_facturas.asp?strRefrescar=C';
							}
					document.forms[0].submit();
				}
			}


		}
		else if (document.forms[0].CB_TIPO.value == 'ELIMINAR_DEUDOR'){
					if (confirm("¿ Está seguro de eliminar los deudores ingresados ? Este proceso es IRREVERSIBLE"))
					{
						if (confirm("¿ Está REALMENTE seguro de eliminar los deudores ingresados ? Este proceso es COMPLETAMENTE IRREVERSIBLE, y eliminará deuda y gestiones asociadas al deudor - cliente"))
						{
							if (intTipo=='G'){
										document.forms[0].action='Utilitario_facturas.asp?strGraba=S';
									}else{
										document.forms[0].action='Utilitario_facturas.asp?strRefrescar=C';
									}
							document.forms[0].submit();
						}
					}


		}
		else
		{
			if (intTipo=='G'){
						document.forms[0].action='Utilitario_facturas.asp?strGraba=S';
					}else{
						document.forms[0].action='Utilitario_facturas.asp?strRefrescar=C';
					}
			document.forms[0].submit();
		}


}

function RefrescaDatos(){
	document.forms[0].submit();
}

</script>


