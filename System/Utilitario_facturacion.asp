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

	<script language="JavaScript">
		function ventanaSecundaria (URL){
		window.open(URL,"DETALLE","width=200, height=200, scrollbars=no, menubar=no, location=no, resizable=no")
	}
	</script>

<%

Response.CodePage=65001
Response.charset ="utf-8"
	
intCliente 		= request("CB_CLIENTE")
strTipoProceso 	= request("CB_TIPO")
strNroFact 		= request("TX_NRO_FACT")
dtmFecFact 		= request("TX_FECHA_FACT")
strObsFact3 	= request("TX_OBS_FACT3")
strGraba 		= request("strGraba")


intIdCuota 		= Replace(Trim(Request("TX_ID_CUOTA")),chr(10),"")
strRut 			= Trim(Request("TX_RUT"))
strNroDoc 		= Trim(Request("TX_NRO_DOC"))
intMonto 		= Trim(Request("TX_MONTO"))
strObsFact 		= Trim(Request("TX_OBSFACT"))

idUsuario 		= session("session_idusuario")
strCodCliente 	= session("ses_codcli")

abrirscg()

	strFecha = TraeFechaHoraActual(Conn)

	strSql = "SELECT ID_FACTURA"
	strSql= strSql & " FROM FACTURACION_CLIENTES"
	strSql= strSql & " WHERE NUMERO_FACTURA = '" & strNroFact & "' AND ESTADO_FACTURA = 3"

	set rsFactDuplicada=Conn.execute(strSql)

	If Not rsFactDuplicada.Eof and strTipoProceso = "INGRESAR FACTURA" Then

		%>
		<SCRIPT>
			alert('Esta ingresando un número de factura que se encuentra ingresada anteriormente y vigente (no anulada), favor ingrese otro número o anule la factura anterior.')
		</SCRIPT>
		<%

CerrarSCG()

	ElseIf Trim(strGraba) = "S" Then

		vID_CUOTA = split(intIdCuota,CHR(13))
		vRut = split(strRut,CHR(13))
		vNroDoc = split(strNroDoc,CHR(13))
		vMonto = split(intMonto,CHR(13))
		vObsFact = split(strObsFact,CHR(13))

		'Response.write "<br>ASC = " & ASC(MID(strRut,11,1))

		intTamvID_CUOTA=ubound(vID_CUOTA)
		intTamvRut=ubound(vRut)
		intTamvNroDoc=ubound(vNroDoc)
		intTamvMonto=ubound(vMonto)
		intTamvObsFact=ubound(vObsFact)

		'Response.write "<br>intTamvID_CUOTA = " & intTamvID_CUOTA
		'Response.write "<br>intTamvValor = " & intTamvValor
		'Response.End

AbrirSCG()

			strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_PROCESO_FACTURACION]') AND type in (N'U'))"
			strSql = strSql & " DROP TABLE [TMP_PROCESO_FACTURACION]"
			Conn.Execute strSql,64

			strSql = " CREATE TABLE TMP_PROCESO_FACTURACION (ID_CUOTA INT NOT NULL,"
			strSql = strSql &"  RUT_DEUDOR VARCHAR(20) NOT NULL,"
			strSql = strSql &"  NRO_DOC VARCHAR(20) NOT NULL,"
			strSql = strSql &"  MONTO VARCHAR(20) NULL,"
			strSql = strSql &"  ID_USUARIO INT NOT NULL,"
			strSql = strSql &"  TIPO_PROCESO VARCHAR(50) NOT NULL)"

			Conn.Execute strSql,64


			For indice = 0 to intTamvID_CUOTA

				'Response.write "<br>intTamvID_CUOTA=" & intTamvID_CUOTA
				'Response.write "<br>intTamvRut=" & intTamvRut
				''Response.write "<br>intTamvNroDoc=" & intTamvNroDoc

				If intTamvID_CUOTA = intTamvRut and intTamvID_CUOTA = intTamvNroDoc Then

					If Trim(Replace(vID_CUOTA(indice), chr(10),"")) <> "" and  Trim(Replace(vRut(indice), chr(10),"")) <> "" and  Trim(Replace(vNroDoc(indice), chr(10),"")) <> "" Then

						if intTamvID_CUOTA <> -1 Then
							intId_cuota = Trim(Replace(vID_CUOTA(indice), chr(10),""))
							intId_cuota = Trim(Replace(intId_cuota, chr(160),""))
						End If

						if intTamvRut <> -1 Then
							strRutDD = Trim(Replace(vRut(indice), chr(10),""))
							strRutDD = Trim(Replace(strRutDD, chr(160),""))
							'Response.write "sss1=" & ASC(mid(strRutDD,11,1))
							'Response.write "sss2=" & mid(strRutDD,11,1)
							strRutDD = Trim(strRutDD)
							'Response.write "sss3=" & LEN(strRutDD)
						End If

						if intTamvNroDoc <> -1 Then
							strNroDocumento = ucase(Trim(Replace(vNroDoc(indice), chr(10),"")))
						End If
						if intTamvMonto <> -1 Then
							intMonto2 = Trim(Replace(vMonto(indice), chr(10),""))
						End If
						if intTamvObsFact <> -1 Then
							strObsFact2 = ucase(Trim(Replace(vObsFact(indice), chr(10),"")))
						End If

						If Trim(IntId_Cuota) <> "" Then
							strSql = "INSERT INTO TMP_PROCESO_FACTURACION (ID_CUOTA,RUT_DEUDOR,NRO_DOC,MONTO,ID_USUARIO,TIPO_PROCESO) VALUES ( " & IntId_Cuota & ",'" & strRutDD & "','" & strNroDocumento & "','" & intMonto2 & "'," & idUsuario & ",'" & strTipoProceso & "')"
							''Response.write strSql
							set rsInsert = Conn.execute(strSql)
						End If
					End If
				Else
					%>
					<script>
							alert('Columnas no contienen la misma cantidad de elementos');
							history.back();
					</script>
					<%
					Validacion="0"
					Response.End
				End If

			Next

CerrarSCG()



AbrirSCG()

			strSql = "SELECT TPF.ID_CUOTA,ISNULL(C.ID_CUOTA,0) AS ESTADO_CUOTA, ROW_NUMBER() OVER(PARTITION BY TPF.ID_CUOTA ORDER BY TPF.ID_CUOTA) AS DUPLIC,"
			strSql = strSql & " (CASE WHEN (TPF.RUT_DEUDOR <> C.RUT_DEUDOR OR TPF.NRO_DOC <> C.NRO_DOC) THEN 1 ELSE 0 END) CONCIDENCIA"
			strSql = strSql & " FROM TMP_PROCESO_FACTURACION TPF LEFT JOIN CUOTA C ON TPF.ID_CUOTA = C.ID_CUOTA"

			set rsInf=Conn.execute(strSql)
			if not rsInf.eof then
				do until rsInf.eof

				intDuplicados = rsInf("DUPLIC")
				intIdNoConincidente = rsInf("ESTADO_CUOTA")
				intCamposNC = rsInf("CONCIDENCIA")

				if intDuplicados > "1" then

					idCuotaDuplicados = rsInf("ID_CUOTA")
					intTotalDuplicados = CStr(intTotalDuplicados) + "," + CStr(idCuotaDuplicados)

				End If

				if intIdNoConincidente = "0" then

					idCuotaNC= rsInf("ID_CUOTA")
					intTotaCuotalNC = CStr(intTotaCuotalNC) + "," + CStr(idCuotaNC)

				End If

				if intCamposNC = "1" then

					idCuotaCamposNC= rsInf("ID_CUOTA")
					intTotaCamposNC = CStr(intTotaCamposNC) + "," + CStr(idCuotaCamposNC)

				End If

				rsInf.movenext
				loop
			end if
			rsInf.close
			set rsInf=nothing

				'Response.write "<br>intTotalDuplicados = " & intTotalDuplicados
				'Response.write "<br>intTotaCuotalNC = " & intTotaCuotalNC
				'Response.write "<br>intTotaCamposNC = " & intTotaCamposNC

				strSql = "SELECT * FROM TMP_PROCESO_FACTURACION"

				set rsUsu=Conn.execute(strSql)
				if not rsUsu.eof then
					do until rsUsu.eof

					strTipoProceso2 = rsUsu("TIPO_PROCESO")
					IntId_Cuota2 = rsUsu("ID_CUOTA")
					strRutDD2 = rsUsu("RUT_DEUDOR")
					strNroDocumento2 = rsUsu("NRO_DOC")
					intMonto2 = rsUsu("MONTO")
					idUsuario2 = rsUsu("ID_USUARIO")


					If strTipoProceso2 = "ENVIO A VISAR" Then
						strSql = "UPDATE CUOTA SET FECHA_ESTADO_FACT = '" & strFecha & "', FECHA_ENVIO_VISAR = '" & strFecha & "', USUARIO_ESTADO_FACT = " & idUsuario2 & ", MONTO_VISACION = '" & intMonto2 & "', ESTADO_FACTURA = '1', OBS_PROCESO_FACT = 'ENVIADO A VISAR' WHERE ID_CUOTA = '" & IntId_Cuota2 & "' AND NRO_DOC = '" & strNroDocumento2 & "' AND RUT_DEUDOR = '" & strRutDD2 &"' AND ISNULL(ESTADO_FACTURA,'100') NOT IN ('1','2','3','6')"

					ElseIf strTipoProceso2 = "ENVIO A FACTURAR" Then
						strSql = "UPDATE CUOTA SET FECHA_ESTADO_FACT = '" & strFecha & "', FECHA_ENVIO_FACTURAR = '" & strFecha & "', USUARIO_ESTADO_FACT = " & idUsuario2 & ", MONTO_FACTURACION = '" & intMonto2 & "', OBSERVACION_FACTURACION = '" & strObsFact2 & "' , ESTADO_FACTURA = '2', OBS_PROCESO_FACT = 'ENVIADO A FACTURAR' WHERE ID_CUOTA = '" & IntId_Cuota2 & "' AND NRO_DOC = '" & strNroDocumento2 & "' AND RUT_DEUDOR = '" & strRutDD2 &"' AND ESTADO_FACTURA = '1'"

					ElseIf strTipoProceso2 = "INGRESAR FACTURA" Then
						strSql = "UPDATE CUOTA SET FECHA_ESTADO_FACT = '" & strFecha & "', FECHA_FACTURACION = CAST('" & dtmFecFact & "'+' '+ REVERSE(SUBSTRING(REVERSE(convert(varchar(10),CAST('" & strFecha & "' AS DATETIME),108)),4,10)) AS DATETIME), NUMERO_FACTURA = '" & strNroFact & "', USUARIO_ESTADO_FACT = " & idUsuario2 & ", ESTADO_FACTURA = '3', OBS_PROCESO_FACT = 'FACTURADO' WHERE ID_CUOTA = '" & IntId_Cuota2 & "' AND NRO_DOC = '" & strNroDocumento2 & "' AND RUT_DEUDOR = '" & strRutDD2 &"' AND MONTO_FACTURACION = '" & intMonto2 & "' AND ESTADO_FACTURA = '2'"

					ElseIf strTipoProceso2 = "ANULAR PROCESO" Then
						strSql = "UPDATE CUOTA SET FECHA_ESTADO_FACT = '" & strFecha & "', USUARIO_ESTADO_FACT = " & idUsuario2 & ", ESTADO_FACTURA = '4',OBS_PROCESO_FACT = ESTADO_FACTURACION.NOM_ESTADO + ' ANULADO' FROM CUOTA INNER JOIN dbo.ESTADO_FACTURACION ON dbo.ESTADO_FACTURACION.CODIGO = CUOTA.ESTADO_FACTURA WHERE ID_CUOTA = '" & IntId_Cuota2 & "' AND NRO_DOC = '" & strNroDocumento2 & "' AND RUT_DEUDOR = '" & strRutDD2 &"' AND ESTADO_FACTURA IN ('1','2')"

					ElseIf strTipoProceso2 = "NO FACTURABLE" Then
						strSql = "UPDATE CUOTA SET FECHA_ESTADO_FACT = '" & strFecha & "', USUARIO_ESTADO_FACT = " & idUsuario2 & ", OBSERVACION_FACTURACION = '" & strObsFact2 & "' ,ESTADO_FACTURA = '6', OBS_PROCESO_FACT = 'NO FACTURABLE' WHERE ID_CUOTA = '" & IntId_Cuota2 & "' AND NRO_DOC = '" & strNroDocumento2 & "' AND RUT_DEUDOR = '" & strRutDD2 &"' AND ISNULL(ESTADO_FACTURA,'100') NOT IN ('3','6')"

					ElseIf strTipoProceso2 = "FACTURABLE NUEVAMENTE" Then
						strSql = "UPDATE CUOTA SET FECHA_ESTADO_FACT = '" & strFecha & "', USUARIO_ESTADO_FACT = " & idUsuario2 & ", OBSERVACION_FACTURACION = '" & strObsFact2 & "' ,ESTADO_FACTURA = '7', OBS_PROCESO_FACT = 'FACTURABLE NUEVAMENTE' WHERE ID_CUOTA = '" & IntId_Cuota2 & "' AND NRO_DOC = '" & strNroDocumento2 & "' AND RUT_DEUDOR = '" & strRutDD2 &"' AND ISNULL(ESTADO_FACTURA,'100') IN ('6') AND ISNULL(ESTADO_FACTURA,'100') NOT IN ('7')"

					End if

					'Response.write "<br>strSql=" & strSql
					'Response.End

					AbrirSCG1()
					set rsUpdate = Conn1.execute(strSql)
					CerrarSCG1()

					rsUsu.movenext
					loop
				end if
				rsUsu.close
				set rsUsu=nothing


				If strTipoProceso2 = "INGRESAR FACTURA" Then

					strSql = "SELECT C.COD_CLIENTE, C.NUMERO_FACTURA, SUM(MONTO_VISACION) AS MONTO_VISACION, SUM(MONTO_FACTURACION) AS MONTO_FACTURA,"
					strSql = strSql & " COUNT(NUMERO_FACTURA) AS CANT_DOC_FACT, USUARIO_ESTADO_FACT"
					strSql = strSql & " FROM TMP_PROCESO_FACTURACION TPF INNER JOIN CUOTA C ON TPF.ID_CUOTA = C.ID_CUOTA"
					strSql = strSql & " GROUP BY C.COD_CLIENTE, C.NUMERO_FACTURA,USUARIO_ESTADO_FACT"

					SET rsFactura=Conn.execute(strSql)

					'Response.write "<br>strSql=" & strSql

					Do While Not rsFactura.Eof
						strCOD_CLIENTE = rsFactura("COD_CLIENTE")
						intMontovisacion= rsFactura("MONTO_VISACION")
						intMontoFactura = rsFactura("MONTO_FACTURA")
						intCantFactura = rsFactura("CANT_DOC_FACT")
						strUsuarioFactura = rsFactura("USUARIO_ESTADO_FACT")


						strSql = "INSERT INTO FACTURACION_CLIENTES (ID_FACTURA, COD_CLIENTE, MONTO_TOTAL_FACTURA, NUMERO_FACTURA, FECHA_FACTURA, OBSERVACION_FACTURA, USUARIO_EMISOR_FACTURA, FECHA_EMISION, TDOC_FACT, ESTADO_FACTURA,FECHA_ESTADO_FACTURA,USUARIO_ESTADO_FACTURA,MONTO_ENVIADO_VISAR)"
						strSql = strSql & " VALUES (ISNULL((SELECT MAX (ID_FACTURA) FROM FACTURACION_CLIENTES),0)+ 1," & strCOD_CLIENTE & "," & intMontoFactura & "," & strNroFact & ",CAST('" & dtmFecFact & "'+' '+ REVERSE(SUBSTRING(REVERSE(convert(varchar(10),CAST('" & strFecha & "' AS DATETIME),108)),4,10)) AS DATETIME),'" & strObsFact3 & "','" & strUsuarioFactura & "','" & strFecha & "'," & intCantFactura & ",'3','" & strFecha & "','" & strUsuarioFactura & "','" & intMontovisacion & "')"

						'Response.write "<br>strSql=" & strSql

						Conn.execute(strSql)

						rsFactura.Movenext
					Loop

				End if

				strSql = " DROP TABLE [TMP_PROCESO_FACTURACION]"
				Conn.Execute strSql,64

CerrarSCG()
		If Validacion <> "0" Then
		%>

		<script>
			alert('Proceso realizado correctamente');
		</script>
		<%
		End if

	  End if%>

	<title>UTILITARIO</title>
	<style type="text/css">
		<!--
		.Estilo37 {color: #FFFFFF}
		-->
	</style>

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
</head>
<body>
	<div class="titulo_informe">UTILITARIO FACTURACIÓN</div>
	<br>
	<table width="90%" align="CENTER" border="0" align="center">
		   <tr>
			<td valign="top">
			<BR>
			<FORM name="datos" method="post">
			<table width="100%" border="0" class="estilo_columnas">
				<thead>
				<tr >
					<td aling = "left" colspan = "2" width="39%">TIPO PROCESO</td>
					<td aling = "left" colspan = "1" width="21%">NUMERO FACTURA</td>
					<td aling = "left" colspan = "1" width="20%">FECHA FACTURA</td>
					<td aling = "left" colspan = "1" width="20%">OBSERVACION FACTURA</td>
				</tr>
				</thead>
				<tr>
					<td Colspan= "2">
						<select name="CB_TIPO">
							<option value="0" <%If strTipoProceso = "0" then response.write "SELECTED"%>>SELECCIONAR</option>
							<option value="ENVIO A VISAR" <%If strTipoProceso = "ENVIO A VISAR" then response.write "SELECTED"%>>ENVIO A VISAR</option>
							<option value="ENVIO A FACTURAR" <%If strTipoProceso = "ENVIO A FACTURAR" then response.write "SELECTED"%>>ENVIO A FACTURAR</option>
							<option value="INGRESAR FACTURA" <%If strTipoProceso = "INGRESAR FACTURA" then response.write "SELECTED"%>>INGRESAR FACTURA</option>
							<option value="NO FACTURABLE" <%If strTipoProceso = "NO FACTURABLE" then response.write "SELECTED"%>>NO FACTURABLE</option>
							<option value="ANULAR PROCESO" <%If strTipoProceso = "ANULAR PROCESO" then response.write "SELECTED"%>>ANULAR PROCESO</option>
							<option value="FACTURABLE NUEVAMENTE" <%If strTipoProceso = "FACTURABLE NUEVAMENTE" then response.write "SELECTED"%>>FACTURABLE NUEVAMENTE</option>
						</select>
					</td>

					<td Colspan= "1">
						<input name="TX_NRO_FACT" type="text" value="<%=strNroFact%>" size="12" maxlength="12">
					</td>

					<td Colspan= "1">
						<input name="TX_FECHA_FACT" type="text" id="TX_FECHA_FACT" readonly="true" value="<%=dtmFecFact%>" size="10" maxlength="10">
							<!--<a href="javascript:showCal('TX_FECHA_FACT');"><img src="../Imagenes/calendario.gif" border="0"></a>-->
					</td>

					<td Colspan= "1">
						<input name="TX_OBS_FACT3" type="text" value="<%=strObsFact3%>" size="45" maxlength="45">
					</td>
				</tr>
			</table>

			<table width="80%" border="0" bordercolor="#FFFFFF" ALIGN="center">
				<TR>

					<TD class=hdr_i>
						ID Cuota<BR><BR>
						<TEXTAREA NAME="TX_ID_CUOTA" ROWS=30 COLS=13><%=intIdCuota%></TEXTAREA>
					</TD>
					<TD class=hdr_i>
						Rut Deudor<BR><BR>
						<TEXTAREA NAME="TX_RUT" ROWS=30 COLS=15><%=strRut%></TEXTAREA>
					</TD>
					<TD class=hdr_i>
						Nro. Documento<BR><BR>
						<TEXTAREA NAME="TX_NRO_DOC" ROWS=30 COLS=15><%=strNroDoc%></TEXTAREA>
					</TD>
					<TD class=hdr_i>
						Monto Visar/Facturar<BR><BR>
						<TEXTAREA NAME="TX_MONTO" ROWS=30 COLS=14><%=intMonto%></TEXTAREA>
					</TD>
					<TD class=hdr_i>
						Observacion Documento Facturado<BR><BR>
						<TEXTAREA NAME="TX_OBSFACT" ROWS=30 COLS=33><%=strObsFact%></TEXTAREA>
					</TD>
				</TR>
				<TR>
					<TD colspan="5" ALIGN="RIGHT">
						<INPUT TYPE="BUTTON" class="fondo_boton_100" value="Procesar" name="B1" onClick="envia('G');return false;">
					</TD>
				</TR>
			</table>

			  </td>
		  </tr>
		</table>

</form>
</body>
</html>
<script type="text/javascript">
$(document).ready(function(){

	$('#TX_FECHA_FACT').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

})
</script>
<script language="JavaScript1.2">

function envia(intTipo)	{

		if ((intTipo=='G') && (document.datos.CB_TIPO.value == 'INGRESAR FACTURA') && ((document.datos.TX_FECHA_FACT.value =='') || (document.datos.TX_NRO_FACT.value =='')) )
		{
			alert("Para ingresar factura debe ingresar el numero y la fecha de emisión de la factura")
			return false
		}
		else if ((intTipo=='G') && (document.datos.CB_TIPO.value == '0'))
		{
			alert("Debe seleccionar el tipo de proceso")
			return false
		}
		else
		{
			document.forms[0].action='Utilitario_facturacion.asp?strGraba=S';
			document.forms[0].submit();
		}
	}


</script>


