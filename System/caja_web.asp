<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
    <link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->	
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->

	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>
    <link href="../css/style.css" rel="stylesheet" type="text/css">

    <script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
    <script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
    <script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
    <script src="../Componentes/jquery.numeric/jquery.numeric.js"></script> 

    <script language="JavaScript " type="text/JavaScript">
    $(document).ready(function(){

        $.prettyLoader();
    })
</script>


<%
	If Trim(Request("Limpiar"))="1" Then
		session("session_RUT_DEUDOR") = ""
		rut = ""
	End if

	rut = request("rut")
	id_convenio = request("id_convenio")

	if trim(rut) <> "" Then
		session("session_RUT_DEUDOR") = rut
	Else
		rut = session("session_RUT_DEUDOR")
	End if

	strRUT_DEUDOR=rut

	intSeq = request("intSeq")
	strGraba = request("strGraba")
	txt_FechaIni = request("txt_FechaIni")
	intSucursal="1"
	fecha= date
	strCOD_CLIENTE = request("CB_CLIENTE")
	strCOD_CLIENTE = session("ses_codcli")
	usuario=session("session_idusuario")

'--Ve si el cliente posee cobranza externa--'

abrirscg()

		strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
		strSql = strSql & " FROM CLIENTE CL"
		strSql = strSql & " WHERE CL.COD_CLIENTE = '" & strCOD_CLIENTE & "'"

		set RsCli=Conn.execute(strSql)
		If not RsCli.eof then
			intUsaCobInterna = RsCli("USA_COB_INTERNA")
		End if
		RsCli.close
		set RsCli=nothing

cerrarscg()

'-------------------------------------------'

	AbrirSCG()

	if Trim(fechahoy) = "" Then
		fechahoy = TraeFechaActual(Conn)
	End If

	intNroBoleta = request("TX_BOLETA")
	intCompIngreso = request("TX_COMPINGRESO")
	intCantDoc = request("TX_CANTDOC")
	intMontoCliente = request("TX_MONTOCLIENTE")
	intFormaPagoCliente = request("CB_FPAGO_CLIENTE")
	intFormaPagoEMP = request("CB_FPAGO_CLIENTE")
	intTipoPago = request("CB_TIPOPAGO")

	strBcoDepEmpresa = request("CB_BEMP")
	strBcoDepCliente = request("CB_BCLIENTE")

	strNroDepCliente = request("TX_NRODEPCLIENTE")
	strNroDepEmpresa = request("TX_NRODEPEMP")

	strClaveDeudor = request("TX_CLAVEDOC")
	dtmFechaPago = request("TX_FECHA_PAGO")
	intPeriodoPagado = request("TX_PERIODO")
	intCapital = request("TX_CAPITAL")
	intReajuste = request("TX_REAJUSTE")
	intInteres = request("TX_INTERES")
	intGravamenes = request("TX_GRAVAMENES")
	intMulta = request("TX_MULTAS")
	intCargos = request("TX_CARGOS")
	intCostas = request("TX_COSTAS")

	strDocCancelados = request("TX_DOCCANCELADOS")
	strObservaciones = request("TX_OBSERVACIONES")

	intTotalCapital = ValNulo(request("TX_DEUDACAPITAL"),"N")
	intIndemComp = ValNulo(request("TX_INDCOM"),"N")
	intHonorarios = ValNulo(request("TX_HONORARIOS"),"N")
	intOtros = ValNulo(request("TX_OTROS"),"N")
	intIntereses = ValNulo(request("TX_INTERESES"),"N")
	intGastosJud = ValNulo(request("TX_GASTOSJUD"),"N")


	intGastosAdmin = ValNulo(request("TX_GASTOSADMIN"),"N")

	intDescuento = ValNulo(request("TX_DESCUENTO"),"N")

	intNroBoleta = ValNulo(intNroBoleta,"N")
	intCompIngreso = ValNulo(intCompIngreso,"C")
	intCantDoc = ValNulo(intCantDoc,"N")
	intMontoCliente = Trim(intMontoCliente)
	intFormaPagoCliente = Trim(intFormaPagoCliente)
	intFormaPagoEMP = Trim(intFormaPagoEMP)
	intTipoPago = Trim(intTipoPago)
	If Trim(intTipoPago) =  "RP" Then intTipoPago =  "RP1"
	strNroDepCliente = Trim(strNroDepCliente)
	strNroDepEmpresa = Trim(strNroDepEmpresa)
	strClaveDeudor = Trim(strClaveDeudor)
	dtmFechaPago = Trim(dtmFechaPago)
	intPeriodoPagado = Trim(intPeriodoPagado)
	intCapital = Trim(intCapital)
	intReajuste = ValNulo(intReajuste,"N")
	intInteres = ValNulo(intInteres,"N")
	intGravamenes = Trim(intGravamenes)
	intMulta = Trim(intMulta)
	intCargos = Trim(intCargos)
	intCostas = Trim(intCostas)
	intOtros = Trim(intOtros)
	strDocCancelados = Trim(strDocCancelados)
	strObservaciones = Trim(strObservaciones)


	strSql=" SELECT TOP 1 ID_CAJA"
	strSql=strSql & " FROM CAJAS_RECAUDACION_USUARIO WHERE ID_USUARIO = '" & session("session_idusuario") & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "'"
	'Response.write "<br>strSql=" & strSql

	set rsCajaAsig=Conn.execute(strSql)
	if not rsCajaAsig.eof then
		do until rsCajaAsig.eof


		strSql="SELECT ISNULL(FECHA_APERTURA,'') AS FECHA_APERTURA, ISNULL(FECHA_CIERRE,'') AS FECHA_CIERRE FROM CAJA_WEB_EMP_CIERRE WHERE COD_USUARIO = " & usuario
		strSql= strSql & " AND FECHA_APERTURA < CAST('" & fecha & "' AS DATETIME) AND FECHA_CIERRE IS NULL AND CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "'"
		'Response.write "<br>strSql=" & strSql

		set rsNoCierreCajaDiaAnterior=Conn.execute(strSql)

		If Not rsNoCierreCajaDiaAnterior.Eof Then
			strNoCierreDiaAnterior="NOK"
		   strCierreHab = "NO"
			%>
			<SCRIPT>
				alert('La Caja no fue cerrada el día de ayer, Favor cierrela antes de ingresar el pago.')
				location.href='apertura_caja.asp?rut=" + rut + "&tipo=1';
			</SCRIPT>
			<%
		Else
				If strRUT_DEUDOR = "" then
					intRutSelNOk = 1
				Else
					strSql="SELECT RUT_DEUDOR FROM DEUDOR"
					strSql= strSql & " WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND '" & strRUT_DEUDOR & "' IN (SELECT RUT_DEUDOR FROM CUOTA INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA=ESTADO_DEUDA.CODIGO WHERE ESTADO_DEUDA.CODIGO=1 AND COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "')"
					'Response.write "<br>strSql=" & strSql

					set rsApertura=Conn.execute(strSql)
					If Not rsApertura.Eof Then
						intRutSelNOk=0
					Else
						intRutSelNOk=1
					End if

				End If

				If intRutSelNOk = 1 and Trim(id_convenio) = "" then%>
					<SCRIPT>
						alert('Favor seleccione un deudor Activo del Cliente para ingresar pagos')
						location.href='principal.asp';
					</SCRIPT>
				<%End if

				strNoCierreDiaAnterior="OK"

				strSql="SELECT ISNULL(FECHA_APERTURA,'') AS FECHA_APERTURA, ISNULL(FECHA_CIERRE,'') AS FECHA_CIERRE FROM CAJA_WEB_EMP_CIERRE WHERE COD_USUARIO = " & usuario
				strSql= strSql & " AND FECHA_CIERRE IS NULL AND CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "'"
				'Response.write "<br>strSql=" & strSql

				set rsApertura=Conn.execute(strSql)
				If Not rsApertura.Eof Then
					strApertura="OK"
					dtmFechaCaja = rsApertura("FECHA_APERTURA")
				Else
					strApertura="NOK"
					strCierreHab = "NO"
					dtmFechaCaja = ""
				End if

				strSql="SELECT ISNULL(FECHA_APERTURA,'') AS FECHA_APERTURA, ISNULL(FECHA_CIERRE,'') AS FECHA_CIERRE FROM CAJA_WEB_EMP_CIERRE WHERE COD_USUARIO = " & usuario
				strSql= strSql & " AND FECHA_CIERRE IS NULL AND CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "'"

				'Response.write "<br>strSql=" & strSql
				set rsCierre=Conn.execute(strSql)

				If Not rsCierre.Eof Then

					strSql1="SELECT ISNULL(FECHA_APERTURA,'') AS FECHA_APERTURA, ISNULL(FECHA_CIERRE,'') AS FECHA_CIERRE FROM CAJA_WEB_EMP_CIERRE WHERE COD_USUARIO = " & usuario
					strSql1= strSql & " AND FECHA_CIERRE IS NULL AND CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "'"
					set strCierreDS=Conn.execute(strSql1)

					If Not strCierreDS.Eof Then
						strCierre="OK"
					Else
						strCierre=""
					End If

				Else

					strCierre="OK"
					strCierreHab = ""

				End if

				'Response.write "strCierreHab=" & strCierreHab
				'Response.write strSql
				'Response.End

				If strApertura="NOK" Then
					strCierreHab = "NO"
					%>
					<SCRIPT>
						alert('La caja no ha sido abierta. Debe abrirla para poder ingresar pagos.')
						location.href='apertura_caja.asp?rut=" + rut + "&tipo=1';
					</SCRIPT>
					<%
				End If

		End if

		rsCajaAsig.movenext
		loop

	Else

		Response.Write ("<script language = ""Javascript"">" & vbCrlf)

		Response.Write (vbTab & "alert('No se puede abrir la caja, Debe solicitar al administrador que le asigne una caja para este cliente');" & vbCrlf)
		Response.Write (vbTab & "location.href='principal.asp?rut=" + rut + "&tipo=1';" & vbCrlf)

		Response.Write ("</script>")

	End if
	rsCajaAsig.close
	set rsCajaAsig=nothing


	If Trim(intTipoPago) =  "CO" OR Trim(intTipoPago) =  "CC" Then
		strSql="SELECT ID_CONVENIO FROM CONVENIO_ENC "
		'strSql= strSql & "WHERE RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "'"
		strSql= strSql & "WHERE ID_CONVENIO = " & id_convenio
		strTitCol = "CONVENIO"
		set rsConvenio=Conn.execute(strSql)
		If Not rsConvenio.Eof Then
			id_convenio = rsConvenio("ID_CONVENIO")
		End If

	ElseIf Trim(intTipoPago) =  "RP" Then

	Else
		strSql="SELECT ID_CONVENIO FROM CONVENIO_ENC "
		strSql= strSql & "WHERE RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "'"

		set rsConvenio=Conn.execute(strSql)
		If Not rsConvenio.Eof Then
			strTieneConvenio = "SI"
		End If

	End if

	If Trim(request("strGraba")) = "SI" Then
		
       ' Response.write "GRABANDO" & intTipoPago
       ' Response.end

		strSql = "EXEC UPD_SEC 'NRO_COMP'"
		set rsNroComp=Conn.execute(strSql)
		If not rsNroComp.Eof then
			intCompIngreso = rsNroComp("SEQ")
		Else
			intCompIngreso = "1"
		End if

		strSql = "EXEC UPD_SEC 'CAJA_WEB_EMP'"

		set rsPago=Conn.execute(strSql)
		If not rsPago.Eof then
			intSeq = rsPago("SEQ")
		End if


		intTotalCliente = Cdbl(ValNulo(intTotalCapital,"N")) + Cdbl(ValNulo(intIndemComp,"N")) + Cdbl(ValNulo(intGastosJud,"N")) + Cdbl(ValNulo(intIntereses,"N")) '''+ intOtros

		
		intTotalEmpresa = ValNulo(intHonorarios,"N") + ValNulo(intOtros,"N") + ValNulo(intGastosAdmin,"N")

        intCompIngreso = replace(replace(intCompIngreso,",",""),".","")
        intTotalCapital= replace(replace(intTotalCapital,",",""),".","")
        intIndemComp= replace(replace(intIndemComp,",",""),".","")
        intHonorarios= replace(replace(intHonorarios,",",""),".","")
        intGastosJud= replace(replace(intGastosJud,",",""),".","")
        intOtros= replace(replace(intOtros,",",""),".","")
        intIntereses= replace(replace(intIntereses,",",""),".","")
        intTotalCliente= replace(replace(intTotalCliente,",",""),".","")
        intTotalEmpresa= replace(replace(intTotalEmpresa,",",""),".","")
        intGastosAdmin  = replace(replace(intGastosAdmin,",",""),".","")


		If Trim(intTipoPago) =  "RP1" Then intTipoPago =  "RP"

		strSql = "INSERT INTO CAJA_WEB_EMP (ID_PAGO, FECHA_PAGO, COD_CLIENTE, RUT_DEUDOR, COMP_INGRESO, NRO_BOLETA, MONTO_CAPITAL, TIPO_PAGO, INDEM_COMP, MONTO_EMP, GASTOS_JUDICIALES, GASTOS_OTROS, INTERES_PLAZO, TOTAL_CLIENTE, TOTAL_EMP, USR_INGRESO, FECHA_INGRESO, DESC_CLIENTE, ID_CONVENIO, OBSERVACIONES, GASTOS_ADMINISTRATIVOS) "
		strSql = strSql & " VALUES (" & intSeq & ", '" & dtmFechaCaja & "','" & strCOD_CLIENTE & "','" & strRUT_DEUDOR & "'," & intCompIngreso & "," & intNroBoleta & "," & intTotalCapital & ",'" & intTipoPago & "'," & intIndemComp & "," & intHonorarios & "," & intGastosJud & "," & intOtros & "," & intIntereses & "," & intTotalCliente & "," & intTotalEmpresa & "," & session("session_idusuario") & ",getdate()," & intDescuento & "," & ValNulo(id_convenio,"N") & ",'" & strObservaciones & "'," & intGastosAdmin & ")"


		If Trim(intTipoPago) =  "RP" Then intTipoPago =  "RP1"
		''Response.write strSql
		''Response.End
		set rsInsertaEnc=Conn.execute(strSql)


		If Trim(intTipoPago) <> "CO" AND Trim(intTipoPago) <> "CC" AND Trim(intTipoPago) <> "RP" Then

			strSql = "SELECT ID_CUOTA, NRO_DOC, CUENTA, FECHA_VENC FROM CUOTA WHERE RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND COD_CLIENTE='" & strCOD_CLIENTE & "' AND SALDO > 0"

            'response.Write(strSql)
            'response.end
			set rsTemp= Conn.execute(strSql)

			intCorrelativo = 1
			Do until rsTemp.eof
				strObjeto = "CH_" & Replace(Trim(rsTemp("ID_CUOTA")),"-","_")
				strObjeto1 = "TX_SALDO_" & Replace(Trim(rsTemp("ID_CUOTA")),"-","_")


				If UCASE(Request(strObjeto)) = "ON" Then

					intSaldo = Request(strObjeto1)
					strNroDoc = rsTemp("NRO_DOC")
					strCuenta = rsTemp("CUENTA")
					strFechaVenc = rsTemp("FECHA_VENC")

					strSql = "INSERT INTO CAJA_WEB_EMP_DETALLE (ID_PAGO, CORRELATIVO, NRO_DOC, CUENTA, FECHA_VENC, CAPITAL, REAJUSTE, INTERESES, ID_CUOTA) "
					strSql = strSql & " VALUES (" & intSeq & ", " & intCorrelativo & ",'" & strNroDoc & "','" & strCuenta & "','" & strFechaVenc & "'," & replace(replace(intSaldo,",",""),".","") & "," & replace(replace(intReajuste,",",""),".","") & "," & replace(replace(intInteres,",",""),".","") & "," & Trim(rsTemp("ID_CUOTA")) & ")"
					'Response.write strSql
					'Response.End
					set rsInsertaDet=Conn.execute(strSql)

					strSql = "UPDATE CUOTA SET ID_PAGO = " & intSeq & ", SALDO = 0, ESTADO_DEUDA = '4', FECHA_ESTADO = GETDATE() WHERE ID_CUOTA = " & Trim(rsTemp("ID_CUOTA"))
					'Response.write strSql
					'Response.End
					set rsUpdate=Conn.execute(strSql)

				End if
			rsTemp.movenext
			intCorrelativo = intCorrelativo + 1
			loop
			rsTemp.close
			set rsTemp=nothing

		Else
			If Trim(intTipoPago)="CO" or Trim(intTipoPago) = "CC"  Then
				strSql = "SELECT ID_CONVENIO, CUOTA, TOTAL_CUOTA, FECHA_PAGO AS FECHA_VENC FROM CONVENIO_DET WHERE ID_CONVENIO = " & id_convenio & " AND ID_CONVENIO IN (SELECT ID_CONVENIO FROM CONVENIO_ENC WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "')"
			Else

			End If

			set rsTemp= Conn.execute(strSql)
			intCorrelativo = 1
			Do until rsTemp.eof
				strObjeto = "CH_" & rsTemp("CUOTA")
				strObjeto1 = "TX_SALDO_" & rsTemp("CUOTA")

				If UCASE(Request(strObjeto)) = "ON" Then

					intSaldo = Request(strObjeto1)
					intNroCuota = rsTemp("CUOTA")
					strFechaVenc = rsTemp("FECHA_VENC")

					strSql = "INSERT INTO CAJA_WEB_EMP_DETALLE (ID_PAGO, CORRELATIVO, NRO_DOC, FECHA_VENC, CAPITAL, REAJUSTE, INTERESES) "
					strSql = strSql & " VALUES (" & intSeq & ", " & intCorrelativo & ",'" & intNroCuota & "','" & strFechaVenc & "'," & replace(replace(intSaldo,",",""),".","")  & "," & replace(replace(intReajuste,",",""),".","")  & "," & replace(replace(intInteres,",",""),".","")  & ")"
					'Response.write strSql
					'Response.End
					set rsInsertaDet=Conn.execute(strSql)

					If Trim(intTipoPago)="CO" or Trim(intTipoPago) = "CC" Then
						strSql = "UPDATE CONVENIO_DET SET ID_PAGO = " & intSeq & ", ID_PAGO_CORR = " & intCorrelativo & ",PAGADA = 'S', FECHA_DEL_PAGO = GETDATE() WHERE ID_CONVENIO = " & id_convenio & " AND CUOTA = " & intNroCuota
					Else

					End if
					'Response.write strSql
					'Response.End
					set rsUpdate=Conn.execute(strSql)
					intCorrelativo = Cdbl(intCorrelativo) + 1

					strNroDoc = id_convenio & "-" & intNroCuota
					strSql = "UPDATE CUOTA SET ESTADO_DEUDA = '4', FECHA_ESTADO = getdate(), SALDO = 0 WHERE RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "' AND NRO_DOC = '" & strNroDoc & "'"
					set rsUpdate=Conn.execute(strSql)
				End if
				rsTemp.movenext
			loop
			rsTemp.close
			set rsTemp=nothing
		End If

		strTipo = request("TXDESTINO")
		strRutCli = request("TXRUTCLI")
		strFormaPago = request("TXFPAGO")
		strcuotas = request("TXCUOTA")
		intMontoCli = replace(replace(request("TXMONTOCLI"),",",""),".","") 
		intFechaCli = request("TXFECVENCLI")
		intBancoCli = request("TXBANCOCLIENTE")
		//intPlazaCli = request("TXPLAZACLIENTE")
		intNroCheCli = request("TXNROCHEQUECLI")
		intNroCtaCli = request("TXNROCTACTECLI")
		strDivide = request("TXDIVIDE")
		strHretenido = request("TXHRETENIDO")

		If intFechaCli = "" then
			intFechaCli = fechahoy
		End if


		'Response.write "<br>intFechaCli=" & intFechaCli
		'Response.write "<br>strTipo=" & strTipo

		sTipo = Trim(strTipo)
		sRut = Trim(strRutCli)
		sForma = Trim(strFormaPago)
		sCuota = Trim(strcuotas)
		sMonto = Trim(intMontoCli)
		sFecha = Trim(intFechaCli)
		sBanco = Trim(intBancoCli)
		sPlaza = Trim(intPlazaCli)
		sNroChe = Trim(intNroCheCli)
		sNroCta = Trim(intNroCtaCli)
		sDivide = Trim(strDivide)
		sHretenido= Trim(strHretenido)

		estado="A"


		vTipo = split(sTipo,"*")
		vRut = split(sRut,"*")
		vForma = split(sForma,"*")

		vMonto = split(sMonto,"*")
		vFecha = split(sFecha,"*")
		vBanco = split(sBanco,"*")
		vDivide = split(sDivide,"*")
		vHretenido = split(sHretenido,"*")

		vPlaza = split(sPlaza,"*")
		vNroChe = split(sNroChe,"*")
		vNroCta = split(sNroCta,"*")

		x=0
		intCorrelativoCH = 0
		if vTipo(0) <> "" and vTipo(0) <> "NULL" then
		for each doc in vTipo

			'Response.write "<br>x=" & vMonto(x)
			'Response.write "<br>FEcha=" & IsNull(vFecha)
			'Response.write "<br>vForma=" & vForma(x)
			'Response.write "<br>vFecha=" & vFecha(x)

			If Trim(vTipo(x)) = "0" Then
				strBancoDep = strBcoDepCliente
				strNroDep = strNroDepCliente
			ElseIf Trim(vTipo(x)) = "1" Then
				strBancoDep = strBcoDepEmpresa
				strNroDep = strNroDepEmpresa
			Else
				strBancoDep = ""
				strNroDep = ""
			End if

			''Response.write "<br>vForma(x)=" & vForma(x)
			''Response.write "<br>vFecha(x)=" & vFecha(x)
			''Response.write "<br>x=" & x


			If vFecha(x) = "" then
				vFecha(x) = fechahoy
			End if

			''Response.write "strSql=" & fechahoy

			If (Trim(vForma(x))="EF" OR Trim(vForma(x))="OC" OR Trim(vForma(x))="TR" OR Trim(vForma(x))="TD" OR Trim(vForma(x))="TC") Then
				SQL = "INSERT INTO CAJA_WEB_EMP_DOC_PAGO (ID_PAGO, CORRELATIVO, MONTO, VENCIMIENTO, COD_BANCO, NRO_DEPOSITO, TIPO_PAGO, FORMA_PAGO, DIVIDE, H_RETENIDO)"
				SQL = SQL & " VALUES (" & intSeq & "," & x + 1 & "," & vMonto(x) & ", '" & vFecha(x) & "' , '" & strBancoDep & "','" & strNroDep & "','" & vTipo(x) & "','" & vForma(x) & "','" &vDivide(x) & "','" &vHretenido(x) & "')"

			ElseIf (Trim(vForma(x))="DP") Then
				SQL = "INSERT INTO CAJA_WEB_EMP_DOC_PAGO (ID_PAGO, CORRELATIVO, MONTO, VENCIMIENTO, COD_BANCO, NRO_DEPOSITO, TIPO_PAGO, FORMA_PAGO, DIVIDE, H_RETENIDO)"
				SQL = SQL & " VALUES (" & intSeq & "," & x + 1 & "," & vMonto(x) & ", '" & vFecha(x) & "' , '" & vBanco(x) & "','" & vNroChe(x) & "','" & vTipo(x) & "','" & vForma(x) & "','" &vDivide(x) & "','" &vHretenido(x) & "')"

			ElseIf (Trim(vForma(x))="CD" or Trim(vForma(x))="CF" or Trim(vForma(x))="VV") Then
				SQL = "INSERT INTO CAJA_WEB_EMP_DOC_PAGO (ID_PAGO, CORRELATIVO, MONTO, VENCIMIENTO, COD_BANCO, NRO_CHEQUE, NRO_CTA_CTE, CODIGO_PLAZA, TIPO_PAGO, FORMA_PAGO,RUT_CHEQUE, DIVIDE, H_RETENIDO)"
				SQL = SQL & " VALUES (" & intSeq & "," & x + 1 & "," & vMonto(x) & ",'" & vFecha(x) & "','" & vBanco(x) & "','" & vNroChe(x) & "','" & vNroCta(x) & "','" & vPlaza11 & "','" & vTipo(x) & "','" & vForma(x) & "','" & vRut(x) & "','" &vDivide(x) & "','" &vHretenido(x) & "')"

			Else
				SQL = "INSERT INTO CAJA_WEB_EMP_DOC_PAGO (ID_PAGO, CORRELATIVO, MONTO, COD_BANCO, NRO_CHEQUE, NRO_CTA_CTE, CODIGO_PLAZA, TIPO_PAGO, FORMA_PAGO,RUT_CHEQUE, DIVIDE, H_RETENIDO)"
				SQL = SQL & " VALUES (" & intSeq & "," & x + 1 & "," & vMonto(x) & ",'" & vBanco(x) & "','" & vNroChe(x) & "','" & vNroCta(x) & "','" & vPlaza11 & "','" & vTipo(x) & "','" & vForma(x) & "','" & vRut(x) & "','" &vDivide(x) & "','" &vHretenido(x) & "')"
			end if

			set rsInser=Conn.execute(SQL)


			If (Trim(vForma(x))="CF11111") Then
				intCorrelativoCH = intCorrelativoCH + 1
				intTipoDocumento = 2 ''Cheque Protestado
				strClaveAdic = strCOD_CLIENTE & "-" & strRUT_DEUDOR & "-" & vNroChe(x) & "-" & intCorrelativoCH

				strSql="SELECT * FROM BANCOS WHERE CODIGO = " & Trim(vBanco(x))
				''Response.write "strSql=" & strSql
				set rsBANC=Conn.execute(strSql)
				if not rsBANC.eof then
					strBancoCH = rsBANC("NOMBRE_B")
				End If

				strSql = "INSERT INTO CUOTA (RUT_DEUDOR, COD_CLIENTE, NRO_DOC, NRO_CUOTA, FECHA_VENC, VALOR_CUOTA, SALDO, TIPO_DOCUMENTO, ESTADO_DEUDA, FECHA_ESTADO , FECHA_CREACION, USUARIO_CREACION, ID_PAGO,  ADIC_2) "
				strSql = strSql & " VALUES ('" & strRUT_DEUDOR & "','" & strCOD_CLIENTE & "','" & vNroChe(x) & "'," &intCorrelativoCH & " ,'" & vFecha(x) & "'," & vMonto(x) & "," & vMonto(x) & ",'" & intTipoDocumento & "','1',getdate(),getdate()," & session("session_idusuario")  & "," & intSeq & ",'" & strBancoCH & "')"

			End If

			'Response.write strSql
			'Response.End
			set rsInsertaCuota=Conn.execute(strSql)


			x = x + 1
		next
		end if

		'If Trim(intTipoPago)="CO" or Trim(intTipoPago) = "CC" Then
		'	strTipoCompArch = "comp_pago_convenio.asp"
		'Else
        If Trim(intTipoPago)="RP" Then
			strTipoCompArch = "comp_pago_repactacion.asp"
		Else
			strTipoCompArch = "comp_pago.asp"
		End If


		strSql = "SELECT COMP_INGRESO FROM CAJA_WEB_EMP WHERE ID_PAGO = " & intSeq
		set rsCompPago = Conn.execute(strSql)
		If Not rsCompPago.Eof Then
			intCompPago = rsCompPago("COMP_INGRESO")
		Else
			intCompPago = 0
		End If


		strSql="SELECT TOP 1 SUCURSAL, ADIC_5 FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "'"
		strSql = strSql & " AND NRO_DOC IN (SELECT NRO_DOC FROM CAJA_WEB_EMP_DETALLE  WHERE ID_PAGO = " & intSeq & ") AND SUCURSAL IS NOT NULL"
		set rsSuc=Conn.execute(strSql)
		if not rsSuc.eof then
			strSucursal = rsSuc("SUCURSAL")
			strInterlocutor = rsSuc("ADIC_5")
		end if
		rsSuc.close
		set rsSuc=nothing


		strRutCliente = TraeCampoId(Conn, "RUT", strCOD_CLIENTE, "CLIENTE", "COD_CLIENTE")

		If Trim(strRutCliente) <> "" AND Not IsNull(strRutCliente) Then
			strRutCSD = FormatNumber(Mid(strRutCliente,1,len(strRutCliente)-2),0)
			strRutCCD = Right(strRutCliente,2)
			strRutClie = strRutCSD & strRutCCD
		End If

		strSql="SELECT B.RUT, B.RAZON_SOCIAL FROM SEDE A, CONVENIO_CORRELATIVO B WHERE A.RUT = B.RUT AND A.COD_CLIENTE = B.COD_CLIENTE AND B.COD_CLIENTE = '" & strCOD_CLIENTE & "' AND A.SEDE = '" & strSucursal & "'"
		''Response.write "strSql=" & strSql
		set rsSede=Conn.execute(strSql)
		if not rsSede.eof then
			strRutClie = rsSede("RUT")
			If Trim(strRutClie) <> "" Then
				strRutCSD = FormatNumber(Mid(strRutClie,1,len(strRutClie)-2),0)
				strRutCCD = Right(strRutClie,2)
				strRutClie = strRutCSD & strRutCCD
			End If
			strRSocialCli = rsSede("RAZON_SOCIAL")
			strDescCli = strRutClie & " " & strRSocialCli
		end if
		rsSede.close
		set rsSede=nothing


		AbrirSCG1()
			strSql = "SELECT SUM(MONTO) AS MONTO , TIPO_PAGO, DIVIDE FROM CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO = " & intSeq & " AND DIVIDE IN ('1','2','4') GROUP BY TIPO_PAGO, DIVIDE"
			'Response.write "strSql=" & strSql
			set rsTemp=Conn.execute(strSql)
			Do While Not rsTemp.eof
				If Trim(rsTemp("TIPO_PAGO")) = "1" and Trim(rsTemp("DIVIDE")) = "1"  Then 'LLACRUZ
					intTotalHonorarios = ValNulo(rsTemp("MONTO"),"N")
				End If
				If Trim(rsTemp("TIPO_PAGO")) = "0" and Trim(rsTemp("DIVIDE")) = "2" Then 'CLIENTE
					intTotalARemesar = ValNulo(rsTemp("MONTO"),"N")
				End If
				If Trim(rsTemp("DIVIDE")) = "4" Then 'CLIENTE
					intTotalGastoAdministrativo = ValNulo(rsTemp("MONTO"),"N")
				End If
				rsTemp.movenext
			Loop
		CerrarSCG1()


		if Trim(intTipoPago) =  "CO" then
		
			AbrirSCG1()
				strSql = "SELECT CE.TOTAL_CONVENIO, CE.PIE FROM dbo.CAJA_WEB_EMP CWE " &_
						"INNER JOIN dbo.CONVENIO_ENC CE " &_
						"ON CWE.ID_CONVENIO = CE.ID_CONVENIO " &_
						"WHERE CWE.ID_PAGO = " & intSeq
				'Response.write "strSql=" & strSql
				set rsTemp=Conn.execute(strSql)
				Do While Not rsTemp.eof

						intTotaSaldoCliente= rsTemp("TOTAL_CONVENIO") - rsTemp("PIE")

					rsTemp.movenext
				Loop
			CerrarSCG1()
		
		else
		
			AbrirSCG1()
				strSql = "SELECT ISNULL(SUM(MONTO),0) AS MONTO FROM CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO = " & intSeq & " AND DIVIDE = 3 "
				'Response.write "strSql=" & strSql
				set rsTemp=Conn.execute(strSql)
				Do While Not rsTemp.eof

						intTotaSaldoCliente= rsTemp("MONTO")

					rsTemp.movenext
				Loop
			CerrarSCG1()
		
		end if

		AbrirSCG1()
			strSql = "SELECT ISNULL(H_RETENIDO,0) as H_RETENIDO FROM CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO = " & intSeq & " AND TIPO_PAGO = 1 "
			'Response.write "strSql=" & strSql
			set rsTemp=Conn.execute(strSql)
			Do While Not rsTemp.eof

					strHonorarioRetenido = rsTemp("H_RETENIDO")

				rsTemp.movenext
			Loop
		CerrarSCG1()
		
		If strHonorarioRetenido = "" then strHonorarioRetenido = "0" End If

		strRutCliente = strRutClie

		If trim(id_convenio) <> "" Then
			strNroOperacion = id_convenio
			strNroOperacion = intCompPago  'Validar esta info. Según nueva definición esto es lo correcto.
		Else
			strNroOperacion = intCompPago
		End If

		
		If Trim(id_convenio) <> "" Then
			strSql = "SELECT * FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO = " & intSeq
			set rsDetCaja=Conn.execute(strSql)
			If Not rsDetCaja.eof Then
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

					strSql="SELECT CUOTA, FECHA_PAGO FROM CONVENIO_DET WHERE CUOTA <> 0 AND ID_CONVENIO = " & id_convenio & " AND CUOTA = " & rsDetCaja("NRO_DOC")

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
		End If
		strGlosaCuota = strCuotas


		strSql = "SELECT TOP 1 D.ID_PAGO, CONVERT(VARCHAR(10),VENCIMIENTO,103) as VENCIMIENTO, (CASE WHEN D.FORMA_PAGO IN ('EF','TR') THEN convert(VARCHAR(10),[dbo].[fun_x_dias_sgtes_habiles] (VENCIMIENTO, 1),103) WHEN D.FORMA_PAGO IN ('DP','VV', 'CD','CF','TD','TC') THEN convert(VARCHAR(10),[dbo].[fun_x_dias_sgtes_habiles] (VENCIMIENTO, 3),103) ELSE VENCIMIENTO END ) as VENCIMIENTO3"
		strSql = strSql & " FROM CAJA_WEB_EMP_DOC_PAGO D, CAJA_WEB_EMP C"
		strSql = strSql & " WHERE C.ID_PAGO = D.ID_PAGO AND VENCIMIENTO IS NOT NULL"
		strSql = strSql & " AND D.ID_PAGO = " & intSeq & " ORDER BY D.TIPO_PAGO DESC, DIVIDE DESC, D.VENCIMIENTO ASC "

		'Response.write "<br>strSql=" & strSql

		set RsFec=Conn.execute(strSql)

			If not RsFec.eof then
				dtmFechaVenc = RsFec("VENCIMIENTO")
				dtmFechaVencRemesar = RsFec("VENCIMIENTO3")
			Else
				dtmFechaVencRemesar = ""
			End if

		RsFec.close
		set RsFec=nothing

		intTotalRecaudacion = Cdbl(intTotalHonorarios) + Cdbl(intTotalARemesar) + Cdbl(intTotalGastoAdministrativo)


		dtmFechaPagoConta = dtmFechaVenc

		If Trim(dtmFechaPagoConta) = "" or Trim(dtmFechaPagoConta) = "01/01/1900" or IsNull(dtmFechaPagoConta) Then
			dtmFechaPagoConta = dtmFechaCaja
		End If

		strSql = "SELECT TOP 1 FORMA_PAGO"
		strSql = strSql & " FROM CAJA_WEB_EMP_DOC_PAGO D"
		strSql = strSql & " WHERE D.ID_PAGO = " & intSeq
		strSql = strSql & " ORDER BY D.TIPO_PAGO DESC, D.DIVIDE DESC "

		set RsFec=Conn.execute(strSql)
		If not RsFec.eof then
			strFormaPago = RsFec("FORMA_PAGO")
		End if
		RsFec.close
		set RsFec=nothing

		if ((strFormaPago = "EF" or strFormaPago = "TR" or strFormaPago = "CD" or strFormaPago = "CF" or strFormaPago = "VV") and Cdbl(intTotalHonorarios) > "0" and Cdbl(intTotalARemesar) = "0") or (strFormaPago = "EF" and Cdbl(intTotalHonorarios) = "0" and Cdbl(intTotalARemesar) > "0") or (strFormaPago = "EF" and Cdbl(intTotalHonorarios) > "0" and Cdbl(intTotalARemesar) > "0") or strCod_Cliente = "1500" Then
			strRemesa = "S"
		Else
			strRemesa = "N"
		End If

		if (strFormaPago = "EF" and Cdbl(intTotalHonorarios) = "0" and Cdbl(intTotalARemesar) > "0") or strCod_Cliente = "1500" Then
			strTraspasado = "T"
		Else
			strTraspasado = "N"
		End If


		if strFormaPago = "" Then
			strFormaPago = "EF"
		End If
		
		if strCod_Cliente = "1500" Then
		
		strTipoProducto = "FACTURA"
		intComTransferencia = "99999"
		intComRemesa = "99999"
		
		else
		
		strTipoProducto = "OTRO"
		intComTransferencia = "null"
		intComRemesa =	"null"	
		End If
		
		strSql="SELECT FECHA AS FECHA, USUARIO.LOGIN AS LOGIN, CR.NOM_CAJA AS NOM_CAJA, CR.COD_CAJA AS COD_CAJA, ISNULL(FECHA_CIERRE,'') AS FECHA_CIERRE"
		strSql= strSql & " 	FROM CAJA_WEB_EMP_CIERRE INNER JOIN USUARIO ON USUARIO.ID_USUARIO = CAJA_WEB_EMP_CIERRE.COD_USUARIO"
		strSql= strSql & " 							 INNER JOIN CAJAS_RECAUDACION AS CR ON CR.ID_CAJA = CAJA_WEB_EMP_CIERRE.SUCURSAL"
		strSql= strSql & "	WHERE COD_USUARIO = " & usuario & " AND FECHA_CIERRE IS NULL AND CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "'"

		''Response.write "strSql=" & strSql

		set rsInforme=Conn.execute(strSql)
		if not rsInforme.eof then

		intCodCaja = rsInforme("COD_CAJA")

		end if
		rsInforme.close
		set rsInforme=nothing

        intTotalHonorarios = Replace(Replace(intTotalHonorarios,",",""),".","")
        intTotalGastoAdministrativo  = Replace(Replace(intTotalGastoAdministrativo,",",""),".","")
        intTotalARemesar = Replace(Replace(intTotalARemesar,",",""),".","")
        intTotalRecaudacion = Replace(Replace(intTotalRecaudacion,",",""),".","")
        intTotaSaldoCliente  = Replace(Replace(intTotaSaldoCliente,",",""),".","")
        strHonorarioRetenido  = Replace(Replace(strHonorarioRetenido,",",""),".","")
        intComTransferencia  = Replace(Replace(intComTransferencia,",",""),".","")

        if intTotalHonorarios = "" then intTotalHonorarios = 0
        if intTotalGastoAdministrativo = "" then intTotalGastoAdministrativo = 0
        if intTotalARemesar = "" then intTotalARemesar = 0
        if intTotalRecaudacion = "" then intTotalRecaudacion = 0
        if intTotaSaldoCliente = "" then intTotaSaldoCliente = 0
        if strHonorarioRetenido = "" then strHonorarioRetenido = 0
        if intComTransferencia = "" then intComTransferencia = 0

        'Response.write "strSql=" & intTotalHonorarios  & "<br/>"
		strSql = "SELECT MAX(NRO_CLIENTE_DEUDOR) AS NRO_CLIENTE_DEUDOR FROM CUOTA WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND COD_CLIENTE='" & strCOD_CLIENTE & "'"
		set rsInsertaEnc=Conn.execute(strSql)
		strNroClienteDeudor = rsInsertaEnc("NRO_CLIENTE_DEUDOR")

		strSql = "INSERT INTO EXP_CONTABILIDAD (NRO_CLIENTE_DEUDOR, FECHA_PAGO, NRO_OPERACION, RUT_CLIENTE, RUT_DEUDOR, FORMA_PAGO, FECHA_VENC_A_REMESAR, NRO_BOLETA, MONTO_HONORARIOS, MONTO_A_REMESAR, TOTAL_RECAUDACION, TRASPASADO, REMESA, ID_PAGO, ID_USUARIO, COD_CAJA, CP_NULO, SALDO_VIGENTE, H_RETENIDO,COM_REMESA,COM_TRANSFERENCIA,TIPO_PRODUCTO) "
		strSql = strSql & " VALUES ('" & strNroClienteDeudor & "', '" & dtmFechaPagoConta & "','" & strNroOperacion & "','" & strRutCliente & "','" & strRUT_DEUDOR& "','" & strFormaPago & "','" & dtmFechaVencRemesar & "'," & intNroBoleta & "," & (Cdbl(intTotalHonorarios) + Cdbl(intTotalGastoAdministrativo)) & "," & Cdbl(intTotalARemesar) & "," & Cdbl(intTotalRecaudacion) & ",'" & strTraspasado & "','" & strRemesa & "'," & intSeq & "," & usuario & "," & intCodCaja & ", 'N', " &  intTotaSaldoCliente & ", '" & strHonorarioRetenido & "'," &  intComTransferencia & "," &  intComRemesa & ", '" & strTipoProducto & "')"
		'Response.write "strSql=" & strSql

		set rsInsertaEnc=Conn.execute(strSql)
		'Response.End

		abrirscg()
			If "1" = "1" Then
				strSql1 = "EXEC Proc_Des_Asignacion_cobradores '" & strCOD_CLIENTE & "'," & session("session_idusuario")
				set rsDesAsig = Conn.execute(strSql1)

				strSql1 = "EXEC Proc_Cambia_Custodio_Deudor '" & strCOD_CLIENTE & "'," & session("session_idusuario")
				set rsCambiaCustodio = Conn.execute(strSql1)
			End If
		cerrarscg()

		strEnlaceImprime = strTipoCompArch & "?strImprime=S&intNroComp=" & intCompPago

		%>
		<SCRIPT>
			alert("Pago Ingresado Correctamente");
		</SCRIPT>
		<%

		%>
		<SCRIPT>
			window.open("<%=strEnlaceImprime%>","INFORMACION","width=800, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes");
		</SCRIPT>
		<%

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
	--> }
	 .hiddencol
        {
            display:none;
        }
	</style>

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>
	<script src="../javascripts/SelCombox.js"></script>
	<script src="../javascripts/OpenWindow.js"></script>


	<script language="JavaScript " type="text/JavaScript">

	function Refrescar(rut)
	{
		if(rut == '')
		{
			return
		}
				datos.action = "caja_web.asp?rut=" + rut + "&tipo=1";
				datos.submit();

	}

	function Refrescar1(rut)
	{
		if(rut == '')
		{
			return
		}
				datos.action = "caja_web.asp?nomas=1&rut=" + rut + "&tipo=1";
				datos.submit();

	}

	</script>


</head>
<body>
<form name="datos" method="post">
<div class="titulo_informe">Módulo de Ingreso de Pagos</div>
<table width="90%" border="0" bordercolor="#999999" cellpadding="2" cellspacing="5" align="center">
  <tr>
    <td valign="top">
	  <%
abrirscg()

	If rut <> "" then

		strNombreDeudor = TraeNombreDeudor(Conn,strRUT_DEUDOR)

		strSql=""
		strSql="SELECT FORMULA_HONORARIOS, FORMULA_INTERESES, INTERES_MORA,COD_TIPODOCUMENTO_HON, PIE_PORC_CAPITAL, HON_PORC_CAPITAL, IC_PORC_CAPITAL, TASA_MAX_CONV, DESCRIPCION, TIPO_INTERES, GASTOS_OPERACIONALES, GASTOS_ADMINISTRATIVOS, GASTOS_OPERACIONALES_CD, GASTOS_ADMINISTRATIVOS_CD FROM CLIENTE WHERE COD_CLIENTE ='" & strCOD_CLIENTE & "'"
		set rsTasa=Conn.execute(strSql)
		if not rsTasa.eof then
			intTasaMax = ValNulo(rsTasa("TASA_MAX_CONV"),"N")/100
			intTasaMensualMora = ValNulo(rsTasa("INTERES_MORA"),"C")
			intPorcPie = ValNulo(rsTasa("PIE_PORC_CAPITAL"),"N")/100
			intPorcHon = ValNulo(rsTasa("HON_PORC_CAPITAL"),"N")/100
			intPorcIc = ValNulo(rsTasa("IC_PORC_CAPITAL"),"N")/100
			strDescripcion = rsTasa("DESCRIPCION")
			strTipoInteres = rsTasa("TIPO_INTERES")
			intGOpeSD1=ValNulo(rsTasa("GASTOS_OPERACIONALES"),"N")
			intGOpeCD1=ValNulo(rsTasa("GASTOS_OPERACIONALES_CD"),"N")
			intGAdmSD1=ValNulo(rsTasa("GASTOS_ADMINISTRATIVOS"),"N")
			intGAdmCD1=ValNulo(rsTasa("GASTOS_ADMINISTRATIVOS_CD"),"N")
			intTipoDocHono = ValNulo(rsTasa("COD_TIPODOCUMENTO_HON"),"C")

			strNomFormHon = ValNulo(rsTasa("FORMULA_HONORARIOS"),"C")
			strNomFormInt = ValNulo(rsTasa("FORMULA_INTERESES"),"C")


		Else
			intTasaMax = 1
			intPorcPie = 1
			intPorcHon = 1
			intPorcIc = 1
			strDescripcion = ""
			strTipoInteres = "S"
		end if
		strTipoInteres = "S"
		rsTasa.close
		set rsTasa=nothing

		intMaximaAnual = intTasaMax * 12
		intMaximaAnualMora = intTasaMensualMora * 12

	Else
		strNombreDeudor=""
	End if

	%>

	<table width="100%" class="estilo_columnas">
	<thead>	
		<tr >
			<td>MANDANTE</td>
			<td>RUT</td>
			<td>NOMBRE O RAZON SOCIAL:</td>
			<td>USUARIO</td>
			<td>SUCURSAL</td>
			<td>FECHA</td>
			<td>&nbsp;</td>
		</tr>
	</thead>
	      <tr bgcolor="#FFFFFF" class="Estilo8">
	      <td>
	      	<select name="CB_CLIENTE">
				<%
					ssql="SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' ORDER BY RAZON_SOCIAL"
					set rsTemp= Conn.execute(ssql)
					if not rsTemp.eof then
						do until rsTemp.eof%>
						<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCOD_CLIENTE)=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("RAZON_SOCIAL")%></option>
						<%
						rsTemp.movenext
						loop
					end if
					rsTemp.close
					set rsTemp=nothing
				%>
			</select>
			</td>

			<td ALIGN="LEFT"><input name="TX_RUT" type="text" size="10" maxlength="10" onChange="Refrescar(this.value)" value="<%=rut%>"></td>
			<td><%=strNombreDeudor%><INPUT TYPE="hidden" NAME="rut" value="<%=rut%>"> </td>
	        <td ALIGN="RIGHT"><%=session("nombre_user")	%></td>

	        <td><%=nom_sucursal%></td>
	        <td><%=DATE%></td>
	        <td>
				<acronym title="LIMPIAR FORMULARIO">
					<input name="li_" class="fondo_boton_100" type="button" onClick="window.navigate('caja_web.asp?Limpiar=1');" value="Limpiar">
				</acronym>
			</td>
	      </tr>
    </table>
	</td>
	</tr>
</table>



<table width="90%" align="center" border="0" >
	<tr>
		<td>
			<font class="subtitulo_informe">> Resumen Pago</font>
		</td>
	</tr>
</table>	

<table width="90%" align="center" border="0" class="estilo_columnas">
<thead>
  <tr  bordercolor="#999999" >
    <td><span class="">Fecha Pago</span></td>
    <td><span class="">N° Comprobante</span></td>
	<td><span class="">N° Boleta</span></td>
	<td><span class="">Tipo Pago</span></td>
   </tr>
</thead>
  <tr class="Estilo8">
	<td><input name="TX_FECHA_PAGO" type="text" READONLY value="<%=dtmFechaCaja%>" size="10" maxlength="10"></td>
    <td><input name="TX_COMPINGRESO" type="text" READONLY value="<%=intNroComp%>" size="10" maxlength="10"></td>
	<td><input name="TX_BOLETA" type="text" value="<%=strNroBoleta%>" size="10" maxlength="7"></td>
	<td>
		<select name="CB_TIPOPAGO">
			<%
			ssql="SELECT * FROM CAJA_TIPO_PAGO"

			If Trim(intTipoPago)="CO" or Trim(intTipoPago)="CC" Then
				ssql = ssql & " WHERE ID_TIPO_PAGO ='CO'"
			ElseIf Trim(intTipoPago)="RP" Then
				ssql = ssql & " WHERE ID_TIPO_PAGO = 'RP'"
			Else

				If Trim(strCOD_CLIENTE) = "1000" Then
					ssql = ssql & " WHERE ID_TIPO_PAGO <> 'CO' AND ID_TIPO_PAGO <> 'AB'"
				Else
					ssql = ssql & "  WHERE  ID_TIPO_PAGO IN( 'AB','PTC','PTE')"
				End If

                ' ''response.write (ssql)

			%>
				<option value="">SELECCIONE</option>
			<%
			End If
			
            set rsCLI=Conn.execute(ssql)

			if not rsCLI.eof then
				do until rsCLI.eof
				%>
				<option value="<%=rsCLI("ID_TIPO_PAGO")%>"
				<%if Trim(intTipoPago)=Trim(rsCLI("ID_TIPO_PAGO")) then Response.Write("SELECTED") end if%> WIDTH="10"><%=ucase(rsCLI("ID_TIPO_PAGO") & " - " & rsCLI("DESC_TIPO_PAGO"))%></option>

				<%rsCLI.movenext
				loop
			end if
			rsCLI.close
			set rsCLI=nothing
			%>
		</select>
	</td>
    </tr>
</table>



<table width="90%" border="0" ALIGN="CENTER">
 <tr>
 	<td>
 		<font class="subtitulo_informe">> Detalle de Deuda</font>
 	</td>
</tr>
</table>


	<% If Trim(id_convenio) = "" Then %>

	<table width="90%" border="0" ALIGN="CENTER" >
				<tr class="Estilo34">
					<td colspan=15 align="LEFT">
					<a href="#" onClick= "marcar_boxes(true);">Marcar todos</a>&nbsp;&nbsp;&nbsp;
					<a href="#" onClick="desmarcar_boxes(true);">Desmarcar todos</a>
					</td>
				</tr>
             

	    <td valign="top">
		<%
		If Trim(rut) <> "" then
		abrirscg()
			strSql="SELECT dbo." & strNomFormInt & "(ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(ID_CUOTA) as HONORARIOS, DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, RUT_DEUDOR, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC, IsNull(datediff(d,FECHA_VENC,getdate()),0) as ANTIGUEDAD, NRO_DOC, IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,IsNull(USUARIO_ASIG,0) as USUARIO_ASIG, NRO_CUOTA, IsNull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS, SUCURSAL , ESTADO_DEUDA, COD_REMESA, CUENTA, NRO_DOC, TIPO_DOCUMENTO, TIPO_DOCUMENTO.NOM_TIPO_DOCUMENTO AS NOMDOCCUMENTO, ID_CUOTA FROM CUOTA LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO WHERE RUT_DEUDOR='"& rut &"' AND COD_CLIENTE='"& strCOD_CLIENTE &"' AND SALDO > 0 AND ESTADO_DEUDA IN ('1','6') ORDER BY CUENTA, FECHA_VENC , NRO_CUOTA "
			''response.Write(strSql)
			'response.End()
			set rsDET=Conn.execute(strSql)
			if not rsDET.eof then
			%>
			  <table style="width:100%;"class="intercalado" id="tbl_Procesa">
			  	<thead>
		        <tr>		        
		          <td >&nbsp;</td>
                  <!--Valor usado para saber cual chekbox esta selecionado esta oculto-->
                  <td class='hiddencol'>ID_CUOTA</td>
		          <td class="Estilo27" >CUENTA</td>
		          <td class="Estilo27" >NRO. DOC</td>
		          <td class="Estilo27" >CUOTA</td>
		          <td class="Estilo27" >F.VENCIM.</td>
		          <td class="Estilo27" >ANTIG.</td>
		          <td class="Estilo27" >TIPO DOC</td>
		          <td class="Estilo27" >ASIG.</td>
                  <td class="Estilo27" align="center">CAPITAL</td>
		          <td class="Estilo27" align="center">INTERES</td>
		          <td class="Estilo27" align="center">HONORARIOS</td>
		          <td class="Estilo27" align="center">PROTESTOS</td>
                  <td class="Estilo27" align="center">SALDO</td>
		        </tr>
		    	</thead>
		    	<tbody>
				<%
				intSaldo = 0
				intValorCuota = 0
				total_ValorCuota = 0
				strArrID_CUOTA=""
				strArrConcepto = ""
				strArrID_CUOTA = ""

				do until rsDET.eof
				'intSaldo = ValNulo(rsDET("SALDO"),"N")
				'intValorCuota = ValNulo(rsDET("VALOR_CUOTA"),"N")


				intSaldo = Round(session("valor_moneda") * ValNulo(rsDET("SALDO"),"N"),0)
				'intSaldo = Round(session("valor_moneda") * ValNulo(rsDET("VALOR_CUOTA"),"N"),0)
                ''response.Write("asdasd" & ValNulo(rsDET("SALDO"),"N")  & "<br>")
                

				strNroDoc = Trim(rsDET("ID_CUOTA"))
				strNroCuota = Trim(rsDET("NRO_CUOTA"))
				strSucursal = Trim(rsDET("SUCURSAL"))
				strEstadoDeuda = Trim(rsDET("ESTADO_DEUDA"))
				strCodRemesa = Trim(rsDET("COD_REMESA"))

				intDiasMora = rsDET("ANTIGUEDAD")

				intGastosProtestos = rsDET("GASTOS_PROTESTOS")

				intHonorarios = rsDET("HONORARIOS")
				intInteresSaldo = rsDET("INTERESES")

                intSaldoTotal = intSaldo +  intInteresSaldo + intHonorarios + intGastosProtestos
				intHonorarios = Round(intHonorarios,0)

				strArrConcepto = strArrConcepto & ";" & "CH_" & rsDET("ID_CUOTA")
				strArrID_CUOTA = strArrID_CUOTA & ";" & rsDET("ID_CUOTA")

				%>

		        <tr>

		          <TD><INPUT TYPE=checkbox id="CH_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" name="CH_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>"  onChange ="suma_capital(this,TX_SALDO_<%=Replace(Replace(rsDET("ID_CUOTA"),"-","_"),".","")%>.value,TX_INTERES_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>.value,TX_HONORARIOS_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>.value,TX_GPROT_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>.value);suma_total_general(0);"></TD>
                  <!--Valor usado para saber cual chekbox esta selecionado esta oculto-->
                  <td  class='hiddencol'><%=rsDET("ID_CUOTA")%></td>
		          <td><div align="right"><%=rsDET("CUENTA")%></div></td>
		          <td><div align="right"><%=rsDET("NRO_DOC")%></div></td>
		          <td><div align="center"><%=rsDET("NRO_CUOTA")%></div></td>
		          <td><div align="right"><%=rsDET("FECHA_VENC")%></div></td>
		          <td><div align="right"><%=rsDET("ANTIGUEDAD")%></div></td>
		          <td><div align="center"><%=rsDET("NOMDOCCUMENTO")%></div></td>
		          <td><div align="center"><%=rsDET("COD_REMESA")%></div></td>
                  <td align="right" ><input name="TX_SALDO_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=FormatNumber(intSaldo,0)%>" size="10" maxlength="10" align="RIGHT" readonly="readonly"></td>
		          <td align="right" ><input name="TX_INTERES_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=FormatNumber(intInteresSaldo,0)%>"  size="10" maxlength="10" align="RIGHT" readonly="readonly"></td>
		          <td align="right" ><input name="TX_HONORARIOS_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=FormatNumber(intHonorarios,0)%>" size="10" maxlength="10" align="RIGHT" onblur="FormatearObjeto(this)" readonly="readonly"></td>
		          <td align="right" ><input name="TX_GPROT_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=FormatNumber(intGastosProtestos,0)%>" size="10" maxlength="10" align="RIGHT" onblur="FormatearObjeto(this)" readonly="readonly"></td>
                  <td align="right" ><input name="TX_SALDOTOTAL_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=FormatNumber(intSaldoTotal,0)%>" size="10" maxlength="10" align="RIGHT" readonly="readonly"></td>
				
				 <%
					total_ValorCuota = total_ValorCuota + intValorCuota
					total_docs = total_docs + 1
				 %>
				 </tr>
				 <%rsDET.movenext
				 loop

				vArrConcepto = split(strArrConcepto,";")
				vArrID_CUOTA = split(strArrID_CUOTA,";")

				intTamvConcepto = ubound(vArrConcepto)

				 %>
				</tbody>
		      </table>
			  <%end if
			  rsDET.close
			  set rsDET=nothing
		  Else
		  %>
			<table width="100%" >
			<tr >
			<td align="center" class="estilo_columna_individual">

			Deudor no posee documentos pendientes
			</td>
			</tr>
			</table>
		  <%end if%>
	    </td>
	  </tr>

	</table>

				<% Else %>

					<table width="90%" border="0" ALIGN="CENTER">
					  <tr>
					    <td valign="top">
						<%
						If Trim(rut) <> "" then
						abrirscg()
							If Trim(intTipoPago)="CO" OR Trim(intTipoPago)="CC" Then
								strSql="SELECT ID_CONVENIO, CUOTA, TOTAL_CUOTA, CONVERT(VARCHAR(10),getdate(),103) as FECHAACTUAL ,  CONVERT(VARCHAR(10),FECHA_PAGO,103) as FECHA_PAGO, IsNull(datediff(d,FECHA_PAGO,getdate()),0) as ANTIGUEDAD FROM CONVENIO_DET WHERE ID_CONVENIO = " & id_convenio & " AND PAGADA IS NULL AND ID_CONVENIO IN (SELECT ID_CONVENIO FROM CONVENIO_ENC WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "')"
							Else

							End If
							'response.Write(strSql)
							'response.End()
							set rsDetConvenio=Conn.execute(strSql)
							if not rsDetConvenio.eof then
							%>
							  <table width="100%" class="estilo_columnas" id="tbl_Procesa">
							  	<thead>
						        <tr >
									<td>&nbsp;</td>
                                    <!--Valor usado para saber cual chekbox esta selecionado-->
									<td class='hiddencol'>ID_CUOTA</td>
                                    <td class="Estilo27"><%=strTitCol%></td>
									<td class="Estilo27">CUOTA</td>
									<td class="Estilo27">ANTIG.</td>
									<td class="Estilo27">F.VENCIM.</td>
									<td class="Estilo27" align="center">INTERESES</td>
									<td class="Estilo27" align="center">GASTOS</td>
									<td class="Estilo27" align="center">HONORARIOS</td>
									<td class="Estilo27" align="center">G.ADMIN</td>
                                    <td class="Estilo27" align="center">VALOR CUOTA</td>
						        </tr>
						    </thead>
						    <tbody>
								<%
								intSaldo = 0
								intValorCuota = 0
								total_ValorCuota = 0
								do until rsDetConvenio.eof
								intSaldo = ValNulo(rsDetConvenio("TOTAL_CUOTA"),"N")


								strFechaVencim = rsDetConvenio("FECHA_PAGO")
								strFechaActual = rsDetConvenio("FECHAACTUAL")

								intDiasMora = rsDetConvenio("ANTIGUEDAD")


								strFechaVencim = Mid(strFechaVencim,7,4) & "/" & Mid(strFechaVencim,1,2)  & "/" & Mid(strFechaVencim,4,2)
								strFechaActual = Mid(strFechaActual,7,4) & "/" & Mid(strFechaActual,1,2)  & "/" & Mid(strFechaActual,4,2)

								intInteresesDoc = 0

								intHonorariosDoc = 0

								intGastosDoc = 0

								intGastosAdmin = 0

								intInteresCuota = InteresCuota(intDiasMora,intMaximaAnualMora/100,intSaldo)

								'Response.write "<br>intMaximaAnualMora=" & intMaximaAnualMora

								If intInteresCuota < 0 Then
									intInteresCuota = 0
								End If


								%>
						        <tr bordercolor="#999999" >

									<TD><INPUT TYPE=checkbox id="CH_<%=rsDetConvenio("CUOTA")%>"  name="CH_<%=rsDetConvenio("CUOTA")%>"  onChange="suma_capital_2(this,TX_SALDO_<%=rsDetConvenio("CUOTA")%>.value,TX_INTERESES_<%=rsDetConvenio("CUOTA")%>.value,TX_GASTOS_<%=rsDetConvenio("CUOTA")%>.value,TX_HONORARIOS_<%=rsDetConvenio("CUOTA")%>.value, 0);suma_total_general(1);";></TD>
                                    <td  class='hiddencol'><%=rsDetConvenio("CUOTA")%></td>
                                    <td><%=rsDetConvenio("ID_CONVENIO")%></td>
									<td><%=rsDetConvenio("CUOTA")%></td>
									<td><%=intDiasMora%></td>
									<td><%=rsDetConvenio("FECHA_PAGO")%></td>
									<td align="right" ><input READONLY name="TX_INTERESES_<%=rsDetConvenio("CUOTA")%>" type="text" value="<%=FormatNumber(intInteresCuota,0)%>" size="8" maxlength="8" align="RIGHT"></td>
									<td align="right" ><input READONLY name="TX_GASTOS_<%=rsDetConvenio("CUOTA")%>" type="text" value="<%=FormatNumber(intGastosDoc,0)%>" size="8" maxlength="8" align="RIGHT"></td>
									<td align="right" ><input READONLY name="TX_HONORARIOS_<%=rsDetConvenio("CUOTA")%>" type="text" value="<%=FormatNumber(intHonorariosDoc,0)%>" size="8" maxlength="8" align="RIGHT"></td>
									<td align="right" ><input READONLY name="TX_GASTOS_ADMIN_<%=rsDetConvenio("CUOTA")%>" type="text" value="<%=FormatNumber(intGastosAdmin,0)%>" size="8" maxlength="8" align="RIGHT"></td>
                                    <td align="right" ><input READONLY="readonly" name="TX_SALDO_<%=rsDetConvenio("CUOTA")%>" type="text" value="<%=FormatNumber(intSaldo,0)%>" size="8" maxlength="8" align="RIGHT"></td>
						         <%
									total_ValorCuota = total_ValorCuota + intValorCuota
									''total_gc = total_gc + clng(rsDetConvenio("TOTAL_CUOTA"))
									total_gc = total_gc + 0
									total_docs = total_docs + 1
								 %>
								 </tr>
								 <%rsDetConvenio.movenext
								 loop
								 %>
								</tbody> 
						      </table>
							  <%end if
							  rsDetConvenio.close
							  set rsDetConvenio=nothing
						  Else
						  %>
							<table width="100%" border="0" bordercolor="#FFFFFF">
							<tr>
							<td align="center" class="estilo_columna_individual">

							Deudor no posee cuotas de convenio pendientes
							</td>
							</tr>
							</table>
						  <%end if%>
					    </td>
					  </tr>

				</table>
		<% End if %>


	<table width="90%" align="center" border="0" class="estilo_columnas" >
	<thead>
		<tr >
			<td><span class="Estilo27">Capital</span></td>
			<td><span class="Estilo27">Intereses</span></td>
			<td><span class="Estilo27">Honorarios</span></td>
			<td><span class="Estilo27">Gastos Prot.</span></td>
			<td><span class="Estilo27">Indem. Compensat.</span></td>
			<td><span class="Estilo27">Gastos Ope.</span></td>
			<td><span class="Estilo27">Gastos Adm.</span></td>
			<td><span class="Estilo27">Descuentos</span></td>
			<td><span class="Estilo27">Total Deuda</span></td>
	    </tr>
	</thead>
	<tr>
		<td><input class="TX_DEUDACAPITAL" name="TX_DEUDACAPITAL" type="text" value="" size="10" maxlength="10" onchange="solonumero(TX_DEUDACAPITAL);suma_total_general(0);"></td>
		<td><input class="TX_INTERESES"  id="TX_INTERESES" name="TX_INTERESES" type="text" value="<%=intInteresesD%>" size="10" maxlength="10" onblur="solonumero(TX_INTERESES);suma_total_general(1);"></td>
		<td><input class="TX_HONORARIOS" id="TX_HONORARIOS" name="TX_HONORARIOS" type="text" value="<%=intHonorariosD%>" size="10" maxlength="10" onblur="solonumero(TX_HONORARIOS);suma_total_general(2);"></td>
		<td><input class="TX_OTROS"      id="TX_OTROS" name="TX_OTROS" type="text" value="<%=intOtrosD%>" size="10" maxlength="10" onblur="solonumero(TX_OTROS);suma_total_general(1);"></td>
		<td><input class="TX_INDCOM"     id="TX_INDCOM" name="TX_INDCOM" type="text" value="<%=intIndemCompensatoriaD%>" size="10" maxlength="10" onblur="solonumero(TX_INDCOM);suma_total_general(1);"></td>
		<td><input class="TX_GASTOSJUD"  id="TX_GASTOSJUD" name="TX_GASTOSJUD" type="text" value="<%=intGastosJudicialesD%>" size="10" maxlength="10" onblur="solonumero(TX_GASTOSJUD);suma_total_general(1);"></td>
		<td><input class="TX_GASTOSADMIN"id="TX_GASTOSADMIN" name="TX_GASTOSADMIN" type="text" value="<%=intGastosAdminD%>" size="10" maxlength="10" onblur="solonumero(TX_GASTOSADMIN);suma_total_general(1);"></td>
		<td><input class="TX_DESCUENTO"  id="TX_DESCUENTO" name="TX_DESCUENTO" type="text" value="" size="10" maxlength="10" onblur="solonumero(TX_DESCUENTO);suma_total_general(1);"></td>
		<td><input name="TX_TOTALGRAL" type="text" value="" size="10" maxlength="10" onblur="solonumero(TX_TOTALGRAL);suma_total_general(0);" readonly="readonly"></td>
        
	  </tr>
	</table>

	<table width="90%" align="center">
	<tr>
		<td class="">OBSERVACIONES: </td>
	<td>
		<INPUT TYPE="TEXT" NAME="TX_OBSERVACIONES" size="100">
	</td>

	<TD <%=strCierreHab%>>
		<INPUT TYPE="BUTTON" <%=strCierreHab%> class="fondo_boton_100" NAME="Guardar" value="Guardar" onClick="envia('<%=perfil%>');" class="Estilo8" >
	</TD>
	</tr>
	</table>


	<div width="90%" style="margin-left:5%;" class="subtitulo_informe">> Detalle documentos a recaudar</div>


	<table width="90%" align="center" border="0" class="estilo_columnas">
		<thead>
	<tr >
		<td><span class="Estilo27">Destinatario</span></td>
		<td><span class="Estilo27">Forma Pago</span></td>
		<td><span class="Estilo27">Concepto</span></td>
        <td><span class="Estilo27">Monto</span></td>
        <td class="Estilo27">Rut Cheque</td>
        <td><span class="Estilo27">Fecha Venc</span></td>
		<td><span class="Estilo27">Banco</span></td>
		<!--td><span class="Estilo27">Plaza</span></td-->
		<td><span class="Estilo27">N° Cheque o Dep</span></td>
		<td><span class="Estilo27">Cta. Cte.</span></td>
		<td><span class="Estilo27">Retención</span></td>
		<TD></TD>
       </tr>
   </thead>
      <tr>
		<td>
			<select name="CB_DESTINO"  style="width:100px;" class="Estilo8" onChange="CargaConcepto(this.value,'<%=strCOD_CLIENTE%>');">
				<option value="">SELECCIONE</option>

			<%If strCOD_CLIENTE <> "1500" then%>
			
				<option value="1">LLACRUZ</option>	
				
			<%End If%>
			
				<option value="0">CLIENTE</option>
				</select>
			</td>
		<td>
			<select name="CB_FPAGO" width="10"  style="width:100px;" maxlength="10" class="Estilo8" onchange="FORMA_PAGO();">
			<option value="">SELECCIONE</option>
			<%
			ssql="SELECT * FROM CAJA_FORMA_PAGO WHERE ACTIVO=1"
			set rsCLI=Conn.execute(ssql)
			if not rsCLI.eof then
				do until rsCLI.eof
				%>
				<option value="<%=rsCLI("ID_FORMA_PAGO")%>"
				<%if Trim(cliente)=Trim(rsCLI("ID_FORMA_PAGO")) then
					response.Write("Selected")
				end if%>
				><%=ucase(rsCLI("DESC_FORMA_PAGO"))%></option>

				<%rsCLI.movenext
				loop
			end if
			rsCLI.close
			set rsCLI=nothing
			%>
			</select>
        </td>

		<td>
			<select id="CB_DIVIDE" name="CB_DIVIDE"  style="width:100px;" onChange="CargaMonto();" class="Estilo8">
			</select>
		</td>
		<td><input id="TX_MONTOCLI" class="TX_MONTOCLI" name="TX_MONTOCLI" type="text" value="" size="8" maxlength="10" class="Estilo8" onchange="solonumero(TX_MONTOCLI);FormatearObjeto(this)" ></td>
        <td><input name="TX_RUTCLI" type="text" value="" size="10" maxlength="10" class="Estilo8"></td>
	    <td><input name="inicio" type="text" id="inicio" value="" size="8" maxlength="10" class="Estilo8" onBlur="muestra_dia()"><a href="javascript:showCal('Calendar7');"><img src="../imagenes/calendario.gif" border="0"></a></td>
		<TD><SELECT name="CB_BANCO_CLIENTE"  style="width:100px;" class="Estilo8">
		<option value="">SELECCIONE</option>
		<%
		ssql="SELECT * FROM BANCOS"
		set rsBANC=Conn.execute(ssql)
		if not rsBANC.eof then
			do until rsBANC.eof
			%>
			<option value="<%=rsBANC("CODIGO")%>"
			<%if Trim(banco)=Trim(rsBANC("CODIGO")) then
				response.Write("Selected")
			end if%>
			><%=ucase(rsBANC("CODIGO") & " - " & Mid(rsBANC("NOMBRE_B"),1,12))%></option>

			<%rsBANC.movenext
			loop
		end if
		rsBANC.close
		set rsBANC=nothing

cerrarscg()
		%>
		</SELECT></TD>
		<td><input name="TX_NROCHEQUECLI" type="text" value="" size="15" maxlength="20" class="Estilo8"></td>
		<td><input name="TX_NROCTACTECLI" type="text" value="" size="12" maxlength="20" class="Estilo8"></td>
		<TD>
			<SELECT name="CB_HRETENIDO" class="Estilo8"  style="width:100px;">
			<option value="0">NORMAL</option>
			<option value="1">HON_RET</option>
			</SELECT>
		</TD>
		<td><input type="button" class="fondo_boton_100" name="ingdoc" value="OK" onClick="metedoccli();" class="Estilo8"></td>
        </tr>


	   <tr>

			<td>
				<select name="DESTINO" size="10" style="width:100px;" id="DESTINO"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select>
			</td>
			<td><select name="FPAGO" size="10" style="width:100px;" id="FPAGO"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>
			<td><select name="LS_DIVIDE" size="10" style="width:100px;" ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>
			<td><select name="MONTOCLI" size="10" style="width:100px;" id="MONTOCLI"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>
			<td><select name="RUTCLI" size="10" style="width:100px;" id="RUTCLI"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>
			<td><select name="FECHACLI" size="10" style="width:100px;" id="FECHACLI"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>
			<td><select name="BANCOCLI" size="10" style="width:100px;" id="BANCOCLI"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>
			<!--td><select name="PLAZACLI" size="10" id="PLAZACLI"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td-->
			<td><select name="NROCHECLI" size="10" style="width:100px;" id="NROCHECLI"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>
			<td><select name="NRCTACTECLI" size="10" style="width:100px;" id="NRCTACTECLI"  ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>
			<td><select name="LS_HRETENIDO" size="10" style="width:100px;" ondblClick=" borra_combos_cli(this.selectedIndex);" onChange="select_combos_cli(this.selectedIndex);" class="Estilo8"></select></td>

			<td></td>
		</tr>


		<tr>
			<td>
			<INPUT TYPE="hidden" NAME="TXPERIODO">
			<INPUT TYPE="hidden" NAME="TXCAPITAL">
			<INPUT TYPE="hidden" NAME="TXREAJUSTE">
			<INPUT TYPE="hidden" NAME="TXINTERES">
			<INPUT TYPE="hidden" NAME="TXGRAVAMENES">
			<INPUT TYPE="hidden" NAME="TXMULTAS">
			<INPUT TYPE="hidden" NAME="TXCARGOS">

			<INPUT TYPE="hidden" NAME="TXFPAGO">
			<INPUT TYPE="hidden" NAME="TXDESTINO">
			<INPUT TYPE="hidden" NAME="TXCUOTA">
			<INPUT TYPE="hidden" NAME="TXMONTOCLI">
			<INPUT TYPE="hidden" NAME="TXRUTCLI">
			<INPUT TYPE="hidden" NAME="TXFECVENCLI">
			<INPUT TYPE="hidden" NAME="TXBANCOCLIENTE">
			<INPUT TYPE="hidden" NAME="TXDIVIDE">
			<INPUT TYPE="hidden" NAME="TXHRETENIDO">

			<!--INPUT TYPE="hidden" NAME="TXPLAZACLIENTE"-->
			<INPUT TYPE="hidden" NAME="TXNROCHEQUECLI">
			<INPUT TYPE="hidden" NAME="TXNROCTACTECLI">
			<%hoy=date%>
			<INPUT TYPE="hidden" NAME="TXFECHAACTUAL" value="<%=hoy%>">
			<INPUT TYPE="hidden" NAME="dtmFechaCaja" value="<%=dtmFechaCaja%>">

			<INPUT TYPE="hidden" NAME="strGraba" value="">
			</td>
		</tr>
	</table>


</form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">

function FORMA_PAGO(){
	if ((datos.CB_FPAGO.value=='EF')||(datos.CB_FPAGO.value=='TC')||(datos.CB_FPAGO.value=='TD'))
	{
		datos.inicio.value='';
		datos.inicio.disabled = true;
		datos.TX_RUTCLI.value = '';
		datos.TX_RUTCLI.disabled = true;
		datos.CB_BANCO_CLIENTE.disabled = true;
		//datos.CB_DIVIDE.disabled = true;
		//datos.CB_PLAZA_CLIENTE.disabled = true;
		datos.TX_NROCHEQUECLI.value='';
		datos.TX_NROCHEQUECLI.disabled = true;
		datos.TX_NROCTACTECLI.value='';
		datos.TX_NROCTACTECLI.disabled = true;
	}else if((datos.CB_FPAGO.value=='AB')||(datos.CB_FPAGO.value=='CU')){
		datos.inicio.value=''
		datos.inicio.disabled = false;
		datos.TX_RUTCLI.value = '';
		datos.TX_RUTCLI.disabled = true;
		datos.CB_BANCO_CLIENTE.disabled = true;
		//datos.CB_DIVIDE.disabled = true;
		//datos.CB_PLAZA_CLIENTE.disabled = true;
			datos.TX_NROCHEQUECLI.disabled = false;
		datos.TX_NROCTACTECLI.value='';
		datos.TX_NROCTACTECLI.disabled = true;
	}else if (datos.CB_FPAGO.value=='DP'){
		datos.TX_RUTCLI.disabled = true;
		datos.inicio.disabled = true;
		datos.CB_BANCO_CLIENTE.disabled = false;
		//datos.CB_DIVIDE.disabled = true;
		//datos.CB_PLAZA_CLIENTE.disabled = false;
		datos.TX_NROCHEQUECLI.disabled = false;
		datos.TX_NROCTACTECLI.disabled = false;
		datos.inicio.value=''
	}
	else {
		datos.TX_RUTCLI.disabled = false;
		datos.inicio.disabled = false;
		datos.CB_BANCO_CLIENTE.disabled = false;
		//datos.CB_DIVIDE.disabled = false;
		//datos.CB_PLAZA_CLIENTE.disabled = false;
		datos.TX_NROCHEQUECLI.disabled = false;
		datos.TX_NROCTACTECLI.disabled = false;
		if (datos.CB_FPAGO.value=='CD'){
			datos.inicio.value=datos.TXFECHAACTUAL.value
		}else{
			datos.inicio.value=''
		}
	}
	function CargaMonto(){
   // alert(2)
            datos.TX_MONTOCLI.value = LimpiaNumeros(datos.TX_MONTOCLI.value)
            datos.TX_HONORARIOS.value =LimpiaNumeros(datos.TX_HONORARIOS.value)
            datos.TX_GASTOSADMIN.value =LimpiaNumeros(datos.TX_GASTOSADMIN.value)

        if (datos.CB_DIVIDE.value == '1') {
				datos.TX_MONTOCLI.value = datos.TX_HONORARIOS.value;
				//alert(datos.CB_DIVIDE.value);
			}
		else if (datos.CB_DIVIDE.value == '4') {
				datos.TX_MONTOCLI.value = datos.TX_GASTOSADMIN.value;
				//alert(datos.CB_DIVIDE.value);
			}
		else if (((datos.CB_FPAGO.value=='EF')||(datos.CB_FPAGO.value=='TC')||(datos.CB_FPAGO.value=='TD')||(datos.CB_FPAGO.value=='DP')||(datos.CB_FPAGO.value=='VV')) && ((datos.CB_DIVIDE.value != '1')||(datos.CB_DIVIDE.value != '4'))) {
			datos.TX_MONTOCLI.value = datos.TX_TOTALGRAL.value - datos.TX_HONORARIOS.value - datos.TX_GASTOSADMIN.value;
			}
        else if (datos.CB_FPAGO.value=='TR' )
            {
                datos.TX_MONTOCLI.value = datos.TX_TOTALGRAL.value - datos.TX_HONORARIOS.value - datos.TX_GASTOSADMIN.value;
            }
		else{
			datos.TX_MONTOCLI.value = ''
			}

            datos.TX_MONTOCLI.value =FormatearNumero(datos.TX_MONTOCLI.value)
            datos.TX_HONORARIOS.value =FormatearNumero(datos.TX_HONORARIOS.value)
            datos.TX_GASTOSADMIN.value =FormatearNumero(datos.TX_GASTOSADMIN.value)
	}
}

function CargaConcepto(destin,strCod_Cliente){

	var comboBox = document.getElementById('CB_DIVIDE');
	comboBox.options.length = 0;

	if (destin =='1')
		{
			var newOption = new Option('SELECCIONE','10');
			comboBox.options[comboBox.options.length] = newOption;

			var newOption = new Option('HONORARIO', '1');
			comboBox.options[comboBox.options.length] = newOption;

			var newOption = new Option('GASTOS ADM.', '4');
			comboBox.options[comboBox.options.length] = newOption;

		}
	else if (strCod_Cliente =='1500'){
		
			//alert(strCod_Cliente);
			var newOption = new Option('SELECCIONE','10');
			comboBox.options[comboBox.options.length] = newOption;
			
			var newOption = new Option('ABONO CLIENTE', '2');
			comboBox.options[comboBox.options.length] = newOption;
			}
			
	else if (destin =='0'){
		
			var newOption = new Option('SELECCIONE','10');
			comboBox.options[comboBox.options.length] = newOption;

			var newOption = new Option('ABONO CLIENTE', '2');
			comboBox.options[comboBox.options.length] = newOption;
				
			var newOption = new Option('SALDO CLIENTE', '3');
			comboBox.options[comboBox.options.length] = newOption;
			}
	
	else{
			var newOption = new Option('SELECCIONE','10');
			comboBox.options[comboBox.options.length] = newOption;
		}
}

function CargaMonto(){

        datos.TX_MONTOCLI.value = LimpiaNumeros(datos.TX_MONTOCLI.value)
        datos.TX_HONORARIOS.value =LimpiaNumeros(datos.TX_HONORARIOS.value)
        datos.TX_GASTOSADMIN.value =LimpiaNumeros(datos.TX_GASTOSADMIN.value)
        datos.TX_TOTALGRAL.value  =LimpiaNumeros(datos.TX_TOTALGRAL.value)

	if (datos.CB_DIVIDE.value == '1') {
			datos.TX_MONTOCLI.value = datos.TX_HONORARIOS.value;
			//alert(datos.CB_DIVIDE.value);
		}
	else if (datos.CB_DIVIDE.value == '4') {
			datos.TX_MONTOCLI.value = datos.TX_GASTOSADMIN.value;
			//alert(datos.CB_DIVIDE.value);
		}
	else if ((datos.CB_DIVIDE.value != '10') && ((datos.CB_FPAGO.value=='EF')||(datos.CB_FPAGO.value=='TC')||(datos.CB_FPAGO.value=='TD')||(datos.CB_FPAGO.value=='DP')||(datos.CB_FPAGO.value=='VV')) && ((datos.CB_DIVIDE.value != '1')||(datos.CB_DIVIDE.value != '4'))) {
		datos.TX_MONTOCLI.value = datos.TX_TOTALGRAL.value - datos.TX_HONORARIOS.value - datos.TX_GASTOSADMIN.value;
		}
        else if (datos.CB_FPAGO.value=='TR' )
        {
        datos.TX_MONTOCLI.value = datos.TX_TOTALGRAL.value - datos.TX_HONORARIOS.value - datos.TX_GASTOSADMIN.value;
        }
	else{
		datos.TX_MONTOCLI.value = ''
		}

            datos.TX_MONTOCLI.value =FormatearNumero(datos.TX_MONTOCLI.value)
            datos.TX_HONORARIOS.value =FormatearNumero(datos.TX_HONORARIOS.value)
            datos.TX_GASTOSADMIN.value =FormatearNumero(datos.TX_GASTOSADMIN.value)
            datos.TX_TOTALGRAL.value  =FormatearNumero(datos.TX_TOTALGRAL.value)

}

function solonumero(valor){
     //Compruebo si es un valor numÃ©rico
     valor.value =  Solo_Numerico(LimpiaNumeros(valor.value));

     if (valor.value.length ==0)
     {
     valor.value = "0";
     }
      if (isNaN(valor.value)) {
            //entonces (no es numero) devuelvo el valor cadena vacia
            valor.value="0";
			return valor.value
      }else{
            //En caso contrario (Si era un nÃºmero) devuelvo el valor
			valor.value
			return valor.value
      }
	  //}
}

function borra_combos_cli(indice){
	borra_opcion(datos.DESTINO,indice);
	borra_opcion(datos.FPAGO,indice);
	borra_opcion(datos.RUTCLI,indice);
	borra_opcion(datos.MONTOCLI,indice);
	borra_opcion(datos.FECHACLI,indice);
	borra_opcion(datos.BANCOCLI,indice);
	borra_opcion(datos.NROCHECLI,indice);
	borra_opcion(datos.NRCTACTECLI,indice);
	borra_opcion(datos.LS_DIVIDE,indice);
	borra_opcion(datos.LS_HRETENIDO,indice);

}

function borra_opcion(combo,indice){
	if (combo.options.length>0){
	//	combo.options[indice]=null;
		for (var e=indice; e< combo.options.length-1; e++) {
			//alert(e);
			combo.options[e].text=combo.options[e+1].text;
			combo.options[e].value=combo.options[e+1].value;
		}
		combo.options[combo.options.length-1]=null;
	}
}

function select_combos_cli(indice){
	datos.DESTINO.selectedIndex=indice;
	datos.FPAGO.selectedIndex=indice;
	datos.RUTCLI.selectedIndex=indice;
	datos.MONTOCLI.selectedIndex=indice;
	datos.FECHACLI.selectedIndex=indice;
	datos.BANCOCLI.selectedIndex=indice;
	//datos.PLAZACLI.selectedIndex=indice;
	datos.NROCHECLI.selectedIndex=indice;
	datos.NRCTACTECLI.selectedIndex=indice;
}

//-------------------------------

function metedoccli(){

	if (datos.CB_DESTINO.value==''){
		alert("Debe seleccionar el destino del pago");
		datos.CB_DESTINO.focus();
	}else if (datos.CB_FPAGO.value==''){
		alert("Debe seleccionar la forma de pago");
		datos.CB_FPAGO.focus();
	}else if(((datos.CB_FPAGO.value=='CD')||(datos.CB_FPAGO.value=='CF') ||(datos.CB_FPAGO.value=='VV'))&&((datos.TX_RUTCLI.value==''))){
		alert("Debe ingresar el Rut");
		datos.TX_RUTCLI.focus();
	}else if (datos.CB_DIVIDE.value=='10'){
		alert("Debe seleccionar el Concepto correspondiente");
		datos.CB_DIVIDE.focus();
	}else if (datos.TX_MONTOCLI.value==''){
		alert("Debe ingresar el Monto");
		datos.TX_MONTOCLI.focus();
	}else if(((datos.CB_FPAGO.value=='CD')||(datos.CB_FPAGO.value=='CF') ||(datos.CB_FPAGO.value=='VV'))&&((datos.inicio.value==''))){
		alert("Debe ingresar la fecha de vencimiento");
		datos.inicio.focus();
	}else if(((datos.CB_FPAGO.value=='DP')||(datos.CB_FPAGO.value=='CF')||(datos.CB_FPAGO.value=='CD')||(datos.CB_FPAGO.value=='TR')||(datos.CB_FPAGO.value=='VV'))&&((datos.CB_BANCO_CLIENTE.value==''))){
		alert("Debe ingresar el Banco al que pertenece el Cheque , Deposito o Transferencia");
		datos.CB_BANCO_CLIENTE.focus();
	}else if(((datos.CB_FPAGO.value=='CD')||(datos.CB_FPAGO.value=='CF')||(datos.CB_FPAGO.value=='VV') ||(datos.CB_FPAGO.value=='DP'))&&((datos.TX_NROCHEQUECLI.value==''))){
		alert("Debe ingresar el Número del cheque o Número Comprobante Deposito")
		datos.TX_NROCHEQUECLI.focus();
    }else if(((datos.CB_FPAGO.value=='CD')||(datos.CB_FPAGO.value=='CF')||(datos.CB_FPAGO.value=='VV'))&&((datos.TX_NROCTACTECLI.value==''))){
		alert("Debe ingresar el Número de la cuenta corriente")
		datos.TX_NROCTACTECLI.focus();
    }else if ((datos.TX_MONTOCLI.value=='')||  (datos.TX_MONTOCLI.value <='0')  ){
		alert("Monto a Recaudar no valido");
		datos.TX_MONTOCLI.focus();
	}else{

		datos.TX_RUTCLI.disabled = false;
		datos.TX_MONTOCLI.disabled = false;
		datos.inicio.disabled = false;
		datos.CB_BANCO_CLIENTE.disabled = false;
		datos.TX_NROCHEQUECLI.disabled = false;
		datos.TX_NROCTACTECLI.disabled = false;
		apilar_combo_combo(datos.CB_DESTINO, datos.DESTINO);
		apilar_combo_combo(datos.CB_FPAGO, datos.FPAGO);
		apilar_textbox_combo(datos.TX_RUTCLI, datos.RUTCLI);
		apilar_textbox_combo(datos.TX_MONTOCLI, datos.MONTOCLI);
		apilar_textbox_combo(datos.inicio, datos.FECHACLI);
		apilar_combo_combo(datos.CB_BANCO_CLIENTE, datos.BANCOCLI);
		//apilar_combo_combo(datos.CB_PLAZA_CLIENTE, datos.PLAZACLI);
		apilar_textbox_combo(datos.TX_NROCHEQUECLI, datos.NROCHECLI);
		apilar_textbox_combo(datos.TX_NROCTACTECLI, datos.NRCTACTECLI);
		apilar_combo_combo(datos.CB_DIVIDE, datos.LS_DIVIDE);
		apilar_combo_combo(datos.CB_HRETENIDO, datos.LS_HRETENIDO);

		if((datos.CB_FPAGO.value != 'CD') && (datos.CB_FPAGO.value  != 'CF') && (datos.CB_FPAGO.value  != 'VV')) {
			datos.CB_FPAGO.value="";
			datos.TX_RUTCLI.value="";
			datos.inicio.value="";
			datos.CB_BANCO_CLIENTE.value="";
			datos.TX_NROCHEQUECLI.value="";
			datos.TX_NROCTACTECLI.value="";
		}

		datos.CB_DIVIDE.value="10";
		datos.TX_MONTOCLI.value="";
		datos.CB_DESTINO.value="";
		//datos.CB_PLAZA_CLIENTE.value="";

		datos.CB_DESTINO.focus();
	}
}

function suma_capital(objeto , intValorSaldoCapital, intValorInteres, intValorHonorarios, intValorGProtestos){

    
     intValorSaldoCapital = LimpiaNumeros(intValorSaldoCapital)
     intValorInteres=  LimpiaNumeros(intValorInteres)
     intValorHonorarios = LimpiaNumeros(intValorHonorarios)
     intValorGProtestos = LimpiaNumeros(intValorGProtestos)
     datos.TX_DEUDACAPITAL.value = LimpiaNumeros(datos.TX_DEUDACAPITAL.value)
     datos.TX_HONORARIOS.value = LimpiaNumeros(datos.TX_HONORARIOS.value)
     datos.TX_INTERESES.value = LimpiaNumeros(datos.TX_INTERESES.value)
     datos.TX_OTROS.value = LimpiaNumeros(datos.TX_OTROS.value)


	if (objeto.checked == true) {
		datos.TX_DEUDACAPITAL.value =  (parseInt(datos.TX_DEUDACAPITAL.value) + parseInt(intValorSaldoCapital));
		datos.TX_HONORARIOS.value =  (parseInt(datos.TX_HONORARIOS.value) + parseInt(intValorHonorarios));
		datos.TX_INTERESES.value =  (parseInt(datos.TX_INTERESES.value) + parseInt(intValorInteres));
		datos.TX_OTROS.value =  (parseInt(datos.TX_OTROS.value) + parseInt(intValorGProtestos));
	}
	else
	{
		datos.TX_DEUDACAPITAL.value =  (eval(datos.TX_DEUDACAPITAL.value) - eval(intValorSaldoCapital));
		datos.TX_HONORARIOS.value = ( eval(datos.TX_HONORARIOS.value) - eval(intValorHonorarios));
		datos.TX_INTERESES.value =  (eval(datos.TX_INTERESES.value) - eval(intValorInteres));
		datos.TX_OTROS.value =  (eval(datos.TX_OTROS.value) - eval(intValorGProtestos));
	}
    
  

    if (datos.TX_DEUDACAPITAL.value <= 0)
    {
        datos.TX_DEUDACAPITAL.value =  0
    }
    if (datos.TX_HONORARIOS.value <= 0)
    {
        datos.TX_HONORARIOS.value =  0
    }
    if (datos.TX_INTERESES.value <= 0)
    {
        datos.TX_INTERESES.value =  0
    }
    if (datos.TX_OTROS.value <= 0)
    {
        datos.TX_OTROS.value =  0
    }

    
     datos.TX_DEUDACAPITAL.value = FormatearNumero(datos.TX_DEUDACAPITAL.value)
     datos.TX_HONORARIOS.value = FormatearNumero(datos.TX_HONORARIOS.value)
     datos.TX_INTERESES.value = FormatearNumero(datos.TX_INTERESES.value)
     datos.TX_OTROS.value = FormatearNumero(datos.TX_OTROS.value)


    /*alert(datos.TX_DEUDACAPITAL.value)*/

}

function suma_capital_2(objeto , intValorSaldoCapital, intValorIntereses, intValorGastos, intValorHonorarios, intValorGastosAdmin){
	
     intValorSaldoCapital = LimpiaNumeros(intValorSaldoCapital)
     intValorIntereses=  LimpiaNumeros(intValorIntereses)
     intValorGastos = LimpiaNumeros(intValorGastos)
     intValorHonorarios = LimpiaNumeros(intValorHonorarios)
     intValorGastosAdmin = LimpiaNumeros(intValorGastosAdmin)

     datos.TX_DEUDACAPITAL.value =  LimpiaNumeros(datos.TX_DEUDACAPITAL.value )
     datos.TX_INTERESES.value=  LimpiaNumeros(datos.TX_INTERESES.value )
     datos.TX_GASTOSJUD.value=  LimpiaNumeros(datos.TX_GASTOSJUD.value )
     datos.TX_GASTOSADMIN.value=  LimpiaNumeros(datos.TX_GASTOSADMIN.value )
     datos.TX_HONORARIOS.value=  LimpiaNumeros(datos.TX_HONORARIOS.value )

	if (objeto.checked == true) {
		datos.TX_DEUDACAPITAL.value = eval(datos.TX_DEUDACAPITAL.value) + eval(intValorSaldoCapital);
		datos.TX_INTERESES.value = eval(datos.TX_INTERESES.value) + eval(intValorIntereses);
		datos.TX_GASTOSJUD.value = eval(datos.TX_GASTOSJUD.value) + eval(intValorGastos);
		datos.TX_GASTOSADMIN.value = eval(datos.TX_GASTOSADMIN.value) + eval(intValorGastosAdmin);
		datos.TX_HONORARIOS.value = eval(datos.TX_HONORARIOS.value) + eval(intValorHonorarios);
	}
	else
	{
		datos.TX_DEUDACAPITAL.value = eval(datos.TX_DEUDACAPITAL.value) - eval(intValorSaldoCapital);
		datos.TX_INTERESES.value = eval(datos.TX_INTERESES.value) - eval(intValorIntereses);
		datos.TX_GASTOSJUD.value = eval(datos.TX_GASTOSJUD.value) - eval(intValorGastos);
		datos.TX_GASTOSADMIN.value = eval(datos.TX_GASTOSADMIN.value) - eval(intValorGastosAdmin);
		datos.TX_HONORARIOS.value = eval(datos.TX_HONORARIOS.value) - eval(intValorHonorarios);
	}
    
    if (datos.TX_DEUDACAPITAL.value <= 0)
    {
        datos.TX_DEUDACAPITAL.value =  0
    }
    if (datos.TX_INTERESES.value <= 0)
    {
        datos.TX_INTERESES.value =  0
    }
     if (datos.TX_GASTOSJUD.value <= 0)
    {
        datos.TX_GASTOSJUD.value =  0
    }
     if (datos.TX_GASTOSADMIN.value <= 0)
    {
        datos.TX_GASTOSADMIN.value =  0
    }
   
    if (datos.TX_HONORARIOS.value <= 0)
    {
        datos.TX_HONORARIOS.value =  0
    }
   
   
   
    
    
    
     datos.TX_DEUDACAPITAL.value =  FormatearNumero(datos.TX_DEUDACAPITAL.value )
     datos.TX_INTERESES.value=  FormatearNumero(datos.TX_INTERESES.value )
     datos.TX_GASTOSJUD.value=  FormatearNumero(datos.TX_GASTOSJUD.value )
     datos.TX_GASTOSADMIN.value=  FormatearNumero(datos.TX_GASTOSADMIN.value )
     datos.TX_HONORARIOS.value=  FormatearNumero(datos.TX_HONORARIOS.value )

     

}

function suma_total_general (origen){


var  TX_TOTALGRAL = datos.TX_TOTALGRAL.value 
var  TX_DESCUENTO = datos.TX_DESCUENTO.value
var  TX_DEUDACAPITAL= (datos.TX_DEUDACAPITAL.value);
var  TX_HONORARIOS=(datos.TX_HONORARIOS.value);
var  TX_INDCOM=(datos.TX_INDCOM.value);
var  TX_OTROS=(datos.TX_OTROS.value);
var  TX_INTERESES=(datos.TX_INTERESES.value);
var  TX_GASTOSJUD=(datos.TX_GASTOSJUD.value);
var  TX_GASTOSADMIN=(datos.TX_GASTOSADMIN.value);
var  TX_DESCUENTO=(datos.TX_DESCUENTO.value);

datos.TX_TOTALGRAL.value    = LimpiaNumeros(datos.TX_TOTALGRAL.value)
datos.TX_DESCUENTO.value   = LimpiaNumeros(datos.TX_DESCUENTO.value)
datos.TX_DEUDACAPITAL.value = LimpiaNumeros(datos.TX_DEUDACAPITAL.value);
datos.TX_HONORARIOS.value   = LimpiaNumeros(datos.TX_HONORARIOS.value);
datos.TX_INDCOM.value       = LimpiaNumeros(datos.TX_INDCOM.value);
datos.TX_OTROS.value        = LimpiaNumeros(datos.TX_OTROS.value);
datos.TX_INTERESES.value    = LimpiaNumeros(datos.TX_INTERESES.value);
datos.TX_GASTOSJUD.value    = LimpiaNumeros(datos.TX_GASTOSJUD.value);
datos.TX_GASTOSADMIN.value  = LimpiaNumeros(datos.TX_GASTOSADMIN.value);




    if (parseInt(datos.TX_TOTALGRAL.value) < parseInt(datos.TX_DESCUENTO.value)  ) 
    {
        alert("Descuento No Debe Ser Mayor a Total Deuda")
        return;
    }
 
 

	if (origen == 1) 
    {
		datos.TX_HONORARIOS.value = Math.round(datos.TX_HONORARIOS.value)
        datos.TX_TOTALGRAL.value =(parseInt(datos.TX_DEUDACAPITAL.value) + parseInt(datos.TX_HONORARIOS.value) + parseInt(datos.TX_INDCOM.value) + parseInt(datos.TX_OTROS.value) + parseInt(datos.TX_INTERESES.value) + parseInt(datos.TX_GASTOSJUD.value) + parseInt(datos.TX_GASTOSADMIN.value) - parseInt(datos.TX_DESCUENTO.value));
	}
	else if (origen == 2) 
    {
		datos.TX_TOTALGRAL.value = eval(datos.TX_DEUDACAPITAL.value) + eval(datos.TX_HONORARIOS.value) + eval(datos.TX_INDCOM.value) + eval(datos.TX_OTROS.value) + eval(datos.TX_INTERESES.value) + eval(datos.TX_GASTOSJUD.value) + eval(datos.TX_GASTOSADMIN.value) - eval(datos.TX_DESCUENTO.value);
	}
	else
	{
		datos.TX_TOTALGRAL.value = eval(datos.TX_DEUDACAPITAL.value) + eval(datos.TX_HONORARIOS.value) + eval(datos.TX_INDCOM.value) + eval(datos.TX_OTROS.value) + eval(datos.TX_INTERESES.value) + eval(datos.TX_GASTOSJUD.value) + eval(datos.TX_GASTOSADMIN.value) - eval(datos.TX_DESCUENTO.value);
	}

         
    datos.TX_DEUDACAPITAL.value= FormatearNumero(datos.TX_DEUDACAPITAL.value);
    datos.TX_HONORARIOS.value=FormatearNumero(datos.TX_HONORARIOS.value);
    datos.TX_INDCOM.value=FormatearNumero(datos.TX_INDCOM.value);
    datos.TX_OTROS.value=FormatearNumero(datos.TX_OTROS.value);
    datos.TX_INTERESES.value=FormatearNumero(datos.TX_INTERESES.value);
    datos.TX_GASTOSJUD.value=FormatearNumero(datos.TX_GASTOSJUD.value);
    datos.TX_GASTOSADMIN.value=FormatearNumero(datos.TX_GASTOSADMIN.value);
    datos.TX_DESCUENTO.value=FormatearNumero(datos.TX_DESCUENTO.value);
    datos.TX_TOTALGRAL.value = FormatearNumero(datos.TX_TOTALGRAL.value)
    
    
}


function apilar_textbox_combo(origen, destino){
//addNew(document.myForm.proceso.options[document.myForm.proceso.selectedIndex].value)
	// Add a new option.
	var ok=false;
	i=destino.length;
	//valor = datos.txt_clavedoc.value.length;
	//alert(valor);
	valor=origen.value.length ;
	valor2=origen.value;
	if (valor>=0){
		texto=origen.value;
		if (texto==''){
		texto='';
		valor2 = '';
		}
	}else{
	texto='';
	valor2='';
	}
	var el = new Option(texto,valor2);
			destino.options[i] = el;
		//alert("ingrese un valor para agregar.");
}
//------------------------------------------------------------------
function apilar_combo_combo(origen, destino){
	// Add a new option.
	var ok=false;
	i=destino.length;
	valor=origen.selectedIndex ;
	valor2=origen.options[valor].value;
	if (valor>=0){
		texto=origen.options[valor].text;
		if (texto=='SELECCIONAR' || texto=='0'){
			texto='';
			valor2='';
		}
	}else{
	texto='';
	valor2='';
	}
	var el = new Option(texto,valor2);
	destino.options[i] = el;
		//alert("Seleccione un valor para agregar.");
}
//////--------------------------------------------------------------------
function disa(){
		datos.Guardar.disabled = true;
}
function habilita(){
		datos.Guardar.disabled = false;
}
function envia(perfilusuario){
	disa()

    var cont = 0
    
     for (var i = 1; i < document.getElementById('tbl_Procesa').rows.length; i++) {
              var Id_Cuota = document.getElementById('tbl_Procesa').rows[i].cells[1].innerHTML;
                   var  chk = document.getElementById("CH_" + Id_Cuota).checked;
                    
	             if (chk == true) {
                      cont = cont + 1
                      }
	           
    }

    if (cont==0)
    {
    alert("Indique Cuota a Cancelar")
    habilita();
    return false;
    }



   if(datos.TX_RUT.value==''){
		alert("Debe ingresar el rut")
		datos.TX_RUT.focus();
		habilita();
	}else if(datos.CB_CLIENTE.value == '0'){
		alert("Debe seleccionar el Cliente");
		habilita();
	}else if (datos.TX_BOLETA.value == ''){
		alert("Debe ingresar Número de Boleta");
		habilita();
	}else if (datos.CB_TIPOPAGO.value == ''){
		alert("Debe seleccionar el Tipo de Pago");
		habilita();
	}else{
		i=datos.DESTINO.length;
		montcli=0
		montemp=0
		monttotal=0
		cli=0
		emp=0
		for (var e=0; e<i;e++){
			if(datos.DESTINO.options[e].value=='0'){
					montcli= eval(LimpiaNumeros(montcli)) + eval(LimpiaNumeros(datos.MONTOCLI.options[e].value));
			}else{
					montemp = eval(LimpiaNumeros(montemp)) + eval(LimpiaNumeros(datos.MONTOCLI.options[e].value));
			}
			monttotal = eval(LimpiaNumeros(monttotal)) + eval(LimpiaNumeros(datos.MONTOCLI.options[e].value));

		}

     

        datos.TX_TOTALGRAL.value = LimpiaNumeros(datos.TX_TOTALGRAL.value)
        
        var monttotal2 = LimpiaNumeros(monttotal)
     
        if (datos.TX_TOTALGRAL.value <=0)
        {
            alert("Total Deuda No valido");
            habilita();
            return;
        } 

        
        if (eval(datos.TX_TOTALGRAL.value) != eval(monttotal2)) {
			alert("Los montos ingresados en el detalle de documentos no son correctos : Total General :" + eval(datos.TX_TOTALGRAL.value) + " , Total Detalle = " + eval(monttotal2));
            datos.TX_TOTALGRAL.value = FormatearNumero(datos.TX_TOTALGRAL.value )
            habilita();
		}else{

        datos.TX_TOTALGRAL.value = FormatearNumero(datos.TX_TOTALGRAL.value )
               
           $.prettyLoader.show();

			JuntaDetalleCliente();
			datos.strGraba.value='SI';
			disa();
			datos.submit();
		}
	}
}

function chkFecha(f) {
  str = f.value
  if (str.length<10){
  	alert("Error - IngresÃ³ una fecha no vÃ¡lida");
  	f.value=''
	f.focus();
  //	f.select();
  }else{
	if ( !formatoFecha(str) ) {
		alert("Debe indicar la Fecha en formato DD/MM/AAAA. Ejemplo: 'Para 20 de Diciembre de 2009 se debe ingresar 20/12/2009'");
    //f.select()
		f.value=''
		f.focus()
		return false
	}
	if ( !validarFecha(str) ) {
    // Los mensajes de error estÃ¡n dentro de validarFecha.
    //f.select()
   f.value=''
	f.focus()
    return false
  }
  }

  // validacion de la fecha


  return true
}

//-----------------------------------------------------------
  function validarFecha(str_fecha){

  var sl1=str_fecha.indexOf("/")
  var sl2=str_fecha.lastIndexOf("/")
  var inday = parseFloat(str_fecha.substring(0,sl1))
  var inmonth = parseFloat(str_fecha.substring(sl1+1,sl2))
  var inyear = parseFloat(str_fecha.substring(sl2+1,str_fecha.length))

  //alert("day:" + inday + ", mes:" + inmonth + ", agno: " + inyear)

  if (inmonth < 1 || inmonth > 12) {
    alert("Mes invÃ¡lido en la fecha");
    return false;
  }
  if (inday < 1 || inday > diasEnMes(inmonth, inyear)) {
    alert("DÃ­a invÃ¡lido en la fecha");
    return false;
  }

  return true
}


//------------------------------------------------------------------

function formatoFecha(str) {
  var sl1, sl2, ui, ddstr, mmstr, aaaastr;

  // El formato debe ser d/m/aaaa, d/mm/aaaa, dd/m/aaaa, dd/mm/aaaa,
  // Las posiciones son a partir de 0
  if (str.length < 8 &&  str.length > 10)    // el tamagno es fijo de 8, 9 o 10
    return false


  sl1=str.indexOf("/")
  if (sl1 < 1 && sl1 > 2 )    // el primer slash debe estar en la 1 o 2
    return false

  sl2=str.lastIndexOf("/")
  if (sl2 < 3 &&  sl2 > 5)    // el Ãºltimo slash debe estar en la 3, 4 o 5
    return false

  ddstr = str.substring(0,sl1)
  mmstr = str.substring(sl1+1,sl2)
  aaaastr = str.substring(sl2+1,str.length)

  if ( !sonDigitos(ddstr) || !sonDigitos(mmstr) || !sonDigitos(aaaastr) )
    return false

  return true
}
function sonDigitos(str) {
  var l, car

  l = str.length
  if ( l<1 )
    return false

  for ( i=0; i<l; i++) {
    car = str.substring(i,i+1)
    if ( "0" <= car &&  car <= "9" )
      continue
    else
      return false
  }
  return true
}

function diasEnMes (month, year)
{
  if (month == 1 || month == 3 || month == 5 || month == 7 || month == 8 || month == 10 || month == 12)
    return 31;
  else if (month == 2)
    // February has 29 days in any year evenly divisible by four,
      // EXCEPT for centurial years which are not also divisible by 400.
      return (  ((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0) ) ) ? 29 : 28 );
  else if (month == 4 || month == 6 || month == 9 || month == 11)
    return 30;
  // En caso contrario:
  alert("diasEnMes: Mes invÃ¡lido");
  return -1;
}
function Valida_Rut(Vrut)
{
	var dig
	Vrut = Vrut.split("-");

	if (!isNaN(Vrut[0]))
	{
		largo_rut = Vrut[0].length;
		if ((largo_rut >= 7 ) && (largo_rut <= 8))
		{
			if (largo_rut > 7)
			{
				multiplicador = 3;
			}
			else
			{
				multiplicador = 2;
			}
			suma = 0;
			contador = 0;
				do
				{
					digito = Vrut[0].charAt(contador);
					digito = Number(digito);
						if (multiplicador == 1)
						{
							multiplicador = 7;
						}

					suma = suma + (digito * multiplicador);
					multiplicador --;
					contador ++;
				}
				while (contador < largo_rut);
			resto = suma % 11
			dig_verificador = 11 - resto;

				if (dig_verificador == 10)
				{
					dig = "k";
				}
				else if (dig_verificador == 11)
				{
					dig = 0
				}
				else
				{
					dig = dig_verificador;
				}

				if (dig != Vrut[1])
				{
					alert ("El Rut es invalido !");
					datos.TX_RUTCLI.value="";
					datos.TX_RUTCLI.focus();
					//return 0;
				}
		}
		else
		{
			datos.TX_RUTCLI.value="";
			datos.TX_RUTCLI.focus();
			alert("El Rut es invalido ! ");

			//return 0;
		}
	}
	else
	{
		alert("El Rut es invalido ! ");
		datos.TX_RUTCLI.value="";
		datos.TX_RUTCLI.focus();
		//return 0;
	}
		//return 1;
}


function muestra_dia(){
//alert(getCurrentDate())
//alert("hola")
	var diferencia=DiferenciaFechas(datos.inicio.value)
	//alert(diferencia)
	if(datos.inicio.value!=''){
		if ((diferencia>=-30)) {
			//alert('Ok')
		}else{
			alert('la fecha de compromiso debe ser mayor a la \nfecha actual')
			datos.inicio.value=''
			datos.inicio.focus()
		}
	}
}


function DiferenciaFechas (CadenaFecha1) {
   var fecha_hoy = getCurrentDate() //hoy


   //Obtiene dia, mes y aÃ±o
   var fecha1 = new fecha( CadenaFecha1 )
   var fecha2 = new fecha(fecha_hoy)

   //Obtiene objetos Date
   var miFecha1 = new Date( fecha1.anio, fecha1.mes, fecha1.dia )
   var miFecha2 = new Date( fecha2.anio, fecha2.mes, fecha2.dia )

   //Resta fechas y redondea
   var diferencia = miFecha1.getTime() - miFecha2.getTime()
   var dias = Math.floor(diferencia / (1000 * 60 * 60 * 24))
   var segundos = Math.floor(diferencia / 1000)
   //alert ('La diferencia es de ' + dias + ' dias,\no ' + segundos + ' segundos.')

   return dias //false
}
//---------------------------------------------------------------------
function fecha( cadena ) {

   //Separador para la introduccion de las fechas
   var separador = "/"

   //Separa por dia, mes y aÃ±o
   if ( cadena.indexOf( separador ) != -1 ) {
        var POSI_1 = 0
        var POSI_2 = cadena.indexOf( separador, POSI_1 + 1 )
        var POSI_3 = cadena.indexOf( separador, POSI_2 + 1 )
        this.dia = cadena.substring( POSI_1, POSI_2 )
        this.mes = cadena.substring( POSI_2 + 1, POSI_3 )
        this.anio = cadena.substring( POSI_3 + 1, cadena.length )
   } else {
        this.dia = 0
        this.mes = 0
        this.anio = 0
   }
}
///----------------------------------------


function JuntaDetalleCliente(){
	datos.TXDESTINO.value = ""
	datos.TXFPAGO.value = ""
	datos.TXRUTCLI.value=""
	datos.TXMONTOCLI.value=""
	datos.TXFECVENCLI.value	=""
	datos.TXBANCOCLIENTE.value=""
	datos.TXDIVIDE.value=""
	datos.TXHRETENIDO.value=""

	//datos.TXPLAZACLIENTE.value=""
	datos.TXNROCHEQUECLI.value=""
	datos.TXNROCTACTECLI.value=""
	for (var e=0; e<datos.DESTINO.options.length;e++){
		if (e!=0) {
		//poner la coma
			datos.TXDESTINO.value=datos.TXDESTINO.value+"*";
			datos.TXFPAGO.value=datos.TXFPAGO.value+"*";
			datos.TXRUTCLI.value=datos.TXRUTCLI.value+"*"	;
			datos.TXMONTOCLI.value=datos.TXMONTOCLI.value+"*";
			datos.TXFECVENCLI.value	=datos.TXFECVENCLI.value+"*";
			datos.TXBANCOCLIENTE.value=datos.TXBANCOCLIENTE.value+"*";
			datos.TXDIVIDE.value=datos.TXDIVIDE.value+"*";
			//datos.TXPLAZACLIENTE.value=datos.TXPLAZACLIENTE.value+"*";
			datos.TXNROCHEQUECLI.value=datos.TXNROCHEQUECLI.value+"*";
			datos.TXNROCTACTECLI.value=datos.TXNROCTACTECLI.value+"*";
			datos.TXHRETENIDO.value=datos.TXHRETENIDO.value+"*";
		}
		//concatenar
		datos.TXDESTINO.value=datos.TXDESTINO.value+datos.DESTINO.options[e].value;
		datos.TXFPAGO.value=datos.TXFPAGO.value+datos.FPAGO.options[e].value;
		datos.TXRUTCLI.value=datos.TXRUTCLI.value+datos.RUTCLI.options[e].value;
		datos.TXMONTOCLI.value=datos.TXMONTOCLI.value+datos.MONTOCLI.options[e].value;
		datos.TXFECVENCLI.value=datos.TXFECVENCLI.value+datos.FECHACLI.options[e].value;
		datos.TXBANCOCLIENTE.value=datos.TXBANCOCLIENTE.value+datos.BANCOCLI.options[e].value;
		datos.TXDIVIDE.value=datos.TXDIVIDE.value+datos.LS_DIVIDE.options[e].value;
		datos.TXHRETENIDO.value=datos.TXHRETENIDO.value+datos.LS_HRETENIDO.options[e].value;


		//datos.TXPLAZACLIENTE.value=datos.TXPLAZACLIENTE.value+datos.PLAZACLI.options[e].value;
		datos.TXNROCHEQUECLI.value=datos.TXNROCHEQUECLI.value+datos.NROCHECLI.options[e].value;
		datos.TXNROCTACTECLI.value=datos.TXNROCTACTECLI.value+datos.NRCTACTECLI.options[e].value;
	}
}

if (datos.TX_DEUDACAPITAL.value == '') datos.TX_DEUDACAPITAL.value = 0;
if (datos.TX_INDCOM.value == '') datos.TX_INDCOM.value = 0;
if (datos.TX_HONORARIOS.value == '') datos.TX_HONORARIOS.value = 0;
if (datos.TX_TOTALGRAL.value == '') datos.TX_TOTALGRAL.value = 0;
if (datos.TX_OTROS.value == '') datos.TX_OTROS.value = 0;
if (datos.TX_INTERESES.value == '') datos.TX_INTERESES.value = 0;
if (datos.TX_GASTOSJUD.value == '') datos.TX_GASTOSJUD.value = 0;
if (datos.TX_GASTOSADMIN.value == '') datos.TX_GASTOSADMIN.value = 0;
if (datos.TX_DESCUENTO.value == '') datos.TX_DESCUENTO.value = 0;
habilita();

function InicializaCombo()
{
		var comboBox = document.getElementById('CB_DIVIDE');
		comboBox.options.length = 0;
		var newOption = new Option('SELECCIONE','10');
		comboBox.options[comboBox.options.length] = newOption;
}

function marcar_boxes(){

	desmarcar_boxes()
	<% For i=1 TO intTamvConcepto %>
			if (document.forms[0].<%=vArrConcepto(i)%>.disabled == false) {
			document.forms[0].<%=vArrConcepto(i)%>.checked=true;
			suma_capital(document.forms[0].<%=vArrConcepto(i)%>,document.forms[0].TX_SALDO_<%=vArrID_CUOTA(i)%>.value,document.forms[0].TX_INTERES_<%=vArrID_CUOTA(i)%>.value,document.forms[0].TX_HONORARIOS_<%=vArrID_CUOTA(i)%>.value,document.forms[0].TX_GPROT_<%=vArrID_CUOTA(i)%>.value);
			suma_total_general(0)
			}
	<% Next %>
}

function desmarcar_boxes(){
		datos.TX_DEUDACAPITAL.value = 0;
		datos.TX_HONORARIOS.value = 0;
		datos.TX_INTERESES.value = 0;
		datos.TX_OTROS.value = 0;
		datos.TX_TOTALGRAL.value = 0;

		<% For i=1 TO intTamvConcepto %>
			document.forms[0].<%=vArrConcepto(i)%>.checked=false;
		<% Next %>
}

InicializaCombo();

</script>

<script type="text/javascript">

    function FormatearNumero(numero) {
        var number = new String(numero);
        var result = '';
        while (number.length > 3) {
            result = '.' + number.substr(number.length - 3) + result;
            number = number.substring(0, number.length - 3);
        }
        result = number + result;
        /*alert(result);*/
        return result;

    };
    function FormatearObjeto(objeto) {
        var number = new String(objeto.value);
        number = number.replace(/\./g, "")

        var result = '';
        while (number.length > 3) {
            result = '.' + number.substr(number.length - 3) + result;
            number = number.substring(0, number.length - 3);
        }
        result = number + result;
        objeto.value = result 
        /*alert(result);*/

    };

    
    function LimpiaNumeros(numero) {
    
        var numero = new String(numero);
        var result = numero.replace(/\./g, "")
        /*alert(numero)
        alert(result)*/
        return result ;
    };

    function Solo_Numerico(variable) {
        Numer = parseInt(variable);
        if (isNaN(Numer)) {
            return "";
        }
        return Numer;
    }

</script>

	<script type="text/javascript">
		$(".TX_DEUDACAPITAL").numeric();
	    $(".TX_INTERESES").numeric();
	    $(".TX_HONORARIOS").numeric();
	    $(".TX_OTROS").numeric();
	    $(".TX_INDCOM").numeric();
	    $(".TX_GASTOSJUD").numeric();
	    $(".TX_GASTOSADMIN").numeric();
	    $(".TX_DESCUENTO").numeric();
	    $(".TX_MONTOCLI").numeric();
	    $("#remove").click(
		function (e) {
		    e.preventDefault();
		    $(".TX_INTERESES,.TX_HONORARIOS,.TX_OTROS,.TX_INDCOM,.TX_GASTOSJUD,.TX_GASTOSADMIN,.TX_DESCUENTO,.TX_MONTOCLI").removeNumeric();
		}
	);
	</script>











