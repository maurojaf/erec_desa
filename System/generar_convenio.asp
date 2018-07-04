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
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
 
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	If Trim(Request("Limpiar"))="1" Then
		session("session_RUT_DEUDOR") = ""
		rut = ""
	End if

	If Trim(Request("TX_RUT")) = "" Then
		strRUT_DEUDOR = session("session_RUT_DEUDOR")
	Else
		strRUT_DEUDOR = Trim(Request("TX_RUT"))
		session("session_RUT_DEUDOR") = strRUT_DEUDOR
	End If

	intOrigen = Request("intOrigen")

	intCliente = session("ses_codcli")

	strSede = Request("CB_SEDE")
	strCiudadSede = strSede

	intPorcDescCapital = ValNulo(Request("desc_capital"),"N")
	intPorcDescIndemComp = ValNulo(Request("desc_indemComp"),"N")
	intPorcDescHonorarios = ValNulo(Request("desc_honorarios"),"N")
	intPorcDescGastos = ValNulo(Request("desc_gastos"),"N")
	intPorcDescInteres = ValNulo(Request("desc_interes"),"N")
	intPorcDescProtestos = ValNulo(Request("desc_protestos"),"N")

	intPie = Request("pie")
	intCuotas = Request("cuotas")
	intDiaDePago = Request("TX_DIAPAGO")

	intOriginalCapital = Round(ValNulo(Request("TX_CAPITAL"),"N"),0)
	''intOriginalIndemComp = Round(ValNulo(Request("hdintIndemComp"),"N"),0)
	intOriginalIntereses = Round(ValNulo(Request("TX_INTERES"),"N"),0)
	intOriginalProtestos = Round(ValNulo(Request("TX_GASTOSPROTESTOS"),"N"),0)
	intOriginalGastos = Round(ValNulo(Request("TX_GASTOS"),"N"),0)
	intOriginalHonorarios = Round(ValNulo(Request("TX_HONORARIOS"),"N"),0)
	intOriginalIndemComp = Round(ValNulo(Request("TX_INDEM_COMP"),"N"),0)




	''intOriginalGastos = Round(ValNulo(Request("hdintGastos"),"N"),0)

	intTotalDeuda = intOriginalCapital + intOriginalIntereses + intOriginalProtestos + intOriginalHonorarios + intOriginalIndemComp + intOriginalGastos

	intDescCapital = Round(intPorcDescCapital,0)
	intDescIndemComp = Round(intPorcDescIndemComp,0)
	intDescHonorarios = Round(intPorcDescHonorarios,0)
	intDescGastos = Round(intPorcDescGastos,0)
	intDescInteres = Round(intPorcDescInteres,0)
	intDescProtestos = Round(intPorcDescProtestos,0)

	intTotalDescuentos = intDescCapital + intDescIndemComp + intDescHonorarios + intDescGastos + intDescInteres
	'intTotalDescuentos = intDescCapital + intDescGastos

	intTotalConvenio = intTotalDeuda - intTotalDescuentos

	intTotDeudaCapital = Round(intOriginalCapital - intDescCapital,0)
	intTotIndemComp = Round(intOriginalIndemComp - intDescIndemComp,0)
	intTotHonorarios = Round(intOriginalHonorarios - intDescHonorarios,0)
	intTotGastos = Round(intOriginalGastos - intDescGastos,0)
	intTotIntereses = Round(intOriginalIntereses - intDescInteres,0)

	intTotProtestos = Round(intOriginalProtestos - intDescProtestos,0)

	AbrirSCG()

	strSql = "SELECT * FROM SEDE S, CONVENIO_CORRELATIVO C WHERE S.COD_CLIENTE = '" & intCliente & "' AND S.SEDE = '" & strSede & "'"
	strSql = strSql & " AND S.RUT = C.RUT AND S.COD_CLIENTE = C.COD_CLIENTE"
	'Response.write "strSql=" & strSql
	set rsSede = Conn.execute(strSql)
	if not rsSede.eof then
		strRazonSocialSede=rsSede("RAZON_SOCIAL")
		strDireccionSede=rsSede("DIRECCION")
		strRutSede=rsSede("RUT")
		strTelefonoSede=rsSede("TELEFONO")
		strNroFolio=rsSede("FOLIO_ACTUAL")
	Else
		%>
		<SCRIPT>
			alert('Parámetros del convenio no han sido configurados, revise configuración.');
			history.back();
		</SCRIPT>
		<%
		Response.End
	End If



	strSql=""
	strSql="SELECT FORMULA_HONORARIOS,FORMULA_INTERESES,TASA_MAX_CONV, DESCRIPCION, TIPO_INTERES FROM CLIENTE WHERE COD_CLIENTE ='" & intCliente & "'"
	set rsTasa=Conn.execute(strSql)
	if not rsTasa.eof then
		intTasaMax = rsTasa("TASA_MAX_CONV")
		strDescripcion = rsTasa("DESCRIPCION")
		strTipoInteres = rsTasa("TIPO_INTERES")
		strNomFormHon = ValNulo(rsTasa("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsTasa("FORMULA_INTERESES"),"C")

	Else
		intTasaMax = 1
		strDescripcion = ""
		strTipoInteres = ""
	end if
	rsTasa.close
	set rsTasa=nothing


			intKapitalInicial = intTotalConvenio-intPie
			'strTipoInteres SIMPLE
			'M= C ( 1+i*n ) M= 3,250(1 + (0.025)(1.5/12)= 3,351.56

			'strTipoInteres COMPUESTO
			'M= C ( 1+i)^n ) M= 3,250(1 + (0.025)(1.5/12)= 3,351.56

			'Response.write "EXPO=" & calcula_base_exponente(3, 4)
			'Response.write "strTipoInteres=" & strTipoInteres

			'Response.write "<br>intTasaMax=" & intTasaMax
			'Response.write "<br>intKapitalInicial=" & intKapitalInicial
			'Response.write "<br>intCuotas=" & intCuotas


			If Trim(strTipoInteres)="C" Then
				intMontoConInteres = intKapitalInicial * calcula_base_exponente((1 + intTasaMax/100),intCuotas)
				'Response.write "<br>intMontoConInteres=" & intMontoConInteres
				'Response.write intCuotas
				calculaBaseExponenteMenosUno = (calcula_base_exponente((1 + intTasaMax/100),intCuotas)-1)
				
				if calculaBaseExponenteMenosUno <= 0 then
					calculaBaseExponenteMenosUno = 1
				end if
				
				intValorCuota=intKapitalInicial*((intTasaMax/100*calcula_base_exponente((1 + intTasaMax/100),intCuotas))/calculaBaseExponenteMenosUno)
				intValorCuota = Round(intValorCuota,0)
				'Response.write "<br>sss=" & sss
				''C36*((0,02*(1,02)^C39)/((1,02)^C39-1))

			Else
				intValorCuota=intKapitalInicial * (1 + ((intTasaMax/100)*intCuotas))
				intValorCuota = Round(intValorCuota/intCuotas,0)
				''intMontoConInteres = intKapitalInicial * (1 + ((intTasaMax/100)*intCuotas))
			End If


			intMonto = intValorCuota

		strSql="SELECT NOMBRE_DEUDOR FROM DEUDOR WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & intCliente & "'"
		set RsDeudor=Conn.execute(strSql)
		if not RsDeudor.eof then
			strNombreDeudor = RsDeudor("NOMBRE_DEUDOR")
		end if
		RsDeudor.close
		set RsDeudor=nothing

		strSql=""
		strSql="SELECT TOP 1 CALLE,NUMERO,COMUNA,RESTO,CORRELATIVO,ESTADO FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND ESTADO <> '2' ORDER BY ESTADO DESC"
		set rsDIR=Conn.execute(strSql)
		if not rsDIR.eof then
			calle_deudor=rsDIR("Calle")
			numero_deudor=rsDIR("Numero")
			comuna_deudor=rsDIR("Comuna")
			resto_deudor=rsDIR("Resto")
			strDirDeudor = calle_deudor & " " & numero_deudor & " " & resto_deudor & " " & comuna_deudor
		end if
		rsDIR.close
		set rsDIR=nothing


		strSql=""
		strSql="SELECT TOP 1 COD_AREA,TELEFONO,CORRELATIVO,ESTADO FROM DEUDOR_TELEFONO WHERE  RUT_DEUDOR='" & strRUT_DEUDOR & "' AND ESTADO <> '2' ORDER BY ESTADO DESC"
		set rsFON=Conn.execute(strSql)
		if not rsFON.eof then
			codarea_deudor = rsFON("COD_AREA")
			Telefono_deudor = rsFON("Telefono")
			strFonoDeudor = codarea_deudor & "-" & Telefono_deudor
		end if
		rsFON.close
		set rsFON=nothing


		strSql=""
		strSql="SELECT TOP 1 RUT_DEUDOR,CORRELATIVO,FECHA_INGRESO,EMAIL,ESTADO FROM DEUDOR_EMAIL WHERE  RUT_DEUDOR='" & strRUT_DEUDOR & "' AND ESTADO <> '2' ORDER BY ESTADO DESC"
		set rsMAIL=Conn.execute(strSql)
		if not rsMAIL.eof then
			strEmail = rsMAIL("EMAIL")
		end if
		rsMAIL.close
		set rsMAIL=nothing



				intMesDePago = Month(date)
				intAnoDePago = Year(date)

				intKapitalInicial = intTotalConvenio-intPie
				'M= C ( 1+it ) M= 3,250(1 + (0.25)(1.5/12)= 3,351.56

				'strTipoInteres = "C"

				If Trim(strTipoInteres)="C" Then
					intMontoConInteres = intKapitalInicial * calcula_base_exponente((1 + intTasaMax/100),intCuotas)
					
					calculaBaseExponenteMenosUno = (calcula_base_exponente((1 + intTasaMax/100),intCuotas)-1)
				
					if calculaBaseExponenteMenosUno <= 0 then
						calculaBaseExponenteMenosUno = 1
					end if
					
					intValorCuota=intKapitalInicial*((intTasaMax/100*calcula_base_exponente((1 + intTasaMax/100),intCuotas))/calculaBaseExponenteMenosUno)
					intValorCuota = Round(intValorCuota,0)
				Else
					intMontoConInteres = intKapitalInicial * (1 + ((intTasaMax/100)*intCuotas))
				End If




				'intValorCuota = intMonto


				intMontoTotal = (intTotalConvenio -  intPie)
				' intMontoTotal = 0
				' For i=1 to intCuotas

					' intTotalGastos=0
					' intCont=1

					' intMesDePago = intMesDePago + 1
					' If intMesDePago = 13 Then
						' intMesDePago = 1
						' intAnoDePago = intAnoDePago + 1
					' End if
					' intCont = intCont + 1

					' dtmFechaPago = PoneIzq(intDiaDePago,"0") & "/" & PoneIzq(intMesDePago,"0") & "/" & intAnoDePago
					' intNroCuota = i

					' If i=1 Then
						' dtmFechaCuota1 = dtmFechaPago
					' End If

					' If Not Isnull(intMontoConInteres/intCuotas) Then
						' intMonto = Round(intMontoConInteres/intCuotas,0)
						' intMonto= intValorCuota
					' End if


					' If Mid(dtmFechaPago,4,2) = "02" and Cdbl(intDiaDePago) > 28 Then
						' dtmFechaPago = "28/02/" & Mid(dtmFechaPago,7,4)
					' End if

					' If Cdbl(intDiaDePago) > 30 Then
						' dtmFechaPago = "30/" & Mid(dtmFechaPago,4,2) & "/" & Mid(dtmFechaPago,7,4)
					' End if
					' intMontoTotal = intMontoTotal + intMonto

				' Next

				dtmFechaCuotaFin = dtmFechaPago










%>

	<title>Documento sin t&iacute;tulo</title>
	<style type="text/css">
		<!--
		.Estilo1 {font-size: 14px;font-weight: bold;}
		.Estilo13 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; }
		.Estilo2 {font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif;}
		.Estilo5 {font-size: 11px; font-weight: bold;}
		.Estilo8 {font-size: 11px}
		.Estilo9 {font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; }
		.Estilo12 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
		.Estilo14 {font-size: 10px}
		.Estilo15 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
		.Estilo26 {font-size: 12px}
		.Estilo28 {font-family: "Courier New", Courier, monospace}
		.Estilo32 {font-size: 13px}
		.Estilo22 {font-size: 13px; font-family: "Courier New", Courier, monospace;}
		.Estilo36 {font-family: Verdana, Arial, Helvetica, sans-serif;font-weight: bold;}
		.Estilo37 {font-family: Verdana, Arial, Helvetica, sans-serif; }
		.Estilo38 {font-size: 13px; font-family: Verdana, Arial, Helvetica, sans-serif; }
		.Estilo40 {	font-family: Verdana, Arial, Helvetica, sans-serif;	color: #FF0000;	font-size: 11px;font-weight: bold;}
		.Estilo41 {color: #FFFFFF}
		.Estilo33 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 14px;font-weight: bold;}
		.Estilo34 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px}

		H1.SaltoDePagina {PAGE-BREAK-AFTER: always}
			.transpa {
			background-color: transparent;
			border: 1px solid #FFFFFF;
			text-align:center
			}
		-->
		.align { text-align: justify; padding:3px }
	</style>


<script language="JavaScript " type="text/JavaScript">

function Grabar()
{
	if (confirm("¿ Está seguro de grabar el <%=session("NOMBRE_CONV_PAGARE")%> ? Si aun no lo ha imprimido, presione cancelar y luego el botón imprimir."))
		{
			datos.action = "grabar_convenio.asp";
			datos.submit();
		}
}
</script>
</head>

<BODY>
<FORM NAME="datos" METHOD="POST" ACTION="grabar_convenio.asp">

<INPUT TYPE="hidden" NAME="strCOD_CLIENTE" value="<%=intCliente%>">
<INPUT TYPE="hidden" NAME="strRUT_DEUDOR" value="<%=strRUT_DEUDOR%>">
<INPUT TYPE="hidden" NAME="intTotalConvenio" value="<%=intTotalConvenio%>">
<INPUT TYPE="hidden" NAME="intTotalCapital" value="<%=intTotDeudaCapital%>">
<INPUT TYPE="hidden" NAME="intIntereses" value="<%=intTotIntereses%>">
<INPUT TYPE="hidden" NAME="intProtestos" value="<%=intTotProtestos%>">


<INPUT TYPE="hidden" NAME="intIndemComp" value="<%=intTotIndemComp%>">
<INPUT TYPE="hidden" NAME="intGastos" value="<%=intTotGastos%>">
<INPUT TYPE="hidden" NAME="intHonorarios" value="<%=intTotHonorarios%>">
<INPUT TYPE="hidden" NAME="intDescTotalCapital" value="<%=intDescCapital%>">
<INPUT TYPE="hidden" NAME="intDescIndemComp" value="<%=intDescIndemComp%>">
<INPUT TYPE="hidden" NAME="intDescGastos" value="<%=intDescGastos%>">
<INPUT TYPE="hidden" NAME="intDescHonorarios" value="<%=intDescHonorarios%>">
<INPUT TYPE="hidden" NAME="intDescProtestos" value="<%=intDescProtestos%>">
<INPUT TYPE="hidden" NAME="intPie" value="<%=intPie%>">
<INPUT TYPE="hidden" NAME="intCuotas" value="<%=intCuotas%>">
<INPUT TYPE="hidden" NAME="intDiaPago" value="<%=intDiaDePago%>">
<INPUT TYPE="hidden" NAME="strObservaciones" value="<%=strObservaciones%>">
<INPUT TYPE="hidden" NAME="strRutSede" value="<%=strRutSede%>">

<TABLE ALIGN="CENTER" WIDTH="650" CELLSPACING=0 CELLPADDING=1>
		<TR>
		  <TD>
		  <table width="650" border="0" cellspacing=0 cellpadding=0>
			<tr>
			  <td width=150 rowspan="3"  class="Estilo13" align="center"><img border="0" src="../imagenes/Logos/<%=session("ses_codcli")%>/logo.jpg"></td>
			  <td width=100>&nbsp;</td>
			  <td></td>
			</tr>
			<tr>
			  <td></td>
			  <td>&nbsp;</td>
			  <td align="right"><span class="Estilo33"> Nº <%=strNroFolio%></span></td>
			</tr>
		  </table>
		  </TD>
 		</TR>

 		<TR>
		 	<TD>
		 		<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
				 		<tr class="Estilo1">
						 	<TD colspan=4 align="center" class="Estilo33">
					 		  PAGARÉ 
					 		</TD>
				 		</TR>
				 		<TR>
						 	<td colspan=4><div align="center"><%=strMandante%></div></TD>
				 		</TR>
				</TABLE>

	 		</TD>
 		</TR>
 		<TR>
			<TD>
				<br>
				<table width="650" border="0">
				<br>
					<tr>
						<td class="Estilo34" align="left"><p class="align">
								Debo y pagaré a la orden de Universidad Mayor, Rut. 71.500.500-K, la suma de 
								<B>$<%=fn(intMontoTotal,0)%></B> pesos.
								<br>
								La cantidad adecuada devengará un interés del 0,0% anual. En caso de mora
								o simple retardo en el pago, se devengará un interés mensual igual al interés
								máximo convencional que la Ley permite para operaciones de crédito de dinero 
								no reajustables, por el periodo comprendido entre la fecha de la mora o simple 
								retardo y el día de su pago.
								<br><br>
								El capital y los intereses adeudados se pagarán  en <b><%=intCuotas%></b> cuotas mensuales,
								cuyos montos y vencimientos serán los siguientes:
								<br>
						</td>
					</tr>
				</table>
				<br>
				<br>
			</TD>
 		</TR>
 	<TR>
 		<TD>
 		<TABLE ALIGN= "CENTER" WIDTH="70%" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
 			<%

						intMesDePago = Month(date)
						intAnoDePago = Year(date)

						intKapitalInicial = intTotalConvenio-intPie
						'M= C ( 1+it ) M= 3,250(1 + (0.25)(1.5/12)= 3,351.56

						'strTipoInteres = "C"

						If Trim(strTipoInteres)="C" Then
							intMontoConInteres = intKapitalInicial * calcula_base_exponente((1 + intTasaMax/100),intCuotas)
						Else
							intMontoConInteres = intKapitalInicial * (1 + ((intTasaMax/100)*intCuotas))
						End If

						'Response.write "<br>intMontoConInteres=" & intMontoConInteres
						'Response.write "<br>intCuotas=" & intCuotas
						%>
						
						<tr>
						<td>
						<table align="center" border="1" cellpadding="0" cellspacing="0">
							<tr>
								<td width="140px" align="center">  NºCuota  </td>
								<td width="200px" align="center">  Monto $  </td>
								<td width="200px" align="center">  Fecha Vencimiento  </td>	
							</tr>
						<%

						For i=1 to intCuotas
						intTotalGastos=0
						intCont=1

						intMesDePago = intMesDePago + 1
						If intMesDePago = 13 Then
							intMesDePago = 1
							intAnoDePago = intAnoDePago + 1
						End if
						intCont = intCont + 1

						dtmFechaPago = PoneIzq(intDiaDePago,"0") & "/" & PoneIzq(intMesDePago,"0") & "/" & intAnoDePago
						intNroCuota = i

						'intMonto = Round(intKapitalInicial/intCuotas,0)
						If Not Isnull(intMontoConInteres/intCuotas) Then
						
							strValorCuota = CStr(intMontoConInteres/intCuotas)
							
							intCantidadCaracteres = InStr(strValorCuota, ",") - 1
							
							if intCantidadCaracteres <= 0 then
							
								intCantidadCaracteres = Len(strValorCuota)
							
							end if
						
							intMonto = CLng(Mid(strValorCuota, 1, intCantidadCaracteres))
							
						End if

						'Response.write "<br>intMonto=" & intMonto
						
						If Mid(dtmFechaPago,4,2) = "02" and Cdbl(intDiaDePago) > 28 Then
							dtmFechaPago = "28/02/" & Mid(dtmFechaPago,7,4)
						End if

						If Cdbl(intDiaDePago) > 30 Then
							dtmFechaPago = "30/" & Mid(dtmFechaPago,4,2) & "/" & Mid(dtmFechaPago,7,4)
						End if
						
						if CInt(intNroCuota) = CInt(intCuotas) then
						
							intInteresConvenio = intMontoConInteres - intMonto * intCuotas
						
							intMontoCuota = FN(intMonto + intInteresConvenio,0)
							
						else
						
							intMontoCuota = FN(intMonto,0)
						
						end if

						%>													
							<tr>
								<td align="center"> <%=intNroCuota%></td>
								<td align="center"> $ <%=intMontoCuota %> </td>
								<td align="center"> <%=dtmFechaPago%> </td>
							</tr>
						
						<% Next %>
						</table>
					</td>
					</tr>	
					
				</TABLE>
			</TD>
 		</TR>

 		<TR>
		 	<TD>
				<TABLE ALIGN= "CENTER" WIDTH="90%" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
					<br>
					<tr>
						<td class="Estilo34" align="left"><p class="align">
						En caso de mora o simple retardo en el pago de cualquiera de las cuotas antes
						señaladas, acepto expresamente que la Universidad Mayor podrá exigir
						anticipadamente el pago íntegro y total  de la deuda insoluta, la que se 
						considerará como de plazo vencido para todos los efectos legales.
						<br><br>
						La obligación contraída en este instrumento es indivisible, para todos los 
						efectos legales, indivisibilidad que pasará a los herederos y sucesores del
						deudor, pudiendo la Universidad Mayor, demandar su cobro a cuales quiera
						de ellos, en conformidad al artículo 1.528 del Código Civil.
						<br><br>
						Libero a la Universidad Mayor de la obligación de protestar el presente pagaré en
						caso de no ser pagado a su vencimiento.
						<br><br>
						Para los efectos legales, judiciales y de eventual protesto de este documento,
						señalo domicilio en Avenida Los Leones esquina San Pio X Nº 2422, Providencia, y me someto a la 
						competencia y jurisdicción de los tribunales ordinarios de Justicia de la ciudad
						de Santiago comuna de Las Condes.
						<br><br>
						Para el caso de no pago del pagaré suscrito o en caso de simple retardo o
						mora en el cumplimiento de las obligaciones contraídas, el compareciente
						autoriza expresamente a Universidad Mayor para el ingreso de los datos
						personales y antecedentes del protesto o incumplimiento en un sistema de
						información comercial facultando expresamente su digitación, procesamiento y
						comunicación en línea o en cualquier otra forma y liberando desde ya a
						Universidad Mayor de dar aviso de pago o extinción de deuda, previo
						otorgamiento de constancia suficiente del pago de la deuda correspondiente,
						según lo dispuesto en la Ley 19.628.
						<br><br>
						<!--Autorización: En este acto, _______________________________ acepta y autoriza a Universidad Mayor, para que en caso de 
						mora o simple retardo en el pago de una cualquiera de las cuotas del presente pagaré, sus datos puedan ser ingresados, 
						procesados, tratados y comunicados a terceros sin restricciones por cualquier forma o medio, incluyendo expresamente 
						la base de datos SICOM y DICOM.
						<br>-->
						<br>
						<br>
						<br>
						<br>
						___________________________________<br>
						Firma Suscriptor:<br><br>
						<br>
						Fecha: _____________________________<br>
						<br>
						Nombre Suscriptor: ________________________________________<br>
						<br>
						Numero Interno: ____________________<br>
						<br>
						RUT: ______________________________<br>
						<br>
						Domicilio: _________________________________________________<br>
						<br>
						Fono: _____________________________<br>
						<br>
						Celular: ___________________________<br>						
						<br>
						<br><br>
						"La Universidad Mayor se encuentra exenta del impuesto de Timbres y
						Estampillas que grava la emisión de Letras de Cambios y Pagarés, y las actas
						de protesto, de acuerdo a lo establecido en los artículos 23 Nº3 y 11 inciso
						segundo del DL 3475, de fecha 29.08.1980."
						<br>
						<br>
						</td>
					</tr>
				</TABLE>
			</TD>
 		</TR>
 </TABLE>




<%

'Capital: $ FN(intTotDeudaCapital,0)
'Intereses: $ FN(intTotIntereses,0)
'Honorarios: FN(intTotHonorarios,0)
'G.Protesto: FN(intTotProtestos,0)
'Total Deuda: FN(intTotalDeuda,0)
'Pie session("NOMBRE_CONV_PAGARE") FN(intPie,0)
'Saldo en session("NOMBRE_CONV_PAGARE") FN(intTotalDeuda -  intPie,0)
'Intereses session("NOMBRE_CONV_PAGARE") FN(intMontoA,0)
'Saldo a Convenir: $ FN((intTotalDeuda -  intPie) + intMontoA,0)
'
'Pie session("NOMBRE_CONV_PAGARE")$ FN(intPie,0)
'Total Pago Caja:FN(intPie,0)

%>

<BR>
   	<H1 class=SaltoDePagina> </H1>
<BR>

<TABLE ALIGN="CENTER" WIDTH="650" BORDER="1" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

		<TR>
			 <TD>
			  <table width="650" border="0" cellspacing=0 cellpadding=0>
				<tr>
				  <td width=150 rowspan="3"  class="Estilo13" align="center"><img border="0" src="../imagenes/Logos/<%=session("ses_codcli")%>/logo.jpg"></td>
				  <td width=100>&nbsp;</td>
				  <td></td>
				</tr>
				<tr>
				  <td></td>
				  <td>&nbsp;</td>
				  <td align="right"><span class="Estilo33"> Nº <%=strNroFolio%></span></td>
				</tr>
			  </table>
			 </TD>
 		</TR>

		<TR>
		 	<TD>
		 		<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
				 		<tr class="Estilo1">
						 	<TD colspan=4 align="center" class="Estilo33">
					 		  DETALLE DEUDA<BR>
					 		</TD>
				 		</TR>
				 		<TR>
						 	<td colspan=4><div align="center"><%=strMandante%></div></TD>
				 		</TR>
				</TABLE>

	 		</TD>
 		</TR>
 		<TR>
			<TD>
				<TABLE ALIGN="CENTER" WIDTH="650" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

					<tr class="Estilo1">
						<TD colspan=4 align="LEFT" class="Estilo34">
						  <br>ANTECEDENTES DE LA DEUDA
						</TD>
					</TR>
				</TABLE>
				<table width="100%" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
					<tr class="Estilo34">
						<td>NRO_DOC</td>
						<td>FECHA_VENC</td>
						<td>DIAS_MORA</td>
						<td>TIPO_DOCUMENTO</td>
						<!--td>ASIGNACION</td-->
						<td>CAPITAL</td>
						<td>INTERES</td>
						<td>PROTESTOS</td>
						<td>HONORARIOS</td>
						<td>ABONO</td>
					</tr>
				<%

				strSql = "SELECT ID_CUOTA, NRO_DOC, TIPO_DOCUMENTO, GASTOS_PROTESTOS, CUENTA, FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD FROM CUOTA WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND COD_CLIENTE='" & intCliente & "' AND SALDO > 0 ORDER BY FECHA_VENC DESC"
				strSql = "SELECT dbo." & strNomFormInt & "(ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(ID_CUOTA) as HONORARIOS, VALOR_CUOTA, ID_CUOTA, NRO_DOC, NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO, GASTOS_PROTESTOS, CUENTA, FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, CUSTODIO, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS , SALDO FROM CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND COD_CLIENTE='" & intCliente & "' AND SALDO > 0 AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO ORDER BY FECHA_VENC DESC"
				set rsTemp= Conn.execute(strSql)

				intTasaMensual = 2/100
				intTasaDiaria = intTasaMensual/30
				intCorrelativo = 1
				strArrID_CUOTA=""
				intTotSelSaldo= 0
				intTotalCapital=0
				intTotSelIntereses= 0
				intTotSelProtestos= 0
				intTotSelHonorarios= 0
				intTotalAbono=0
				Do until rsTemp.eof
					strObjeto = "CH_" & rsTemp("ID_CUOTA")
					strObjeto1 = "TX_SALDO_" & rsTemp("ID_CUOTA")
					If UCASE(Request(strObjeto)) = "ON" Then

						'intSaldo = Request(strObjeto1)
                        intSaldo = Round(session("valor_moneda") * ValNulo(rsTemp("SALDO"),"N"),0)
						strNroDoc = rsTemp("NRO_DOC")
						strFechaVenc = rsTemp("FECHA_VENC")
						strTipoDoc = rsTemp("TIPO_DOCUMENTO")

						intAntiguedad = ValNulo(rsTemp("ANTIGUEDAD"),"N")

						intIntereses = rsTemp("INTERESES")
						intHonorarios = rsTemp("HONORARIOS")

						intValorCapital = rsTemp("VALOR_CUOTA")
						intAbono = intValorCapital - intSaldo 

						intProtestos = ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")

						strArrID_CUOTA = strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")

						intTotalAbono = intTotalAbono + intAbono
						intTotalCapital = intTotalCapital + intValorCapital
						intTotSelSaldo= intTotSelSaldo+intSaldo
						intTotSelIntereses= intTotSelIntereses+intIntereses
						intTotSelProtestos= intTotSelProtestos+intProtestos
						intTotSelHonorarios= intTotSelHonorarios+intHonorarios

						%>
						<tr class="Estilo34">
						<td><%=strNroDoc%></td>
						<td><%=strFechaVenc%></td>
						<td><%=intAntiguedad%></td>
						<td><%=strTipoDoc%></td>
						<td ALIGN="RIGHT"><%=FN(intValorCapital,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intIntereses,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intProtestos,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intHonorarios,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intAbono,0)%></td>
						</tr>
						<%

					End if
				%>

				<%
				rsTemp.movenext
				intCorrelativo = intCorrelativo + 1
				loop
				rsTemp.close
				set rsTemp=nothing

				strArrID_CUOTA = Mid(strArrID_CUOTA,2,len(strArrID_CUOTA))
		%>
			<tr class="Estilo34">
				<td colspan = 4>Totales</td>
				<td ALIGN="RIGHT"><%=FN(intTotalCapital,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotSelIntereses,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotSelProtestos,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotSelHonorarios,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotalAbono,0)%></td>
			</tr>

			<INPUT TYPE="HIDDEN" NAME="strArrID_CUOTA" VALUE="<%=strArrID_CUOTA%>">
			</table>

			<br>
					<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0">
						<TR class="Estilo1">
							<TD colspan=4 align="CENTER" class="Estilo38">
							  DETALLE <%=UCASE(session("NOMBRE_CONV_PAGARE"))%><br><br>
							</TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD ALIGN="RIGHT"><b>Conceptos</b></TD><TD align="right"><b>Monto Original</b></TD>
							<TD align="right"><b>Descuentos</b></TD><TD align="right"><b>Total</b></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD ALIGN="RIGHT">Capital: </TD><TD align="right">$ <%=FN(intOriginalCapital,0)%></TD>
							<TD align="right">$ <%=FN(intDescCapital,0)%></TD><TD align="right">$ <%=FN(intTotDeudaCapital,0)%></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right">Intereses: </TD><TD align="right">$ <%=FN(intOriginalIntereses,0)%></TD>
							<TD align="right">$ <%=FN(intDescInteres,0)%></TD><TD align="right">$ <%=FN(intTotIntereses,0)%></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right">Protestos: </TD><TD align="right">$ <%=FN(intOriginalProtestos,0)%></TD>
							<TD align="right">$ <%=FN(intDescProtestos,0)%></TD><TD align="right">$ <%=FN(intTotProtestos,0)%></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right">Honorarios: </TD><TD align="right">$ <%=FN(intOriginalHonorarios,0)%></TD>
							<TD align="right">$ <%=FN(intDescHonorarios,0)%></TD><TD align="right">$ <%=FN(intTotHonorarios,0)%></TD>
						</TR>
						<!--<TR HEIGHT=15 class="Estilo38">
							<TD align="right">Indem.Comp.: </TD><TD align="right">$ <%=FN(intOriginalIndemComp,0)%></TD>
							<TD align="right">$ <%=FN(intDescIndemComp,0)%></TD><TD align="right">$ <%=FN(intTotIndemComp,0)%></TD>
						</TR>-->
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right">Gastos: </TD><TD align="right">$ <%=FN(intOriginalGastos,0)%></TD>
							<TD align="right">$ <%=FN(intDescGastos,0)%></TD><TD align="right">$ <%=FN(intTotGastos,0)%></TD>
						</TR>
						<TR HEIGHT=15>
							<TD colspan=3>&nbsp</TD><TD align="right">______________</TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>Total Deuda:</TD><TD align="right"><strong>$ <%=FN(intTotalConvenio,0)%></strong></TD>
						</TR>
						<!--TR HEIGHT=15 class="Estilo38">
							<TD align="right">Total Descuentos:</TD><TD align="right"><strong>$ <%=FN(intTotalDeuda,0)%></strong></TD>
						</TR-->

						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>Pie <%=session("NOMBRE_CONV_PAGARE")%>:</TD><TD align="right"><strong>-  $ <%=FN(intPie,0)%></strong></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD colspan=3>&nbsp</TD><TD align="right">______________</TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>Saldo:</TD><TD align="right"><strong>$ <%=FN((intTotalConvenio -  intPie),0)%></strong></TD>
						</TR>
					</table>
				</tr>
			</TD>
 		</TR>
 </TABLE>
 	<INPUT TYPE="hidden" NAME="intIntConvenio" value="<%=intMontoTotal%>">
	<INPUT TYPE="hidden" NAME="intValorCuota" value="<%=Round(intValorCuota,0)%>">
	<INPUT TYPE="hidden" NAME="strSede" value="<%=strSede%>">

</FORM>

<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
			<TR class="Estilo1">
				<TD align="LEFT" class="Estilo34">
					<acronym title="GRABAR <%=UCASE(session("NOMBRE_CONV_PAGARE"))%>">
						<input name="BT_GRABAR" type="button" onClick="Grabar();" value="Grabar">
					</acronym>
				</TD>
				<TD align="RIGHT" class="Estilo34">
					<acronym title="IMPRIMIR <%=UCASE(session("NOMBRE_CONV_PAGARE"))%>">
						<input name="BT_IMPRIMIR" type="button" onClick="window.print();" value="Imprimir">
					</acronym>
				</TD>

			</TR>
	</TABLE>





&nbsp;&nbsp;
</body>
</html>