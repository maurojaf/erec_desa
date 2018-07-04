<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
<!--#include file="../lib/lib.asp"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	AbrirSCG()
		cod_caja=session("session_idusuario")
		strRutSede = Request("strRutSede")
		strSede = Request("strSede")

		strArrID_CUOTA=Replace(Request("strArrID_CUOTA"),";",",")
		''Response.write "strArrID_CUOTA=" & strArrID_CUOTA
		vArrID_CUOTA = split(strArrID_CUOTA,";")
		intTamvID_CUOTA=ubound(vArrID_CUOTA)


		strCOD_CLIENTE = ValNulo(Request("strCOD_CLIENTE"),"C")
		strRUT_DEUDOR = ValNulo(Request("strRUT_DEUDOR"),"C")
		intTotalConvenio = ValNulo(Request("intTotalConvenio"),"N")
		intTotalCapital = ValNulo(Request("intTotalCapital"),"N")
		intIndemComp = ValNulo(Request("intIndemComp"),"N")
		intIntereses = ValNulo(Request("intIntereses"),"N")
		intProtestos = ValNulo(Request("intProtestos"),"N")
		intGastos = ValNulo(Request("intGastos"),"N")
		intHonorarios = ValNulo(Request("intHonorarios"),"N")
		intDescTotalCapital = ValNulo(Request("intDescTotalCapital"),"N")
		intDescIndemComp = ValNulo(Request("intDescIndemComp"),"N")
		intDescGastos = ValNulo(Request("intDescGastos"),"N")
		intDescHonorarios = ValNulo(Request("intDescHonorarios"),"N")
		intDescProtestos = ValNulo(Request("intDescProtestos"),"N")
		intPie = ValNulo(Request("intPie"),"N")
		intCuotas = ValNulo(Request("intCuotas"),"N")
		intDiaDePago = Request("intDiaPago")
		strObservaciones = Request("strObservaciones")
		intIntConvenio = Request("intIntConvenio")
		intValorCuota = Request("intValorCuota")


		strSql = "EXEC UPD_SEC 'CONVENIO_ENC'"
		set rsPago=Conn.execute(strSql)
		If not rsPago.Eof then
			intSeq = rsPago("SEQ")
		End if

		strSql = "EXEC UPD_CONVENIO_CORRELATIVO '" & strCOD_CLIENTE & "','" & strRutSede & "'"
		set rsPago=Conn.execute(strSql)
		If not rsPago.Eof then
			intFolio = rsPago("CORRELATIVO")
		Else
			%>
				<script>
				alert("No se ha definido CORRELATIVO para cliente y sede seleccionada");
				history.back();
				</script>
			<%
		End if

		'Response.write "<BR>intFolio = " & intFolio
		'Response.write "<BR>intSeq = " & intSeq
		'Response.write "<BR>intIntConvenio = " & intIntConvenio
		'Response.write "<BR>Cuotas = " & Round((intTotalConvenio-intPie+intIntConvenio)/intCuotas,0)
		'Response.End


		strSql = "INSERT INTO CONVENIO_ENC (ID_CONVENIO, FECHA_INGRESO, COD_CLIENTE, RUT_DEUDOR, USR_INGRESO,  TOTAL_CONVENIO, CAPITAL, INTERESES, GASTOS, PROTESTOS, IC, HONORARIOS, DESC_CAPITAL, DESC_INTERESES, DESC_GASTOS, DESC_HONORARIOS, DESC_PROTESTOS, PIE, CUOTAS, DIA_PAGO, OBSERVACIONES, FOLIO, COD_ESTADO_FOLIO, SEDE) "
		strSql = strSql & " VALUES (" & intSeq & ", getdate(),'" & strCOD_CLIENTE & "','" & strRUT_DEUDOR & "'," & session("session_idusuario") & "," &  intTotalConvenio & ","
		strSql = strSql & intTotalCapital & "," & intIntereses & "," & intGastos & "," & intProtestos & "," & intIndemComp & "," & intHonorarios & "," & intDescTotalCapital & "," &  intDescIndemComp & "," & intDescGastos & "," & intDescHonorarios & "," & intDescProtestos & "," & intPie & "," & intCuotas & "," & intDiaDePago & ",'" & strObservaciones & "'," & intFolio & ",1,'" & strSede & "')"

		'Response.write strSql
		'Response.End

		set rsInsertaEnc=Conn.execute(strSql)

		strSql = "INSERT INTO CONVENIO_DET (ID_CONVENIO, CUOTA, TOTAL_CUOTA, FECHA_PAGO) "
		strSql = strSql & " VALUES (" & intSeq & ", 0," & intPie & ",getdate())"

		'Response.write strSql
		'Response.End
		set rsInsertaDet=Conn.execute(strSql)

		intMesDePago = Month(date)
		intAnoDePago = Year(date)
		For i=1 to intCuotas
			intMesDePago = intMesDePago + 1
			If intMesDePago = 13 Then
				intMesDePago = 1
				intAnoDePago = intAnoDePago + 1
			End if
			dtmFechaPago = PoneIzq(intDiaDePago,"0") & "/" & PoneIzq(intMesDePago,"0") & "/" & intAnoDePago
			intNroCuota = i
			''intMonto = Round((intTotalConvenio-intPie+intIntConvenio)/intCuotas,0)

			''Response.write "<br>dtmFechaPago=" & dtmFechaPago
			''Response.write "<br>intMontoConInteres=" & intMontoConInteres

			If Not Isnull(intIntConvenio/intCuotas) Then
				
				strValorCuota = CStr(intIntConvenio/intCuotas)
				
				intCantidadCaracteres = InStr(strValorCuota, ",") - 1
							
				if intCantidadCaracteres <= 0 then
				
					intCantidadCaracteres = Len(strValorCuota)
				
				end if
			
				intMonto = CLng(Mid(strValorCuota, 1, intCantidadCaracteres))
				
			End if


			If Mid(dtmFechaPago,4,2) = "02" and Cdbl(intDiaDePago) > 28 Then
				dtmFechaPago = "28/02/" & Mid(dtmFechaPago,7,4)
			End if

			If Cdbl(intDiaDePago) > 30 Then
				dtmFechaPago = "30/" & Mid(dtmFechaPago,4,2) & "/" & Mid(dtmFechaPago,7,4)
			End if
			
			if CInt(intNroCuota) = CInt(intCuotas) then
			
				intInteresConvenio = intIntConvenio - intMonto * intCuotas
						
				intMonto = intMonto + intInteresConvenio
				
			end if

			strSql = "INSERT INTO CONVENIO_DET (ID_CONVENIO, CUOTA, TOTAL_CUOTA, FECHA_PAGO) "
			strSql = strSql & " VALUES (" & intSeq & ", " & intNroCuota & "," & intMonto & ",'" & dtmFechaPago & "')"
			'Response.write strSql
			'Response.End
			set rsInsertaDet=Conn.execute(strSql)



			strNroDocumento = intSeq & "-" & intNroCuota
			intTipoDocumento = "5"
			strClaveAdic = strCOD_CLIENTE & "-" & strRUT_DEUDOR & "-" & strNroDocumento & "-1"

			strSql = "INSERT INTO CUOTA (RUT_DEUDOR, COD_CLIENTE, NRO_DOC, NRO_CUOTA, FECHA_VENC, VALOR_CUOTA, SALDO, TIPO_DOCUMENTO, ESTADO_DEUDA, FECHA_ESTADO , FECHA_CREACION, USUARIO_CREACION) "
			strSql = strSql & " VALUES ('" & strRUT_DEUDOR & "','" & strCOD_CLIENTE & "','" & strNroDocumento & "',1,'" & dtmFechaPago & "'," & intMonto & "," & intMonto & ",'" & intTipoDocumento & "','1',getdate(),getdate()," & session("session_idusuario")  & ")"

			''strSql = strSql & " VALUES (" & intSeq & ", " & intNroCuota & "," & intMonto & ",'" & dtmFechaPago & "')"
			''Response.write strSql
			'Response.End
			'set rsInsertaCuota=Conn.execute(strSql)
		Next


		strSql = "INSERT INTO CONVENIO_CUOTA (ID_CONVENIO, ID_CUOTA, RUT_DEUDOR,COD_CLIENTE,NRO_DOC,NRO_CUOTA, SALDO) SELECT " & intSeq & ",  ID_CUOTA, RUT_DEUDOR,COD_CLIENTE,NRO_DOC,NRO_CUOTA, SALDO FROM CUOTA "
		strSql = strSql & " WHERE RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & strCOD_CLIENTE &  "'"
		strSql = strSql & " AND ID_CUOTA IN (" & strArrID_CUOTA  & ")"
		'Response.write strSql
		'Response.end

		set rsInsertaConvCuota=Conn.execute(strSql)

		strSql = "UPDATE CUOTA SET ESTADO_DEUDA = 10, SALDO = 0, FECHA_ESTADO = getdate() "
		strSql = strSql & " WHERE RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & strCOD_CLIENTE &  "'"
		strSql = strSql & " AND ID_CUOTA IN (" & strArrID_CUOTA  & ")"

		set rsInsertaConvCuota=Conn.execute(strSql)


		%>
			<script>alert("Convenio grabado correctamente.")</script>
		<%
		Response.Write ("<script language = ""Javascript"">" & vbCrlf)
		Response.Write (vbTab & "location.href='simulacion_convenio.asp?rut=" & rut & "&tipo=1'" & vbCrlf)
		Response.Write ("</script>")
		%>
	</td>
   </tr>
  </table>

</form>
