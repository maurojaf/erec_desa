<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
<!--#include file="../lib/comunes/rutinas/GrabaAuditoria.inc" -->
<!--#include file="../lib/lib.asp"-->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	rut = request("TX_RUT")
	strRUT_DEUDOR=rut
	cod_pago = request("cod_pago")
	fecha=date
	usuario=session("session_idusuario")

	AbrirScg()

		strSql = "SELECT COD_CLIENTE, RUT_DEUDOR, TIPO_PAGO FROM CAJA_WEB_EMP WHERE ID_PAGO = " & cod_pago

		set rsCabecera=Conn.execute(strSql)
		If not rsCabecera.eof then
			strCOD_CLIENTE = rsCabecera("COD_CLIENTE")
			strRUT_DEUDOR = rsCabecera("RUT_DEUDOR")
			strTipoPago = rsCabecera("TIPO_PAGO")
		End if

		If strTipoPago = "CO" Then

			strSql = "SELECT CAPITAL, ID_CUOTA, NRO_DOC, CORRELATIVO FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO = " & cod_pago
           
			set rsDetalle=Conn.execute(strSql)
			If not rsDetalle.eof then
				Do until rsDetalle.eof
					strCuota = rsDetalle("NRO_DOC")
					intIdPagoCorr = rsDetalle("CORRELATIVO")

                	strSql = "UPDATE CONVENIO_DET SET PAGADA = NULL, ID_PAGO = NULL, ID_PAGO_CORR = NULL, FECHA_DEL_PAGO = NULL "
					strSql = strSql & " WHERE ID_PAGO = " & cod_pago & " AND ID_PAGO_CORR = " & intIdPagoCorr
					''Response.write "<br>ROLOstrSqlROLO = " & strSql
                    set rsUpdate=Conn.execute(strSql)

					rsDetalle.MoveNext
				Loop

                    strSql = "UPDATE EXP_CONTABILIDAD "
					strSql = strSql & " SET CP_NULO = 'S',TRASPASADO = 'T', REMESA = 'S'"
					strSql = strSql & " WHERE ID_PAGO = " & cod_pago & " "
                    set rsUpdate=Conn.execute(strSql)
			End If

		ElseIf strTipoPago = "RP" Then
			strSql = "SELECT NRO_DOC, CORRELATIVO FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO = " & cod_pago
			set rsDetalle=Conn.execute(strSql)
			If not rsDetalle.eof then
				Do until rsDetalle.eof
					strCuota = rsDetalle("NRO_DOC")
					intIdPagoCorr = rsDetalle("CORRELATIVO")

					set rsUpdate=Conn.execute(strSql)
					rsDetalle.MoveNext
				Loop
			End If
		Else
			strSql = "SELECT CAPITAL, ID_CUOTA, NRO_DOC, CORRELATIVO FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO = " & cod_pago & " "
			set rsDetalle=Conn.execute(strSql)
			If not rsDetalle.eof then
				Do until rsDetalle.eof
					intCapital = rsDetalle("CAPITAL")
					strNroDoc = rsDetalle("NRO_DOC")
					intIdCuota = rsDetalle("ID_CUOTA")

					strSql = "UPDATE CUOTA SET SALDO = SALDO + " & intCapital & ", ESTADO_DEUDA = '1', FECHA_ESTADO = GETDATE() "
					strSql = strSql & " WHERE RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CUOTA = " & intIdCuota & " "
					''Response.write "<br>strSql = " & strSql
					set rsUpdate=Conn.execute(strSql)

					strSql = "UPDATE EXP_CONTABILIDAD "
					strSql = strSql & " SET CP_NULO = 'S',TRASPASADO = 'T', REMESA = 'S'"
					strSql = strSql & " WHERE ID_PAGO = " & cod_pago & " "

					''Response.write "<br>ROLOstrSqlROLO = " & strSql

					set rsUpdate2=Conn.execute(strSql)

					rsDetalle.MoveNext
				Loop
			End If
		End If


		strSql = "INSERT INTO REVERSO_CAJA_WEB_EMP_DETALLE SELECT * FROM CAJA_WEB_EMP_DETALLE WHERE ID_PAGO =" & cod_pago & " "
		set rsIinserta=Conn.execute(strSql)

		strSql = "DELETE CAJA_WEB_EMP_DETALLE WHERE ID_PAGO=" & cod_pago
		set rsBorra=Conn.execute(strSql)

		aa = GrabaAuditoria("BORRAR", "ID_PAGO=" & cod_pago, "reversar_pago.asp", "CAJA_WEB_EMP_DETALLE")

		strSql = "INSERT INTO REVERSO_CAJA_WEB_EMP_DOC_PAGO SELECT * FROM CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO =" & cod_pago & " "
		set rsIinserta=Conn.execute(strSql)

		strSql = "DELETE CAJA_WEB_EMP_DOC_PAGO WHERE ID_PAGO=" & cod_pago
		set rsBorra=Conn.execute(strSql)
		aa = GrabaAuditoria("BORRAR", "ID_PAGO=" & cod_pago, "reversar_pago.asp", "CAJA_WEB_EMP_DOC_PAGO")

		strSql = "INSERT INTO REVERSO_CAJA_WEB_EMP SELECT * FROM CAJA_WEB_EMP WHERE ID_PAGO =" & cod_pago & " "
		'Response.write strSql
		'Response.End
		set rsIinserta=Conn.execute(strSql)

		strSql = "DELETE CAJA_WEB_EMP WHERE ID_PAGO =" & cod_pago & " "
		set rsBorra=Conn.execute(strSql)
		aa = GrabaAuditoria("BORRAR", "ID_PAGO =" & cod_pago, "reversar_pago.asp", "CAJA_WEB_EMP")

		'Response.Write strSql
		'Response.End

		CerrarScg()
	%>
	<script>alert("El pago fue reversado correctamente")</script>
	<%
	Response.Write ("<script language = ""Javascript"">" & vbCrlf)

	If strTipoPago = "CO" Then
		Response.Write (vbTab & "location.href='detalle_caja.asp?CB_TIPOPAGO=CO'" & vbCrlf)
	ElseIf strTipoPago = "RP" Then
		Response.Write (vbTab & "location.href='detalle_caja.asp?CB_TIPOPAGO=RP'" & vbCrlf)
	Else
		Response.Write (vbTab & "location.href='detalle_caja.asp?CB_TIPOPAGO=NOCO'" & vbCrlf)
	End if
	Response.Write ("</script>")
%>
