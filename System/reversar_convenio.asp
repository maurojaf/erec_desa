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
	intIdConvenio = request("cod_convenio")
	fecha=date
	usuario=session("session_idusuario")

	AbrirScg()

		strSql = "SELECT COD_CLIENTE, RUT_DEUDOR  FROM CONVENIO_ENC WHERE ID_CONVENIO = " & intIdConvenio
		set rsCabecera=Conn.execute(strSql)
		If not rsCabecera.eof then
			strCOD_CLIENTE = rsCabecera("COD_CLIENTE")
			strRUT_DEUDOR = rsCabecera("RUT_DEUDOR")
		End if

		strSql = "SELECT SALDO, NRO_DOC, ID_CUOTA FROM CONVENIO_CUOTA WHERE ID_CONVENIO = " & intIdConvenio
		'Response.write "<br>strSql = " & strSql
		set rsDetalle=Conn.execute(strSql)
		If not rsDetalle.eof then
			Do until rsDetalle.eof
				intCapital = rsDetalle("SALDO")
				strNroDoc = rsDetalle("NRO_DOC")
				intID_CUOTA = rsDetalle("ID_CUOTA")

				strSql = "UPDATE CUOTA SET SALDO = SALDO + " & intCapital & ", ESTADO_DEUDA = '1', FECHA_ESTADO = GETDATE() "
				strSql = strSql & " WHERE ID_CUOTA = " & intID_CUOTA
				'Response.write "<br>strSql = " & strSql
				set rsUpdate=Conn.execute(strSql)
				rsDetalle.MoveNext
			Loop
		End If



		strSql = "INSERT INTO REVERSO_CONVENIO_DET (ID_CONVENIO,CUOTA,TOTAL_CUOTA,FECHA_PAGO,PAGADA,ID_PAGO,ID_PAGO_CORR,FECHA_DEL_PAGO ) " 
        strSql =  strSql  & " SELECT ID_CONVENIO,CUOTA,TOTAL_CUOTA,FECHA_PAGO,PAGADA,ID_PAGO,ID_PAGO_CORR,FECHA_DEL_PAGO " 
        strSql =  strSql  & " FROM CONVENIO_DET WHERE ID_CONVENIO=" & intIdConvenio
		set rsIinserta=Conn.execute(strSql)

		strSql = "DELETE CONVENIO_DET WHERE ID_CONVENIO=" & intIdConvenio
		set rsBorra=Conn.execute(strSql)

		aa = GrabaAuditoria("BORRAR", "ID_CONVENIO=" & intIdConvenio, "reversar_convenio.asp","CONVENIO_DET")

		strSql = "DELETE CONVENIO_CUOTA WHERE ID_CONVENIO=" & intIdConvenio
		set rsBorra=Conn.execute(strSql)

		aa = GrabaAuditoria("BORRAR", "ID_CONVENIO=" & intIdConvenio, "reversar_convenio.asp","CONVENIO_CUOTA")


		strSql = "INSERT INTO REVERSO_CONVENIO_ENC (ID_CONVENIO ,COD_CLIENTE,RUT_DEUDOR,USR_INGRESO,FECHA_INGRESO,TOTAL_CONVENIO,CAPITAL,INTERESES"
        strSql = strSql  &  ",GASTOS,HONORARIOS,DESC_CAPITAL,DESC_INTERESES,DESC_GASTOS,DESC_HONORARIOS,PIE,CUOTAS,DIA_PAGO,OBSERVACIONES,FOLIO,COD_ESTADO_FOLIO,SEDE"
        strSql = strSql  &  ",IC,DESC_IC,PROTESTOS,DESC_PROTESTOS )"
        strSql = strSql  &  " SELECT ID_CONVENIO ,COD_CLIENTE,RUT_DEUDOR,USR_INGRESO,FECHA_INGRESO,TOTAL_CONVENIO,CAPITAL,INTERESES"
        strSql = strSql  &  ",GASTOS,HONORARIOS,DESC_CAPITAL,DESC_INTERESES,DESC_GASTOS,DESC_HONORARIOS,PIE,CUOTAS,DIA_PAGO,OBSERVACIONES,FOLIO,COD_ESTADO_FOLIO,SEDE"
        strSql = strSql  &  ",IC,DESC_IC,PROTESTOS,DESC_PROTESTOS FROM CONVENIO_ENC WHERE ID_CONVENIO=" & intIdConvenio
        
        set rsIinserta=Conn.execute(strSql)

		strSql = "DELETE CONVENIO_ENC WHERE ID_CONVENIO=" & intIdConvenio
		set rsBorra=Conn.execute(strSql)

		aa = GrabaAuditoria("BORRAR", "ID_CONVENIO=" & intIdConvenio, "reversar_convenio.asp","CONVENIO_ENC")

		'Response.End

		CerrarScg()
	%>
	<script>alert("El convenio fue reversado correctamente")</script>
	<%
	Response.Write ("<script language = ""Javascript"">" & vbCrlf)
	Response.Write (vbTab & "location.href='detalle_convenio.asp'" & vbCrlf)
	Response.Write ("</script>")
%>
