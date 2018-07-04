<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"



cuotas_rut 					=request.querystring("cuotas_rut")
rut 						=request.querystring("rut")
NOMBRE_DEUDOR				=request.querystring("NOMBRE_DEUDOR")
fecha_generar_documentos 	=request.querystring("fecha_generar_documentos")

fecha_generar_documentos    =replace(fecha_generar_documentos,"/","-")


AbrirSCG()


	strSql=" SELECT USA_SUBCLIENTE, USA_INTERESES, USA_HONORARIOS, USA_PROTESTOS, FORMULA_HONORARIOS, " 
	strSql= strSql & " FORMULA_INTERESES,TASA_MAX_CONV, DESCRIPCION, TIPO_INTERES "
	strSql= strSql & " FROM CLIENTE WHERE COD_CLIENTE ='" & session("ses_codcli") & "'"

	set rsTasa=Conn.execute(strSql)
	if not rsTasa.eof then
		intTasaMax 			= rsTasa("TASA_MAX_CONV")
		strDescripcion 		= rsTasa("DESCRIPCION")
		strTipoInteres 		= rsTasa("TIPO_INTERES")
		strNomFormHon 		= ValNulo(rsTasa("FORMULA_HONORARIOS"),"C")
		strNomFormInt 		= ValNulo(rsTasa("FORMULA_INTERESES"),"C")
		strUsaSubCliente	= rsTasa("USA_SUBCLIENTE")
		strUsaInteres 		= rsTasa("USA_INTERESES")
		strUsaHonorarios 	= rsTasa("USA_HONORARIOS")
		strUsaProtestos 	= rsTasa("USA_PROTESTOS")

	Else
		intTasaMax = 1
		strDescripcion = ""
		strTipoInteres = ""
	end if


	fecha =fecha_generar_documentos&"_detalle_deuda.csv"
		
	terceroCSV = TRIM(session("ses_ruta_sitio"))&"\Archivo\BibliotecaNotificacionDeudores\"&trim(session("ses_codcli"))&"\"&rut&"\"&fecha



	set confile = createObject("scripting.filesystemobject")
	set fichCA = confile.CreateTextFile(terceroCSV)




	strTextoTercero="RUT_CLIENTE;NOMBRE_CLIENTE;RUT;NOMBRE;TIPO_DOCUMENTO;N_DOCUMENTO;N_CUOTA;FEC_VENCIMINETO;DIAS_MORA;CAPITAL;ABONO;"
	If Trim(strUsaProtestos)="1" Then
		strTextoTercero= strTextoTercero & "PROTESTOS;"
	end if
	if trim(strUsaInteres)="1" then
		strTextoTercero= strTextoTercero & "INTERES;" 
	END IF 
	if trim(strUsaHonorarios)="1" then
		strTextoTercero= strTextoTercero & "HONORARIOS;"
	end if
	strTextoTercero= strTextoTercero & "SALDO;FECHA_PAGO;MONTO_PAGO;OBSERVACION"

	fichCA.writeline(strTextoTercero)



		strSql = "SELECT RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, convert(int,VALOR_CUOTA) VALOR_CUOTA , convert(int,isnull(dbo." & trim(strNomFormInt) & "(ID_CUOTA),0)) as INTERESES, convert(int,isnull(dbo." & trim(strNomFormHon) & "(ID_CUOTA),0)) as HONORARIOS, ID_CUOTA, NRO_DOC, NRO_CUOTA, NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO, convert(int, isnull(GASTOS_PROTESTOS,0)) GASTOS_PROTESTOS, CUENTA, convert(char, FECHA_VENC, 103) FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, CUSTODIO, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, convert(int,SALDO) SALDO, ( "
		strSql = strSql & " 		SELECT MIN(FECHA_VENC) "
		strSql = strSql & " 		FROM CUOTA, TIPO_DOCUMENTO  "
		strSql = strSql & " 		WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & session("ses_codcli") & "' AND SALDO > 0  "
		strSql = strSql & " 		AND ESTADO_DEUDA IN (	SELECT ESTADO_DEUDA  "
		strSql = strSql & " 								FROM ESTADO_DEUDA "
		strSql = strSql & " 								WHERE ACTIVO = 1) "
		strSql = strSql & " 		AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO "
		strSql = strSql & " 		AND CUOTA.ID_CUOTA IN ("&cuotas_rut&") "
		strSql = strSql & " 		GROUP BY CUOTA.RUT_DEUDOR ) MAX_FECHA_VENC "

		strSql = strSql & " FROM CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & session("ses_codcli") & "' AND SALDO > 0 AND ESTADO_DEUDA IN (SELECT ESTADO_DEUDA FROM ESTADO_DEUDA WHERE ACTIVO = 1) AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO AND CUOTA.ID_CUOTA IN ("&cuotas_rut&")"
		strSql = strSql & " ORDER BY CONVERT(DATETIME, CUOTA.FECHA_VENC) ASC "

		set rsTemp= Conn.execute(strSql)

		Do while not rsTemp.eof
		
			intSaldo 			= rsTemp("SALDO")
			strNroDoc 			= rsTemp("NRO_DOC")
			strFechaVenc 		= rsTemp("FECHA_VENC")
			strTipoDoc 			= rsTemp("TIPO_DOCUMENTO")
			strNroCuota 		= rsTemp("NRO_CUOTA")
			intAntiguedad 		= ValNulo(rsTemp("ANTIGUEDAD"),"N")
			intIntereses 		= rsTemp("INTERESES")
			intHonorarios 		= rsTemp("HONORARIOS")
			intValorCapital 	= rsTemp("VALOR_CUOTA")
			intAbono 			= (intValorCapital) - (intSaldo)
			intProtestos 		= ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")
			strArrID_CUOTA 		= strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")
			MAX_FECHA_VENC 		= rsTemp("MAX_FECHA_VENC")
			intTotDoc 			= (intSaldo)+(intIntereses)+(intProtestos)+(intHonorarios)
			intTotSelSaldo 		= (intTotSelSaldo)+(intSaldo)
			intTotSelIntereses 	= (intTotSelIntereses)+(intIntereses)
			intTotSelProtestos 	= (intTotSelProtestos)+(intProtestos)
			intTotSelHonorarios = (intTotSelHonorarios)+(intHonorarios)
			intTotSelValorAbono = (intTotSelValorAbono)+(intAbono)
			intTotSelDoc 		= (intTotSelDoc)+(intTotDoc)
			intTotValorCapital  = (intTotValorCapital) + (intValorCapital)
  


			strTextoTercero =rsTemp("RUT_SUBCLIENTE")& ";" &rsTemp("NOMBRE_SUBCLIENTE")& ";" &rut& ";" &NOMBRE_DEUDOR& ";" &strTipoDoc& ";" &strNroDoc& ";" &strNroCuota&";"& strFechaVenc & ";" &intAntiguedad& ";" &FN(intValorCapital,0)& ";" &FN(intAbono,0)


			If Trim(strUsaProtestos)="1" Then
				strTextoTercero= strTextoTercero & ";"&FN(intProtestos,0)
			end if
			if trim(strUsaInteres)="1" then
				strTextoTercero= strTextoTercero & ";" &FN(intIntereses,0)
			END IF 
			if trim(strUsaHonorarios)="1" then
				strTextoTercero= strTextoTercero & ";"&FN(intHonorarios,0)
			end if

			strTextoTercero= strTextoTercero & ";"&FN(intTotDoc,0)

			fichCA.writeline(strTextoTercero)


		rsTemp.movenext
		loop


	'conectamos con el FSO
	set confile = createObject("scripting.filesystemobject")
	'creamos el objeto TextStream

	'response.write "terceroCSV=" & terceroCSV
	'response.End

	''set fichCA = confile.CreateTextFile(terceroCSV)
	''fichCA.write(strTextoTercero)
	fichCA.close()


%>