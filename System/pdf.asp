<!--#include file="arch_utils.asp"-->
<!--#include file="../Componentes/fpdf/fpdf.asp"--> 
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
<%
'Response.CodePage = 65001
'Response.charset="iso-8859-1"

cuotas_rut 					=request.querystring("cuotas_rut")
rut 						=request.querystring("rut")
NOMBRE_DEUDOR 				=request.querystring("NOMBRE_DEUDOR")
fecha_generar_documentos	=request.querystring("fecha_generar_documentos")

fecha_generar_documentos    =replace(fecha_generar_documentos,"/","-")

AbrirSCG()
%>
<html>
<head>
	<title>Generar PDF con ASP</title>
	<meta charset="utf-8"> 
</head>
<body> 
<%


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
		strSql = strSql & " ORDER BY CONVERT(DATETIME, CUOTA.FECHA_VENC)  ASC "
		set rsTemp= Conn.execute(strSql)


	Set pdf=CreateJsObject("FPDF")
	pdf.CreatePDF
	pdf.SetPath("../Componentes/fpdf/fpdf/")
	pdf.SetFont "Arial","",12
	pdf.Open()
	pdf.AddPage()
	pdf.Cell 100,5,"________________________________________________________________________________",0, 0,"L"
	pdf.Ln()
	pdf.SetFont "Arial","B",8
	pdf.Cell 18,5,"LLACRUZ sPa",0 
	pdf.Image "../Imagenes/Logos/"&session("ses_codcli")&"/Logo.jpg",150,20,30,0
	pdf.Ln()
	pdf.SetFont "Arial","",8
	pdf.Cell 18,5,"Direccion :MONSENOR FELIX CABRERA 42, OFICINA 2 PROVIDENCIA",0
	pdf.Ln()
	pdf.Cell 18,5,"R.U.T: 78859630-8",0
	pdf.Ln()
	pdf.Cell 18,4,"Telefono : 28978900",0
	pdf.SetFont "Arial","",12
	pdf.Ln()
	pdf.Cell 100,1,"________________________________________________________________________________",0, 0,"L"
	pdf.Ln()
	pdf.Ln()
	pdf.Ln()
	pdf.Ln()	
	pdf.SetFont "Arial","B",9
	pdf.Cell 80,5,"Nombre : " & trim(NOMBRE_DEUDOR),0
	pdf.Ln()
	pdf.Cell 80,5,"RUT : " & trim(rut),0
	pdf.Ln()
	pdf.Cell 80,5,"Fecha - Hora : " &now(),0
	pdf.Ln()
	pdf.SetFont "Arial","",12
	pdf.Cell 100,5,"________________________________________________________________________________",0, 0,"L"
	pdf.SetFont "Arial","B",10
	pdf.Ln()
	pdf.Cell 185,10,"ANTECEDENTES DE LA DEUDA",0, 0,"C"
	pdf.Ln()
	pdf.SetFont "Arial","",5	
	pdf.Cell 15,5,"RUT CLIENTE",1	
	pdf.Cell 27,5,"NOMBRE CLIENTE",1	
	pdf.Cell 10,5,"N DOC",1	
	pdf.Cell 10,5,"CUOTA",1	
	pdf.Cell 15,5,"FEC.VENC.",1	
	pdf.Cell 10,5,"ANT.",1	
	pdf.Cell 18,5,"TIPO DOC.",1	
	pdf.Cell 15,5,"CAPITAL",1

	if trim(strUsaInteres)="1" then
		pdf.Cell 15,5,"INTERES",1	
	end if
	If Trim(strUsaProtestos)="1" Then
		pdf.Cell 15,5,"PROTESTOS",1	
	end if
	If Trim(strUsaHonorarios)="1" Then
		pdf.Cell 18,5,"HONORARIOS",1	
	end if

	pdf.Cell 15,5,"ABONO",1	
	pdf.Cell 15,5,"SALDO",1	
	pdf.Ln()	
	pdf.SetFont "Arial","I",5


		'Response.write "strSql=" & strSql

		intTasaMensual = 2/100
		intTasaDiaria = intTasaMensual/30
		intCorrelativo = 0
		strArrID_CUOTA=""
		intTotSelSaldo= 0
		intTotSelIntereses= 0
		intTotSelProtestos= 0
		intTotSelHonorarios= 0
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


			pdf.Cell 15,5,rsTemp("RUT_SUBCLIENTE"),1	
			pdf.Cell 27,5,(mid(rsTemp("NOMBRE_SUBCLIENTE"),1,20)),1	
			pdf.Cell 10,5,rsTemp("NRO_DOC"),1	
			pdf.Cell 10,5,rsTemp("NRO_CUOTA"),1	
			pdf.Cell 15,5,rsTemp("FECHA_VENC"),1	
			pdf.Cell 10,5,ValNulo(rsTemp("ANTIGUEDAD"),"N"),1	
			pdf.Cell 18,5,rsTemp("TIPO_DOCUMENTO"),1	
			pdf.Cell 15,5,FN(rsTemp("SALDO"),0),1	

			if trim(strUsaInteres)="1" then
				pdf.Cell 15,5,FN(rsTemp("INTERESES"),0),1
			end if

			If Trim(strUsaProtestos)="1" Then	
				pdf.Cell 15,5,FN(rsTemp("GASTOS_PROTESTOS"),0),1	
			end if

			If Trim(strUsaHonorarios)="1" Then
				pdf.Cell 18,5,FN(rsTemp("HONORARIOS"),0),1	
			end if

			pdf.Cell 15,5, FN(intValorCapital - intSaldo,0),1	
			pdf.Cell 15,5, FN(intTotDoc,0),1	
			pdf.Ln() 
	
	
			intCorrelativo = intCorrelativo + 1

		rsTemp.movenext		
		loop

		formateado_intTotValorCapital 	=FN(intTotValorCapital,0)
		formateado_intTotSelProtestos	=FN(intTotSelProtestos,0)
		formateado_intTotSelIntereses	=FN(intTotSelIntereses,0)	
		formateado_intTotSelHonorarios	=FN(intTotSelHonorarios,0)
		formateado_intTotSelValorAbono 	=FN(intTotSelValorAbono,0)
		formateado_total 				=FN((intTotSelSaldo +(intTotSelIntereses)+(intTotSelHonorarios) + (intTotSelProtestos)), 0)

	pdf.Ln()
	pdf.SetFont "Arial","",10
	pdf.Cell 185,10,"DETALLE",0, 0,"C"
	pdf.Ln()
	pdf.SetFont "Arial","B",9
	pdf.Cell 80,5,"Total documentos: " & intCorrelativo,0
	pdf.Ln()
	pdf.Cell 80,5,"Vencimiento mayor:" & MAX_FECHA_VENC,0
	pdf.Ln()
	pdf.Cell 80,5,"Capital:" & FN(formateado_intTotValorCapital,0),0
	pdf.Ln()
	IF trim(strUsaProtestos)="1" then	
	pdf.Cell 80,5,"Protesto:" & FN(formateado_intTotSelProtestos,0),0
	pdf.Ln()
	end if
	if trim(strUsaInteres)="1" then
		pdf.Cell 80,5,"Interes:" & FN(formateado_intTotSelIntereses,0),0
		pdf.Ln()
	end if
	If Trim(strUsaHonorarios)="1" Then
		pdf.Cell 80,5,"Gastos de cobranza:" & FN(formateado_intTotSelHonorarios,0),0		
		pdf.Ln()    
	end if
	pdf.Cell 80,5,"Abonado:" & FN(formateado_intTotSelValorAbono,0),0		
	pdf.Ln()
	pdf.SetFont "Arial","",12
	pdf.Cell 100,5,"________________________________________________________________________________",0, 0,"L"
	pdf.Ln()
	pdf.SetFont "Arial","B",9
	pdf.Cell 80,5,"Total a pagar:" & FN(formateado_total,0),0		
	 
	filewrite=Server.MapPath("../Archivo/BibliotecaNotificacionDeudores/"&trim(session("ses_codcli"))&"/"&rut&"/"&fecha_generar_documentos&"_detalle_deuda.pdf")
	pdf.Output filewrite
	'pdf.Output()
	pdf.Close()

%> 
</body>
</html>