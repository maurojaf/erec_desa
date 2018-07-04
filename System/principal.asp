<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->
	<!--#include file="../lib/asp/comunes/general/SoloNumeros.inc" -->
<%
	Response.CodePage 	=65001
	Response.charset	="utf-8"
%>
	<link rel="stylesheet" href="../css/style_generales_sistema.css">

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
	<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<script language="javascript">

$(document).ready(function(){

	$.prettyLoader();
 	
})

function saveScroll()
{
var pos = document.getElementById('hdnScrollPos');
pos.value = event.srcElement.scrollTop;
//alert(pos.value);
}

function gridLoad(gridID, posID)
{
var el = document.getElementById(gridID);
el.scrollTop = document.getElementById(posID).value;
}
function ventanaConvenio (URL){
window.open(URL,"INFORMACION","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}
function envia_agendamiento(){
$.prettyLoader.show(2000);
datos.BT_CARTERA.disabled = true;
datos.action='modulo_agendamientos.asp?';
datos.submit();
}

</script>

<form name="datos" id="datos" method="post">
<%
		If Request("cliente") <> "" then 
			intCodCliente = Request("cliente")
			session("ses_codcli") = intCodCliente
		else
			intCodCliente 	= session("ses_codcli")
		end if
		
		strUsuario		= session("session_idusuario")
		strRutDeudor 	= Trim(Request("TX_RUT"))
		strSelOP 		= Request("OP_SEL")

		If Trim(strSelOP) = ""  Then strSelOP 	= "1"
		If Trim(strSelOP) = "1" Then strSel1 	= "checked"
		If Trim(strSelOP) = "2" Then strSel2 	= "checked"

		strRutSubClienteSel = Request("strRutSubCliente")
		strRUT_DEUDORNSel 	= Request("strRUT_DEUDORSel")
		tipo_rut 			= Request("tipo_rut")

		if trim(tipo_rut)="" then
			tipo_rut	="3"
		end if

		if trim(tipo_rut)="1" then		

			Cadena1=strRutDeudor
			Cadena2="-"
			If InStr(Cadena1,Cadena2)<0 then
				strRutDeudor = mid(TRIM(strRutDeudor), 1 ,len(TRIM(strRutDeudor))-1) &"-"& mid(TRIM(strRutDeudor), len(TRIM(strRutDeudor)) , 1)
			Else
				strRutDeudor = replace(strRutDeudor,"-","")

				strRutDeudor = mid(TRIM(strRutDeudor), 1 ,len(TRIM(strRutDeudor))-1) &"-"& mid(TRIM(strRutDeudor), len(TRIM(strRutDeudor)) , 1)
			End if


			tipo_rut	="3"

		end if

		if trim(tipo_rut)="2" then

			Cadena1=strRutDeudor
			Cadena2="-"
			If InStr(Cadena1,Cadena2)>0 then
				strRutDeudor = mid(strRutDeudor,1,len(strRutDeudor)-2)

				tur=strreverse(strRutDeudor) 
				mult = 2 

				for i = 1 to len(tur) 
					if mult > 7 then mult = 2 end if 

					suma = mult * mid(tur,i,1) + suma 
					mult = mult +1 
				next 

				valor = 11 - (suma mod 11) 

				if valor = 11 then 
					codigo_veri = "0" 
				elseif valor = 10 then 
					codigo_veri = "k" 
				else 
					codigo_veri = valor 
				end if 

				strRutDeudor =strRutDeudor&"-"&codigo_veri

				tipo_rut	="3"

			else

				tur =strreverse(strRutDeudor) 
				mult = 2 

				for i = 1 to len(tur) 
					if mult > 7 then mult = 2 end if 

					suma = mult * mid(tur,i,1) + suma 
					mult = mult +1 
				next 

				valor = 11 - (suma mod 11) 

				if valor = 11 then 
					codigo_veri = "0" 
				elseif valor = 10 then 
					codigo_veri = "k" 
				else 
					codigo_veri = valor 
				end if 

				strRutDeudor =strRutDeudor&"-"&codigo_veri

				tipo_rut	="3"

			End if
			
		end if

		if trim(tipo_rut)="3" then
			strRutDeudor	= strRutDeudor	
		end if		


		if trim(strRutDeudor) <> "" Then
			session("session_RUT_DEUDOR") = strRutDeudor
		Else
			strRutDeudor 	= session("session_RUT_DEUDOR")
		End if

		If Trim(Request("Limpiar"))="1" Then
			session("session_RUT_DEUDOR") = ""
			strRutDeudor = ""
		End if

		strFiltro 	= Request("strFiltro")
		If Trim(strFiltro) = "" Then strFiltro ="0"

		If Trim(strFiltro) = "1" Then
			strTipoDeuda = "SALDADA"
		Else
			strTipoDeuda = "ACTIVA"
		End if

		direccion_val 	=request("radiodir")
		fono_val 		=request("radiofon")
		email_val 		=request("radiomail")
		cor_tel 		=request("correlativo_fono")
		cor_dir 		=request("correlativo_direccion")
		cor_cor 		=request("correlativo_mail")
		var_dir 		=request("dir")
		var_fon 		=request("fon")
		var_mail 		=request("mail")



		AbrirSCG()

		If var_fon <> "" then
			if fono_val <> "" and Not IsNull(fono_val) then
				strSql=""
				strSql="UPDATE DEUDOR_TELEFONO SET estado= " & fono_val  & ", FECHA_REVISION = getdate(), USR_REVISION = '" & session("session_login") & "' WHERE CORRELATIVO= " & cint(cor_tel)

					strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "

				'Response.write strSql
				Conn.execute(strSql)
			end If
		End if

		If var_dir <> "" then
			if direccion_val <> "" and Not IsNull(direccion_val) then
				strSql=""
				strSql="UPDATE DEUDOR_DIRECCION SET estado= " & direccion_val & ", FECHA_REVISION = getdate(), USR_REVISION = '" & session("session_login") & "' CORRELATIVO='"&cint(cor_dir)&"'"

					strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "

				Conn.execute(strSql)
			end If
		End if

		If var_mail <> "" then
			if email_val <> "" and Not IsNull(email_val) then
				strSql=""
				strSql="UPDATE DEUDOR_EMAIL SET estado= "& email_val  & ", FECHA_REVISION = getdate(), USR_REVISION = '" & session("session_login") & "' WHERE CORRELATIVO='"&cint(cor_cor)&"'"

					strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "

				Conn.execute(strSql)
			end If
		End If


		AbrirSCG1()
			strSql = "SELECT ID_DEUDOR ,RUT_DEUDOR,NOMBRE_DEUDOR, RESP_EMAIL, OBSERVACIONES_CONF, FECHA_UG_TITULAR, FECHA_CONF, USUARIO_CONF , IsNull(datediff(minute,FECHA_CONF,IsNull(FECHA_UG_TITULAR,'01/01/1900')),0) as DIFMINUTOS FROM DEUDOR WHERE COD_CLIENTE = '" & intCodCliente & "' "

					strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "


			'Response.write "<br>strSql=" & strSql
			set RsDeudor=Conn1.execute(strSql)

			If not RsDeudor.eof then
				existe = "si"
				strRespPriorizacion = RsDeudor("RESP_EMAIL")
				rut_deudor = RsDeudor("RUT_DEUDOR")
				nombre_deudor = RsDeudor("NOMBRE_DEUDOR")
				strObsConf = RsDeudor("OBSERVACIONES_CONF")
				strFechaConf = RsDeudor("FECHA_CONF")
				strUsuarioConf = RsDeudor("USUARIO_CONF")
				intMinDif = RsDeudor("DIFMINUTOS")
				intId_Deudor = RsDeudor("ID_DEUDOR")

				If Trim(strFechaConf) <> "" and Trim(strUsuarioConf) <> "" then
					strTextoConf = "Fecha : " & strFechaConf & " , Usuario : " & strUsuarioConf & ", Obs : "
				End If

			End If

		CerrarSCG1()

		AbrirSCG1()
			strSql = "SELECT PR.ID_PRIORIZACION,PR.OBSERVACION_PRIORIZACION, USUARIO.LOGIN, PR.FECHA_PRIORIZACION,TSP.NOM_TIPO_SOLICITUD"
			strSql= strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
			strSql= strSql & " 					 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
			strSql= strSql & " 					 INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
			strSql= strSql & " 					 INNER JOIN USUARIO ON PR.ID_USUARIO_PRIORIZACION = USUARIO.ID_USUARIO"
			strSql= strSql & " 					 INNER JOIN TIPO_SOLICITUD_PRIORIZACION TSP ON TSP.ID_TIPO_SOLICITUD = PR.ID_TIPO_SOLICITUD"



					strSql = strSql & " WHERE PR.RUT_DEUDOR='"&strRutDeudor&"' "



			strSql = strSql &  " AND PRC.ESTADO_PRIORIZACION = 0 AND ESTADO_DEUDA.ACTIVO = 1"
			strSql= strSql & " 					 GROUP BY PR.ID_PRIORIZACION,PR.OBSERVACION_PRIORIZACION,USUARIO.LOGIN, PR.FECHA_PRIORIZACION, TSP.NOM_TIPO_SOLICITUD"

			'Response.write "<br>strSql=" & strSql
			set RsPrio=Conn1.execute(strSql)

			If not RsPrio.eof then

				intEstadoPrior = 0

				Do While Not RsPrio.Eof

				intIdPriorizacion = RsPrio("ID_PRIORIZACION")

				strTotalDoc = ""

				AbrirSCG2()
							strSql = "SELECT CUOTA.NRO_DOC, (CASE WHEN (DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0) OR [dbo].[fun_trae_fecha_ult_gestion] (CUOTA.ID_CUOTA) < PR.FECHA_PRIORIZACION THEN 1 ELSE 0 END) AS AGEND_PRIO"
							strSql= strSql & " FROM PRIORIZACION PR INNER JOIN PRIORIZACIONES_CUOTA PRC ON PR.ID_PRIORIZACION = PRC.ID_PRIORIZACION"
							strSql= strSql & " 					 INNER JOIN CUOTA ON CUOTA.ID_CUOTA = PRC.ID_CUOTA"
							strSql= strSql & " 					 INNER JOIN USUARIO ON PR.ID_USUARIO_PRIORIZACION = USUARIO.ID_USUARIO"

							strSql= strSql & " WHERE PRC.ID_PRIORIZACION = '" & intIdPriorizacion & "' AND PRC.ESTADO_PRIORIZACION = 0"

							'Response.write "<br>strSql=" & strSql
							set RsPrioDoc=Conn2.execute(strSql)

							If not RsPrioDoc.eof then

							intAgendPrio = 0

								Do While Not RsPrioDoc.Eof

									strDoc = RsPrioDoc("NRO_DOC")
									strTotalDoc = strTotalDoc & "-" & strDoc
									intAgendPrio = intAgendPrio + RsPrioDoc("AGEND_PRIO")

									RsPrioDoc.movenext
								Loop
							End If

				CerrarSCG2()

					strObsPrio = RsPrio("OBSERVACION_PRIORIZACION")
					strUsuarioPrio = RsPrio("LOGIN")
					strFechaPrio = RsPrio("FECHA_PRIORIZACION")
					strTipoSol = RsPrio("NOM_TIPO_SOLICITUD")

					If Trim(strTotalDoc) <> "" Then
						strTotalDoc = "Doc: " & Mid(strTotalDoc,2,Len(strTotalDoc))
					End If

					If Trim(strFechaPrio) <> "" and Trim(strUsuarioPrio) <> "" then
						strTextoPrio = "Fecha: " & strFechaPrio & " , Usuario : " & strUsuarioPrio & chr(13) & "Tipo Sol: " & strTipoSol & chr(13) & "Obs : " & strObsPrio & chr(13) & strTotalDoc & chr(13) & chr(13)

						strTextoPrioF = strTextoPrioF & strTextoPrio
					End If

					RsPrio.movenext
				Loop
			End If

		CerrarSCG1()


		AbrirSCG1()
			strSql = "SELECT TOP 1 [dbo].[fun_PrioridadCuotaDocActivo] ('" & intCodCliente & "', RUT_DEUDOR, 0) as PRIORIDAD  FROM DEUDOR "

					strSql = strSql & " WHERE DEUDOR.RUT_DEUDOR='"&strRutDeudor&"' "



			set rsPrioridad=Conn1.execute(strSql)
			If not rsPrioridad.eof then
				strPrioridad = CCur(rsPrioridad("PRIORIDAD"))
			Else
				strPrioridad = 99
			End if
			rsPrioridad.close
			set rsPrioridad=nothing
		CerrarSCG1()


		'---CUENTA LA CANTDAD DE SUBCLIENTES ACTIVOS----'

		AbrirSCG1()

			strSql="SELECT IsNull(COUNT(DISTINCT RUT_SUBCLIENTE),0) AS CANT FROM VW_DEUDOR_SUBCLIENTE WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE = '" & intCodCliente & "' AND RUT_SUBCLIENTE IN (SELECT RUT_SUBCLIENTE FROM CUOTA "

					strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "



			strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1))"
			'Response.write "strSql=" &strSql
			set RsDeudor=Conn1.execute(strSql)
			If not RsDeudor.eof then
				intCantidadRSAct = RsDeudor("CANT")
			Else
				intCantidadRSAct = 0
			End if
			RsDeudor.close
			set RsDeudor=nothing

		CerrarSCG1()

		'---CUENTA LA CANTDAD DE SUBCLIENTES NO ACTIVOS----'

		AbrirSCG1()

			strSql="SELECT IsNull(COUNT(DISTINCT RUT_SUBCLIENTE),0) AS CANT FROM VW_DEUDOR_SUBCLIENTE WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE = '" & intCodCliente & "' AND RUT_SUBCLIENTE IN (SELECT RUT_SUBCLIENTE FROM CUOTA "

					strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "
	
			strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 0))"

			'REsponse.write "strSql=" &strSql
			set RsDeudor=Conn1.execute(strSql)
			If not RsDeudor.eof then
				intCantidadRSNAct = RsDeudor("CANT")
			Else
				intCantidadRSNAct = 0
			End if
			RsDeudor.close
			set RsDeudor=nothing

		CerrarSCG1()

		'---CUENTA LA EL TOTAL DE SUBCLIENTES----'

		AbrirSCG1()

			strSql="SELECT IsNull(COUNT(DISTINCT RUT_SUBCLIENTE),0) AS CANT FROM VW_DEUDOR_SUBCLIENTE "


					strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "


			strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "'"
			'REsponse.write "strSql=" &strSql
			set RsDeudor=Conn1.execute(strSql)
			If not RsDeudor.eof then
				intTotalRS = RsDeudor("CANT")
			Else
				intTotalRS = 0
			End if
			RsDeudor.close
			set RsDeudor=nothing

		CerrarSCG1()

		'---CUENTA LA CANTDAD DE SUBCLIENTES ACTIVOS/NO ACTIVOS SEGUN FILTRO DE DEUDA (strFiltro = 0: DEUDA ACTIVA, strFiltro = 1: DEUDA NO ACTIVA)----'

		AbrirSCG1()

			strSql="SELECT IsNull(COUNT(DISTINCT RUT_SUBCLIENTE),0) AS CANT FROM VW_DEUDOR_SUBCLIENTE "


					strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "


			strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "'"

			If Trim(strFiltro) = "0" Then
				strSql= strSql & " AND RUT_SUBCLIENTE IN (SELECT RUT_SUBCLIENTE FROM CUOTA "


					strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "


				strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1))"
			Else
				strSql= strSql & " AND RUT_SUBCLIENTE IN (SELECT RUT_SUBCLIENTE FROM CUOTA "


					strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "



				strSql = strSql &  " AND COD_CLIENTE = '" & intCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 0))"
			End If

			'Response.write "strSql=" &strSql

			set RsDeudor=Conn1.execute(strSql)

			If not RsDeudor.eof then
				intCantidadRSFiltro = RsDeudor("CANT")
			Else
				intCantidadRSFiltro = 0
			End if
			RsDeudor.close
			set RsDeudor=nothing

		CerrarSCG1()


		'---CUENTA LA EL TOTAL DE DOCUMENTOS ASOCIADOS A ESE DEDUDOR CLIENTE----'

		AbrirSCG1()

			strSql="SELECT IsNull(COUNT(*),0) AS CANT FROM CUOTA "

				strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "
		

			strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "'"
			'REsponse.write "strSql=" &strSql
			set RsDeudor=Conn1.execute(strSql)
			If not RsDeudor.eof then
				intTotalDoc = RsDeudor("CANT")
			Else
				intTotalDoc = 0
			End if
			RsDeudor.close
			set RsDeudor=nothing

		CerrarSCG1()


		'---CUENTA EL TOTAL DE DOCUMENTOS ACTIVOS----'

		AbrirSCG1()

			strSql="SELECT IsNull(COUNT(*),0) AS CANT FROM CUOTA "

				strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "
		

			strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
			'REsponse.write "strSql=" &strSql
			set RsDeudor=Conn1.execute(strSql)
			If not RsDeudor.eof then
				intCantidadDocAct = RsDeudor("CANT")
			Else
				intCantidadDocAct = 0
			End if
			RsDeudor.close
			set RsDeudor=nothing

		CerrarSCG1()


		'---CUENTA EL TOTAL DE DOCUMENTOS NO ACTIVOS----'

		AbrirSCG1()

			strSql="SELECT IsNull(COUNT(*),0) AS CANT FROM CUOTA "


				strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "
	

			strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 0)"
			'REsponse.write "strSql=" &strSql
			set RsDeudor=Conn1.execute(strSql)
			If not RsDeudor.eof then
				intCantidadDocNoAct = RsDeudor("CANT")
			Else
				intCantidadDocNoAct = 0
			End if
			RsDeudor.close
			set RsDeudor=nothing

		CerrarSCG1()

		'---TRAE LA INFORMACION ASOCIADA A LA FECHA DEL DEUDOR - SUB_CLIENTE----'

		AbrirSCG1()

			strSql="SELECT OBSERVACION FROM INFORMACION_DEUDOR_SUBCLIENTE "



				strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "

									

			strSql = strSql & " AND COD_CLIENTE = '" & intCodCliente & "' AND OBSERVACION IS NOT NULL"
			'REsponse.write "strSql=" &strSql
			set RsInfDS=Conn1.execute(strSql)
			If not RsInfDS.eof then
				rsInfDeudor = "1" 
			Else
				rsInfDeudor = "0" 
			End if
			RsInfDS.close
			set RsInfDS=nothing

		CerrarSCG1()
		
		If intCantidadRSFiltro = 0 Then

			strTamanoHScrool=70

		ElseIf intCantidadRSFiltro >= 1 Then

			intTotalTamano = 30 + intCantidadRSFiltro * 17

			If intTotalTamano <= 110 then

			strTamanoHScrool= intTotalTamano

			Else

			strTamanoHScrool= 110

			End If

		End If

		'Response.write "<br>intCantidadRSFiltro =========" & intCantidadRSFiltro
		'Response.write "<br>intAgendPrio=========" & intAgendPrio
 
		If strRutDeudor = "" then
			intRutSelNOk = "1"
		Else			
			strSql="SELECT RUT_DEUDOR FROM DEUDOR" 
			strSql= strSql & " WHERE COD_CLIENTE = '" & intCodCliente & "' "


				strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "

						

			'Response.write "<br>strSql=" & strSql

			set rsApertura=Conn.execute(strSql)
			If Not rsApertura.Eof Then
				intRutSelNOk="0"
			Else
				intRutSelNOk="1"
			End if
			
		End If
		
		abrirSCG1()
			strSql="SELECT * FROM CLIENTE WHERE COD_CLIENTE = '" & intCodCliente & "'"
			'REsponse.write "strSql=" & strSql
			set rsCliente=Conn1.execute(strSql)
			if not rsCliente.eof then

				strUsaSubCliente = ValNulo(rsCliente("USA_SUBCLIENTE"),"N")
				strUsaInteres = ValNulo(rsCliente("USA_INTERESES"),"N")
				strUsaHonorarios = ValNulo(rsCliente("USA_HONORARIOS"),"N")
				strUsaProtestos = ValNulo(rsCliente("USA_PROTESTOS"),"N")


				intTasaMensual = ValNulo(rsCliente("INTERES_MORA"),"C")
				intMesHon = ValNulo(rsCliente("MESES_TD_HON"),"N")
				strNomFormHon = ValNulo(rsCliente("FORMULA_HONORARIOS"),"C")
				strNomFormInt = ValNulo(rsCliente("FORMULA_INTERESES"),"C")
				
				strNomAdic1Deudor =  ValNulo(rsCliente("ADIC1_DEUDOR"),"C")
				strNomAdic2Deudor =  ValNulo(rsCliente("ADIC2_DEUDOR"),"C")
				strNomAdic3Deudor =  ValNulo(rsCliente("ADIC3_DEUDOR"),"C")
				strUsaCustodio = Trim(rsCliente("USA_CUSTODIO"))
				intTipoNegocio = Trim(rsCliente("TIPO_NEGOCIO"))
				
				'Response.write "intTipoNegocio====" &intTipoNegocio
				
			end if
			If intTasaMensual = "" Then
				%>
				<SCRIPT>alert('No se ha definido tasa de interes de mora, se ocupara una tasa del 2%, favor parametrizar')</SCRIPT>
				<%
				intTasaMensual = "2"
			End If
			rsCliente.close
			set rsCliente=nothing
		cerrarSCG1()
								
%>



	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
        <td colspan="2" align="right">
          <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
            <script type="text/javascript" language="JavaScript1.2" src="frameset/stm31.js"></script>
            <tr valign="TOP">
              <td colspan="3">
                  <table width="100%" height="335" border="0" cellpadding="0" cellspacing="0" >
                    <tr>
                      <td height="331" valign="top" background="../imagenes/fondo_111.jpg">
                      <table width="100%" border="0">
                          <tr>
                            <td>
								<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
									<tr>
										<TD height="20" colspan="2" class="subtitulo_informe">
											> RUT DE BUSQUEDA
										</TD>
										<TD ALIGN="RIGHT"></TD>
									</tr>
								</table>
								<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">

									 <tr>
										<td align="left">
											<acronym title="">
											  <input name="TX_RUT" type="text" id="TX_RUT" value="<%=strRutDeudor%>" size="15" maxlength="10">
											</acronym>
											&nbsp;&nbsp; 
											<acronym title="DESPLEGAR DATOS DEL DEUDOR ASOCIADO AL RUT DE B&Uacute;SQUEDA">
											  <input name="me_"  class="fondo_boton_100" type="button" id="me_" onClick="envia();" value="Buscar">
											</acronym>
											&nbsp;&nbsp; 
											<acronym title="LIMPIAR FORMULARIO">
											  <input name="li_"  class="fondo_boton_100" type="button" onClick="window.navigate('principal.asp?Limpiar=1');" value="Limpiar">
											 </acronym>
										</td>
										<td align="right">
										<%
												'Response.write "<br>intEstadoPrior=" & intEstadoPrior
												'Response.write "<br>Comp=" & (Cdbl(intMinDif) >= 0)

										If intRutSelNOk = "0" Then
										
										strColorPrio="boton_azul"
										strTextoBPrior="Priorizar caso"

											If intAgendPrio > "0" or  intEstadoPrior = "0"  then 
												strColorPrio="boton_rojo"
												strTextoBPrior="Caso Priorizado"
											End If%>
											
											<%If intAgendPrio > "0" then %>
												<abbr title="<%=strTextoPrioF%>">
												<img src="../imagenes/priorizar_urgente.png" border="0">
												<abbr>
											<%ElseIf intEstadoPrior = "0" Then%>
												<abbr title="<%=strTextoPrioF%>">
												<img src="../imagenes/priorizar_normal.png" border="0">
											<%End If%>

											<input name="BT_PRIORIZAR" class="<%=strColorPrio%> fondo_boton_130" type="button" id="BT_PRIORIZAR" onClick="priorizar_caso();" value="<%=strTextoBPrior%>">

										<%End if%>																					
										</td>

									</tr>
									<tr>
										<td align="left" height="25">
											<input type="radio" name="tipo_rut" <%if trim(tipo_rut)="3" then response.write " checked " end if%> id="tipo_rut" value="3">Con  guión
											&nbsp;&nbsp;
											<input type="radio" name="tipo_rut" <%if trim(tipo_rut)="1" then response.write " checked " end if%> id="tipo_rut" value="1">Sin guión
											&nbsp;&nbsp;
											<input type="radio" <%if trim(tipo_rut)="2" then response.write " checked " end if%> name="tipo_rut" id="tipo_rut" value="2">Sin dígito verificador
											&nbsp;&nbsp;
										</td>
										<td>&nbsp;</td>
									</tr>
									<tr>
		
										<TD  align="right" height="20" colspan="2">

										<%If intRutSelNOk = "0" then %>

											<TABLE class="" height="30" width="800" border="0" cellpadding="0" cellspacing="0" >
												<TR >
													<TD width="100%" COLSPAN="5" height="10" style="vertical-align: top;">
													<table border="0" cellpadding="0" cellspacing="0" width="100%">
														<TR>

															<TD align="right">
																<input name="BT_FichaDeudor" class="fondo_boton_100" type="button" id="BT_FichaDeudor" onClick="javascript:VentanaFichaDeudor('FichaDeudor.asp?CodigoCliente=<%=intCodCliente%>&RutDeudor=<%=strRutDeudor%>&CodigoUsuario=<%=strUsuario%>');" value="Ficha Deudor">
																<script language="JavaScript" type="text/JavaScript">
																	function VentanaFichaDeudor(URL){
																		window.open(URL,"DATOS","width=818, height=500, scrollbars=yes, menubar=no, location=no, resizable=no")
																	}
																</script>
															</TD>	

															<td align="right"><input name="BT_BackOffice" class="fondo_boton_100" type="button" id="BT_BackOffice" onClick="envia_backoffice();" value="BackOffice"><td>

															<%'If (TraeSiNo(session("perfil_full")) = "Si" or TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_emp")) = "Si" or TraeSiNo(session("perfil_caja")) = "Si") and (intCodCliente = "1070" or intCodCliente = "1500")  Then%>
																<!--<td align="right"><input name="BT_PAGOS" class="fondo_boton_100" type="button" id="BT_PAGOS" onClick="envia_caja();" value="Ing.Pagos"><td>-->

															<%'End If%>

															<%If TraeSiNo(session("perfil_full")) = "Si" or TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_emp")) = "Si" or TraeSiNo(session("perfil_caja")) = "Si" Then%>
																<td align="right"><input name="BT_INFO" class="fondo_boton_100" type="button" id="BT_INFO" onClick="envia_info();" value="Info. Pago"><td>
																<%If TraeSiNo(session("perfil_caja")) = "Si" Then%>
																	<td align="right"><input name="BT_CONVENIO" class="fondo_boton_100" type="button" id="BT_CONVENIO" onClick="envia_convenio();" value="<%=session("NOMBRE_CONV_PAGARE")%>"><td>
																<%End If%>
															<%End If%>

																<td align="right"><input name="BT_PLANPAGO" class="fondo_boton_100" type="button" id="BT_PLANPAGO" onClick="envia_plandepago();" value="Plan de Pago"><td>
															
															
															<td align="right"><input name="BT_BIBLIOTECA" class="fondo_boton_100" type="button" id="BT_BIBLIOTECA" onClick="javascript:ventanaBiblioteca('biblioteca_deudores.asp?strRut=<%=strRutDeudor%>');" value="Biblioteca"><td>
															<td align="right"><input name="BT_CARTERA" class="fondo_boton_100" type="button" id="BT_CARTERA" onClick="envia_agendamiento();" value="Agendamiento"><td>
														</TR>
													</table>
													</TD>
												</TR>
											</TABLE>
										<% Else %>

										<% End if %>
										</TD>
                          			</tr>
								</table>
                            </td>
                          </tr>

                        </table>

        <%
		if strRutDeudor <> "" and not isnull(strRutDeudor) then

			if existe = "si" and intTotalDoc > 0 then


				if Trim(rut_deudor) <> ""  then %>
				<div class="subtitulo_informe">> INFORMACIÓN DEL CLIENTE</div>

				<INPUT id="hdnScrollPos" type="hidden" NAME="hdnScrollPos">
				<div id="divScroll" style="overflow:auto; width:100%; height:<%=strTamanoHScrool%>px;" onscroll="saveScroll()">


				<table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">


				<%If intCantidadRSFiltro > 0 then%>
					<thead>
					<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<td width="10">&nbsp;</td>
						<td align="left" width="40">RUT</td>
						<td align="left" width="500">NOMBRE O RAZ&Oacute;N SOCIAL</td>
					</tr>
					</thead>

				<%Else%>


					<tr>
						<td width="100%" height="20" align="center" class="estilo_columna_individual">
							<FONT SIZE='2' color="#FF0000" ><B>DEUDOR NO POSEE CLIENTES CON DEUDA <%=strTipoDeuda%></B></FONT>
						</td>
					</tr>


				<%End If

						'----VERIFICA SI EL RUT SUBCLIENTE ESTA CONTENIDO EN LA SELECCION SEGÚN FILTRO DE DEUDA----'

						AbrirSCG()

						strSql = "SELECT COUNT(*) AS CUENTA FROM VW_DEUDOR_SUBCLIENTE"
			

							strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "


						strSql= strSql & "  AND RUT_SUBCLIENTE = '" & strRutSubClienteSel & "'"

						If strFiltro ="0" Then

						strSql= strSql & " AND RUT_SUBCLIENTE IN (SELECT RUT_SUBCLIENTE FROM CUOTA INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO WHERE COD_CLIENTE = '" & intCodCliente & "' "


								strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "

						

							strSql = strSql & " AND ESTADO_DEUDA.ACTIVO = 1)"

						Else

						strSql= strSql & " AND RUT_SUBCLIENTE IN (SELECT RUT_SUBCLIENTE FROM CUOTA INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO WHERE COD_CLIENTE = '" & intCodCliente & "' "


								strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "
	

						 strSql = strSql & " AND ESTADO_DEUDA.ACTIVO = 0)"

						End If

						'Response.write "strSql=" & strSql

						set RsSCliChek=Conn.execute(strSql)

						intRurSubClienteCont = RsSCliChek("CUENTA")

						CerrarSCG()


						'Response.write "intRurSubClienteCont=====" & intRurSubClienteCont
						'Response.write "intCantidadRSFiltro=====" & intCantidadRSFiltro

						AbrirSCG()

						strSql = "SELECT RUT_SUBCLIENTE, MAX(NOMBRE_SUBCLIENTE) as NOMBRE_SUBCLIENTE FROM VW_DEUDOR_SUBCLIENTE"

								strSql = strSql & " WHERE RUT_DEUDOR='"&strRutDeudor&"' "


						If strFiltro ="0" Then

						strSql= strSql & " AND RUT_SUBCLIENTE IN (SELECT RUT_SUBCLIENTE FROM CUOTA INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO WHERE COD_CLIENTE = '" & intCodCliente & "'  "


								strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "


						 strSql = strSql & " AND ESTADO_DEUDA.ACTIVO = 1)"

						Else

						strSql= strSql & " AND RUT_SUBCLIENTE IN (SELECT RUT_SUBCLIENTE FROM CUOTA INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO WHERE COD_CLIENTE = '" & intCodCliente & "' "


								strSql = strSql & " AND RUT_DEUDOR='"&strRutDeudor&"' "


						 strSql = strSql & " AND ESTADO_DEUDA.ACTIVO = 0)"

						End If

						strSql= strSql & " GROUP BY RUT_SUBCLIENTE"

						'Response.write "strSql=" & strSql

						set RsSCli=Conn.execute(strSql)

						intPrimero = 1
						Total = 0

						Do While Not RsSCli.eof
							strRutSubCliente = RsSCli("RUT_SUBCLIENTE")
							strNombreSubCliente = RsSCli("NOMBRE_SUBCLIENTE")

							Total = total + 1

							'Response.write "<br>strRutSubClienteSel=====" & strRutSubClienteSel
							'Response.write "<br>strRutSubCliente=====" & strRutSubCliente

							If Trim(strRutSubClienteSel) = "" or intCantidadRSFiltro = 0 or intRurSubClienteCont = 0 Then
								If Trim(intPrimero) = 1 Then
									strColorSubCliente = "6AC982"
									strRutSubClienteSel = strRutSubCliente
								Else
									strColorSubCliente = session("COLTABBG2")
								End If
							Else
								If strRutSubClienteSel = strRutSubCliente Then
									strColorSubCliente = "6AC982"
									strRutSubClienteSel = strRutSubCliente
								Else
									strColorSubCliente = session("COLTABBG2")
								End If
							End If
					%>

					<tr bgcolor="#<%=strColorSubCliente%>" class="Estilo8">
						<td width="20" whith align="CENTER">
							<a href="#" onClick="Seleccionar('<%=strRutSubCliente%>','<%=strFiltro%>');return false;" style="text-decoration: none;"><img src="../imagenes/icon_aceptar.gif" border="0" width="12" height="12"></a>
						</td>
						<td align="left"><%=strRutSubCliente%></td>
						<td align="left"><%=strNombreSubCliente%></td>

					</tr>

					<%
							intPrimero = 2
							RsSCli.movenext
						Loop
						CerrarSCG()
					%>
					</table>


				</div>
				<%
				abrirSCG()

				strSql = "proc_Parametros_Tabla_Deudor '"&TRIM(intCodCliente)&"','"&TRIM(strRutDeudor)&"'"

					''REsponse.write "<br>strSql=====" & strSql

					set RsDeudor=Conn.execute(strSql)

					''REsponse.write "<br><br>rut_deudor=" & rut_deudor
				%>

				<table width="100%" border="0" bordercolor="#FFFFFF">
					<tr>
						<TD><div class="subtitulo_informe">> INFORMACIÓN DEL DEUDOR (PRIORIDAD <%=strPrioridad%>)</div>
						</TD>
						<TD height="20" ALIGN="RIGHT" class="">
							<% If Trim(strFiltro) = "1" Then
							%>
								
								<input type="button" class="fondo_boton_130" name="" value="Ver deuda activa" onclick="location.href='principal.asp?strFiltro=0&strRutSubCliente=<%=strRutSubClienteSel%>'">
								
							<%Else
							%>
								<input type="button" class="fondo_boton_130" name="" value="Ver deuda saldada" onclick="location.href='principal.asp?strFiltro=1&strRutSubCliente=<%=strRutSubClienteSel%>'">


							<%End if%>
						</TD>

					</tr>
				</table>
				<table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">
					<thead>
					<tr>
						<td>&nbsp;</td>
						<td align="left" width="">ID</td>
						<td align="left" width="80">RUT</td>
						<td width="300" align="left">NOMBRE O RAZÓN SOCIAL</td>
						<td width="100" align="left">RUT REP.LEG.</td>
						<td width="300" align="left">NOMBRE RUT REPRESENTANTE LEGAL</td>
						<td align="center"><b>CAMPAÑA</b></td>
						<td align="center"><b>TRAMO VENCIMIENTO</b></td>
						<td align="center"><b>TRAMO MONTO</b></td>
						<td align="center"><b>TRAMO ASIGNACIÓN</b></td>

					</tr>
					</thead>

					<%

					intPrimeroDeudor = 1


					Do While not RsDeudor.eof

						strRutRepLegal = RsDeudor("REPLEG_RUT")
						strNombreRepLegal = RsDeudor("REPLEG_NOMBRE")
						strCampana = RsDeudor("NOMBRE_CAMPANA")
						nombre_deudor = RsDeudor("NOMBRE_DEUDOR")
						rut_deudor = RsDeudor("RUT_DEUDOR")
						intIdDeudor = RsDeudor("ID_DEUDOR")
						strRUT_DEUDORN = RsDeudor("RUT_DEUDOR")
						strNombreDeudorN = RsDeudor("NOMBRE_DEUDOR")
						strTrmoVencimiento 	= Ucase(RsDeudor("TRAMO_VENC"))
						strTramoMonto 	= Ucase(RsDeudor("TRAMO_MONTO"))
						strTramoAsignacion 	= Ucase(RsDeudor("TRAMO_ASIG"))
		
						'Response.write "strRUT_DEUDORNSel=====" & strRUT_DEUDORNSel
						'Response.write "strRUT_DEUDORN=====" & strRUT_DEUDORN

						If Trim(strRUT_DEUDORNSel) = "" Then
							If Trim(intPrimeroDeudor) = 1 Then
								strColorDeudor = "6AC982"
								strRUT_DEUDORNSel = strRUT_DEUDORN
								rut_deudor = strRUT_DEUDORNSel
							Else
								strColorDeudor = session("COLTABBG2")
							End If
						Else
							If strRUT_DEUDORNSel = strRUT_DEUDORN Then
								strColorDeudor = "6AC982"
								strRUT_DEUDORNSel = strRUT_DEUDORN
								rut_deudor = strRUT_DEUDORNSel
							Else
								strColorDeudor = session("COLTABBG2")
							End If
						End If

						rut_deudor = strRUT_DEUDORNSel
						'Response.write "rut_deudor======" & rut_deudor
						'Response.write "strRUT_DEUDORNSel======" & strRUT_DEUDORNSel

					%>

					<tr bgcolor="#<%=strColorDeudor%>" class="Estilo8">
						<td align="CENTER">
							
						</td>
						<td align="left"><%=intIdDeudor%></td>
						<td align="left"><%=strRUT_DEUDORN%></td>

						<acronym title="<%=nombre_deudor%>">

						<td align="left"><%=Mid(nombre_deudor,1,50)%></td>

						</acronym>

						<td align="left"><%=strRutRepLegal%></td>

						<acronym title="<%=strNombreRepLegal%>">

						<td align="left"><%=Mid(strNombreRepLegal,1,40)%></td>

						</acronym>

						<td width="150" align="center">
							<%If Trim(strCampana)<>"" Then%>
							<img src="../imagenes/campana.gif" border="0" width="15" height="15">&nbsp;&nbsp;<B><%=ucase(strCampana)%></B>
							<%End If%>
                     	</td>
						<td width="150" align="center" ><%=strTrmoVencimiento%></td>
						<td width="150" align="center" ><%=strTramoMonto%></td>
						<td width="150" align="center" ><%=strTramoAsignacion%></td>
		
		
					</tr>

				<%
						intPrimeroDeudor = 2
					RsDeudor.movenext
					Loop

						RsDeudor.close
						set RsDeudor=nothing
						strDicom = "N/I"
						cerrarSCG()
				%>
				</table>
				<%

								abrirSCG1()
									strSql=""
									strSql="SELECT TOP 1 ID_DIRECCION, Calle,Numero,Comuna,Resto,CORRELATIVO,Estado FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR='" & rut_deudor & "' and ESTADO <>'2' ORDER BY ESTADO DESC , FECHA_INGRESO DESC"
									set rsDIR=Conn1.execute(strSql)
									if not rsDIR.eof then
										intIdDireccion=rsDIR("ID_DIRECCION")
										calle_deudor=rsDIR("Calle")
										numero_deudor=rsDIR("Numero")
										comuna_deudor=rsDIR("Comuna")
										resto_deudor=rsDIR("Resto")
										correlativo_deudor=rsDIR("CORRELATIVO")
										estado_direccion=rsDIR("Estado")
										If estado_direccion="1" then
											estado_direccion="VALIDA"
										ElseIf estado_direccion="2" then
											estado_direccion="NO VALIDA"
										Else
											estado_direccion="SIN AUDITAR"
										End if

										abrirSCG2()
										strSql="SELECT TOP 1 UPPER(CONTACTO) AS CONTACTO FROM DIRECCION_CONTACTO WHERE ID_DIRECCION = " & intIdDireccion & " ORDER BY ID_CONTACTO DESC"
										set rsContacto=Conn2.execute(strSql)
										if not rsContacto.eof Then
											strContactoDir = rsContacto("CONTACTO")
										End If
										cerrarSCG2()

									end if



									rsDIR.close
									set rsDIR=nothing
								cerrarSCG1()

								abrirSCG()
									strSql="SELECT TOP 1 ID_TELEFONO, COD_AREA,TELEFONO,CORRELATIVO,ESTADO, ISNULL(TELEFONO_DAL,0) AS TELEFONO_DAL FROM DEUDOR_TELEFONO WHERE  RUT_DEUDOR='" & rut_deudor & "' and ESTADO <> '2' ORDER BY ESTADO DESC, FECHA_INGRESO DESC"
									set rsFON=Conn.execute(strSql)
									if not rsFON.eof then
										intIdTelefono = rsFON("ID_TELEFONO")
										codarea_deudor = rsFON("COD_AREA")
										Telefono_deudor = rsFON("TELEFONO")
										strTelefonoDal = rsFON("TELEFONO_DAL")
										Correlativo_deudor2 = rsFON("CORRELATIVO")
										estado_fono = rsFON("Estado")

										if estado_fono="1" then
											estado_fono="VALIDO"
										elseif estado_fono="2" then
											estado_fono="NO VALIDO"
										else
											estado_fono="SIN AUDITAR"
										end if

										abrirSCG2()
										strSql="SELECT TOP 1 UPPER(CONTACTO) AS CONTACTO FROM TELEFONO_CONTACTO WHERE ID_TELEFONO = " & intIdTelefono & " ORDER BY ID_CONTACTO DESC"
										set rsContacto=Conn2.execute(strSql)
										if not rsContacto.eof Then
											strContacto = rsContacto("CONTACTO")
										End If
										cerrarSCG2()

									end if
									rsFON.close
									set rsFON=nothing
								cerrarSCG()

								abrirSCG()
									strSql=""
									strSql="SELECT TOP 1 ID_EMAIL,RUT_DEUDOR,CORRELATIVO,FECHA_INGRESO,EMAIL,ESTADO FROM DEUDOR_EMAIL WHERE  RUT_DEUDOR='" & rut_deudor & "' and ESTADO<>'2' ORDER BY ESTADO DESC , FECHA_INGRESO DESC"
									set rsMAIL=Conn.execute(strSql)
									if not rsMAIL.eof then
										intIdEmail = rsMAIL("ID_EMAIL")
										email = rsMAIL("EMAIL")
										Correlativo_deudor3 = rsMAIL("CORRELATIVO")
										estado_mail = rsMAIL("ESTADO")

										if estado_mail="1" then
											estado_mail="VALIDO"
										elseif estado_mail="2" then
											estado_mail="NO VALIDO"
										else
											estado_mail="SIN AUDITAR"
										end if

										abrirSCG2()
										strSql="SELECT TOP 1 UPPER(CONTACTO) AS CONTACTO FROM EMAIL_CONTACTO WHERE ID_EMAIL = " & intIdEmail & " ORDER BY ID_CONTACTO DESC"
										set rsContacto=Conn2.execute(strSql)
										if not rsContacto.eof Then
											strContactoEmail = rsContacto("CONTACTO")
										End If
										cerrarSCG2()

									end if

									rsMAIL.close
									set rsMAIL=nothing
								cerrarSCG()

								strDireccion = calle_deudor & " " & numero_deudor & " " & resto_deudor & " " & comuna_deudor
								strDomicilio = calle_deudor & " " & numero_deudor & " " & resto_deudor
								if(codarea_deudor = "0") then
									strTelefono = telefono_deudor
								else
									strTelefono = codarea_deudor & " " & telefono_deudor
								end if
								strEmail = email


								AbrirSCG()
								strSql="SELECT RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, COUNT(CUOTA.NRO_DOC) AS NUMDOC, SUM(CUOTA.VALOR_CUOTA) AS VALORORIGINAL, SUM(ISNULL(CUOTA.SALDO,0)) AS MONTODOC, MAX(IsNull(datediff(d,FECHA_VENC,getdate()),0)) as ANTIGUEDAD, CLIENTE.COD_CLIENTE, TIPO_DOCUMENTO, CLIENTE.DESCRIPCION AS DESCRIPCION , ESTADO_DEUDA, ESTADO_DEUDA.DESCRIPCION AS NOMESTADODEUDA, NOM_TIPO_DOCUMENTO, COD_TIPODOCUMENTO_HON"
								
								
								strSql= strSql & " FROM CUOTA	INNER JOIN CLIENTE ON CUOTA.COD_CLIENTE=CLIENTE.COD_CLIENTE"
								strSql= strSql & " 				INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
								strSql= strSql & " 				LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"

								strSql= strSql & " WHERE CUOTA.RUT_DEUDOR= '" & strRUT_DEUDORNSel & "'"
								strSql= strSql & "  AND CUOTA.COD_CLIENTE = '" & intCodCliente & "'"

								If Trim(strFiltro) = "1" Then
									strSql= strSql & " AND ESTADO_DEUDA.ACTIVO <> 1"
								Else
									strSql= strSql & " AND ESTADO_DEUDA.ACTIVO = 1"
								End If

								If Trim(strRutSubClienteSel) <> "" Then
									strSql= strSql & " AND CUOTA.RUT_SUBCLIENTE = '" & strRutSubClienteSel &"'"
								End if

								strSql= strSql & " GROUP BY RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, CUOTA.COD_CLIENTE,CLIENTE.COD_CLIENTE,CLIENTE.DESCRIPCION,TIPO_DOCUMENTO, ESTADO_DEUDA, ESTADO_DEUDA.DESCRIPCION, NOM_TIPO_DOCUMENTO,COD_TIPODOCUMENTO_HON"

								''Response.write strSql

								set rsDEU=Conn.execute(strSql)
								If not rsDEU.eof then

									monto=0

									%>
									<table width="100%" >
										<tr>
											<TD><div class="subtitulo_informe">> INFORMACIÓN DE LA DEUDA <%=strTipoDeuda%></div>
											</TD>
										</tr>
									</table>

									<table width="100%" border="0" class="estilo_columnas">
									<thead>
									<tr class="">
									  <td align="left">ESTADO</td>
									  <td align="left">PRODUCTO</td>
									  <td align="left">ANTIG.</td>
									  <td align="left">CAPITAL</td>
									  <%If Trim(strUsaInteres)="1" Then%>
									  	  <td align="left">INTERESES</td>
	          						  <%End If%>
									  <%If Trim(strUsaHonorarios)="1" Then%>
										  <td align="left">HONORARIOS</td>
									  <%End If%>
									  <%If Trim(strUsaProtestos)="1" Then%>
									  	  <td align="left">PROTESTOS</td>
									  <%End If%>
									  <td align="left">SALDO</td>
									  <td align="left">DOCS</td>
									  <td align="left">EJECUTIVO</td>
									  <td align="left">DETALLE</td>
									</tr>
									</thead>
									<%

									intTasaMensual = intTasaMensual/100
									intTasaDiaria = intTasaMensual/30


									Do until rsDEU.eof
										intTipoDocHono = ValNulo(rsDEU("COD_TIPODOCUMENTO_HON"),"C")
										intTotHonorarios = 0
										intTotIntereses = 0
										intTotProtesto = 0

										intValorOriginal = Round(session("valor_moneda") * ValNulo(rsDEU("VALORORIGINAL"),"N"),0)

										intCodDetCliente=rsDEU("COD_CLIENTE")
										intAntiguedad=rsDEU("ANTIGUEDAD")
										''intCuotaEnc=rsDEU("ID_CUOTA_ENC")




											AbrirSCG1()
											strSql = "SELECT dbo." & strNomFormInt & "(ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(ID_CUOTA) as HONORARIOS, ID_CUOTA,DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES,RUT_DEUDOR, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC, IsNull(datediff(d,FECHA_VENC,getdate()),0) as ANTIGUEDAD, IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,IsNull(USUARIO_ASIG,0) as USUARIO_ASIG, NRO_CUOTA, IsNull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS, SUCURSAL , ESTADO_DEUDA, COD_REMESA, CUENTA, NRO_DOC, TIPO_DOCUMENTO, CONVERT(VARCHAR(10),FECHA_ESTADO,103) AS FEC_ESTADO, ADIC_1, CUSTODIO FROM CUOTA "
											strSql = strSql & " WHERE RUT_DEUDOR='" & rut_deudor & "' AND COD_CLIENTE='" & intCodDetCliente & "' AND TIPO_DOCUMENTO = '" & rsDEU("TIPO_DOCUMENTO") & "' and estado_deuda = '1'"
											If Trim(strRutSubClienteSel) <> "" Then
												strSql= strSql & " AND CUOTA.RUT_SUBCLIENTE = '" & strRutSubClienteSel &"'"
											End if

											'Response.write "strSql=" & strSql
											set rsCalculo=Conn1.execute(strSql)
											Do While Not rsCalculo.Eof

												intSaldo = Round(session("valor_moneda") * ValNulo(rsCalculo("SALDO"),"N"),0)
												intProtesto = Round(session("valor_moneda") * ValNulo(rsCalculo("GASTOS_PROTESTOS"),"N"),0)
												intAntiguedad = ValNulo(rsCalculo("ANTIGUEDAD"),"N")

												intIntereses = rsCalculo("INTERESES")
												intHonorarios = rsCalculo("HONORARIOS")

												intTotHonorarios = intTotHonorarios + intHonorarios
												intTotIntereses = intTotIntereses + intIntereses
												intTotProtesto = intTotProtesto + intProtesto

												'Response.write "<br>intHonorarios=" & intHonorarios

												rsCalculo.movenext
											Loop
											CerrarSCG1()

										intTotalProducto = ValNulo(rsDEU("MONTODOC"),"N") + Round(intTotHonorarios,0) + Round(intTotIntereses,0) + Round(intTotProtesto,0)
									%>
									<tr bgcolor="#<%=session("COLTABBG2")%>" >
										<td align="left"><%=rsDEU("NOMESTADODEUDA")%></td>
										<td height="20" align="left"><%=rsDEU("NOM_TIPO_DOCUMENTO")%></td>
										<td height="20" align="left"><%=intAntiguedad%></td>

										<td align="left">$<%=FN(intValorOriginal,0)%></td>
										<%If Trim(strUsaInteres)="1" Then%>
											<td align="left">$<%=FN(intTotIntereses,0)%></td>
										<%End If%>
										<%If Trim(strUsaHonorarios)="1" Then%>
											<td align="left">$<%=FN(intTotHonorarios,0)%></td>
										<%End If%>
										<%If Trim(strUsaProtestos)="1" Then%>
											<td align="left">$<%=FN(intTotProtesto,0)%></td>
										<%End If%>
										<td align="left">$<%=FN(intTotalProducto,0)%></td>

									  <td align="left"><%=rsDEU("NUMDOC")%></td>
									  <td align="left">
									   <%

											AbrirSCG1
											strSql="SELECT TOP 1 ISNULL(USUARIO_ASIG,0) as USUARIO_ASIG FROM CUOTA WHERE RUT_DEUDOR='" & rut_deudor &"' AND COD_CLIENTE = '" & intCodDetCliente & "' AND ESTADO_DEUDA = '" & rsDEU("ESTADO_DEUDA") & "'"
											'reSPONSE.WRITE strSql
											'reSPONSE.eND
											set rsEJ=Conn1.execute(strSql)
											if not rsEJ.eof then
												strCob = TraeCampoId(Conn, "LOGIN", rsEJ("USUARIO_ASIG"), "USUARIO", "ID_USUARIO")
											else
												strCob = "SIN ASIG."
											end if
											rsEJ.close
											set rsEJ=nothing
											CerrarSCG1
										%>
										<%=UCASE(strCob)%>
										</td>
									  <td align="left">
										<acronym title="DETALLE DE LA DEUDA DEL DEUDOR <%=strRutDeudor%> CON EL CLIENTE <%=rsDEU("DESCRIPCION")%>">

										<a href="javascript:ventanaDetalle('detalle_deuda.asp?intCodEstado=<%=Trim(rsDEU("ESTADO_DEUDA"))%>&rut=<%=rut_deudor%>&strRutSubCliente=<%=strRutSubClienteSel%>&cliente=<%=intCodDetCliente%>')">DETALLE</a>
										<a href="javascript:ventanaDetalle('detalle_deuda.asp?rut=<%=rut_deudor%>&strRutSubCliente=<%=strRutSubClienteSel%>&cliente=<%=intCodDetCliente%>')">T</a>

										</acronym>
									  </td>

									</tr>
									<%
										''Response.write "MONTODOC=" &rsDEU("MONTODOC")
										if not isnull(rsDEU("MONTODOC")) then
											monto=monto + Cdbl(rsDEU("MONTODOC"))
										end if
									 rsDEU.movenext
									 Loop
									 %>
									 </table>
									 <%
									 strSinDeudaActiva="N"
								Else
								%>
									<table width="100%" border="0" bordercolor="#FFFFFF">
										<tr>
											<TD><div class="subtitulo_informe">> INFORMACIÓN DE LAS DEUDAS</div>
											</TD>
										</tr>
										<tr>
											<TD width="100%" height="20" align="center" class="estilo_columna_individual">
											<% If intCantidadDocAct = 0 or intCantidadDocNoAct=0 Then %>
												<FONT SIZE='2' color="#FF0000" ><B>DEUDOR NO POSEE DEUDA <%=strTipoDeuda%></B></FONT>
											<% Else %>
												<FONT SIZE='2' color="#FF0000"><B>CLIENTE NO POSEE DEUDA <%=strTipoDeuda%></B></FONT>
											<% End If %>

											</TD>
										</tr>
									</table>
								<%
									strSinDeudaActiva="S"
								End if
								rsDEU.close
								set rsDEU=nothing

							cerrarSCG()


						AbrirSCG1
							strSql="SELECT [dbo].[fun_ubicabilidad_tipo]  ('FONO','" & strRutDeudor & "') AS DETALLEUBICFON,[dbo].[fun_ubicabilidad_tipo]  ('MAIL','" & strRutDeudor & "') AS DETALLEUBICMAIL,[dbo].[fun_ubicabilidad_tipo]  ('DIRECCION','" & strRutDeudor & "') AS DETALLEUBICDIR"
							set rsuBIC= Conn1.execute(strSql)

								strDetalleUbicF = rsuBIC("DETALLEUBICFON")
								strDetalleUbicM = rsuBIC("DETALLEUBICMAIL")
								strDetalleUbicD = rsuBIC("DETALLEUBICDIR")

							rsuBIC.close
							set rsuBIC=nothing
						CerrarSCG1


				'If TraeSiNo(session("perfil_emp")) <> "Si" Then %>

				<table width="100%" border="0" bordercolor="#FFFFFF">
				<tr>
					<TD width="50%" >
						<div class="subtitulo_informe">
						> ULTIMA GESTIÓN
						</div>
					</TD>
					<TD width="50%" ALIGN="RIGHT">
						<div class="subtitulo_informe" style="text-align:right;">
							<a href="detalle_gestiones.asp?rut=<%=rut_deudor%>&cliente=<%=intCodCliente%>&strRutSubCliente=<%=strRutSubClienteSel%>&area_con=<%=area_con%>&fono_con=<%=fono_con%>&strNuevaGestion=S">> NUEVA GESTIÓN</a>
						</div>
					</TD>
				</tr>
				</table>
				<table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">
					<%

				strSql 	="SELECT TOP 1 COD_SUB_CATEGORIA, COD_CATEGORIA, COD_GESTION, G.FECHA_INGRESO, "
				strSql	=strSql & " FECHA_COMPROMISO,  G.fecha_agendamiento, CONVERT(VARCHAR(10), "
				strSql	=strSql & " G.FECHA_INGRESO,108) AS HORA_INGRESO, ID_USUARIO,FECHA_COMPROMISO, "
				strSql	=strSql & " HORA_INGRESO, OBSERVACIONES, ID_MEDIO_GESTION, TIPO_MEDIO_GESTION, "
				strSql	=strSql & " CASE "  
				strSql	=strSql & " WHEN TIPO_MEDIO_GESTION = 1 then ( "
				strSql	=strSql & " 	SELECT CONVERT(VARCHAR, COD_AREA)+'-'+CONVERT(VARCHAR, TELEFONO) "
				strSql	=strSql & " 	FROM DEUDOR_TELEFONO DD "
				strSql	=strSql & " 	WHERE DD.ID_TELEFONO= G.ID_MEDIO_GESTION "	
				strSql	=strSql & " ) "
				strSql	=strSql & " WHEN TIPO_MEDIO_GESTION = 2 then ( "
				strSql	=strSql & " 	SELECT EMAIL "
				strSql	=strSql & " 	FROM DEUDOR_EMAIL DD "
				strSql	=strSql & " 	WHERE DD.ID_EMAIL= G.ID_MEDIO_GESTION "	
				strSql	=strSql & " ) "
				strSql	=strSql & " WHEN TIPO_MEDIO_GESTION = 3 then ( "
				strSql	=strSql & " 	SELECT ISNULL(UPPER(CALLE),'')+' '+ISNULL(UPPER(NUMERO),'')+' '+ISNULL(UPPER(RESTO),'')+' '+ISNULL(UPPER(COMUNA),'') "
				strSql	=strSql & " 	FROM DEUDOR_DIRECCION DD "
				strSql	=strSql & " 	WHERE DD.ID_DIRECCION= G.ID_MEDIO_GESTION "		
				strSql	=strSql & " ) "
				strSql	=strSql & " END NOM_MEDIO_GESTION "
				strSql	=strSql & " FROM GESTIONES G INNER JOIN GESTIONES_CUOTA GC "
				strSql	=strSql & " ON G.Id_Gestion = GC.Id_Gestion "
				strSql	=strSql & " INNER JOIN CUOTA C "
				strSql	=strSql & " ON GC.Id_Cuota = C.ID_CUOTA "
				strSql	=strSql & " INNER JOIN ESTADO_DEUDA ED "
				strSql	=strSql & " ON C.ESTADO_DEUDA = ED.CODIGO "
				strSql	=strSql & " WHERE G.RUT_DEUDOR= '"&TRIM(strRutDeudor)&"' AND G.COD_CLIENTE = '"& intCodCliente &"' "
				strSql	=strSql & " AND ED.ACTIVO = '1' "
				strSql	=strSql & " ORDER BY G.FECHA_INGRESO DESC,G.Id_Gestion DESC" 

				'Response.Write "strSql= " & strSql

					AbrirSCG()
					set rsUltGest=Conn.execute(strSql)
					If not rsUltGest.eof then
					%> 
					
					<thead>
					  <tr bordercolor="#FFFFFF" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						  <td width="40" 	align="left" 	class="Estilo4">FECHA</td>
						  <td width="30" 	align="left" 	class="Estilo4">HORA</td>
						  <td width="218"	align="left" 	class="Estilo4">GESTION</td>
						  <td width="35" 	align="left" 	class="Estilo4">F.COMP.</td>
						  <td width="35" 	align="left" 	class="Estilo4">F.AGEND</td>
						  <td width="220" 	align="left" 	class="Estilo4">OBSERVACIONES</td>
						  <td width="10" 	align="center" 	class="Estilo4">MEDIO</td>
						  <td width="10" 	align="left" 	class="Estilo4">EJECUTIVO</td>
					  </tr>
					</thead>
					
					<%
						Obs 	=UCASE(LTRIM(RTRIM(rsUltGest("OBSERVACIONES"))))
						If Obs="" then
							Obs="SIN INFORMACION ADICIONAL"
						End if

						intCategoria 	= Trim(rsUltGest("COD_CATEGORIA"))
						intSubCategoria = Trim(rsUltGest("COD_SUB_CATEGORIA"))
						intGestion 		= Trim(rsUltGest("COD_GESTION"))
						strNomGestion 	= trim(rsUltGest("NOM_MEDIO_GESTION"))
						intTipo_gestion	= trim(rsUltGest("TIPO_MEDIO_GESTION"))

							AbrirSCG1
								strSql="SELECT DESCRIPCION FROM GESTIONES_TIPO_SUBCATEGORIA WHERE COD_CATEGORIA = " & intCategoria & " AND COD_SUB_CATEGORIA = " & intSubCategoria
								set rsCAT= Conn1.execute(strSql)
								If not rsCAT.eof then
									subcategoria_nombre = rsCAT("Descripcion")
								End if
								rsCAT.close
								set rsCAT=nothing
							CerrarSCG1

							AbrirSCG1
								strSql="SELECT DESCRIPCION FROM GESTIONES_TIPO_GESTION WHERE COD_CATEGORIA = '" & intCategoria & "' AND COD_SUB_CATEGORIA = '" & intSubCategoria & "' AND COD_GESTION = '" & intGestion & "' AND COD_CLIENTE = '" &intCodCliente & "'"
								set rsCAT= Conn1.execute(strSql)
								If not rsCAT.eof then
									gestion_nombre = rsCAT("Descripcion")
								End if
								rsCAT.close
								set rsCAT=nothing
							CerrarSCG1

						''strGestion = TraeCampoId(Conn, "DESCRIPCION", intCategoria, "GESTIONES_TIPO_CATEGORIA", "COD_CATEGORIA") & " - " & TraeCampoId(Conn, "DESCRIPCION", rsUltGest("COD_CATEGORIA"), "GESTIONES_TIPO_CATEGORIA", "COD_CATEGORIA") & " " & subcategoria_nombre & " " & gestion_nombre
						AbrirSCG1
							strGestion = TraeCampoId(Conn1, "DESCRIPCION", rsUltGest("COD_CATEGORIA"), "GESTIONES_TIPO_CATEGORIA", "COD_CATEGORIA") & "-" & subcategoria_nombre & "-" & gestion_nombre
						CerrarSCG1

						%>
						<tr bordercolor="#FFFFFF" class="Estilo8">
						  <td align="left" class="Estilo4"><%=rsUltGest("FECHA_INGRESO")%></td>
						  <td align="left" class="Estilo4"><%=rsUltGest("HORA_INGRESO")%></td>

						  <td align="left" class="Estilo4"><%=strGestion%></td>

						  <td align="left" class="Estilo4"><%=rsUltGest("FECHA_COMPROMISO")%></td>
						  <td align="left" class="Estilo4"><%=rsUltGest("FECHA_AGENDAMIENTO")%></td>
						  <td align="left" class="Estilo4"><acronym title="<%=Obs%>"><%=Mid(Obs,1,200)%></acronym></td>
						  <td align="center" class="Estilo4">
							<%if trim(intTipo_gestion)=1 then%>
								<%=strNomGestion%>

							<%elseif trim(intTipo_gestion)=2 then%>
								<img src="../imagenes/Arroa.png" border="0" title="<%=strNomGestion%>">

							<%elseif trim(intTipo_gestion)=3 then%>
								<img src="../imagenes/mod_direccion_va.png" title="<%=strNomGestion%>">

							<%else%>
								&nbsp;
							<%end if%>					  

						  </td>

						   <%
								If trim(rsUltGest("ID_USUARIO")) <> "" Then
								AbrirSCG1
									strNomUsuario = TraeCampoId(Conn1, "LOGIN", rsUltGest("ID_USUARIO"), "USUARIO", "ID_USUARIO")
								CerrarSCG1
								End If

							%>

						  <td align="left" class="Estilo4"><%=UCASE(strNomUsuario)%></td>
						</tr>
						<%
					 else
						response.write("<tr class='estilo_columna_individual'><td class='Estilo4' colspan='8' align='center'><FONT SIZE='2' color='#FF0000'><B>NO EXISTE GESTIÓN RELACIONADA</B></td></tr>")
					 End If
					 CerrarSCG()
					 %>
				</table>

				<% 'End If ' del permiso del cliente%>


				<div class="subtitulo_informe">
					> DATOS DEL CONTACTO
				</div>

				<table cellpadding="0" cellspacing="0" class="estilo_columnas" style="width:100%;" align="center" >
				<thead>
				<tr bordercolor="#999999">
					<td align="left" colspan="2">DIRECCIÓN MÁS RECIENTE </td>
					<td align="left" colspan="2">TELÉFONO MÁS RECIENTE </td>
					<td align="left" colspan="2">EMAIL MÁS RECIENTE </td>
				</tr>
				</thead>
				<body>
				<tr height="25" class="Estilo8">
					<td align="left">
						<acronym title="<%=strDireccion%>">

						 &nbsp;<%=Mid(strDireccion,1,40)%>

						</acronym>

						<br>
						&nbsp;<%=strContactoDir%>
					</td>
					<td ALIGN="CENTER">
						<a href="mas_direcciones.asp?rut=<%=rut_deudor%>"><acronym title="VER TODAS LAS DIRECCIONES DEL DEUDOR">Más</acronym></a>
						&nbsp;
						<a href="nueva_dir.asp?rut=<%=rut_deudor%>"><acronym title="INGRESAR UNA NUEVA DIRECCIÓN">&nbsp;Nuevo&nbsp;</acronym></a>
				  		<input name="correlativo_direccion" type="hidden" id="correlativo_direccion" value="<%=correlativo_deudor%>">						
					</td>
					<td align="left">
						&nbsp;
						<a href="sip:<%=SoloNumeros(strTelefonoDal)%>">
							<%=SoloNumeros(strTelefono)%>
						</a>
						<br>
						&nbsp;<%=strContacto%>
					</td>
					<td ALIGN="CENTER">
						<a href="mas_telefonos.asp?rut=<%=rut_deudor%>"> <acronym title="VER TODOS LOS TELÉFONOS DEL DEUDOR">Más</acronym></a>
						&nbsp;
						<a href="nuevo_tel.asp?rut=<%=rut_deudor%>"> <acronym title="INGRESAR UN NUEVO TELÉFONO">&nbsp;Nuevo&nbsp;</acronym></a>
						<input name="correlativo_fono" type="hidden" id="correlativo_fono" value="<%=correlativo_deudor2%>">						
					</td>
					<td align="left">
						&nbsp;<%=strEmail%>
						<br>
						&nbsp;<%=strContactoEmail%>
					</td>
					<td ALIGN="CENTER">
						<a href="mas_correos.asp?rut=<%=rut_deudor%>"> <acronym title="VER TODOS LOS EMAIL DEL DEUDOR">Más</acronym></a>
						&nbsp;
						<a href="nuevo_cor.asp?rut=<%=rut_deudor%>"> <acronym title="INGRESAR UN NUEVO EMAIL">&nbsp;Nuevo&nbsp;</acronym>
						 <input name="correlativo_mail" type="hidden" id="correlativo_mail" value="<%=correlativo_deudor3%>">						
					</td>
				</tr>
				<tr class="totales">
					<td align="left" HEIGHT="20" WIDTH="266" >
						&nbsp;<%=strDetalleUbicD%>
				  	</td>
				  	<td align="center"></td>
					<td align="left" HEIGHT="20" WIDTH="266" >
						&nbsp;<%=strDetalleUbicF%>
					</td>
				  	<td align="center">
					</td>
					<td align="left" HEIGHT="20" WIDTH="266" >
						&nbsp;<%=strDetalleUbicM%>
					</td>
				  	<td align="center">
				  		
					</td>
				</tr>
				</tbody>
				</table>

				<%

				end if
			end if
		end if
		%>
	</table>
	</td>
	                </tr>
                </table>
                </td>
            </tr>
          </table>

        </td>
    </tr>

</table>
<script language="javascript">
	//gridLoad('divScroll','hdnScrollPos');
</script>
</form>
	<%
		
		g =  Request.QueryString("g")
		
		
		%>
	<script type="text/javascript">

var tipo_softphone = '<%=session("tipo_softphone")%>';
var g = '<%= g %>';



if(tipo_softphone == "1"){
if(g != 'si'){
			function getParameterByName(name) {
				name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
				var regex = new RegExp("[\\#&]" +
					name + "=([^&#]*)"),
					results = regex.exec(location.hash);
				return results === null ? "" : decodeURIComponent(results[1].replace(/\+/g,
					" "));
			}
			
			
			if (window.location.hash) {
				console.log(location.hash);
				var token = getParameterByName('access_token');
				//location.hash = ''
				alert(token);
				
				window.location.replace('FuncionesAjax/modulo_token_purecloud.asp?tokenID='+token)	
				
			} else {
				var queryStringData = {
					response_type: "token",
					client_id: "92014f51-547f-4f04-af6d-0649e54ba879",
					redirect_uri: "http://sistemas.llacruz.cl/erec_desa/System/principal.asp?EmpresaId=3"
				}
				console.log(queryStringData);
				console.log(jQuery.param(queryStringData));
				window.location.replace("https://login.mypurecloud.com/oauth/authorize?" + jQuery.param(queryStringData));
			}
}
	}else{
//alert("Sin PC");
}		
	</script>
</body>
</html>

<!--#include file="../lib/comunes/js_css/bottom_tooltip.inc" -->

<script language="JavaScript" type="text/JavaScript">
function envia(){
datos.action='principal.asp';
datos.submit();
}

function Priorizar()
{

	if (confirm("¿ Está seguro de que desea priorizar el caso ? "))
	{
		datos.action='principal.asp?strPriorizar=1';
		datos.submit();
	}


}


function Seleccionar(strRutSubCliente,strFiltro){
	var pos = document.getElementById('hdnScrollPos');
	pos.value = event.srcElement.scrollTop;
	//alert(pos.value);
	datos.action='principal.asp?strFiltro=' + strFiltro + '&strRutSubCliente=' + strRutSubCliente;
	datos.submit();
}

function Seleccionar2(strRutSubCliente,strRUT_DEUDORSel) {
	datos.action='principal.asp?strRutSubCliente=' + strRutSubCliente + '&strRUT_DEUDORSel=' + strRUT_DEUDORSel;
	datos.submit();
}

function envia_caja(){
	datos.action='ingreso_pagos.asp?intOrigen=IP';
	datos.submit();
}

function envia_backoffice(){
	datos.action='enviar_backoffice.asp?intOrigen=IP';
	datos.submit();
}

function priorizar_caso(){
	datos.action='priorizar_caso.asp?strRut=<%=strRutDeudor%>';
	datos.submit();
}

function envia_info(){
datos.action='ingreso_pagos.asp?intOrigen=ID';
datos.submit();
}

function envia_plandepago(){
datos.action='simulacion_convenio.asp?intOrigen=PP';
datos.submit();
}

function envia_convenio(){
datos.action='simulacion_convenio.asp?intOrigen=CO';
datos.submit();
}

function envia_biblioteca(){
datos.action='biblioteca_deudores.asp?strRut=<%=strRutDeudor%>';
datos.submit();
}

function envia_dir(){
datos.action='principal.asp?dir=si';
datos.submit();
}

function envia_fon(){
datos.action='principal.asp?fon=si';
datos.submit();
}

function envia_mail(){
datos.action='principal.asp?mail=si';
datos.submit();
}

function paga(){
datos.action='detalle_pago.asp';
datos.submit();
}

function ventanaSecundaria (URL){
window.open(URL,"INFORMACION","width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ventanaBiblioteca (URL){
window.open(URL,"INFORMACION","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ventanaPriorizar (URL){
	window.open(URL,"INFORMACION","width=600, height=280, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ventanaDetalle (URL){
window.open(URL,"DETALLEDEUDA","width=1400, height=600, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ventanaIngresoG (URL){
window.open(URL,"INFORMACION","width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ventanaMas (URL){
window.open(URL,"DATOS","width=818, height=500, scrollbars=no, menubar=no, location=no, resizable=yes")
}

</script>


<%
   ' response.write MID(request.servervariables("PATH_INFO"),2, Instr(MID(request.servervariables("PATH_INFO"),2, LEN(request.servervariables("PATH_INFO"))),"/"))
%>
