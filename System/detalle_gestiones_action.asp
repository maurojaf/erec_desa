<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
<!--#include file="../lib/lib.asp"-->

<%

	Response.CodePage = 65001
	Response.charset="utf-8"

	strFonoAgestionar 	= Request("strFonoAgestionar")
	strContactoSel 		= Request("strContactoSel")
	rut 				= request("rut")
	strCodCliente 		= session("ses_codcli")

	categoria 			= request("cmbcat")
	subcategoria 		= request("cmbsubcat")
	gestion 			= request("cmbgest")
	OBSERVACIONES 		= Mid(Replace(request("OBSERVACIONES"),";"," "),1,599)
	strGestionCliente1 	= Request("GESTION_CLIENTE_1")
	strGestionCliente2 	= Request("GESTION_CLIENTE_2")
	strArchivoAsp 		= "principal.asp?a=1"
	cuotas_deudor 		= request("cuotas_deudor")
	IntIdFonoAgend 		= request("CB_FONO_AGEND")

'**/Inicialización de variables/**'

	Id_Forma_Recaudacion 	=NULL
	strLugarPago 			=NULL
	intMontoCancelado 		=0
	intIdContactoEmail 		=0
	intIdContacto 			=0
	strGestionCliente1 		=NULL
	strGestionCliente2 		=NULL
	IntIdDireccionAsociada  =0

	'split_cuotas_deudor 	=split(cuotas_deudor,",")
	'total_cuotas_deudor 	=ubound(split_cuotas_deudor)

'**Creación de variables compuestas**'

	strCodUltGest = categoria & "*" & subcategoria & "*" & gestion

'**/Captura las variables asociadas al tipo gestion Compromiso de pago/**'

		If Trim(request("hd_tipo_gestion")) = "1" Then

			dtmFecCompromiso  	=Replace(request("TX_FEC_COMPROMISO"),"'","")
			strHoraDesde 		=request("TX_HORADESDE")
			strHoraHasta 		=request("TX_HORAHASTA")
			strFormaPago 		=request("CB_FORMAPAGO_CP")
			strCbLugarPago 		=request("CB_LUGARPAGO_CP")
			intIdContactoCP 	=request("CB_CONTACTO_ASOCIADO_CP") 'ALG
			strFonoCP 			=request("TX_FONO_CP")
			intMontoCancelado 	=Replace(request("TX_MONTO_COMPROMISO"),".","")

			if trim(strCbLugarPago)<>"" and trim(strCbLugarPago)<>"0" then
				cadena 			= split(strCbLugarPago, "-")
				ID				= cadena(0)
				TIPO			= cadena(1)
			end if

			IF TRIM(TIPO)="DIRECCION" THEN
				strLugarPago = ID
				Id_Forma_Recaudacion = NULL

			ElseIf TRIM(TIPO)="FORMA_RECAUDACION" then
				Id_Forma_Recaudacion = ID
				strLugarPago = NULL
			END IF

		End If

'**/Captura las variables asociadas al tipo gestion Compromiso de pago ruta/**'

		If Trim(request("hd_tipo_gestion")) = "2" Then

	'**/Documentos Gestion/**'

			AbrirSCG1()
				strSql = "SELECT * FROM TIPO_DOCUMENTOS_GESTION WHERE COD_CLIENTE ='" & strCodCliente & "'"
				'response.write  strSql& "<br>"
				set rsDoc=Conn1.execute(strSql)
				strDocGestion = ""
				Do until rsDoc.eof
					strObjeto = "CH_TD2_" & rsDoc("COD_DOCUMENTO")
					If UCASE(Request(strObjeto)) <> "" Then
						strDocGestion = strDocGestion & " - " & Request(strObjeto)
					End If
					rsDoc.movenext
				Loop

				If Trim(strDocGestion) <> "" Then
					strDocGestion = Mid(strDocGestion,3,len(strDocGestion))
				End If

				strFNC = Request("TX_FNC_RUTA")
				If Trim(strFNC) <> "" Then
					strDocGestion = strDocGestion & " - " & strFNC
				End If

				strDocGestion = ucase(Mid(strDocGestion,1,250))

				'Response.write "<br>strFNC=" & strFNC

			CerrarSCG1()
	
			dtmFecCompromiso 	=Replace(request("TX_FEC_COMPROMISO_RUTA"),"'","")
			strHoraDesde 		=request("TX_HORADESDE_RUTA")
			strHoraHasta 		=request("TX_HORAHASTA_RUTA")
			strFormaPago 		=request("CB_FORMAPAGO_CP_RUTA")
			strCbLugarPago 		=request("CB_LUGARPAGO_CP_RUTA")
			intIdContactoCP 	=request("CB_CONTACTO_ASOCIADO_CP_RUTA")
			strFonoCP 			=request("CB_FONO_CP_RUTA")
			intMontoCancelado 	=Replace(request("TX_MONTO_COMPROMISO_RUTA"),".","")

			if trim(strCbLugarPago)<>"" and trim(strCbLugarPago)<>"0" then
				cadena 			= split(strCbLugarPago, "-")
				ID				= cadena(0)
				TIPO			= cadena(1)
			end if

			IF TRIM(TIPO)="DIRECCION" THEN
				strLugarPago = ID
				Id_Forma_Recaudacion = NULL

			ElseIf TRIM(TIPO)="FORMA_RECAUDACION" then
				Id_Forma_Recaudacion = ID
				strLugarPago = NULL

			END IF

		End If

'**/Captura las variables asociadas al tipo gestion Indica que pago/**'

		If Trim(request("hd_tipo_gestion")) = "3" Then
			dtmFecNormalizacion =request("TX_FEC_NORM")
			strFormaPago		=request("CB_FORMAPAGO_NORM")
			strCbLugarPago 		=request("CB_LUGAR_NORM")
			strComprobante 		=request("comprobante")
			intMontoCancelado 	=Replace(request("TX_MONTOCANCELADO_NORM"),".","")
			strEnvioHRD 		=request("CB_ENVIO_HRD")

			if trim(strCbLugarPago)<>"" and trim(strCbLugarPago)<>"0" then
				cadena 			= split(strCbLugarPago, "-")
				ID				= cadena(0)
				TIPO			= cadena(1)
			end if

			IF TRIM(TIPO)="DIRECCION" THEN
				strLugarPago = ID
				Id_Forma_Recaudacion = NULL

			ElseIf TRIM(TIPO)="FORMA_RECAUDACION" then
				Id_Forma_Recaudacion = ID
				strLugarPago = NULL

			END IF

		End If

'**/Captura las variables asociadas al tipo gestion Expone requerimiento/**'

		If Trim(request("hd_tipo_gestion")) = "4" Then

			dtmFecNormalizacion =request("TX_FEC_NORM2")
			strFormaPago		=request("CB_FORMAPAGO_NORM2")
			strCbLugarPago		=request("CB_LUGAR_NORM2")
			strComprobante 		=request("comprobante2")
			intMontoCancelado 	=Replace(request("TX_MONTOCANCELADO_NORM2"),".","")
			strEnvioHRD 		=request("CB_ENVIO_HRD2")

			if trim(strCbLugarPago)<>"" and trim(strCbLugarPago)<>"0" then
				cadena 			= split(strCbLugarPago, "-")
				ID				= cadena(0)
				TIPO			= cadena(1)
			end if

			IF TRIM(TIPO)="DIRECCION" THEN
				strLugarPago = ID
				Id_Forma_Recaudacion = NULL

			ElseIf TRIM(TIPO)="FORMA_RECAUDACION" then
				Id_Forma_Recaudacion = ID
				strLugarPago = NULL

			END IF

		End If

'**/Gestión Verificación en terreno/**'

		If Trim(request("hd_tipo_gestion")) = "5" Then

	'**/Documentos Gestion/**'

			AbrirSCG1()
				strSql = "SELECT * FROM TIPO_DOCUMENTOS_GESTION WHERE COD_CLIENTE ='" & strCodCliente & "'"
				'response.write  strSql& "<br>"
				set rsDoc=Conn1.execute(strSql)
				strDocGestion = ""
				Do until rsDoc.eof
					strObjeto = "CH_TD_" & rsDoc("COD_DOCUMENTO")
					If UCASE(Request(strObjeto)) <> "" Then
						strDocGestion = strDocGestion & " - " & Request(strObjeto)
					End If
					rsDoc.movenext
				Loop

				If Trim(strDocGestion) <> "" Then
					strDocGestion = Mid(strDocGestion,3,len(strDocGestion))
				End If

				strFNC = Request("TX_FNC_RUTA")
				If Trim(strFNC) <> "" Then
					strDocGestion = strDocGestion & " - " & strFNC
				End If

				strDocGestion = ucase(Mid(strDocGestion,1,250))

				'Response.write "<br>strFNC=" & strFNC

			CerrarSCG1()
	
			dtmFecCompromiso	=request("TX_FEC_GESTION_TERRENO")
			strHoraDesde		=request("TX_HORADESDE_TERRENO")
			strHoraHasta		=request("TX_HORAHASTA_TERRENO")
			strCbLugarPago		=request("CB_DIRECCION_TERRENO")
			strFonoCP			=request("CB_FONO_TERRENO")
			intIdContactoCP		=request("CB_CONTACTO_ASOCIADO_TERRENO")

			if trim(strCbLugarPago)<>"" and trim(strCbLugarPago)<>"0" then
				cadena 			= split(strCbLugarPago, "-")
				ID				= cadena(0)
				TIPO			= cadena(1)
			end if

			IF TRIM(TIPO)="DIRECCION" THEN
				strLugarPago = ID
				Id_Forma_Recaudacion = NULL

			ElseIf TRIM(TIPO)="FORMA_RECAUDACION" then
				Id_Forma_Recaudacion = ID
				strLugarPago = NULL

			END IF
			
		End If


'**/Captura las variables asociadas al tipo Gestion Simple/**'

		If Trim(request("hd_tipo_agend")) = "1" Then
			dtmFechaAgendamiento 	= request("TX_FEC_AGEND")
			strHoraAgend 			= request("TX_HORAAGEND")
		End If

'**/Captura las variables asociadas al tipo Gestion Telefonico/**'

		If Trim(request("hd_tipo_agend")) = "2" Then
			IntIdTelGest 			=request("CB_FONO_GESTION")
			dtmFechaAgendamiento 	=request("TX_FEC_AGEND_TEL")
			strHoraAgend 			=request("TX_HORAAGEND_TEL")
			intIdContacto 			=request("CB_CONTACTO_ASOCIADO")
		End If

'**/Captura las variables asociadas al tipo Gestion Email/**'

		If Trim(request("hd_tipo_agend")) = "3" Then
			IntIdEmailGest 			=request("CB_EMAIL_GESTION")
			dtmFechaAgendamiento 	=request("TX_FEC_AGEND_EMAIL")
			strHoraAgend 			=request("TX_HORAAGEND_EMAIL")
			intIdContactoEmail 		=request("CB_CONTACTO_ASOCIADO_EMAIL")
		End If


		If Trim(request("hd_tipo_agend")) = "4" Then
			IntIdDireccionAsociada 	=request("CB_DIRECCION_GESTION")
			dtmFechaAgendamiento 	=request("TX_FEC_AGEND_DIRECCION")
			strHoraAgend 			=request("TX_HORAAGEND_DIRECCION")
			intIdContacto 			=request("CB_CONTACTO_ASOCIADO_DIRECCION")
		End If


'**/Modifica las variables según condición/**'

			If Trim(strFonoCP) = "" Then strFonoCP = "NULL"

			If Trim(intIdContactoCP) = "" Then intIdContactoCP = "NULL"

			if trim(IntIdTelGest)="" or trim(IntIdTelGest)="0" then
				IntIdTelGest ="NULL"
			end if

			if trim(intMontoCancelado)="" or trim(intMontoCancelado)="0" then
				intMontoCancelado ="NULL"
			end if

			if trim(IntIdEmailGest)="" or trim(IntIdEmailGest)="0" then
				IntIdEmailGest ="NULL"
			end if

			if trim(IntIdFonoAgend)="" or trim(IntIdFonoAgend)="0" then
				IntIdFonoAgend ="NULL"
			end if

			if Trim(strFormaPago) <> "" and Trim(strFormaPago) <> "NULL" then
				strFormaPago = "'" + strFormaPago + "'"
			else
				strFormaPago="NULL"
			end if

			if Trim(strDocGestion) <> "" and Trim(strDocGestion) <> "NULL" then
				strDocGestion = "'" + strDocGestion + "'"
			else
				strDocGestion="NULL"
			end if

			if Trim(dtmFechaAgendamiento) <> "" and Trim(dtmFechaAgendamiento) <> "NULL" then
				dtmFechaAgendamiento = "'" + dtmFechaAgendamiento + "'"
			else
				dtmFechaAgendamiento="NULL"
			end if

			if Trim(strHoraAgend) <> "" and Trim(strHoraAgend) <> "NULL" then
				strHoraAgend = "'" + strHoraAgend + "'"
			else
				strHoraAgend="NULL"
			end if

			if Trim(dtmFecNormalizacion) <> "" then
				dtmFecNormalizacion= "'" & dtmFecNormalizacion & "'"
			else
				dtmFecNormalizacion="NULL"
			end if

			if Trim(dtmFecCompromiso) <> "" then
				dtmFecCompromiso= "'" & dtmFecCompromiso & "'"
			else
				dtmFecCompromiso="NULL"
			end if

			if Trim(strComprobante) <> "" then
				strComprobante= "'" & strComprobante & "'"
			else
				strComprobante="NULL"
			end if

			if Trim(strHoraDesde) <> "" and Trim(strHoraDesde) <> "NULL" then
				strHoraDesde = "'" + strHoraDesde + "'"
			else
				strHoraDesde="NULL"
			end if

			if Trim(strHoraHasta) <> "" and Trim(strHoraHasta) <> "NULL" then
				strHoraHasta = "'" + strHoraHasta + "'"
			else
				strHoraHasta="NULL"
			end if

'**/Consulta que inicializa variables de la tabla DEUDOR/**'

			AbrirScg1()

				strSql="SELECT ISNULL(ID_CAMPANA,0) as ID_CAMPANA"
				strSql = strSql & " FROM DEUDOR WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & strCodCliente & "'"
				'Response.write "<br>strSql=" & strSql

				set rsDeudor = Conn1.execute(strSql)
				if not rsDeudor.eof then
					intIdCampana=rsDeudor("ID_CAMPANA")
				end if

				rsDeudor.close
				set rsDeudor=nothing

			CerrarScg1()

'**/Crea el correlativo de la nueva gestión a ser ingresada en la tabla gestiones/**'

			AbrirScg1()
			ssql2="SELECT ISNULL((MAX(CORRELATIVO) + 1),0) AS CORRELATIVO FROM GESTIONES WHERE RUT_DEUDOR='"&rut&"' AND COD_CLIENTE='"&strCodCliente&"'"
			'Response.write "ssql2 = " & ssql2

			set rsCOR = Conn1.execute(ssql2)
			if not rsCOR.eof then
				intCorrelativo=rsCOR("CORRELATIVO")
			else
				intCorrelativo= "1"
			end if
			rsCOR.close
			set rsCOR=nothing
			CerrarScg1()

'**/Consulta que inicializa variables de la tabla GESTIONES_TIPO_GESTION/**'

	'**Prioridad de la gestión y el módulo asociado a la gestión**'

			AbrirScg1()

			strSql = "SELECT GESTION_MODULOS, PRIORIDAD FROM GESTIONES_TIPO_GESTION "
			strSql= strSql & " WHERE COD_CLIENTE = " & strCodCliente& " AND COD_CATEGORIA  = " & categoria & " AND COD_SUB_CATEGORIA  = " & subcategoria & " AND COD_GESTION  = " & gestion

			set rsGesTipoGes = Conn1.execute(strSql )
			If Not rsGesTipoGes.eof Then
				intGestionModulos = rsGesTipoGes("GESTION_MODULOS")
				intPrioridadGestion = Cdbl(rsGesTipoGes("PRIORIDAD"))
			Else
				intGestionModulos = 0
				intPrioridadGestion = 0
			End If

			CerrarSCG1()


			AbrirSCG1()
			strSql = "SELECT MAX(ID_GESTION) AS ID_GESTION FROM GESTIONES WHERE RUT_DEUDOR  = '" & rut & "' AND COD_CLIENTE = '" & strCodCliente & "' AND CORRELATIVO  = " & intCorrelativo

			set rsGestion = Conn1.execute(strSql )
			If Not rsGestion.eof Then
				intIdGestion = rsGestion("ID_GESTION")
			Else
				intIdGestion = 0
			End If

			'response.write  strLugarPago& "<br>" & strLugarPago
			'response.end

			'For indice = 0 to total_cuotas_deudor 

				sql_insert_cuota ="EXEC proc_Ingreso_Gestion '" & rut & "','" & strCodCliente & "','" & strCodUltGest & "','" & session("session_idusuario") & "'," & dtmFecCompromiso & "," & strComprobante & "," & dtmFecNormalizacion & ",'" & UCASE(OBSERVACIONES) & "'," & dtmFechaAgendamiento & "," & strHoraAgend & ","

				sql_insert_cuota = sql_insert_cuota & intIdCampana & ","&intIdContacto&"," & intIdContactoEmail & "," & strFormaPago & "," & strHoraDesde  & "," & strHoraHasta  & "," & strDocGestion & ",'" & strGestionCliente1 & "','" & strGestionCliente2 & "'," & ValNulo(intMontoCancelado,"N") & "," 

				sql_insert_cuota = sql_insert_cuota & ValNulo(strEnvioHRD,"N") & "," & strFonoCP & "," & intIdContactoCP & ","&trim(IntIdEmailGest)&","&trim(IntIdTelGest)&",'"&trim(IntIdDireccionAsociada)&"','"&TRIM(Id_Forma_Recaudacion)&"',"&TRIM(IntIdFonoAgend)&",'"&TRIM(strLugarPago)&"','"&cuotas_deudor&"'"

				response.write  sql_insert_cuota& "<br>" & sql_insert_cuota
				response.end
				
				set rs_insert_cuota =conn.execute(sql_insert_cuota)
				if rs_insert_cuota.eof then
					response.write "HA OCURRIDO UN ERROR"
					response.end()
				else
					intIdGestion =rs_insert_cuota("ID_GESTION")
				end if

'**/Consulta la prioridad calculada de la cuota/**'

					AbrirSCG1()

					strSql = "SELECT CAST(ISNULL(PRIORIDAD_CUOTA,11) AS NUMERIC(4,1)) AS PRIORIDAD_CUOTA, PRIORIDAD_CUOTA_CAL = CASE"
					strSql = strSql & "    WHEN [dbo].[fun_FonosDias] (RUT_DEUDOR,2) >= 1 AND [dbo].[fun_dias_atencion_telefonica] (RUT_DEUDOR, GETDATE(),0)>=1 AND CUOTA.COD_CLIENTE = 1100"
					strSql = strSql & "    THEN 4"
					strSql = strSql & "    WHEN CAST(GETDATE() - FECHA_VENC  as int) >= 30 AND CUOTA.COD_CLIENTE = 1100"
					strSql = strSql & "    THEN 5"
					strSql = strSql & "    WHEN CAST(GETDATE() - FECHA_VENC  as int) >= 10 AND CUOTA.COD_CLIENTE = 1100"
					strSql = strSql & "    THEN 6"
					strSql = strSql & "    WHEN SALDO >=100000000 AND CUOTA.COD_CLIENTE = 1100"
					strSql = strSql & "    THEN 7"
					strSql = strSql & "    WHEN CAST(GETDATE() - FECHA_VENC  as int) >= 0 AND CUOTA.COD_CLIENTE = 1100"
					strSql = strSql & "    THEN 9"
					strSql = strSql & "    WHEN CUOTA.ID_ULT_GEST IS NULL AND CUOTA.COD_CLIENTE = 1100"
					strSql = strSql & "    THEN 10"
					strSql = strSql & "    WHEN CUOTA.COD_CLIENTE = 1100"
					strSql = strSql & "    THEN 11"
					strSql = strSql & "    ELSE 100"
					strSql = strSql & " END"
					strSql = strSql & " FROM dbo.CUOTA INNER JOIN dbo.ESTADO_DEUDA ON dbo.ESTADO_DEUDA.CODIGO = cuota.ESTADO_DEUDA"
					strSql = strSql & " WHERE ESTADO_DEUDA.ACTIVO = 1 AND ID_CUOTA IN (" & cuotas_deudor & ")"

					'Response.write strSql &"<br>"

					set rsTmp = Conn1.execute(strSql)

					If Not rsTmp.eof Then
						intPrioridadCal = rsTmp("PRIORIDAD_CUOTA_CAL")
						intPrioridadCuota = Cdbl(rsTmp("PRIORIDAD_CUOTA"))
					End if

					CerrarSCG1()

					If (intPrioridadCal <= intPrioridadGestion AND intPrioridadGestion <> 8) or (intPrioridadCuota >= 8 AND intPrioridadGestion = 8) Then
						intPrioridadFinal = intPrioridadCal
					ElseIf intPrioridadCuota < 8 AND intPrioridadGestion = 8 Then
					   	intPrioridadFinal = intPrioridadCuota
					Else
					   	intPrioridadFinal = intPrioridadGestion
					End If

					If ((intPrioridadCuota  > 5) and UCASE(Request("CH_PRIORITARIA")) = "ON") Then
						intPrioridadFinal = "2.2"
					End If


					AbrirSCG1()
						strSql = "UPDATE CUOTA SET PRIORIDAD_CUOTA = " & Replace(intPrioridadFinal,",",".")
						strSql = strSql & " WHERE ID_CUOTA in ("& cuotas_deudor & ")"
						''Response.write "<br>" & strSql
						Conn1.execute(strSql)
					CerrarSCG1()



'**/Redirige según variable a distintas partes del sistema al ingresar gestión/**'

			AbrirSCG1()

			If Trim(intGestionModulos) = "11" Then
				If Trim(dtmFecCompromiso) <> "" and Trim(dtmFecCompromiso) <> "01/01/1900" and Trim(dtmFecCompromiso) <> "NULL" Then
					strArchivoAsp = "confirmar_cp.asp?id_gestion=" & intIdGestion & "&dtmFecCompGest=" & dtmFecCompromiso
				End If
			End If

			CerrarSCG1()

			AbrirSCG1()
				strSql = "SELECT ID_CUOTA FROM CUOTA WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & strCodCliente & "'"
				strSql = strSql & " AND DATEDIFF(MINUTE,(ISNULL(FECHA_AGEND_ULT_GES,GETDATE()-200)+ convert(varchar(10),ISNULL(HORA_AGEND_ULT_GES,'00:00'),108)),GETDATE()) >= 0 "
				strSql = strSql & " AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
				strSql = strSql & " AND COD_ULT_GEST IN (SELECT cast(COD_CATEGORIA as varchar(2))+ '*' + cast(COD_SUB_CATEGORIA as varchar(2)) + '*' + cast(COD_GESTION as varchar(2))"
				strSql = strSql & " FROM GESTIONES_TIPO_GESTION WHERE VER_AGEND = 1 AND COD_CLIENTE ='" & strCodCliente & "')"

				'Response.write "<br><br>strSql=" & strSql

				set rsValida = Conn1.execute(strSql )
				If Not rsValida.Eof Then
					strArchivoAsp = "detalle_gestiones.asp?strFonoAgestionar=" & IntIdTelGest & "&strContactoSel=" & intIdContacto
				End If

			CerrarSCG1()

			Response.Redirect strArchivoAsp & "&rut=" & rut & "&cliente=" & strCodCliente

		CerrarSCG()

	%>


