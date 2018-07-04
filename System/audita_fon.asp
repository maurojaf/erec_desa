<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/lib.asp"-->

<%
Response.CodePage=65001
Response.charset ="utf-8"
	
rut 				=request("rut")
cliente 			=session("ses_codcli")
strOrigen 			=request("strOrigen")
strFonoAgestionar 	=request("strFonoAgestionar")


abrirscg()
ssql="SELECT CORRELATIVO, IsNull(cast(COD_AREA as varchar(3)) + '-' + telefono,'') as TELEFONO FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" & rut & "'"
''Response.write "<br>ssql=" & ssql
set rsDIR=Conn.execute(ssql)
do until rsDIR.eof
	estado_correlativo=request("radiofon" & Trim(rsDIR("CORRELATIVO")))

	''Response.write "<br>estado_correlativo=" & estado_correlativo
	CORRELATIVO=rsDIR("CORRELATIVO")

	strAnexo=Trim(request("TX_ANEXO_" & Trim(rsDIR("CORRELATIVO"))))
	strDesde=Trim(request("TX_DESDE_" & Trim(rsDIR("CORRELATIVO"))))
	strHasta=Trim(request("TX_HASTA_" & Trim(rsDIR("CORRELATIVO"))))
	strDiasAtencion=Trim(request("CH_DIAS_" & Trim(rsDIR("CORRELATIVO"))))

	If Trim(estado_correlativo) <> "" and Not IsNull(estado_correlativo) Then
		strSql="SELECT ESTADO FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" & rut & "' and CORRELATIVO='" & CORRELATIVO & "'"
		set rsValida=Conn.execute(strSql)
		If Not rsValida.eof Then
			intEstadoFono = rsValida("ESTADO")
		End If
		If Trim(intEstadoFono) <> Trim(estado_correlativo) Then
			strSql = "UPDATE DEUDOR_TELEFONO SET ESTADO='" & cint(estado_correlativo) & "', FECHA_REVISION = getdate(), USR_REVISION = '" & session("session_login") & "' WHERE RUT_DEUDOR = '" & rut & "' and CORRELATIVO='" & CORRELATIVO & "'"
			''REsponse.write "<br>strSql=" & strSql
		End If
		Conn.execute(strSql)

		If Trim(estado_correlativo) = "2" Then

				strSql = "SELECT ID_CUOTA FROM CUOTA WHERE RUT_DEUDOR = '" & rut & "' AND COD_CLIENTE = '" & cliente & "'"
				strSql = strSql & " AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) AND FONO_AGEND_ULT_GES = '" & Trim(rsDIR("TELEFONO")) & "'"
				''REsponse.write "strSql=" & strSql
				set rsTemp= Conn.execute(strSql)

				Do until rsTemp.eof

					strSql = "SELECT TOP 1 FONO_AGEND, FECHA_AGENDAMIENTO , HORA_AGENDAMIENTO, COD_CATEGORIA, COD_SUB_CATEGORIA, COD_GESTION "
					strSql = strSql & " FROM GESTIONES G, GESTIONES_CUOTA GC"
					strSql = strSql & " WHERE G.ID_GESTION = GC.ID_GESTION"
					strSql = strSql & " AND G.RUT_DEUDOR='" & rut & "' AND G.COD_CLIENTE='" & cliente & "'"

					strSql = strSql & " AND FONO_AGEND IN (SELECT CAST(COD_AREA AS VARCHAR(2) )+'-'+TELEFONO FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" & rut & "' AND ESTADO IN (0,1))"

					strSql = strSql & " AND G.ID_GESTION IN "
					strSql = strSql & "(SELECT DISTINCT MAX(G.ID_GESTION) FROM GESTIONES G, GESTIONES_CUOTA GC"
					strSql = strSql & " WHERE G.ID_GESTION = GC.ID_GESTION"
					strSql = strSql & " AND G.RUT_DEUDOR='" & rut & "' AND G.COD_CLIENTE='" & cliente & "' AND FONO_AGEND IS NOT NULL"
					strSql = strSql & " AND FONO_AGEND <> '' AND GC.ID_CUOTA = " & rsTemp("ID_CUOTA") & " GROUP BY FONO_AGEND)"

					strSql = strSql & " ORDER BY FECHA_AGENDAMIENTO ASC"
					'Response.write "<br>" & strSql

					set rsGestiones = Conn.execute(strSql)
					If Not rsGestiones.Eof Then
						strFonoAgend = rsGestiones("FONO_AGEND")
						dtmFecAgend = rsGestiones("FECHA_AGENDAMIENTO")
						TX_HORAAGEND = rsGestiones("HORA_AGENDAMIENTO")
					End If


						strSql = "UPDATE CUOTA SET FONO_AGEND_ULT_GES = '" & strFonoAgend & "', FECHA_AGEND_ULT_GES = '" & dtmFecAgend & "', HORA_AGEND_ULT_GES = '" & TX_HORAAGEND & "'"
						strSql = strSql & " WHERE ID_CUOTA = " & rsTemp("ID_CUOTA")
						'Response.write "strSql=" & strSql
						Conn.execute(strSql)

						strSql = "SELECT FECHA_AGEND_ULT_GES FROM CUOTA WHERE ID_CUOTA = " & rsTemp("ID_CUOTA")
						set rsTempF= Conn.execute(strSql)
						If Not rsTempF.eof Then
							dtmFecAgend = rsTempF("FECHA_AGEND_ULT_GES")
						Else
							dtmFecAgend = ""
						End If

						strSql = "SELECT CONDICION =  CASE WHEN CONVERT(VARCHAR(10),FECHA_AGEND_ULT_GES,103) = CONVERT(VARCHAR(10),GETDATE(),103) AND "
						strSql = strSql & " 						   (CAST (SUBSTRING(HORA_AGEND_ULT_GES,0,3) AS INT) > CAST(SUBSTRING(CONVERT(VARCHAR(10),GETDATE(),108),0,3) AS INT) OR "
						strSql = strSql & " 						   HORA_AGEND_ULT_GES IS NULL)"
						strSql = strSql & " 					  THEN 1"
						strSql = strSql & " 					  WHEN CAST (CONVERT(VARCHAR(10),FECHA_AGEND_ULT_GES,103) AS DATETIME) > CAST (CONVERT(VARCHAR(10),GETDATE(),103) AS DATETIME)"
						strSql = strSql & " 					  THEN 1"
						strSql = strSql & " 				 ELSE 0"
						strSql = strSql & " END"
						strSql = strSql & " FROM  CUOTA WHERE ID_CUOTA = " & rsTemp("ID_CUOTA")

						''Response.write "<br>" & strSql

						set rsCondicion = Conn.execute(strSql )
						If Not rsCondicion.eof Then
							intCondicion = rsCondicion("CONDICION")
						Else
							intCondicion = 0
						End If


						'*************************************************************************************
						'*  Cuento los telefonos validos o sin auditar que no tienen gestiones al documento  *
						'*************************************************************************************

						strSql = "SELECT COUNT(*) AS CANT FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" & rut & "' AND ESTADO IN (0,1)"
						strSql = strSql & " AND CAST(COD_AREA AS VARCHAR(2)) + '-' + TELEFONO NOT IN (SELECT FONO_AGEND FROM GESTIONES G, GESTIONES_CUOTA GC"
						strSql = strSql & " WHERE G.ID_GESTION = GC.ID_GESTION AND G.RUT_DEUDOR = '" & rut & "' AND G.COD_CLIENTE='" & cliente & "' AND GC.ID_CUOTA = " & rsTemp("ID_CUOTA") & ")"
						set rsAnalisis = Conn.execute(strSql )
						If Not rsAnalisis.eof Then
							intCantidad = rsAnalisis("CANT")
						Else
							intCantidad = 9999999
						End If

						If (intCantidad = 0) and (intCondicion = 1) Then
							intMostrar = 0
						Else
							intMostrar = 1
						End If

						'Response.End




					rsTemp.movenext
					intCorrelativo = intCorrelativo + 1
				Loop
				rsTemp.close
				set rsTemp=nothing






		End If

		'Response.write "<br>estado_correlativo=" & estado_correlativo
		'Response.write "<br>intEstadoFono=" & intEstadoFono


		strSql = "UPDATE DEUDOR_TELEFONO SET HORA_DESDE = '" & strDesde & "',ANEXO = '" & strAnexo & "', HORA_HASTA = '" & strHasta & "', DIAS_ATENCION = '" & strDiasAtencion & "' WHERE RUT_DEUDOR = '" & rut & "' AND CORRELATIVO = '" & CORRELATIVO & "'"
		'Response.write "strSql=" & strSql
		Conn.execute(strSql)
	End If

rsDIR.movenext
loop

''Response.end

rsDIR.close
set rsDIR=nothing
cerrarscg()
%>


<script language="JavaScript" type="text/JavaScript">
	alert('Datos de telefonos actualizados');

	<%If strOrigen="deudor_telefonos" Then%>
		//window.navigate('deudor_telefonos.asp?strRUT_DEUDOR=<%=rut%>&strFonoAgestionar=<%=strFonoAgestionar%>');
	<%Else%>
		//window.navigate('mas_telefonos.asp?rut=<%=rut%>');
	<%End if%>
</script>

