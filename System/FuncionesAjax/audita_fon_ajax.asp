<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/lib.asp"-->

<%

Response.CodePage = 65001
Response.charset="utf-8"

rut 				=request("rut")
strOrigen 			=request("strOrigen")
strFonoAgestionar 	=request("strFonoAgestionar")
strAnexo  			=request("strAnexo")
IF(request("strTipoContacto") = "") then strTipoContacto = "null" else strTipoContacto = request("strTipoContacto") end if
estado_correlativo 	=request("estado_correlativo")
CORRELATIVO 		=request("CORRELATIVO")
strHasta 			=request("TX_HASTA")
strDesde 			=request("TX_DESDE")
strDiasAtencion 	=request("strDiasAtencion")

cliente 			=session("ses_codcli")
intIdUsuario        =session("session_idusuario")

abrirscg()




	If Trim(estado_correlativo) <> "" and Not IsNull(estado_correlativo) Then
		strSql="SELECT ESTADO, IsNull(cast(COD_AREA as varchar(3)) + '-' + telefono,'') as TELEFONO  FROM DEUDOR_TELEFONO WHERE RUT_DEUDOR = '" & rut & "' and CORRELATIVO='" & CORRELATIVO & "'"
		set rsValida=Conn.execute(strSql)
		
		If Not rsValida.eof Then
			intEstadoFono 	= rsValida("ESTADO")
			TELEFONO 		= rsValida("TELEFONO")
		End If

		strSql = "UPDATE DEUDOR_TELEFONO SET ESTADO='" & cint(estado_correlativo) & "', HORA_DESDE = '" & strDesde & "',ANEXO = '" & strAnexo & "',IdTipoContacto = "& strTipoContacto &", HORA_HASTA = '" & strHasta & "', DIAS_ATENCION = '" & strDiasAtencion & "', FECHA_REVISION = GETDATE(), USR_REVISION = " & intIdUsuario
		strSql = strSql & " WHERE RUT_DEUDOR = '" & rut & "' AND CORRELATIVO = '" & CORRELATIVO & "'"
		'Response.write "strSql=" & strSql
		
		Conn.execute(strSql)
		
		If Trim(estado_correlativo) = "2" Then

				strSql = "SELECT ID_CUOTA FROM CUOTA WHERE RUT_DEUDOR = '" & rut & "' AND COD_CLIENTE = '" & cliente & "'"
				strSql = strSql & " AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) AND FONO_AGEND_ULT_GES = '" & Trim(TELEFONO) & "'"
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

	End If





cerrarscg()
%>


