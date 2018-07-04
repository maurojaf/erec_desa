<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>


<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<META HTTP-EQUIV="Cache-Control" CONTENT ="no-cache">
    <meta charset="utf-8">
	
   	<!--#include file="arch_utils.asp"-->
	<!--#include file="sesion_inicio.asp"-->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	rut_deudor = request("rut_deudor")
	intID_CUOTA = request("intID_CUOTA")
	strNroDoc = request("strNroDoc")
	strNroCuota = request("strNroCuota")


%>


<%
	nombre =  replace(Time(),":","")
	fileName ="Informe_Doc_" & nombre  & ".xls"
    Response.AddHeader "content-disposition", "attachment; filename=" & fileName
    Response.ContentType = "application/octet-stream"
    Response.Flush()
%>


</head>
<body>
	
		
			<table width="100%" border="1"  align="center">
	        <tr >
				<td  bgcolor="#989898"  align="center"><font color="#ffffff" > F.INGRESO.</font></td>
				<td bgcolor="#989898" align="center"><font color="#ffffff" align="center"> H.INGRESO</td>
				<td bgcolor="#989898" align="center"><font color="#ffffff" align="center"> DMI</td>
				<td bgcolor="#989898" align="center"><font color="#ffffff" align="center"> GESTION</td>
				<td bgcolor="#989898" align="center"><font color="#ffffff" align="center"> F.AGEND</td>
				<td bgcolor="#989898" align="center"><font color="#ffffff" align="center"> H.AGEND</td>
				<td bgcolor="#989898" align="center"><font color="#ffffff" align="center"> OBS</td>
				<td bgcolor="#989898" align="center"><font color="#ffffff" align="center"> CONTACTO</td>
				<td bgcolor="#989898" align="center"><font color="#ffffff" align="center"> MEDIO</td>
				<td bgcolor="#989898" align="center"><font color="#ffffff"> EJECUTIVO</td>
	        </tr>
	    	<tbody>
	     <%

				strSql = "SELECT G.COD_CATEGORIA,  G.COD_SUB_CATEGORIA, G.COD_GESTION, TC.CONTACTO, replace(replace(G.OBSERVACIONES,char(13),''),CHAR(9), '') OBSERVACIONES,GC.ID_CUOTA,G.ID_GESTION, "
				strSql = strSql & " G.FECHA_INGRESO, CONVERT(VARCHAR(5),convert(datetime, G.HORA_INGRESO), 108) HORA_INGRESO, G.FECHA_RETIRO,G.HORA_DESDE, G.HORA_HASTA, CONVERT(VARCHAR(10),G.FECHA_COMPROMISO,103) AS FECHA_COMPROMISO,"
				strSql = strSql & " G.FECHA_AGENDAMIENTO, G.HORA_AGENDAMIENTO,DMI = DATEDIFF(DAY,C.FECHA_VENC,G.FECHA_INGRESO), "

				strSql = strSql & " case "
				strSql = strSql & "	when G.TIPO_MEDIO_GESTION = 1 then "
				strSql = strSql & "		(SELECT UPPER(CONTACTO) AS CONTACTO FROM TELEFONO_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION) "
				strSql = strSql & "	when G.TIPO_MEDIO_GESTION = 2 then "
				strSql = strSql & "		(SELECT UPPER(CONTACTO) AS CONTACTO FROM EMAIL_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION) "
				strSql = strSql & "	when G.TIPO_MEDIO_GESTION = 3 then "
				strSql = strSql & "		(SELECT UPPER(CONTACTO) AS CONTACTO FROM DIRECCION_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION)  "
				strSql = strSql & " else '' end NOM_CONTACTO_GESTION "

				strSql = strSql & " , G.ID_USUARIO, DE.EMAIL EMAIL_ASOCIADO,convert(varchar,dd.COD_AREA) +'-'+ DD.TELEFONO TELEFONO_ASOCIADO "

				strSql = strSql & " FROM GESTIONES G INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION AND GC.ID_CUOTA = " & intID_CUOTA 
				strSql = strSql & "					 INNER JOIN CUOTA C ON GC.ID_CUOTA = C.ID_CUOTA"
				strSql = strSql & " 				 LEFT JOIN TELEFONO_CONTACTO TC ON G.ID_CONTACTO_GESTION = TC.ID_CONTACTO"
				strSql = strSql & " 				 LEFT JOIN DEUDOR_TELEFONO DD ON G.ID_MEDIO_GESTION = DD.ID_TELEFONO"
				strSql = strSql & "					 LEFT JOIN DEUDOR_EMAIL DE ON DE.ID_EMAIL = G.ID_MEDIO_GESTION"	

				strSql = strSql & " WHERE G.COD_CLIENTE='" & session("ses_codcli") & "'"
				strSql = strSql & " ORDER BY G.FECHA_INGRESO DESC , G.ID_GESTION DESC"
				'Response.write "<br>Sql=" & strSql
				AbrirSCG1()
				set rsGestCuota= Conn1.execute(strSql)
				Do until rsGestCuota.eof
					Obs=rsGestCuota("OBSERVACIONES")
					strContacto=rsGestCuota("CONTACTO")
					
					If Obs="" then
						Obs="SIN INFORMACION ADICIONAL"
					End if
         
					
		 
		 
						strContacto =rsGestCuota("NOM_CONTACTO_GESTION")

					strCodGestion= rsGestCuota("COD_CATEGORIA")& "-" & rsGestCuota("COD_SUB_CATEGORIA")& "-" &rsGestCuota("COD_GESTION")

					AbrirSCG2()
					strSql = "SELECT G.COD_CATEGORIA, G.COD_SUB_CATEGORIA, G.COD_GESTION,C.DESCRIPCION + '-' + S.DESCRIPCION + '-' +  G.DESCRIPCION as DESCRIP,"
					strSql = strSql + " ISNULL(G.COMUNICA,2) AS COMUNICA, G.GESTION_MODULOS, "
					strSql = strSql + " ISNULL(G.PRIORIDAD_GTEL,0) AS PRIORIDAD_GTEL, ISNULL(G.PRIORIDAD_GMAIL,0) AS PRIORIDAD_GMAIL "
					strSql = strSql & " FROM GESTIONES_TIPO_CATEGORIA C, GESTIONES_TIPO_SUBCATEGORIA S, GESTIONES_TIPO_GESTION G"
					strSql = strSql & " WHERE C.COD_CATEGORIA = S.COD_CATEGORIA"
					strSql = strSql & " AND C.COD_CATEGORIA = G.COD_CATEGORIA"
					strSql = strSql & " AND S.COD_SUB_CATEGORIA = G.COD_SUB_CATEGORIA"
					strSql = strSql & " AND G.COD_CLIENTE='" & session("ses_codcli") & "'"
					strSql = strSql & " AND CAST(G.COD_CATEGORIA AS VARCHAR(2)) + '-' + CAST(G.COD_SUB_CATEGORIA AS VARCHAR(2)) + '-' + CAST(G.COD_GESTION AS VARCHAR(2)) = '" & Trim(strCodGestion) & "'"

					'Response.write "strSql=" &strSql
					'Response.End
					SET rsNomGestion1=Conn2.execute(strSql)
					If Not rsNomGestion1.Eof Then
						strGestion = rsNomGestion1("DESCRIP")

						intGestionComunica = rsNomGestion1("COMUNICA")
						intGestionGtel = rsNomGestion1("PRIORIDAD_GTEL")
						intGestionGmail = rsNomGestion1("PRIORIDAD_GMAIL")
						strTipoGestion = rsNomGestion1("GESTION_MODULOS")

					Else
						strGestion = ""
					End If
					
					If intGestionGmail = 1 Then
					  strMedio = rsGestCuota("EMAIL_ASOCIADO")
					 ElseIf intGestionGtel = 1 Then
					  strMedio = rsGestCuota("TELEFONO_ASOCIADO")
					 else
					 strMedio =""
					 end if 
					
					CerrarSCG2()

					AbrirSCG2()
					If Trim(rsGestCuota("ID_USUARIO")) <> "" Then
						strLoginCobradorGest = TraeCampoId(Conn2, "LOGIN", rsGestCuota("ID_USUARIO"), "USUARIO", "ID_USUARIO")
					End If
					CerrarSCG2()

	     %>

				<tr>
					<td ><%=rsGestCuota("FECHA_INGRESO")%></td>
					<td ><%=rsGestCuota("HORA_INGRESO")%></td>
					<td ><%=rsGestCuota("DMI")%></td>
					<td ><%=strGestion%></td>
					<td ><%=rsGestCuota("FECHA_AGENDAMIENTO")%></td>
					<td ><%=rsGestCuota("HORA_AGENDAMIENTO")%></td>
					<td >
						<%=Obs%>
					</td>
					<td><%=UCASE(strContacto)%></td>
					<td><%=UCASE(strMedio)%></td>
					<td ><%=UCASE(strLoginCobradorGest)%></td>
				</tr>
	      <%
	      		Response.Flush
	      		rsGestCuota.movenext
			Loop
				CerrarSCG1()


	      %>
	      	</tbody>
	        </table>
</body>
</html>

