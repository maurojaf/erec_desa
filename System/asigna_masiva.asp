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
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->

	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	strCodCliente=	session("ses_codcli")
	strRutdeudor = Trim(Request("TA_RUT"))
	strUsuario = Trim(Request("TA_USUARIO"))
	intUsuarioSel = Trim(Request("CB_EJECUTIVO"))
	resp = Request("resp")

	strArrUsuario = Trim(Request("strArrUsuario"))
	intCuenta=0
	
	'Response.write "<br>intUsuarioSel=" & intUsuarioSel
	
	If Trim(strRutdeudor) <> "" then
		if ASC(RIGHT(strRutdeudor,1)) = 10 then strRutdeudor = Mid(strRutdeudor,1,len(strRutdeudor)-1)
		if ASC(RIGHT(strRutdeudor,1)) = 13 then strRutdeudor = Mid(strRutdeudor,1,len(strRutdeudor)-1)
	End if

	If Trim(strUsuario) <> "" then
		if ASC(RIGHT(strUsuario,1)) = 10 then strUsuario = Mid(strUsuario,1,len(strUsuario)-1)
		if ASC(RIGHT(strUsuario,1)) = 13 then strUsuario = Mid(strUsuario,1,len(strUsuario)-1)
	End if

	vRut = split(strRutdeudor,CHR(13))
	intTamvRut=ubound(vRut)

	vUsuario = split(strUsuario,CHR(13))
	intTamvUsuario=ubound(vUsuario)

	If Trim(Request("strRUT_DEUDOR")) <> "" Then session("IdCliente") = Trim(Request("strRUT_DEUDOR"))
	If Trim(strRUT_DEUDOR) = "" Then strRUT_DEUDOR = Trim(Request("strRUT_DEUDOR"))

	If Trim(strCodCliente) = "" Then strCodCliente=session("ses_codcli")

'--Calcula Objetos relacionados al tipo de cobranza (Interna, Externa), CB_COBRANZA Y CB_EJECUTIVO--'

	strCobranza = Request("CB_COBRANZA")

	abrirscg()

			strSql = "SELECT ISNULL(USA_COB_INTERNA,0) AS USA_COB_INTERNA"
			strSql = strSql & " FROM CLIENTE CL"
			strSql = strSql & " WHERE CL.COD_CLIENTE = '" & strCodCliente & "'"

			set RsCli=Conn.execute(strSql)
			If not RsCli.eof then
				intUsaCobInterna = RsCli("USA_COB_INTERNA")
			End if
			RsCli.close
			set RsCli=nothing

	cerrarscg()

	intVerCobExt = "1"
	intVerEjecutivos = "1"

	If TraeSiNo(session("perfil_emp")) = "Si" and strCobranza = "" and intUsaCobInterna = "1" Then
		strCobranza="INTERNA"
	ElseIf TraeSiNo(session("perfil_emp")) = "No" and strCobranza = "" then
		strCobranza="EXTERNA"
	End If

	If TraeSiNo(session("perfil_emp")) = "Si" Then

		intVerEjecutivos="0"
		intVerCobExt = "0"

	End If

	If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then
		sinCbUsario="0"
	End If

	'Response.write "<br>intTipoInforme=" & intTipoInforme
'---Fin codigo tipo de cobranza---'

'---Trae los rut de los deudores asignados al id_usuario obtenido de CB_EJECUTIVO---'

	If resp = "1" then
	
	abrirscg()
	
		strSql="SELECT D.RUT_DEUDOR"
		strSql= strSql & " FROM DEUDOR D INNER JOIN CUOTA C ON D.RUT_DEUDOR = C.RUT_DEUDOR AND D.COD_CLIENTE = C.COD_CLIENTE"
		strSql= strSql & " 			     INNER JOIN ESTADO_DEUDA ED ON ED.CODIGO = C.ESTADO_DEUDA"
		strSql= strSql & " 			     LEFT JOIN USUARIO U ON D.USUARIO_ASIG = U.ID_USUARIO"

		strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

		If Trim(strCobranza) = "INTERNA" Then
			strSql = strSql & " AND D.CUSTODIO IS NOT NULL"
		End if

		If Trim(strCobranza) = "EXTERNA" Then
			strSql = strSql & " AND D.CUSTODIO IS NULL"
		End if
		
		If Trim(intUsuarioSel) <> "-1" Then
			strSql= strSql & " AND (ISNULL(D.USUARIO_ASIG,0) = " & intUsuarioSel
			strSql= strSql & " OR (" & intUsuarioSel & " = 0 AND ISNULL(D.USUARIO_ASIG,0) NOT IN (SELECT ID_USUARIO FROM USUARIO WHERE ACTIVO=1 AND PERFIL_COB=1)))"
		End if
		
		strSql= strSql & " GROUP BY D.RUT_DEUDOR"	
		
		'Response.write "<br>strSql=" & strSql
		
		set RsCobsel=Conn.execute(strSql)
		
			strRutdeudor = ""
			intCuenta = 0
		
			If not RsCobsel.eof then
			
				Do While Not RsCobsel.Eof
			
					strRutdeudor1 = RsCobsel("RUT_DEUDOR")
					intcuenta = intcuenta + 1
					
					strRutdeudor = strRutdeudor1 + CHR(13) + strRutdeudor
				
					RsCobsel.movenext
				Loop								
			End if
			RsCobsel.close
			set RsCobsel=nothing
			
			'Response.write "<br>strSql=" & intcuenta
	
	cerrarscg()
	
		If intcuenta = 0 and intUsuarioSel = "0" then
		
			%>
				<script>
					alert('No existen deudores SIN ASIGNAR');
				</script>
			<%
		
		ElseIf intcuenta = 0 then
		
			%>
				<script>
					alert('Usuario seleccionado no posee deudores asignados');
				</script>
			<%
		
		End If
		
	End IF
	
%>
<%strTitulo="PANTALLA PRINCIPAL DE ASIGNACION"%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">
<%

	If Trim(Request("Asignar"))="1" Then

			If Trim(strArrUsuario) <> "" Then
				strArrUsuario = Mid(strArrUsuario,2,len(strArrUsuario))
				vArrUsuario = split(strArrUsuario,";")
				intTamvArrUsuario=ubound(vArrUsuario)
			End If

			For i1 = 0 to intTamvArrUsuario
				intSaldo = Request(strObjeto1)
				strObjeto = "CH_USUARIO_" & vArrUsuario(i1)
				If UCASE(Request(strObjeto)) <>  "" Then
					intUsuario = vArrUsuario(i1)
					intUsuarioCH = vArrUsuario(i1)
				End If

			Next

		AbrirScg()
			if strProceso = "" Then
				For indice = 0 to intTamvRut
					strIdRut = Trim(Replace(vRut(indice), chr(10),""))
					If Trim(strIdRut) <> "" Then

						'Response.write "<br>CH_USUARIO=====" & Request("CH_USUARIO")

						If Trim(intUsuarioCH) = "" Then
							intUsuario = ucase(Trim(Replace(vUsuario(indice), chr(10),"")))
						End If

						If intUsuario = "0" Then
							intUsuario = "NULL"
						End If

						'Response.write "<br>intUsuario=====" & intUsuario

						strSql = "EXEC Proc_Asignacion_cobradores '" & strCodCliente & "','" & strIdRut & "'," & session("session_idusuario") & ", " & intUsuario
						set rsAsigna = Conn.execute(strSql)
						''Response.write "<br>strSql=" & strSql

					End If
				Next
						%>
							<script>
								alert('Asignacion realizada correctamente');
							</script>
						<%
			End If
		CerrarScg()
	End If

	AbrirSCG()

				'--Obtiene la información relacionada con la deuda total--'

				strSql = "SELECT COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS, SUM(DOCUMENTOS_CASO) AS TOTAL_DOC, SUM(CAPITAL_CASO) AS MONTO_TOTAL,"
				strSql= strSql & " MIN(CAPITAL_CASO) AS MIN_CAPITAL_CASO, MAX(CAPITAL_CASO) AS MAX_CAPITAL_CASO,MIN(MIN_CAPITAL_DOC) AS MIN_CAPITAL_DOC,MAX(MAX_CAPITAL_DOC) AS MAX_CAPITAL_DOC"
				strSql= strSql & " FROM (SELECT D.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"
				strSql= strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO"
				strSql= strSql & " 				INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"

				strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'"

				If Trim(strCobranza) = "INTERNA" Then
					strSql = strSql & " AND D.CUSTODIO IS NOT NULL"
				End if

				If Trim(strCobranza) = "EXTERNA" Then
					strSql = strSql & " AND D.CUSTODIO IS NULL"
				End if

				strSql= strSql & " GROUP BY D.RUT_DEUDOR) AS PP"

				'Response.write "<br>strSql=" & strSql
				'response.end()
				set RsInf=Conn.execute(strSql)

				intTotalCasos= "0"
				intMinCaso= "0"
				intMaxCaso= "0"
				intTotalDoc= "0"
				intTotalMonto= "0"
				intMinDoc= "0"
				intMaxDoc= "0"

				If not RsInf.eof then

				intTotalCasos= RsInf("TOTAL_CASOS")
				intMinCaso= RsInf("MIN_CAPITAL_CASO")
				intMaxCaso= RsInf("MAX_CAPITAL_CASO")
				intTotalDoc= RsInf("TOTAL_DOC")
				intTotalMonto= RsInf("MONTO_TOTAL")
				intMinDoc= RsInf("MIN_CAPITAL_DOC")
				intMaxDoc= RsInf("MAX_CAPITAL_DOC")

				intPromMontoCaso= intTotalMonto/intTotalCasos

				intPromDocCaso= intTotalDoc/intTotalCasos

				intPromMontoDcumento= intTotalMonto/intTotalDoc

				End If

	CerrarSCG()

	AbrirSCG()

				'--Obtiene la información relacionada con los casos sin asignar--'

				strSql = "SELECT COUNT(DISTINCT RUT_DEUDOR) AS TOTAL_CASOS_SA, SUM(DOCUMENTOS_CASO) AS TOTAL_DOC_SA, SUM(CAPITAL_CASO) AS MONTO_TOTAL_SA"
				strSql= strSql & " FROM (SELECT D.RUT_DEUDOR,SUM(VALOR_CUOTA) AS CAPITAL_CASO,"
				strSql= strSql & " COUNT(ID_CUOTA) AS DOCUMENTOS_CASO,MIN(VALOR_CUOTA) AS MIN_CAPITAL_DOC,MAX(VALOR_CUOTA) AS MAX_CAPITAL_DOC"
				strSql= strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA= ED.CODIGO "
				strSql= strSql & " INNER JOIN DEUDOR D ON C.COD_CLIENTE=D.COD_CLIENTE AND C.RUT_DEUDOR=D.RUT_DEUDOR"

				strSql= strSql & " WHERE ED.ACTIVO=1 AND D.COD_CLIENTE= '" & strCodCliente & "'" 
				strSql= strSql & " AND (ISNULL(D.USUARIO_ASIG,0)=0 OR ISNULL(D.USUARIO_ASIG,0) NOT IN (SELECT ID_USUARIO FROM USUARIO WHERE ACTIVO=1 AND PERFIL_COB=1))"
				
				If Trim(strCobranza) = "INTERNA" Then
					strSql = strSql & " AND D.CUSTODIO IS NOT NULL"
				End if

				If Trim(strCobranza) = "EXTERNA" Then
					strSql = strSql & " AND D.CUSTODIO IS NULL"
				End if

				strSql= strSql & " GROUP BY D.RUT_DEUDOR) AS PP"


				set RsInf2=Conn.execute(strSql)

				intTotalCasosSA= "0"
				intTotalDocSA= "0"
				intTotalMontoSA= "0"

				If not RsInf2.eof then

				intTotalCasosSA= RsInf2("TOTAL_CASOS_SA")
				intTotalDocSA= RsInf2("TOTAL_DOC_SA")
				intTotalMontoSA= RsInf2("MONTO_TOTAL_SA")

				End If

	CerrarSCG()

	AbrirSCG()

				'--Obtiene el total de cobradores--'

				strSql = "SELECT COUNT(*) AS TOTAL_COBRADORES"
				strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO=UC.ID_USUARIO AND UC.COD_CLIENTE= '" & strCodCliente & "'"
				strSql= strSql & " WHERE ACTIVO = 1 AND PERFIL_COB = 1 AND PERFIL_SUP=0"

				If Trim(strCobranza) = "INTERNA" Then
					strSql = strSql & " AND PERFIL_EMP=1"
				End if

				If Trim(strCobranza) = "EXTERNA" Then
					strSql = strSql & " AND PERFIL_EMP=0"
				End if

				'Response.write "<br>strSql=" & strSql

				set RsInf3=Conn.execute(strSql)

				If not RsInf3.eof then

				intTotalCobradores=RsInf3("TOTAL_COBRADORES")

				End If

	CerrarSCG()
%>

<div class="titulo_informe">ASIGNACIÓN MASIVA</div>
<br>
	<table width="90%" border="0" align="center" class="estilo_columnas">
		<thead>
		<tr>
			<td colspan="1" height="20">&nbsp;COBRANZA</td>
			<td colspan="2" height="20">&nbsp;ASIGNACIÓN</td>
		 </tr>
		</thead>

		<tr>
			<td>
				<select name="CB_COBRANZA" id="CB_COBRANZA" onChange="envia(0);">

					<%If Trim(intVerCobExt) = "1" and Trim(intUsaCobInterna) = "1" Then%>
						<option value="0" <%If Trim(strCobranza) ="" Then Response.write "SELECTED"%>>TODOS</option>
					<%End If%>

					<%If Trim(intUsaCobInterna) = "1" Then%>
						<option value="INTERNA" <%If Trim(strCobranza) ="INTERNA" Then Response.write "SELECTED"%>>INTERNA</option>
					<%End If%>

					<%If Trim(intVerCobExt) = "1" Then%>
						<option value="EXTERNA" <%If Trim(strCobranza) ="EXTERNA" Then Response.write "SELECTED"%>>EXTERNA</option>
					<%End If%>

				</select>
			</td>
	
			<td>
				<%

					AbrirSCG()
						strSql= "proc_Inf_AsignacionEjecutivos '"&strCodCliente&"'"
						

				%>
				<select name="CB_EJECUTIVO" ID="CB_EJECUTIVO"  onChange="envia(1);">
					<option value="">SELECCIONE</option>
					<option value="-1" <%if Trim(intUsuarioSel)="-1" then response.Write("Selected") end if%>>TODOS</option>
					<%

				
						set rsUsu=Conn.execute(strSql)
							if not rsUsu.eof then
								do until rsUsu.eof
								%>
								<option value="<%=rsUsu("ID_USUARIO")%>"
								<%if Trim(intUsuarioSel)=Trim(rsUsu("ID_USUARIO")) then
									response.Write("Selected")
								end if%>>
								<%=ucase(rsUsu("USUARIO_ASIG"))%></option>

								<%rsUsu.movenext
								loop
							end if
							rsUsu.close
							set rsUsu=nothing
					CerrarSCG()
					''Response.End
					%>
				</select>
			</td>
			
			<td width="10%" Align="right">
				<input name="me_" type="button" id="me_" class="fondo_boton_100" onClick="Asignar();" value="Procesar">
			</td>
		</tr>
		
	</table>

<table align="center" style="width:90%;" >
	<tr>
		<td valign="top" width="50%">
		<table width="100%">
		<tr>
			<td valign="top">
				<table width="100%" border="0" ALIGN="CENTER" class="estilo_columnas">
				<thead>
			   	<tr HEIGHT="20" >
					<TD Colspan="6" ALIGN="Left" class="subtitulo_informe">> Resumen Cartera</TD>
				</tr>
				
					<tr >
						<TD WIDTH="20%" ALIGN="CENTER">
							Casos
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Documentos
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Monto
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Ejecutivos
						</TD>
					</tr>
				</thead>
					<tr >
						<TD ALIGN="CENTER">
							&nbsp;<%=FN(intTotalCasos,0)%>
						</TD>
						<TD ALIGN="CENTER">
							&nbsp;<%=FN(intTotalDoc,0)%>
						</TD>
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intTotalMonto,0)%>
						</TD>
						<TD ALIGN="CENTER">
							<%=FN(intTotalCobradores,0)%>
						</TD>
					</tr>
					<thead>
					<tr >
						<TD WIDTH="20%" ALIGN="CENTER">
							Monto / Caso
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Monto / Documento
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Documento / Caso
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							&nbsp;
						</TD>
					</tr>
					</thead>
					<tr >
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intPromMontoCaso,0)%>
						</TD>
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intPromMontoDcumento,0)%>
						</TD>
						<TD ALIGN="CENTER">
							<%=FN(intPromDocCaso,2)%>
						</TD>
						<TD ALIGN="CENTER">
							&nbsp;
						</TD>
					</tr>
					<thead>
					<tr >
						<TD WIDTH="20%" ALIGN="CENTER">
							Casos Sin Asignar
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Doc Sin Asignar
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							Monto Sin Asignar
						</TD>
						<TD WIDTH="20%" ALIGN="CENTER">
							&nbsp;
						</TD>
					</tr>
					</thead>
					<tr >
						<TD ALIGN="CENTER">
							<%=FN(intTotalCasosSA,0)%>
						</TD>
						<TD ALIGN="CENTER">
							&nbsp;<%=FN(intTotalDocSA,0)%>
						</TD>
						<TD ALIGN="CENTER">
							$&nbsp;<%=FN(intTotalMontoSA,0)%>
						</TD>
						<TD ALIGN="CENTER">
							&nbsp;
						</TD>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td valign="top">
			<table  width="100%" border="0" class="intercalado" style="width:100%;" >
			<thead>
			   <tr HEIGHT="20">
					<TD Colspan="7" ALIGN="Left" class="subtitulo_informe">> Asiganción por ejecutivo</TD>
				</tr>
				<tr >
					<TD ALIGN="CENTER">
						Marcar
					</TD>
					<TD width="23%" ALIGN="CENTER">
						Ejecutivo
					</TD>
					<TD ALIGN="CENTER">
						Estado
					</TD>
					<TD ALIGN="CENTER">
						Codigo
					</TD>
					<TD ALIGN="CENTER">
						Casos
					</TD>
					<TD ALIGN="CENTER">
						Documentos
					</TD>
					<TD ALIGN="CENTER">
						Monto
					</TD>
					<TD ALIGN="CENTER">
						Rut / Doc
					</TD>
				</tr>
			</thead>
			<tbody>
<%
	AbrirSCG()
			strSql= "proc_Inf_AsignacionEjecutivos '"&strCodCliente&"'"

				'Response.write "<br>strSql=" & strSql

				set RsInf=Conn.execute(strSql)

				intTamvConcepto = 0
				intTTCasos=0
				intTTDoc=0
				intTTMonto=0

				if not RsInf.eof then
					strArrUsuario = ""
					do until RsInf.eof

					strNomUsuario = RsInf("USUARIO_ASIG")
					intTotalCasos= RsInf("TOTAL_CASOS")
					intTotalDoc= RsInf("TOTAL_DOC")
					intTotalMonto= RsInf("MONTO_TOTAL")
					intIdUsuario= RsInf("ID_USUARIO")
					intActivo = RsInf("ACTIVO")

					intTTCasos= intTTCasos + intTotalCasos
					intTtDoc= intTtDoc + intTotalDoc
					intTTMonto= intTTMonto + intTotalMonto

					If intTotalCasos>"0" then
					intDocRut= intTotalDoc/intTotalCasos
					Else
					intDocRut="0"
					End If

					'Response.write "<br>intActivo=" & intActivo
					
					If intActivo="1" then
					intActivo= "Activo"
					Else
					intActivo= "No Activo"
					End If

					strArrUsuario =strArrUsuario & ";" & intIdUsuario

					%>
					<tr >

						<TD><INPUT TYPE="checkbox" NAME="CH_USUARIO_<%=intIdUsuario%>" value="<%=intIdUsuario%>" onClick="desmarcar_boxes(this);"></TD>

						<TD ALIGN="left">
							<%=strNomUsuario%>
						</td>
						<TD ALIGN="left">
							<%=intActivo%>
						</td>
						<TD ALIGN="RIGHT">
							<%=intIdUsuario%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTotalCasos,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							<%=FN(intTotalDoc,0)%>
						</TD>
						<TD ALIGN="RIGHT">
							$&nbsp;<%=FN(intTotalMonto,0)%>
						</TD>
						<TD ALIGN="CENTER">
							<%=FN(intDocRut,1)%>
						</TD>
					</tr>

					<%RsInf.movenext


					strArrConcepto = strArrConcepto & ";" & "CH_USUARIO_" & Trim(intIdUsuario)
					//intTamvConcepto = intTamvConcepto + 1
					loop
				end if
				RsInf.close
				set RsInf=nothing


				vArrConcepto = split(strArrConcepto,";")
				intTamvConcepto = ubound(vArrConcepto)

				If intTTCasos >"0" then
				intDocRutGeneral = intTtDoc / intTTCasos
				Else
				intDocRutGeneral=0
				End If

				If intTtDoc >"0" then
				intPromDocGeneral = intTTMonto / intTtDoc
				Else
				intPromDocGeneral=0
				End If

				If intTTCasos >"0" then
				intPromRutGeneral = intTTMonto / intTTCasos
				Else
				intPromRutGeneral=0
				End If

	CerrarSCG()
%>
		</tbody>
				<tr class="totales">
					<TD colspan="4" ALIGN="CENTER">
						Totales
					</TD>
					<TD ALIGN="RIGHT">
						<%=FN(intTTCasos,0)%>
					</TD>
					<TD ALIGN="RIGHT">
						<%=FN(intTtDoc,0)%>
					</TD>
					<TD ALIGN="RIGHT">
						<%=FN(intTTMonto,0)%>
					</TD>
					<TD ALIGN="CENTER">
						<%=FN(intDocRutGeneral,1)%>
					</TD>
				</tr>

			</table>
			</td>
		</tr>
		</table>
		</td>
		<td width="50%">
			<table width="100%" ALIGN="left" cellSpacing="0" cellPadding="0">
				<TR>
					<TD class=" subtitulo_informe" width="50%" HEIGHT="20" style="text-align:center;">
						> RUT&nbsp;&nbsp;(<%=intcuenta%>&nbsp;REGISTROS)
					</TD>
					<TD class=" subtitulo_informe" width="50%" style="text-align:center;">
						> USUARIO (CODIGO)
					</TD>
				</TR>
				<TR>
					<TD class="hdr_i" width="50%" style="text-align:center; background-color:#989898;">
						<TEXTAREA NAME="TA_RUT" ROWS="25" COLS="30"><%=strRutdeudor%></TEXTAREA>
					</TD>
					<TD class="hdr_i estulo_columna_individual" width="50%" style="text-align:center; border-left:5px solid #FFFFFF;  background-color:#989898;">
						<TEXTAREA NAME="TA_USUARIO" ROWS="25" COLS="30"  onfocus="PorUsuario()"><%=strUsuario%></TEXTAREA>
					</TD>
				</TR>
			</table>
		</td>
	</tr>
</table>
<INPUT TYPE="hidden" NAME="strArrUsuario" value="<%=strArrUsuario%>">
</form>



</body>
</html>


<script language="JavaScript" type="text/JavaScript">
function envia(dato){
	if (dato == '0') {
	datos.action='asigna_masiva.asp?TA_RUT=&CB_EJECUTIVO=';
	datos.submit();
	}
	if (dato == '1') {
	datos.action='asigna_masiva.asp?resp=1';
	datos.submit();
	}
}
function Asignar(){
	//alert(validaCheckbox());
	//alert(IndexOf(document.forms[0].TA_USUARIO.value.length));

	//var strArrayRut = document.forms[0].TA_RUT.value;
	//var strArrayUsr = document.forms[0].TA_USUARIO.value;



	//alert(trim(document.forms[0].TA_RUT.value));
	//alert(trim(document.forms[0].TA_USUARIO.value));

	var strArrayRut = trim(document.forms[0].TA_RUT.value);
	var strArrayUsr = trim(document.forms[0].TA_USUARIO.value);



	var arrayRUT = strArrayRut.split(String.fromCharCode(13));
	var arrayUSR = strArrayUsr.split(String.fromCharCode(13));

	//alert(arrayRUT.length);
	//alert(arrayUSR.length);


	if(document.forms[0].TA_RUT.value == ''){
		alert("Error de ingreso: Debe Ingresar al menos 1 rut");
		return false;
	}

	else if((document.forms[0].TA_USUARIO.value == '') && (validaCheckbox()==false)){
		alert("Error de ingreso: Debe Seleccionar un usuario o ingresar un codigo de usuario");
		return false;
	}
	else if((validaCheckbox()==false) && (arrayRUT.length != arrayUSR.length)){
		alert("Error de ingreso: Las columnas Rut y Usuario no contienen los mismos elementos.");
		return false;
	}
	datos.me_.disabled = true;
	datos.action='asigna_masiva.asp?Asignar=1&CB_EJECUTIVO=';
	datos.submit();
}

function validaCheckbox(){
	var intValidacion = false;
		<% For i=1 TO intTamvConcepto %>
			if (document.forms[0].<%=vArrConcepto(i)%>.checked == true)
				intValidacion = true;
		<% Next %>

		return intValidacion;
}

function PorUsuario(){
	<% For i=1 TO intTamvConcepto %>
		document.forms[0].<%=vArrConcepto(i)%>.checked=false;
	<% Next %>

}

function desmarcar_boxes(objeto){
	<% For i=1 TO intTamvConcepto%>
		if (document.forms[0].<%=vArrConcepto(i)%>.checked == true) {
			document.forms[0].<%=vArrConcepto(i)%>.checked=false;
			<% If i = intTamvConcepto Then %>
			objeto.checked=true;
			<% End If %>
		}
		else
		{
			objeto.checked=true;
		}

	<% Next %>

	document.forms[0].TA_USUARIO.value = '';

	//}
}
function trim(myString)
{
return myString.replace(/^\s+/g,'').replace(/\s+$/g,'')
}
</script>
