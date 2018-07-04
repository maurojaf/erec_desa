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
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	<title>VENTAS WEB</title>
</head>

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	''strCOD_CLIENTE=Request("CB_CLIENTE")
	AgregarCasos=Request("AgregarCasos")
	EliminarCasos=Request("EliminarCasos")
	EliminarCampana=Request("EliminarCampana")

	strCOD_CLIENTE=session("ses_codcli")
	If trim(strCOD_CLIENTE) = "" Then
		strCOD_CLIENTE = request("strCOD_CLIENTE")
	End if

	intCodCampana=Request("CB_CAMPANA")

	If Trim(intCodCampana) = "" Then
		strNombreCampana = ""
		strDescCampana = ""
	End if

	intSoloEmpresa=Request("CH_EMPRESA")
	strCampana=Request("CH_CAMPANA")

	strRut = Trim(Request("TA_RUT"))
	strUsuario = Trim(Request("TA_USUARIO"))

	strCota = Trim(Request("TX_COTA"))
	If Trim(strCota) <> "" Then
		strTop = "TOP " & strCota
	End If




	If Trim(strRut) <> "" then
		if ASC(RIGHT(strRut,1)) = 10 then strRut = Mid(strRut,1,len(strRut)-1)
		if ASC(RIGHT(strRut,1)) = 13 then strRut = Mid(strRut,1,len(strRut)-1)
	End if

	If Trim(strUsuario) <> "" then
		if ASC(RIGHT(strUsuario,1)) = 10 then strUsuario = Mid(strUsuario,1,len(strUsuario)-1)
		if ASC(RIGHT(strUsuario,1)) = 13 then strUsuario = Mid(strUsuario,1,len(strUsuario)-1)
	End if

	vRut = split(strRut,CHR(13))
	intTamvRut=ubound(vRut)

	vUsuario = split(strUsuario,CHR(13))


	If UCASE(intSoloEmpresa)="ON" Then strChecked = "checked"

	strTabla=Request("TX_TABLA")


	If Trim(Request("strRUT_DEUDOR")) <> "" Then session("IdCliente") = Trim(Request("strRUT_DEUDOR"))
	If Trim(strRUT_DEUDOR) = "" Then strRUT_DEUDOR = Trim(Request("strRUT_DEUDOR"))



	If Trim(strCOD_CLIENTE) = "" Then strCOD_CLIENTE=session("ses_codcli")
	If Trim(intCodCampana) = "" Then intCodCampana = "99999999990"


	AbrirScg()
	strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana
	set rsCampana=Conn.execute(strSql)
	If not rsCampana.eof Then
		strNombreCampana = rsCampana("NOMBRE")
		strDescCampana = rsCampana("DESCRIPCION")
	End if
	CerrarScg()



	intUsuarios =  Request("CH_USUARIO")

	VUsuarios = Split(intUsuarios, ",")
	n=0
	For Each XX in VUsuarios
		 n=n+1
	Next

	'Response.write "<br>intCodCampana=" & intCodCampana
	'Response.End
%>
<%strTitulo="PANTALLA PRINCIPAL DE ASIGNACION"%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">
<%

	If Trim(strCampana)="1" Then
		strNombreCampana = request("TX_NOMCAMPANA")
		strDescCampana = request("TX_DESCCAMPANA")
		AbrirScg()
			strSql = "INSERT INTO CAMPANA (COD_CLIENTE, COD_REMESA, FECHA_CREACION, NOMBRE, DESCRIPCION) "
			strSql = strSql & "VALUES ('" & strCOD_CLIENTE & "'," & intCodCampana & ", getdate(), '" & strNombreCampana & "','" & strDescCampana & "')"
			''Response.write "strSql=" & strSql
			set rsInsert = Conn.execute(strSql)
		CerrarScg()

		AbrirScg()
			strSql = "SELECT MAX(ID_CAMPANA) as MAXIDCAMPANA FROM CAMPANA "
			set rsCampana = Conn.execute(strSql)
			If Not rsCampana.Eof Then
				intIdCampana = rsCampana("MAXIDCAMPANA")
			End If
		CerrarScg()

			For indice = 0 to intTamvRut
				strIdRut = Trim(Replace(vRut(indice), chr(10),""))
				intCodigo = ucase(Trim(Replace(vUsuario(indice), chr(10),"")))

				intusuario = intCodigo

				AbrirScg()
					strSql = "UPDATE CUOTA SET USUARIO_ASIG = " & intusuario & " , FECHA_ASIGNACION = getdate() WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)
				CerrarScg()

				AbrirScg()
					strSql = "UPDATE DEUDOR SET ID_CAMPANA = " & intIdCampana & " WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)
				CerrarScg()

				'Response.write "<br>strSql=" & strSql
				'set rsUpdate = Conn.execute(strSql)
			Next


	End If


	If Trim(AgregarCasos) = "1" Then

			intIdCampana = Request("hdidCampana")
			For indice = 0 to intTamvRut
				strIdRut = Trim(Replace(vRut(indice), chr(10),""))
				intCodigo = ucase(Trim(Replace(vUsuario(indice), chr(10),"")))

				intusuario = intCodigo
				AbrirScg()
					strSql = "UPDATE CUOTA SET USUARIO_ASIG = " & intusuario & " , FECHA_ASIGNACION = getdate() WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)
				CerrarScg()

				AbrirScg()
					strSql = "UPDATE DEUDOR SET ID_CAMPANA = " & intIdCampana & " WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)
				CerrarScg()

			Next


	End If

	If Trim(EliminarCasos) = "1" Then

			intIdCampana = Request("hdidCampana")
			For indice = 0 to intTamvRut
				strIdRut = Trim(Replace(vRut(indice), chr(10),""))
				AbrirScg()
					strSql = "UPDATE CUOTA SET USUARIO_ASIG = NULL , FECHA_ASIGNACION = getdate() WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)
				CerrarScg()

				AbrirScg()
					strSql = "UPDATE DEUDOR SET ID_CAMPANA = NULL WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)
				CerrarScg()
			Next
	End If

	If Trim(EliminarCampana) = "1" Then

		intIdCampana = Request("hdidCampana")
		AbrirScg()
			strSql = "UPDATE DEUDOR SET ID_CAMPANA = NULL WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intIdCampana
			set rsUpdate = Conn.execute(strSql)

			strSql = "DELETE FROM CAMPANA WHERE ID_CAMPANA = " & intIdCampana
			set rsUpdate = Conn.execute(strSql)
		CerrarScg()

	End If





	If Trim(Request("Asignar"))="1" Then
		intSel = request("OP_SEL")

		If Trim(intSel) = "3" Then


				strSql = "UPDATE CUOTA SET USUARIO_ASIG = NULL , FECHA_ASIGNACION = getdate() WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
				If UCASE(intSoloEmpresa)="ON" Then
					strSql = strSql & " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) > 50000000"
				End if

				If Trim(intCodCampana) <> "0" Then
					strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
				End if

				If trim(intUsuarios) <> "" Then
					strSql = strSql & " AND USUARIO_ASIG IN (" & intUsuarios & ")"
					'Response.write "<br>strSql=" & strSql
					'Response.End

					AbrirScg()
					set rsUpdate = Conn.execute(strSql)
					CerrarScg()
					%>
					<script>
							alert('Asignacion eliminada correctamente');
					</script>
					<%
				Else
					%>
					<script>
							alert('Debe seleccionar al menos un usuario para poder eliminar');
					</script>
					<%
				End If
		End If

		If Trim(intSel) = "1" or Trim(intSel) = "2" Then
				if Trim(intSel) = "1" Then
					strCondicion = " AND USUARIO_ASIG is NULL "
				End if
				If Trim(strTabla) <> "" Then
					strCondicion2 = " AND RUT_DEUDOR IN (SELECT RUT FROM " & strTabla & ")"
				End if
				If UCASE(intSoloEmpresa)="ON" Then
					strCondicion3 = " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) > 50000000"
				End if

				strSql = "SELECT DISTINCT " & strTop & " RUT_DEUDOR , SUM(SALDO) FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"

				If Trim(intCodCampana) <> "0" Then
					strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
				End if


				strSql = strSql & " " & strCondicion & strCondicion2 & " GROUP BY RUT_DEUDOR ORDER BY SUM(SALDO) DESC"

				'Response.write "<br>PRINCIPAL strSql=" & strSql


				Server.ScriptTimeout = 9000
				Conn2.ConnectionTimeout = 9000
				Conn1.ConnectionTimeout = 9000
				AbrirScg2()

				set rsAsigna2 = Conn2.execute(strSql)

				intUsuarioAsig = ""
				n = 0
				If UBOUND(VUsuarios) >= 0 Then
					Do While Not rsAsigna2.Eof
						intUsuarioAsig = VUsuarios(n)

						strSql1 = "UPDATE CUOTA SET USUARIO_ASIG = " & intUsuarioAsig & " , FECHA_ASIGNACION = getdate() WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"

						If Trim(intCodCampana) <> "0" Then
							strSql1 = strSql1 & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & rsAsigna2("RUT_DEUDOR") & "' AND ID_CAMPANA = " & intCodCampana & ")"
						End if
						strSql1 = strSql1 & " AND RUT_DEUDOR = '" & rsAsigna2("RUT_DEUDOR") & "'"

						'strSql1 = "EXEC proc_EjecutaSentencia '11'"

						'RESPONSE.WRITE "<br>VUsuarios=" & UBOUND(VUsuarios)
						'RESPONSE.WRITE "<br>Conn1=" & Conn1

						'RESPONSE.WRITE "<br>strSql=" & strSql1
						''RESPONSE.End

						AbrirSCG1()
							set rsModif = Conn1.execute(strSql1)
						CerrarSCG1()

						If Trim(strCampana)="1" Then

							strSql = "UPDATE DEUDOR SET ID_CAMPANA = " & intIdCampana & " WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & rsAsigna2("RUT_DEUDOR") & "'"
							AbrirSCG1()
								set rsUpdate2 = Conn1.execute(strSql)
							CerrarSCG1()

						End If

						''Response.write "<br>Asignando rut :" & rsAsigna2("RUT_DEUDOR")

						rsAsigna2.MoveNext
						n=n+1
						If UBOUND(VUsuarios) + 1 = n Then
							n=0
						End If

					Loop
				End If
				%>
					<script>
							alert('Asignacion realizada correctamente');
					</script>
				<%
			CerrarScg2()
		End If

	End If

	intContacto = Trim(request("contacto"))
	if Trim(strRUT_DEUDOR) <> "" and Trim(strCOD_CLIENTE) <> "" Then


		strSql = "SELECT RUT_DEUDOR, USUARIO_ASIG FROM CUOTA "
		strSql = strSql & "WHERE COD_CLIENTE = " & strCOD_CLIENTE & " AND "
		strSql = strSql & "RUT_DEUDOR = '" & strRUT_DEUDOR & "'"

		AbrirScg()
			set rsDEU=Conn.execute(strSql)
			if not rsDEU.eof then
				'Response.write "<br>strSql = " & strSql
				strRUT_DEUDOR = rsDEU("RUT_DEUDOR")
				strEjeAsig = rsDEU("USUARIO_ASIG")

				strSql = "SELECT NOMBRE_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND "
				strSql = strSql & "RUT_DEUDOR = '" & strRUT_DEUDOR & "'"

				AbrirScg1()
				set rsTemp=Conn.execute(strSql)
				If not rsTemp.eof then
					strNombreAMostrar = rsTemp("NOMBRE_DEUDOR")
				Else
					strNombreAMostrar = ""
				End If
				CerrarScg1()
				existe = "si"
			else
				strRUT_DEUDOR = ""
				strEjeAsig = ""
				existe = "si"
			end if

			rsDEU.close
			set rsDEU=nothing
		CerrarScg()

	End If


%>
<div class="titulo_informe">ASIGNACIÓN MASIVA</div>
<br>

<table width="90%" border="0" ALIGN="CENTER">
	<tr>
		<td height="331" valign="top">
			<table width="100%" border="0" class="estilo_columnas">
			<thead>
			<tr>
			<td class="">
				<strong><font color="">CLIENTE</font></strong>
			</td>
			<td class="">
				<strong><font color="">CAMPAÑA</font></strong>
			</td>
			<td class="">
				<strong><font color="">SOLO EMPRESAS</font></strong>
			</td>

			<td class="Estilo38">
				&nbsp
			</td>
			</tr>
			</thead>
			<tr>
			<td>
				<select name="CB_CLIENTE" id="CB_CLIENTE" OnChange="Refrescar();">
					<%
					abrirscg()
					ssql="SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' ORDER BY RAZON_SOCIAL"

					set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
						do until rsTemp.eof%>
						<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCOD_CLIENTE)=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("RAZON_SOCIAL")%></option>
						<%
						rsTemp.movenext
						loop
					end if
					rsTemp.close
					set rsTemp=nothing
					cerrarscg()
					%>
				</select>
			</td>
			<td>
				<select name="CB_CAMPANA" id="CB_CAMPANA" OnChange="Refrescar();">
				<option value="0">NUEVA</option>
				<%
				AbrirSCG()
					If Trim(strCOD_CLIENTE) <> "" Then
						strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
						set rsCampana=Conn.execute(strSql)
						Do While not rsCampana.eof
						If Trim(intCodCampana)=Trim(rsCampana("ID_CAMPANA")) Then strSelCam = "SELECTED" Else strSelCam = ""
						%>
						<option value="<%=rsCampana("ID_CAMPANA")%>" <%=strSelCam%>> <%=rsCampana("ID_CAMPANA") & " - " & rsCampana("NOMBRE")%></option>
						<%
						rsCampana.movenext
						Loop
						rsCampana.close
						set rsCampana=nothing
					End if
				CerrarSCG()
				''Response.End
				%>
				</select>
			</td>
			<td>
				<INPUT TYPE="checkbox" NAME="CH_EMPRESA" <%=strChecked%>>
			</td>
			<td>
				<acronym title="REFRESCAR">
					<input name="BT_REFRESCAR" class="fondo_boton_100" type="button" id="BT_REFRESCAR" onClick="Refrescar();" value="REFRESCAR">
				</acronym>
			 </td>
			<td>

		   <input name="strCOD_CLIENTE" type="hidden" value="<%=strCOD_CLIENTE%>">
		   <input name="hdidCampana" type="hidden" value="<%=intCodCampana%>">
			  <input name="strRUT_DEUDOR" type="hidden" value="<%=strRUT_DEUDOR%>">
			  <input name="ANI" type="hidden" id="ANI" value="<%=ani%>">
			 </td>
			<td>
		</td>
	</tr>
</table>

<%


	strSql="SELECT IsNull(COUNT(DISTINCT RUT_DEUDOR),0) as CRUT, IsNull(SUM(SALDO),0) as CSALDO FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ESTADO_DEUDA = '1' AND SALDO > 0 "
	If Trim(intCodCampana) <> "0" Then
		strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
	End if
	If UCASE(intSoloEmpresa)="ON" Then
		strSql = strSql & " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) > 50000000"
	End if
	'Response.write strSql
	'Response.eND

	AbrirScg()
	set rsDET= Conn.execute(strSql)
	if Not rsDET.eof Then
		intCrut = Trim(rsDET("CRUT"))
		intCsaldo = Trim(rsDET("CSALDO"))
	End if
	CerrarScg()


	strSql="SELECT IsNull(COUNT(DISTINCT RUT_DEUDOR),0) as CRUT, IsNull(SUM(SALDO),0) as CSALDO FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
	If Trim(intCodCampana) <> "0" Then
		strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
	End if
	If UCASE(intSoloEmpresa)="ON" Then
		strSql = strSql & " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) > 50000000"
	End if
	'Response.write strSql
	'Response.eND

	AbrirScg()
	set rsDET= Conn.execute(strSql)
	if Not rsDET.eof Then
		intCrutCampana = Trim(rsDET("CRUT"))
		intCsaldoCampana = Trim(rsDET("CSALDO"))
	End if
	CerrarScg()


	strSql="SELECT IsNull(COUNT(DISTINCT RUT_DEUDOR),0) as CRUT, IsNull(SUM(SALDO),0) as CSALDO FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
	If Trim(intCodCampana) <> "0" Then
		strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
	End if
	If UCASE(intSoloEmpresa)="ON" Then
		strSql = strSql & " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) >= 50000000"
	End if
	strSql = strSql & " AND (USUARIO_ASIG IS NULL OR USUARIO_ASIG = 0)"

	AbrirScg()
	set rsDET1= Conn.execute(strSql)
	'Response.write "<BR>strSql1="&strSql
	'Response.write "<BR>eof="& Not rsDET1.eof
	if Not rsDET1.eof Then
		'Response.write "<BR>DDD1="&rsDET1("CRUT")
		intCrutSA = Trim(rsDET1("CRUT"))
		intCsaldoSA = Trim(rsDET1("CSALDO"))
	End if
	CerrarScg()

	strSql="SELECT IsNull(COUNT(DISTINCT ID_USUARIO),0) as CUSUARIO FROM USUARIO WHERE ACTIVO = 1 AND (PERFIL_COB = 1 OR PERFIL_SUP = 1)"

	AbrirScg()
	set rsDET= Conn.execute(strSql)
	if Not rsDET.eof Then
		intTotalEjecutivo = Trim(rsDET("CUSUARIO"))
	End if
	CerrarScg()

	strSql="SELECT TOP 1 SUM(SALDO) AS MAXSALDO, RUT_DEUDOR FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
	If Trim(intCodCampana) <> "0" Then
		strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
	End if
	If UCASE(intSoloEmpresa)="ON" Then
		strSql = strSql & " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) > 50000000"
	End if
	strSql = strSql & " GROUP BY RUT_DEUDOR ORDER BY SUM(SALDO) DESC"

	'Response.write "<br>strSql===" & strSql

	AbrirScg()
	set rsDET= Conn.execute(strSql)
	if Not rsDET.eof Then
		intMaxSaldo = Trim(rsDET("MAXSALDO"))
	End if
	CerrarScg()

	strSql="SELECT TOP 1 SUM(SALDO) AS MINSALDO, RUT_DEUDOR FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
	If UCASE(intSoloEmpresa)="ON" Then
		strSql = strSql & " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) > 50000000"
	End if
	If Trim(intCodCampana) <> "0" Then
		strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
	End if

	strSql = strSql & " GROUP BY RUT_DEUDOR ORDER BY SUM(SALDO) ASC"

	AbrirScg()
	set rsDET= Conn.execute(strSql)
	if Not rsDET.eof Then
		intMinSaldo = Trim(rsDET("MINSALDO"))
	End if
	CerrarScg()

	strSql="SELECT SUM(SALDO)/COUNT(DISTINCT RUT_DEUDOR) AS PROMSALDO FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
	If UCASE(intSoloEmpresa)="ON" Then
		strSql = strSql & " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) > 50000000"
	End if
	If Trim(intCodCampana) <> "0" Then
		strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
	End if

	AbrirScg()
	set rsDET= Conn.execute(strSql)
	if Not rsDET.eof Then
		intPromSaldo = Trim(rsDET("PROMSALDO"))
	End if
	CerrarScg()




%>

<table width="100%" border="0" ALIGN="CENTER">
	<tr>
		<td valign="top">
			<table width="100%" border="0" ALIGN="CENTER">
				<tr>
					<TD >
						GRABAR CAMPAÑA
					</TD>
					<TD >
						Nombre :
					</TD>
					<TD ALIGN="LEFT">
						<input name="TX_NOMCAMPANA" type="text" value="<%=strNombreCampana%>" size="20" maxlength="20" onchange="">
						<acronym title="ASIGNAR">
						<input name="BT_GC"	class="fondo_boton_100" type="button" onClick="Generar();" value="Gen.Campaña">
						<input name="BT_AC" class="fondo_boton_100" type="button" onClick="AgregarCasos();" value="Agreg.Casos">
						<input name="BT_EC"	class="fondo_boton_100" type="button" onClick="EliminarCasos();" value="Elim.Casos">
						<input name="BT_ECA"class="fondo_boton_100" type="button" onClick="EliminarCampana();" value="Elim.Campaña">
						</acronym>
					</TD>
				</tr>
				<tr>
					<TD >
					<input name="CH_CAMPANA" type="checkbox" value="1" onchange="">
					</TD>
					<TD ALIGN="LEFT">
						Descripción
					</TD>
					<TD ALIGN="LEFT">
					<input name="TX_DESCCAMPANA" type="text" value="<%=strDescCampana%>" size="50" maxlength="120" onchange="">
					</TD>
				</tr>
			</table>
		</TD>
	</tr
</table>

<table width="100%" border="0" bordercolor="#FFFFFF" ALIGN="CENTER">
	<TR>
		<TD class="hdr_i" width="200">
			RUT
		</TD>
		<TD class="hdr_i" width="">
			USUARIO (CODIGO)
		</TD>
	</TR>
	<TR>
		<TD class="hdr_i" width="">
			<TEXTAREA NAME="TA_RUT" ROWS="10" COLS="20"><%=strRut%></TEXTAREA>
		</TD>
		<TD class="hdr_i" width="">
			<TEXTAREA NAME="TA_USUARIO" ROWS="10" COLS="20"><%=strUsuario%></TEXTAREA>
		</TD>
	</TR>
</table>

<table width="100%" border="0" ALIGN="CENTER">
	<tr>
		<td valign="top">
			<table width="100%" border="0" ALIGN="CENTER" class="estilo_columnas">
				<thead>
				<tr >
					<TD WIDTH="20%" ALIGN="CENTER">
						Deuda Maxima
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						Deuda Minima
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						Deuda Promedio
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						Rut Deuda Activa
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						Rut Campaña
					</TD>
				</tr>
				</thead>
				<tr bordercolor="#999999">
					<TD ALIGN="CENTER">
						$&nbsp;<%=FN(intMaxSaldo,0)%>
					</TD>
					<TD ALIGN="CENTER">
						$&nbsp;<%=FN(intMinSaldo,0)%>
					</TD>
					<TD ALIGN="CENTER">
						$&nbsp;<%=FN(intPromSaldo,0)%>
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						<A HREF="cartera_asignada.asp?CB_CAMPANA=<%=intCodCampana%>">
							<%=FN(intCrut,0)%>
						</A>
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						<A HREF="cartera_asignada.asp?CB_CAMPANA=<%=intCodCampana%>">
							<%=FN(intCrutCampana,0)%>
						</A>
					</TD>
				</tr>
				<thead>
				<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					<TD WIDTH="20%" ALIGN="CENTER">
						Rut No Asignados
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						Monto No Asignados
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						Ejecutivos
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						Monto Total Deuda Activa
					</TD>
					<TD WIDTH="20%" ALIGN="CENTER">
						Monto Total Campaña
					</TD>
				</tr>
				</thead>
				<tr bordercolor="#999999">
					<TD ALIGN="CENTER">
						<%=FN(intCrutSA,0)%>
					</TD>
					<TD ALIGN="CENTER">
						$&nbsp;<%=FN(intCsaldoSA,0)%>
					</TD>
					<TD ALIGN="CENTER">
						<%=FN(intTotalEjecutivo,0)%>
					</TD>
					<TD ALIGN="CENTER">
						$&nbsp;<%=FN(intCsaldo,0)%>
					</TD>
					<TD ALIGN="CENTER">
						$&nbsp;<%=FN(intCsaldoCampana,0)%>
					</TD>
				</tr>
			</table>
 		</td>
	</tr>
</table>


<table width="100%" border="0" ALIGN="CENTER">
	<tr>
		<td width="100%" valign="top">
			<TABLE BORDER="0" ALIGN="CENTER">
				<tr bordercolor="#999999">
					<TD><input name="TX_COTA" type="text" value="<%=strCota%>" size="5" maxlength="5" onchange=""></TD>
					<TD><INPUT TYPE=Radio VALUE="1" NAME="OP_SEL"></TD>
					<td>Asignar solo casos sin asignacion</td>
					<TD><INPUT TYPE=Radio VALUE="2" NAME="OP_SEL"></TD>
					<td>Reasignar Todo</td>
					<TD><INPUT TYPE=Radio VALUE="3" NAME="OP_SEL"></TD>
					<TD>Eliminar Asignación</td>
					<TD><input name="BT_ASIGNAR" class="fondo_boton_100" type="button" onClick="Asignar();" value="Procesar"></TD>
				</tr>
			</TABLE>
 		</td>
 	</tr>
 	<tr>
		<td width="100%" valign="top">
			<table border="0" ALIGN="CENTER" class="intercalado" style="width:100%;">
				<thead>
				<TR BORDERCOLOR="#999999">
					<TD WIDTH="10%" align="LEFT">
					Marcar
					</TD>
					<TD WIDTH="45%" align="LEFT">
					Ejecutivo
					</TD>
					<TD WIDTH="20%" align="CENTER">
					Rut Asig
					</TD>
					<TD WIDTH="25%" align="CENTER">
					Monto Asig.
					</TD>
				</TR>
				</thead>
				<tbody>
			<%
				AbrirScg()
				strSql="SELECT ID_USUARIO, rut_usuario, nombres_usuario, isnull(apellido_paterno,'') apellido_paterno, "
				strSql= strSql & " isnull(apellido_materno,'') apellido_materno, fecha_nacimiento, correo_electronico, telefono_contacto, "
				strSql= strSql & " perfil, LOGIN, CLAVE, PERFIL_ADM, perfil_cob, ACTIVO, perfil_proc, perfil_sup, "
				strSql= strSql & " PERFIL_CAJA, perfil_emp, PERFIL_FULL, perfil_back, gestionador_preventivo, "
				strSql= strSql & " anexo, observaciones_usuario "
 				strSql= strSql & " FROM USUARIO WHERE ACTIVO = 1 AND (PERFIL_COB = 1 OR PERFIL_SUP = 1) "
				set rsDET= Conn.execute(strSql)
				IntTotalMonto = 0
				IntTotalRut = 0
				do until rsDET.eof
					strNombre = Trim(rsDET("NOMBRES_USUARIO")) & " " & Trim(rsDET("APELLIDO_PATERNO")) & " " & Trim(rsDET("APELLIDO_MATERNO"))
					strNombre = UCASE(strNombre)

					%>

					<tr bordercolor="#999999">
					<TD><INPUT TYPE=checkbox NAME="CH_USUARIO" value="<%=rsDET("ID_USUARIO")%>"></TD>
					<td><div ALIGN="LEFT"><%=rsDET("ID_USUARIO")%> - <%=strNombre%></div></td>
					<%
					total_ValorCuota = total_ValorCuota + intValorCuota
					total_docs = total_docs + 1

					strSql="SELECT COUNT(DISTINCT RUT_DEUDOR) AS CRUT, IsNull(SUM(SALDO),0) AS CMONTO FROM CUOTA WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
					If UCASE(intSoloEmpresa)="ON" Then
						strSql = strSql & " AND CAST(SUBSTRING(RUT_DEUDOR,1,LEN(RUT_DEUDOR)-2) AS INT) > 50000000"
					End if
					If Trim(intCodCampana) <> "" Then
						strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND ID_CAMPANA = " & intCodCampana & ")"
					End if
					strSql = strSql & " AND USUARIO_ASIG = " & rsDET("ID_USUARIO")

					''Response.write "strSql=" & strSql

					set rsDUsr= Conn.execute(strSql)
					If Not rsDUsr.Eof Then
						intRut = rsDUsr("CRUT")
						intMonto = rsDUsr("CMONTO")
					Else
						intRut = "0"
						intMonto = "0"
					End If

					IntTotalRut = IntTotalRut + intRut
					IntTotalMonto = IntTotalMonto + intMonto

					%>
					<td><div ALIGN="RIGHT"><%=intRut%></div></td>
					<td><div ALIGN="RIGHT">$&nbsp;<%=FN(intMonto,0)%></div></td>
					</tr>
					<%rsDET.movenext
				loop
				CerrarScg()
			%>
				</tbody>
				<thead>
			<tr>
				<td><div ALIGN="RIGHT">&nbsp</div></td>
				<td><div ALIGN="RIGHT">&nbsp</div></td>
				<td><div ALIGN="RIGHT"><%=FN(IntTotalRut,0)%></div></td>
				<td><div ALIGN="RIGHT">$&nbsp;<%=FN(IntTotalMonto,0)%></div></td>
			</tr>
			<thead>
			</table>
 		</td>
	</tr>
</table>

</table>
	</td>
	</tr>
</table>
</form>
</body>
</html>


<script language="JavaScript" type="text/JavaScript">
function envia(){
	datos.action='genera_campanas.asp?Limpiar=1';
	datos.submit();
}

function Asignar(){
	//alert(datos.OP_SEL.value);
	datos.action='genera_campanas.asp?Asignar=1';
	datos.submit();
}

function Generar(){
	if (datos.CB_CAMPANA.value != '0') {
			alert('Debe ingresar Campaña NUEVA');
			datos.CB_CAMPANA.focus();
			return;
	}

	if (datos.TX_NOMCAMPANA.value == '') {
		alert('Debe ingresar Nombre Campaña');
		datos.TX_NOMCAMPANA.focus();
		return;
	}

	if (datos.TX_DESCCAMPANA.value == '') {
		alert('Debe ingresar descripcion Campaña');
		datos.TX_DESCCAMPANA.focus();
		return;
	}

	datos.CH_CAMPANA.value = 1;
	datos.action='genera_campanas.asp?Asignar=1idd';
	datos.submit();
}

function AgregarCasos(){
	if (datos.CB_CAMPANA.value == '0') {
			alert('Debe seleccionar campaña');
			datos.CB_CAMPANA.focus();
			return;
	}

	if (datos.TX_NOMCAMPANA.value == '') {
		alert('Debe ingresar Nombre Campaña');
		datos.TX_NOMCAMPANA.focus();
		return;
	}

	if (datos.TX_DESCCAMPANA.value == '') {
		alert('Debe ingresar descripcion Campaña');
		datos.TX_DESCCAMPANA.focus();
		return;
	}

	datos.action='genera_campanas.asp?AgregarCasos=1';
	datos.submit();
}

function EliminarCampana(){
	if (datos.CB_CAMPANA.value == '0') {
			alert('Debe seleccionar campaña');
			datos.CB_CAMPANA.focus();
			return;
	}

	if (datos.TX_NOMCAMPANA.value == '') {
		alert('Debe ingresar Nombre Campaña');
		datos.TX_NOMCAMPANA.focus();
		return;
	}

	if (datos.TX_DESCCAMPANA.value == '') {
		alert('Debe ingresar descripcion Campaña');
		datos.TX_DESCCAMPANA.focus();
		return;
	}

	datos.action='genera_campanas.asp?EliminarCampana=1';
	datos.submit();
}

function EliminarCasos(){
	if (datos.CB_CAMPANA.value == '0') {
			alert('Debe seleccionar campaña');
			datos.CB_CAMPANA.focus();
			return;
	}

	if (datos.TX_NOMCAMPANA.value == '') {
		alert('Debe seleccionar campaña');
		datos.TX_NOMCAMPANA.focus();
		return;
	}

	if (datos.TX_DESCCAMPANA.value == '') {
		alert('Debe seleccionar campaña');
		datos.TX_DESCCAMPANA.focus();
		return;
	}

	datos.action='genera_campanas.asp?EliminarCasos=1';
	datos.submit();
}

function Refrescar(){
	datos.CH_CAMPANA.value = ''
	datos.action='genera_campanas.asp?Refrescar=1';
	datos.submit();
}

</script>
