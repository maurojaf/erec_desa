<%

Function TraeCantGestiones (strConex, dtmInicio, dtmFin, intCat, intSubCat, intGestion, intCliente)
	strSql = "select count(*) as cant from gestiones where cod_cliente = '" & intCliente & "' and fecha_ingreso >= '" & dtmInicio & "' and fecha_ingreso <= '" & dtmFin & "' and cod_categoria = " & intCat & " and cod_sub_categoria = " & intSubCat & " and cod_gestion = " & intGestion
	
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeCantGestiones = rsTmp("cant")
	else
		TraeCantGestiones = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function

Function TraeCantGestionesUG2 (strConex, dtmInicio, dtmFin, intCat, intSubCat, intGestion, intCliente)
	strSql = "SELECT COUNT(*) AS CANT FROM DEUDOR WHERE cod_cliente = '" & intCliente & "' AND ULTIMA_GESTION = '" & intCat & "-" & intSubCat & "-" & intGestion & "'"
	
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeCantGestionesUG2 = rsTmp("cant")
	else
		TraeCantGestionesUG2 = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function


Function TraeSaldoGestiones (strConex, dtmInicio, dtmFin, intCat, intSubCat, intGestion, intCliente)
	strSql = "select IsNull(sum(saldo),0) as saldo from cuota where cod_cliente = '" & Trim(intCliente) & "' and saldo > 0 and estado_deuda = '1' "
	strSql = strSql & " and RUT_DEUDOR in (select distinct RUT_DEUDOR from gestiones where cod_cliente = '" & Trim(intCliente) 
	strSql = strSql & "' and fecha_ingreso >= '" & Trim(dtmInicio) & "' and fecha_ingreso <= '" & Trim(dtmFin) 
	strSql = strSql & "' and cod_categoria = " & Trim(intCat) 
	strSql = strSql & " and cod_sub_categoria = " & Trim(intSubCat) 
	strSql = strSql & " and cod_gestion = " & Trim(intGestion) & ")"
	
	'Response.write "strSql=" & strSql
	
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeSaldoGestiones = rsTmp("saldo")
	else
		TraeSaldoGestiones = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function

Function TraeCantGestionesUsuario (strConex, dtmInicio, dtmFin, intCat, intSubCat, intGestion, intCliente, intUsuario)
	strSql = "select count(*) as cant from gestiones where cod_cliente = '" & Trim(intCliente) & "' and fecha_ingreso >= '" & dtmInicio & "' and fecha_ingreso <= '" & dtmFin & "' and cod_categoria = " & intCat & " and cod_sub_categoria = " & intSubCat & " and cod_gestion = " & intGestion & " and id_usuario = " & intUsuario
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeCantGestionesUsuario = rsTmp("cant")
	else
		TraeCantGestionesUsuario = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function

Function TraeSaldoGestionesUsuario (strConex, dtmInicio, dtmFin, intCat, intSubCat, intGestion, intCliente, intUsuario)
	strSql = "select IsNull(sum(saldo),0) as saldo from cuota where cod_cliente = '" & intCliente & "' and saldo > 0 and estado_deuda = '1' and RUT_DEUDOR in (select distinct RUT_DEUDOR from gestiones where cod_cliente = '" & intCliente & "' and fecha_ingreso >= '" & dtmInicio & "' and fecha_ingreso <= '" & dtmFin & "' and cod_categoria = " & intCat & " and cod_sub_categoria = " & intSubCat & " and cod_gestion = " & intGestion & " and id_usuario = " & intUsuario & ")"
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeSaldoGestionesUsuario = rsTmp("saldo")
	else
		TraeSaldoGestionesUsuario = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function

Function TraeSaldoGestionesUsuarioCampana (strConex, intCat, intSubCat, intGestion, intCliente, intUsuario, intCampana, dtmInicio, dtmFin)
	strSql = "SELECT ISNULL(SUM(SALDO),0) AS SALDO FROM CUOTA WHERE cod_cliente = '" & intCliente & "' AND SALDO > 0 AND ESTADO_DEUDA IN ('1','6') AND RUT_DEUDOR IN (SELECT DISTINCT RUT_DEUDOR FROM GESTIONES WHERE cod_cliente = '" & intCliente & "' AND IDCAMPANA = " & intCampana & " AND cod_categoria = " & intCat & " AND cod_sub_categoria = " & intSubCat & " AND cod_gestion = " & intGestion & " AND id_usuario = " & intUsuario & " AND fecha_ingreso >= '" & dtmInicio & "' AND fecha_ingreso <= '" & dtmFin & "')"
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeSaldoGestionesUsuarioCampana = rsTmp("saldo")
	else
		TraeSaldoGestionesUsuarioCampana = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function


Function TraeCantGestionesJud (strConex, dtmInicio, dtmFin, intCat, intSubCat, intGestion, intCliente)
	strSql = "select count(*) as cant from gestiones_nuevas_judicial where cod_cliente = '" & intCliente & "' and fecha_ingreso >= '" & dtmInicio & " 00:00:00' and fecha_ingreso <= '" & dtmFin & " 23:59:59' and cod_categoria = " & intCat & " and cod_sub_categoria = " & intSubCat & " and cod_gestion = " & intGestion
	
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeCantGestionesJud = rsTmp("cant")
	else
		TraeCantGestionesJud = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function

Function TraeCantGestionesCampana(strConex, intCampana, intCat, intSubCat, intGestion, intCliente, dtmFecIni, dtmFecFin)
	strSql = "select count(*) as cant from gestiones where cod_cliente = '" & intCliente & "' and idcampana = " & intCampana & " and cod_categoria = " & intCat & " and cod_sub_categoria = " & intSubCat & " and cod_gestion = " & intGestion & " AND fecha_ingreso >= '" & dtmFecIni & " 00:00:00' AND fecha_ingreso <= '" & dtmFecFin & " 23:59:59'"
	'Response.write "<BR>"&strSql
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeCantGestionesCampana = rsTmp("cant")
	else
		TraeCantGestionesCampana = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function

Function TraeCantGestionesUG(strConex, strGestion, intCliente, intCampana)
	strSql = "SELECT COUNT(*) AS CANT FROM DEUDOR WHERE cod_cliente = '" & intCliente & "' AND ULTIMA_GESTION = '" & strGestion & "'"
	If Trim(intCampana) <> "0" Then
		strSql = strSql & " AND IDCAMPANA = " & intCampana
	End if
	'If Trim(intUsuario) <> "" Then
	'	strSql = strSql & " AND COD_EJECUTIVO = " & SESSION("session_cod_ejecutivo") 
	'End if
	strSql = strSql & " and RUT_DEUDOR in (select RUT_DEUDOR from gestiones where cod_cliente = '" & intCliente & "' and idcampana = " & intCampana & " and cod_categoria + '-' +  cod_sub_categoria + '-' + cod_gestion = " & intGestion & ")"
	'Response.write "<BR>"&strSql
	'Response.eND
	
	
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeCantGestionesUG = rsTmp("cant")
	else
		TraeCantGestionesUG = 0
	End if
	rsTmp.close
	set rsTmp=nothing
End Function

Function TraeCantGestionesCampanaUG(strConex, strGestion, intCliente, intCampana, dtmInicio, dtmFin)
	strSql = "SELECT COUNT(*) AS CANT FROM DEUDOR WHERE cod_cliente = '" & intCliente & "' AND ULTIMA_GESTION = '" & strGestion & "'"
	If Trim(intCampana) <> "0" Then
		strSql = strSql & " AND IDCAMPANA = " & intCampana
	End if
	'If Trim(intUsuario) <> "" Then
	'	strSql = strSql & " AND COD_EJECUTIVO = " & SESSION("session_cod_ejecutivo") 
	'End if
	strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM GESTIONES WHERE cod_cliente = '" & intCliente & "' AND IDCAMPANA = " & intCampana & " AND cast(cod_categoria as varchar(2)) + '-' + cast(cod_sub_categoria  as varchar(2))+ '-' + cast(cod_gestion  as varchar(2)) = '" & strGestion & "' and fecha_ingreso >= '" & dtmInicio & " 00:00:00' and fecha_ingreso <= '" & dtmFin & " 23:59:59')"
	'Response.write "<BR>"&strSql
	'Response.eND
	
	
	set rsTmp = strConex.execute(strSql)
	if not rsTmp.eof then
		TraeCantGestionesCampanaUG = rsTmp("cant")
	else
		TraeCantGestionesCampanaUG = 0
	End if
	'Response.write "--TraeCantGestionesCampanaUG="&TraeCantGestionesCampanaUG
	rsTmp.close
	set rsTmp=nothing
End Function


%>