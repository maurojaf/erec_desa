<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../lib/lib.asp"-->

<html>
<head>
<title>EMPRESA S.A.</title>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
</head>

<body>
<%
Response.CodePage=65001
Response.charset ="utf-8"


	
login=request("login")
clave=request("clave")
hora_ingreso=time
existe=0
avisa_of=0
AbrirSCG()

session("COLTABBG") = TraeCampoId(Conn, "COLOR_TABLA_BG", 1, "PARAMETRO_SISTEMA", "ID")
session("COLTABBG2") = TraeCampoId(Conn, "COLOR_TABLA_BG_2", 1, "PARAMETRO_SISTEMA", "ID")

session("ses_PathFisicoSistema") = "D:\app\crm_ErecProd"

''Response.write "col2=" & session("COLTABBG2")


strSql="SELECT VALOR FROM UNIDAD_FOMENTO WHERE CONVERT(VARCHAR(10),FECHA,103)=CONVERT(VARCHAR(10),GETDATE(),103)"
set rsUF=Conn.execute(strSql)
if not rsUF.eof then
	session("valor_uf") = rsUF("VALOR")
	session("valor_moneda") = rsUF("VALOR")
Else
	session("valor_uf") = 22000
	session("valor_moneda") = 22000
End If

'Response.Write "valor_moneda=" & session("valor_moneda")
'Response.End

strSql="SELECT IsNull(PERMITE_NO_VALIDAR_FONOS,'N') as PERMITE_NO_VALIDAR_FONOS FROM PARAMETROS"
set rsParam=Conn.execute(strSql)
if not rsParam.eof then
	session("permite_no_validar_fonos") = rsParam("PERMITE_NO_VALIDAR_FONOS")
Else
	session("permite_no_validar_fonos") = "S"
End If

ssql="SELECT *, CONVERT(VARCHAR(8),GETDATE(),108) AS HH , CONVERT(VARCHAR(10),GETDATE(),103) AS FH FROM USUARIO WHERE LOGIN='" & login & "' AND CLAVE = '" & clave & "' AND ID_USUARIO IN (SELECT ID_USUARIO FROM USUARIO_CLIENTE WHERE COD_CLIENTE = '" & request("CB_CLIENTE") & "')"

set rsUSU=Conn.execute(ssql)
	if not rsUSU.eof then
		existe=1
		ssqlok="SELECT ACTIVO FROM USUARIO WHERE LOGIN='" & login & "' AND CLAVE= '" & clave & "'"
		set rsOK=Conn.execute(ssqlok)
		'reSPONSE.WRITE "ssqlok=" & ssqlok
		'rESPONSE.eND
		ok=TraeSiNo(Trim(rsOK("ACTIVO")))
		rsOK.close
		set rsOK=nothing
		if ok="Si" then
			avisa_of=0
			session("ses_clave")=clave
			session("ses_codcli")=request("CB_CLIENTE")
			session("session_idusuario")=rsUSU("ID_USUARIO")
			session("session_user")=rsUSU("RUT_USUARIO")
			session("session_login")=rsUSU("LOGIN")
			session("session_tipo")=rsUSU("PERFIL")
			session("perfil_adm")=rsUSU("PERFIL_ADM")
			session("perfil_caja")=rsUSU("PERFIL_CAJA")
			session("perfil_emp")=rsUSU("PERFIL_EMP")
			session("perfil_sup")=rsUSU("PERFIL_SUP")
			session("perfil_full")=rsUSU("PERFIL_FULL")
			session("perfil_cob")=rsUSU("PERFIL_COB")
			session("nombre_user")=TRIM(rsUSU("NOMBRES_USUARIO")) & " " & TRIM(rsUSU("APELLIDOS_USUARIO"))
			session("iniciosesion")=rsUSU("FH") & " - " & rsUSU("HH")

			strSql = "SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario")
			set rsClientes=Conn.execute(strSql)
			strClientes=""
			Do While Not rsClientes.eof
				strClientes = strClientes & rsClientes("COD_CLIENTE") & ","
				rsClientes.movenext
			Loop

			strClientes = Mid(strClientes,1,len(strClientes)-1)
			session("strCliUsuarios") = strClientes

			strSql="SELECT NOMBRE_CONV_PAGARE,COD_MONEDA FROM CLIENTE WHERE COD_CLIENTE = '" & request("CB_CLIENTE") & "'"
			set rsCliente=Conn.execute(strSql)
			if not rsCliente.eof then
				'Response.write "NOMBRE_CONV_PAGARE=" & strSql
				session("NOMBRE_CONV_PAGARE") = rsCliente("NOMBRE_CONV_PAGARE")
				session("COD_MONEDA") = rsCliente("COD_MONEDA")
			Else
				session("NOMBRE_CONV_PAGARE") = "CONVENIO"
				session("COD_MONEDA") = 1
			End If
		else
			avisa_of=1
		end if
	else
		existe=0
	end if

rsUSU.close
set rsUSU=nothing

'Response.write "NOMBRE_CONV_PAGARE=" & session("NOMBRE_CONV_PAGARE")
'Response.End

	if avisa_of=0 and existe=1 then


		strSql = "INSERT INTO LOG_CRMCOBROS (ID_USUARIO, LOGIN, FECHA, IP, IP_HOST, IP_LOCAL, IP_CLIENTE)"
		strSql = strSql & " Values (" & session("session_idusuario") & ",'" & session("session_login") & "',getdate(),'" & Mid(request.servervariables("REMOTE_ADDR"),1,19) & "','" & Mid(request.servervariables("REMOTE_HOST"),1,19) & "','" & Mid(request.servervariables("LOCAL_ADDR"),1,19) & "','" & Mid(request.servervariables("HTTP_CLIENT_IP"),1,19) & "')"
		set rsInserta=Conn.execute(strSql)

		response.Redirect("default.asp")




	end if


cerrarSCG()
%>

</body>
<script language="JavaScript">
existe=<%=existe%>;
avisa_of=<%=avisa_of%>;
if(existe=='0'){
alert('USUARIO NO VALIDO');
window.location.href = "../index.asp";
}

if(avisa_of=='1'){
alert('SU CUENTA DE ACCESO HA SIDO DESACTIVADA');
window.location.href = "../index.asp";
}

</script>
</html>
