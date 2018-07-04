<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
    <link href="../css/style_generales_sistema.css" rel="stylesheet">
    <!--#include file="arch_utils.asp"-->
	<!--#include file="sesion_inicio.asp"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>


	<style type="text/css">
	<!--
	.Estilo36 {
		color: #FF0000;
		font-weight: bold;
		font-size: 18px;
	}
	.Estilo38 {
		color: #008000;
		font-weight: bold;
		font-size: 20px;
	}

	.Estilo37 {font-size: 11px; color: #000066; font-family: tahoma; }
	body {
		background-image: url();
	}
	-->
	</style>

	<script type="text/javascript">
	function Refrescar(strCliente){

		datos.HD_CLIENTE.value = strCliente;

		//datos.action='default.asp?CB_CLIENTE=' + strCliente;
		datos.action='default.asp';
		datos.target='_parent';
		datos.submit();
	}
	</SCRIPT>	
</head>
<body>

<%


	AbrirSCG()
		strSql="SELECT DESCRIPCION, TIPO_CLIENTE, ISNULL(COD_MONEDA,0) AS COD_MONEDA FROM CLIENTE WHERE COD_CLIENTE='" & session("ses_codcli") & "'"
		set rsCOR = Conn.execute(strSql)

		If not rsCOR.eof then
			strNomCliente = Trim(rsCOR("DESCRIPCION"))
			strCodMoneda = Trim(rsCOR("COD_MONEDA"))
			session("tipo_cliente")=rsCOR("TIPO_CLIENTE")
		Else
			strNomCliente = ""
			strCodMoneda = ""
			session("tipo_cliente") = ""
		End if

		If Trim(strCodMoneda) <> 2 Then
			strParamMoneda="N"
			session("valor_moneda") = 1
		Else
			session("valor_moneda") = session("valor_uf")
			strSql="SELECT * FROM MONEDA WHERE COD_MONEDA = " & strCodMoneda
			set rsMon = Conn.execute(strSql)
			If not rsMon.eof then
				session("COD_MONEDA") = Trim(rsMon("COD_MONEDA"))
				session("strSimboloMoneda") = Trim(rsMon("SIMBOLO"))
			Else
				strParamMoneda="N"
			End If
		End If

		strSql="SELECT * FROM PARAMETROS"
		set rsParam = Conn.execute(strSql)
		If not rsParam.eof then
			strNomLogo = Trim(rsParam("NOMBRE_LOGO_TOP_IZQ"))
			strNomSistema = Trim(rsParam("NOMBRE_SISTEMA"))
		End if



%>

<form name="datos" method="post">
<INPUT TYPE='hidden' NAME="HD_CLIENTE" VALUE="">
<%
	SERVIDOR= MID(request.servervariables("PATH_INFO"),2, (Instr(MID(request.servervariables("PATH_INFO"),2, LEN(request.servervariables("PATH_INFO"))),"/"))-1)
	'response.write SERVIDOR
	if ucase(SERVIDOR)="EREC" then
		style ="background: url(../Imagenes/Top_prod.jpg) no-repeat center top; background-size:100% 80px;"

	elseif ucase(SERVIDOR)="EREC_DEMO" then
		style ="background: url(../Imagenes/Top_demo.jpg) no-repeat center top; background-size:100% 80px;"

	elseif ucase(SERVIDOR)="EREC_DESA" then
		style ="background: url(../Imagenes/Top_desa.jpg) no-repeat center top; background-size:100% 80px;"

	end if


%>
<table width="100%" height="80" style="background: url(../Imagenes/Fondo_Top.jpg) no-repeat center top;" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr valign="TOP">
<td align="center" class="Estilo37" height="80" style="<%=style%>">
	<br>
	<br>
	<table width="100%" border="0">
		<tr height="25" >
			<td class="Estilo37" align="RIGHT">
				
				<select name="CB_CLIENTE" id="CB_CLIENTE" onChange="Refrescar(this.value)" class="select_individual">
				<%
				ssql="SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
				set rsTemp= Conn.execute(ssql)
				if not rsTemp.eof then
					do until rsTemp.eof%>
					<option value="<%=rsTemp("COD_CLIENTE")%>" <%if Trim(session("ses_codcli"))=trim(rsTemp("COD_CLIENTE")) then response.Write " Selected " End If%>><%=rsTemp("RAZON_SOCIAL")%></option>
					<%
					rsTemp.movenext
					loop
				end if
				rsTemp.close
				set rsTemp=nothing
				%>
				</select>

			</td>
		</tr>
		<tr height="25" >
			<td align="RIGHT" class="Estilo37">
				U.F.:<%=session("valor_uf")%>&nbsp;&nbsp;&nbsp;
				<%=strNomSistema%>&nbsp;&nbsp;&nbsp;
				Usuario: <%=UCASE(session("nombre_user"))%>&nbsp;&nbsp;&nbsp;
			</td>
		</tr>
	</table>
</td>
</tr>

</table>
</form>



</body>
</html>
	<% CerrarSCG() %>