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
	
login=Replace(Replace(request("login"),"'",""),"\","")
clave=Replace(Replace(request("clave"),"'",""),"\","")

AbrirSCG()

strSql =" SELECT COD_CLIENTE FROM USUARIO A,  USUARIO_CLIENTE B "
strSql = strSql & " WHERE A.ID_USUARIO = B.ID_USUARIO "
strSql = strSql & " AND LOGIN = '" & login & "' AND CLAVE = '" & clave & "'"

strCOD_CLIENTEs = ""

set rsUSU=Conn.execute(strSql)
If not rsUSU.eof then
	existe=1
	Do While not rsUSU.Eof
		strCOD_CLIENTEs = strCOD_CLIENTEs & "'" & rsUSU("COD_CLIENTE") & "',"
		rsUSU.movenext
	Loop
	strCOD_CLIENTEs = Mid(strCOD_CLIENTEs,1,len(strCOD_CLIENTEs)-1)
Else
	existe=0
End if

rsUSU.close
set rsUSU=nothing



cerrarSCG()
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Login CRM Cobros</title>
<style type="text/css">
<!--
body {
	background-image: url(../Imagenes/texturafondo.jpg);
	background-repeat: repeat-x;
	background-color: #FFFFFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}

.boton { 
	font-family: Tahoma, Helvetica, sans-serif; 
	font-size: 11px; 
	font-weight: bold;   
	background-color: #16428B; 
	color: #FFFFFF;
	border:1px #16428B solid;
	cursor: pointer;
} 

-->
</style>
<script type="text/javascript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
</head>

<body>
<% If trim(strCOD_CLIENTEs) <> "" Then %>
<table width="643" height="373" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" background="../Imagenes/base.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="200">&nbsp;</td>
      </tr>
      <tr>
        <td height="90"><div align="center">
          <table width="250" height="90" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td>

              <FORM name="datos" method="post">
              	<INPUT TYPE="hidden" NAME="login" value="<%=login%>">
              	<INPUT TYPE="hidden" NAME="clave" value="<%=clave%>">
			  		<TABLE height=100 cellSpacing=0 width=300 align=right border=0 valign="top">
			  			<TR>
							<TD ALIGN="CENTER"><span class="Estilo22"> SELECIONE MANDANTE </span></TD>
						</TR>
			             <TR>

			  				<TD ALIGN="CENTER">
			  				<select name="CB_CLIENTE" onChange="">
			  					<%
			  					abrirscg()
			  					ssql="SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE IN (" & strCOD_CLIENTEs & ") ORDER BY RAZON_SOCIAL"
			  					set rsTemp= Conn.execute(ssql)
			  					if not rsTemp.eof then
			  						do until rsTemp.eof%>
			  						<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCliente)=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("RAZON_SOCIAL")%></option>
			  						<%
			  						rsTemp.movenext
			  						loop
			  					end if
			  					rsTemp.close
			  					set rsTemp=nothing
			  					cerrarscg()
			  					%>
			  				</select>
			  				</TD>
			            </TR>
			            <TR>
						  <TD ALIGN="CENTER"><input name="Submit" type="button" class="boton" id="Submit" onClick="envia();" value="Ingresar"></TD>
						</TR>

			            <!--TR>
			              <TD colSpan=2 align="CENTER"> <p><a href="http://<%=strSitioWebEmpresa%>"><%=strSitioWebEmpresa%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			              <%=strDirEmpresa%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Telefonos <%=strTelEmpresa%>
			                </p>
			                </TD>
			            </TR-->
			          </TABLE>

		</FORM>


              </td>
            </tr>
          </table>
        </div></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
<% End If%>
</html>

<script language="JavaScript" type="text/JavaScript">
function envia(){
	if(datos.login.value==''){
		alert('DEBE INGRESAR SU NOMBRE DE USUARIO');
		datos.login.focus();
	}else if(datos.clave.value==''){
		alert('DEBE INGRESAR SU CONTRASEÑA');
		datos.clave.focus();
	}else{
			datos.action='cbdd01.asp';
			datos.submit();
		}
}

function inicio(){
	window.location.href = '';
}
</script>


</body>
<script language="JavaScript">
existe=<%=existe%>;
if(existe=='0'){
alert('USUARIO INVALIDO, CONTRASEÑA INCORRECTA O NO TIENE ASIGNADO MANDANTES');
window.location.href = '../index.asp';
}
</script>
</html>
