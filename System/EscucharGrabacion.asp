<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

	<!--#include file="arch_utils.asp"-->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%


If session("session_idusuario")="" then
	Response.write "SESION HA EXPIRADO, CIERRE EL NAVEGADOR Y VUELVA A INGRESAR"
	Response.End
End if


	Response.CodePage=65001
	Response.charset ="utf-8"

Private Sub DownloadFile(file)
	'--declare variables
	Dim strAbsFile
	Dim strFileExtension
	Dim objFSO
	Dim objFile
	Dim objStream
	'-- set absolute file location
	''strAbsFile = Server.MapPath(file)
	strAbsFile = file
	'-- create FSO object to check if file exists and get properties
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	'-- check to see if the file exists
	If objFSO.FileExists(strAbsFile) Then
		Set objFile = objFSO.GetFile(strAbsFile)
		'-- first clear the response, and then set the appropriate headers
		Response.Clear
		'-- the filename you give it will be the one that is shown
		' to the users by default when they save
		Response.AddHeader "Content-Disposition", "attachment; filename=" & objFile.Name
		Response.AddHeader "Content-Length", objFile.Size
		Response.ContentType = "application/octet-stream"
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
		'-- set as binary
		objStream.Type = 1
		Response.CharSet = "UTF-8"
		'-- load into the stream the file
		objStream.LoadFromFile(strAbsFile)
		'-- send the stream in the response
		Response.BinaryWrite(objStream.Read)
		objStream.Close
		Set objStream = Nothing
		Set objFile = Nothing
	Else 'objFSO.FileExists(strAbsFile)
		Response.Clear
		Response.Write("No such file exists.")
	End If
	Set objFSO = Nothing
End Sub






strTelefono = Replace(request("strTelefono"),"-","")

strFecIngreso = request("strFecIngreso")
strHoraIngreso = request("strHoraIngreso")
if len(strHoraIngreso)=4 Then strHoraIngreso = "0" & strHoraIngreso
intIdusuario = request("intIdusuario")
strAnexo = request("strAnexo")

strFecIngreso = Mid(strFecIngreso,7,4) & "-" & Mid(strFecIngreso,4,2) & "-" & Mid(strFecIngreso,1,2)

strIdGrab = Trim(request("strIdGrab"))

AbrirSCG()
strSql="SELECT EsInterno,EsExterno FROM USUARIO WHERE ID_USUARIO =  " & session("session_idusuario")
set rsTemp= Conn.execute(strSql)


if not rsTemp.eof then
	EsInterno = Trim(rsTemp("EsInterno"))
    EsExterno = Trim(rsTemp("EsExterno"))
end if
rsTemp.close
set rsTemp=nothing
CerrarSCG()


if EsExterno  = "Falso"  and  EsInterno  = "Falso" then 
    	%>
		<script>
		    alert('Usted no está catalogado como usuario externo ni interno, favor póngase en contacto con el administrador para que le den los privilegios necesarios');
		    window.close();
		</script>
	<%
    response.End
end if



Dim ConnMySql
Set ConnMySql = Server.CreateObject("ADODB.Connection")
ConnMySql.Open "Driver={MySQL ODBC 5.1 Driver}; Server=192.168.2.20; Database=asteriskcdrdb; UID=serverCRM30; PWD=Llacruzhuelen164"

strSql = "select uniqueid,ROUND((billsec/60),2) duration , dst, DATE_FORMAT(calldate,'%Y%m%d') as calldate, DATE_FORMAT( calldate,  '%Y%m%d %h%i%s' ) AS orderCalldate from asteriskcdrdb.cdr "

if strIdGrab  = "" then

    strSql = strSql & " where src =  '" & strAnexo & "'"
    strSql = strSql & " and (calldate BETWEEN DATE_ADD('" & strFecIngreso & " " & strHoraIngreso & ":00.000', INTERVAL -(1800 + duration) SECOND) AND '" & strFecIngreso & " " & strHoraIngreso & ":59.000')"
    strSql = strSql & " and dst = '" & strTelefono & "' and disposition = 'ANSWERED' order by orderCalldate desc"

else
  strSql = strSql &  "where uniqueid =  '" & strIdGrab & "'" 
end if 

set rsTemp= ConnMySql.execute(strSql)

'AQUI SE ESTA IMPRIMIENDO EL UNIQUEID DE LA ULTIMA GRABACION REALIZADA.
'EL PROBLEMA SURGE CUANDO EL SISTEMA ASIGNA A MAS DE UNA GESTION LA ULTIMA LLAMADA REALIZADA.
'Response.write(rsTemp("uniqueid"))
'Response.End

if not rsTemp.eof then

	strIdGrab = Trim(rsTemp("uniqueid"))
	intDuracion = Trim(rsTemp("duration"))
	strTelefono = Trim(rsTemp("dst"))
	strFecha = Trim(rsTemp("calldate"))

end if

rsTemp.close
set rsTemp=nothing
origen_llave_mapa = request.ServerVariables("SERVER_NAME")
strArchivoMp3=strFecha & "-" & strIdGrab&".mp3"
strRutaMp3 = "\\llacruz.cl\grabaciones\asterisk\"& strArchivoMp3

dim tipoUsuario
tipoUsuario = "INTERNO"

if EsExterno  = "Verdadero" then 
    tipoUsuario = "EXTERNO"
end if

%>
<input type="hidden" id="tipoUsuario" name="tipoUsuario" value="<%=tipoUsuario%>" />
<%

if Trim(strIdGrab) = "" Then
	%>
		<script  type="text/javascript" lenguage="javascript">
		    var tipoUsuario = document.getElementById("tipoUsuario").value;
		    alert('No Existen Grabaciones Asociadas. Usted está configurado como Usuario ' + tipoUsuario);
		    window.close();
		</script>
	<%

Else



if EsExterno  = "Verdadero" then 
    Url_MP3= "http://grabaciones.llacruz.cl/"
else if EsInterno = "Verdadero"  then 
    Url_MP3= "http://192.168.2.16/"  
else 
    Url_MP3= ""
    strIdGrab=""
end if 
end if 


strRutaFinalArchivoMp3 = Url_MP3 + strArchivoMp3


dim fs
set fs=Server.CreateObject("Scripting.FileSystemObject")
dim existe

existe = true 
 
if  not fs.FileExists(strRutaMp3) then
    existe = false
end if

set fs=nothing



if existe = false Then
	%>
		<script type="text/javascript" lenguage="javascript">
		    var tipoUsuario = document.getElementById("tipoUsuario").value;
		    alert('No Se Puede Encontrar El Archivo. Usted está configurado como Usuario ' + tipoUsuario);
		    window.close();
		</script>
	<%
end if 


%>
	<table ALIGN="CENTER" WIDTH="100%" border="0" CLASS="tabla1">
		<tr BGCOLOR="#FFFFFF">
			<td class="hdr_i">Estado:</td>
			<td class="hdr_d">Contestado</td>
		</tr>
		<tr BGCOLOR="#FFFFFF">
			<td class="hdr_i">Telefono:</td>
			<td class="hdr_d"><%=strTelefono%></td>
		</tr>
		<tr BGCOLOR="#FFFFFF">
			<td class="hdr_i">Fecha:</td>
			<td class="hdr_d"><%=strFecha%></td>
		</tr>
		<tr BGCOLOR="#FFFFFF">
			<td class="hdr_i">Duracion:</td>
			<td class="hdr_d"><%=intDuracion%>&nbsp;Min.</td>
		</tr>
		<tr BGCOLOR="#FFFFFF">
			<td class="hdr_i" colspan="2">

            <audio controls style="width:470px" autoplay>
			    <source src="<%=strRutaFinalArchivoMp3%>" type="audio/mpeg">
		    </audio>

			<!--<object id="MediaPlayer"
			type="application/x-oleobject" height="42" standby="Instalando Windows Media Player ..." width="270" align="absMiddle" classid="CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95">
			<param name="FileName" value=
			</param><param name="AutoStart" value="true">
			</param><param name="volume" value="3">
			</param><param name="EnableContextMenu" value="1">
			</param><param name="TransparentAtStart" value="false">
			</param><param name="AnimationatStart" value="false">
			</param><param name="ShowControls" value="true">
			</param><param name="ShowDisplay" value="false">
			</param><param name="ShowStatusBar" value="false">
			</param><param name="autoSize" value="false">
			</param><param name="displaySize" value="true">
			</param>
			</object>-->

			</td>
		</tr>
	</table>
<%End If%>
<br>

<a href="<%=strRutaFinalArchivoMp3%>" target='Contenido'><img src="../imagenes/download_for_windows.png" border="0">&nbsp; Descargar</a>
&nbsp;&nbsp;Presionar botón derecho, guardar destino como.

<%Response.Buffer = True%>

</body>
</html>