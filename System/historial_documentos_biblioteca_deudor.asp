<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/lib.asp"-->

	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<%
	Response.CodePage 	=65001
	Response.charset 	="utf-8"


	ID_CUOTA 			= request("ID_CUOTA")
    ruta 				= Request("ruta")
    accion_archivo		= Request("accion_archivo")
	

	if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(accion_archivo) = "descarga" then

	'response.write ruta
	'response.end
	
		response.write DownloadFile(ruta)
		response.write ruta&""

	End if	

%>

	<title>DETALLE HISTORICO</title>


	<script type="text/javascript">
	$(document).ready(function(){
		
		$('.td_hover').hover(function(){
			$(this).css('background-color','#CEE3F6')
		}, function(){
			$(this).css('background-color','')
		})

		$(document).tooltip();	
	})
	</script>

<%
AbrirSCG()

%>
</head>
<body>
<FORM name="frmSend" id="frmSend" accept-charset="utf-8" method="POST"  action="historial_documentos_biblioteca_deudor.asp">

<div class="titulo_informe" >Detalle Historico</div>
<br>

<%
	SQL_SEL	="SELECT convert(varchar(10), car.FECHA_CARGA, 103) +' '+CONVERT(VARCHAR(5),car.FECHA_CARGA, 108) FECHA_CARGA, "
	SQL_SEL	= SQL_SEL & " isnull(usu.nombres_usuario,'')+' '+isnull(usu.apellido_paterno,'')+' '+isnull(usu.apellido_materno,'') nombre_usuario, login, NOMBRE_ARCHIVO, car.ID_ARCHIVO, OBSERVACION_CARGA_ARCHIVO, "
	SQL_SEL	= SQL_SEL & " CASE "
	SQL_SEL	= SQL_SEL & " WHEN arch.ACTIVO=1 THEN 'CARGADO' "
	SQL_SEL	= SQL_SEL & " WHEN arch.ACTIVO=0 THEN 'ELIMINADO' END ESTADO, C.COD_CLIENTE, car.RUT "
	SQL_SEL	= SQL_SEL & " FROM CARGA_ARCHIVOS_CUOTA car "
	SQL_SEL	= SQL_SEL & " INNER JOIN USUARIO usu ON usu.ID_USUARIO=car.USUARIO_CARGA "
	SQL_SEL	= SQL_SEL & " inner join CARGA_ARCHIVOS arch ON arch.ID_ARCHIVO=car.ID_ARCHIVO  "
	SQL_SEL	= SQL_SEL & " INNER JOIN CUOTA C ON CAR.ID_CUOTA = C.ID_CUOTA "
	SQL_SEL	= SQL_SEL & " WHERE CAR.ID_CUOTA ="&trim(ID_CUOTA)'&" AND car.activo=1 "
	SQL_SEL	= SQL_SEL & " ORDER BY arch.ID_ARCHIVO DESC "
	'response.write SQL_SEL
	
	set rs_sel = conn.execute(SQL_SEL)
	if err then
		response.write SQL_SEL & " / ERROR : " & err.description
		response.end()
	end if
%>



  <table class="intercalado" align="center">
  	<thead>
    <tr>
      <td width="110px">ID COMPROBANTE</td>
      <td width="200px">NOMBRE COMPROBANTE</td>
      <td width="120px">FECHA CARGA</td>
      <td width="120px">USUARIO CARGA</td>
      <td width="300px">OBSERVACIONES</td>	
      <td width="100px">ESTADO</td>	         
    </tr>
   	</thead>
   	<tbody>
	<%if not rs_sel.eof then%>
		<%do while not rs_sel.eof
		   If ( i Mod 2 )= 1 Then
				bgcolor = "#F0F0F0"
		   Else
				bgcolor = "#FFFFFF"
		   End If
		   i = i + 1					

			if trim(rs_sel("OBSERVACION_CARGA_ARCHIVO"))="" then
          		OBSERVACIONES = "Sin información adicional"
          	else
          		OBSERVACIONES = Mid(rs_sel("OBSERVACION_CARGA_ARCHIVO"),1,40)
          	end if

          	if trim(rs_sel("login"))="" then
          		nombre_usuario ="Sin información"
          	else
          		nombre_usuario = Mid(rs_sel("login"),1,20)
          	end if

		%>
        <tr class="td_hover" BGCOLOR="<%=bgcolor%>">
          <td height="18"><%=trim(rs_sel("ID_ARCHIVO"))%></td>
          <td height="18" title="<%=rs_sel("NOMBRE_ARCHIVO")%>">
          	
          	<%if trim(rs_sel("ESTADO"))="CARGADO" then%>	
				<a href="#" onclick="bt_descargar('../Archivo/BibliotecaDeudores/<%=trim(rs_sel("cod_cliente"))%>/<%=trim(rs_sel("rut"))%>/<%=trim(rs_sel("nombre_archivo"))%>')"><%=trim(rs_sel("NOMBRE_ARCHIVO"))%></a>

			<%else%>
				<%=trim(Mid(rs_sel("NOMBRE_ARCHIVO"),1,30))%>
			<%end if%>


          </td>
          <td height="18"><%=trim(rs_sel("FECHA_CARGA"))%></td>
          <td height="18" title="<%=trim(rs_sel("nombre_usuario"))%>"><%=trim(ucase(nombre_usuario))%></td>	 
          <td height="18" title="<%=trim(OBSERVACIONES)%>"><%=trim(OBSERVACIONES)%></td>
          <td height="18"><%=trim(rs_sel("ESTADO"))%></td>		 
                  
        </tr>
        <%rs_sel.movenext
        loop%>
	<%else%>

		<tr>
          <td colspan="6">Sin documentos asociados</td>
        </tr>			
    <%end if%>
	</tbody>
  </table>

<%
cerrarSCG()
%>
</form>
</body>
</html>

<script type="text/javascript">
function bt_descargar(ruta){

	frmSend.action = "historial_documentos_biblioteca_deudor.asp?ruta="+ruta+"&accion_archivo=descarga&ID_CUOTA=<%=trim(ID_CUOTA)%>";
	frmSend.submit();

}
</script>
