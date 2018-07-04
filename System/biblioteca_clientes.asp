<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->


	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->

	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	IntId 			= session("ses_codcli")
	accion_archivo	= request("accion_archivo")
	ruta 			= Request("ruta") 
	AbrirSCG()

	if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(accion_archivo) = "descarga" then

		response.write DownloadFile(ruta)
		'response.write ruta&""

	End if	

	intCorrelativo = 0
	VarPath=  Server.mapPath("../Archivo/BibliotecaClientes") & "\" & IntId


%>

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script type="text/javascript">
		$(document).ready(function(){
			$('.div_historial_eliminados').toggle(function(){

				var IntId 		= $('#IntId').val()				
				var criterios 	="alea="+Math.random()+"&accion_ajax=mostrar_archivos_eliminados_biblioteca_clientes&IntId="+IntId

				$('#div_historial_eliminados').load('FuncionesAjax/mostrar_archivos_eliminados_ajax.asp', criterios, function(){
					$('#img_historial_eliminados').attr('src','../Imagenes/flecha_arriba.png')
				})

			}, function(){

				var criterios 	="alea="+Math.random()+"&accion_ajax=mostrar_archivos_eliminados_vacio"

				$('#div_historial_eliminados').load('FuncionesAjax/mostrar_archivos_eliminados_ajax.asp', criterios, function(){
					$('#img_historial_eliminados').attr('src','../Imagenes/flecha_abajo.png')
				})

			})

			$('.td_hover').hover(function(){
				$(this).css('background-color','#CEE3F6')
			}, function(){
				$(this).css('background-color','')
			})

		})

		function bt_descargar(ruta){
				frmSend.action = "biblioteca_clientes.asp?ruta="+ruta+"&accion_archivo=descarga";
				frmSend.submit();
		}

	</script>
	<style type="text/css">
		form{
			text-align: center;
			width: 100%;			
		}
		.titulo_principal{
			background-color:#380ACD; 
			height: 20px;				
			text-align: center;
		}
		.subir_archivo{
			width: 80%;	
			margin: 0 auto;
		}
		.subir_archivo label{
			margin:20px;
		}
	

		.div_archivos_cargados{
			text-align: center;
			width:100%;
		}
		.div_archivos_cargados div{
			width:90%;
			margin: 0 auto;
		}

		.tabla_archivos_cargados{
  			border-collapse:collapse;
  			border: 2px solild #ccc;
    		padding: 0; 
    		text-align: left;   	
    		width: 70%;
    		min-width: 700px;
    		margin:0 auto;


		}

		.tabla_archivos_cargados th{
			background-color: #A9E2F3; 		
			color: #000;
			font: 14px bold tahoma;			
		}
		.tabla_archivos_cargados td{
			text-align: left; 			
		}
		.contenedor_historial_eliminados{
			text-align: left;
			width: 80%;
			margin: 0 auto;
		}

		.div_historial_eliminados{
			height: 20px;
			text-align: left;
			width: 260px;
		}

		.div_historial_eliminados span{
			height: 20px;
			text-align: center;
			width: 80%;
		}
		#div_historial_eliminados{
			text-align: center;
			width: 100%;		
		}

		#div_historial_eliminados div{
			border:1px solid #16428B;
		}

		.contenedor_tabla{
			width: 100%;
			border:2px solid #ccc;
			text-align: center;

		}

	</style>

</HEAD>


<BODY BGCOLOR='FFFFFF'>
<input type="hidden" name="IntId" 	id="IntId" 	Value="<%=trim(IntId)%>">
<div class="titulo_informe">BIBLIOTECA DOCUMENTOS</div>
<FORM name="frmSend" id="frmSend" method="POST" action="biblioteca_clientes.asp">
	<%	AbrirSCG()
		SQL_SEL ="SELECT id_archivo, nombre_archivo, cod_cliente, rut, convert(varchar(10), FECHA_CARGA, 103) +' '+CONVERT(VARCHAR(5),FECHA_CARGA, 108) FECHA_CARGA, "
		SQL_SEL = SQL_SEL & "ID_USUARIO_CARGA,  " 
		SQL_SEL = SQL_SEL & "isnull(nombres_usuario,'')+' '+isnull(apellido_paterno,'')+' '+isnull(apellido_materno,'') nombre_usuario "
		SQL_SEL = SQL_SEL & "FROM CARGA_ARCHIVOS car " 
		SQL_SEL = SQL_SEL & "INNER JOIN USUARIO usu ON usu.ID_USUARIO=car.id_usuario_carga "
		SQL_SEL = SQL_SEL & "WHERE car.activo =1 AND cod_cliente="&trim(IntId)&" AND origen = 1 "
		SQL_SEL = SQL_SEL & " ORDER BY id_archivo desc"
		set rs_sql_sel = Conn.execute(SQL_SEL)			
	%>
		<div class="div_archivos_cargados">
		<%if not rs_sql_sel.eof then%>	
			
				
			<table class="intercalado" border="1" align="center" style="width:90%;" >
				<thead>
					<tr>
						<th></th>
						<th>Nombre archivo</th>
						<th>Fecha carga</th>
						<th>Usuario carga</th>
					</tr>
				</thead>
				<tbody>
				<%

					do while not rs_sql_sel.eof

					   If ( i Mod 2 )= 1 Then
							bgcolor = "#F0F0F0"
					   Else
							bgcolor = "#FFFFFF"
					   End If
					   i = i + 1

				%>
						<tr class="td_hover" BGCOLOR="<%=bgcolor%>">

						<td>
							&nbsp;&nbsp;
						</td>

						<td>
						
							<a href="#" onclick="bt_descargar('../Archivo/BibliotecaClientes/<%=trim(rs_sql_sel("cod_cliente"))%>/<%=trim(rs_sql_sel("nombre_archivo"))%>')"><%=trim(rs_sql_sel("nombre_archivo"))%></a>

						</td>

						<td align="center"><%=trim(rs_sql_sel("FECHA_CARGA"))%></td>

						<td align="center"><%=trim(rs_sql_sel("nombre_usuario"))%></td>

						</tr>
				<%
					rs_sql_sel.movenext 
					loop
				%>
				</tbody>
			</table>


		<%else%>
			<div class="div_archivos_cargados">
				<br>
				<div class='estilo_columna_individual'>
					Sin archivos cargados
				</div>
				<br>
			</div>
		<%end if%>

			</div>



</FORM>
</BODY>
</HTML>




