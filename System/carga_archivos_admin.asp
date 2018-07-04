<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	
	<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/freeaspupload.asp" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	 
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<%
	Session.CodePage  = 65001

	Response.CodePage=65001
	Response.charset ="utf-8"

	IntId	=session("ses_codcli")
	ruta 	=Request("ruta")
	archivo =Request("archivo")

	    Dim DestinationPath
			DestinationPath = Server.mapPath("../Archivo/CargaArchivosAdmin") & "\" & IntId

	'		Response.write "<br>DestinationPath=" & DestinationPath

			Dim uploadsDirVar
			uploadsDirVar = DestinationPath

			function SaveFiles


				Dim Upload, fileName, fileSize, ks, i, fileKey, resumen
				Set Upload = New FreeASPUpload
				Upload.Save(uploadsDirVar)

				If Err.Number <> 0 then Exit function
				SaveFiles = ""
				ks = Upload.UploadedFiles.keys
				if (UBound(ks) <> -1) then
					resumen = "<B>Archivos subidos:</B> "
					for each fileKey in Upload.UploadedFiles.keys
						resumen = resumen & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
					archivo = Upload.UploadedFiles(fileKey).FileName
					next

				else
				End if

				''Response.End
			End function

		If Request.ServerVariables("REQUEST_METHOD") = "POST" and archivo <> "100" then

			response.write SaveFiles()

			AbrirSCG()
			strSql = "EXEC Proc_Audita_Archivo 1, 4, "&trim(session("session_idusuario"))&",null, '"&trim(IntId)&"', '"&trim(archivo)&"', '',0 "
			'response.write strSql
			Conn.execute(strSql)
			CerrarSCG()	 


		End if

		if Request.ServerVariables("REQUEST_METHOD") = "POST" and archivo = "100" then

			response.write DownloadFile(ruta)
            response.Flush
			response.End
			
		End if		



	If TraeSiNo(session("perfil_full")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then

		strColspan="3"
	Else
		strColspan="2"
	End If


	Set Obj_FSO = createobject("scripting.filesystemobject")


	If not Obj_FSO.FolderExists(Server.mapPath("../Archivo/CargaArchivosAdmin") & "\" & IntId) = True Then ' verifica la existencia del archivo
		Obj_FSO.CreateFolder(Server.mapPath("../Archivo/CargaArchivosAdmin") & "\" & IntId) 
	End if


%>

	<TITLE>Informes especificos</TITLE>
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>

	<style type="text/css">
		form{
			text-align: center
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
		.input_file{
			float:left;
		}	
		.boton_carga { 
			font-family: Tahoma, Helvetica, sans-serif; 
			font-size: 11px; 
			font-weight: bold;   
			background-color: #16428B; 
			color: #FFFFFF;
			border:1px #16428B solid;
			cursor: pointer;
		} 

		.input_boton_carga{
			font-family: Tahoma, Helvetica, sans-serif; 
			font-size: 11px; 
			font-weight: bold;   
			background-color: #16428B; 
			color: #FFFFFF;
			border:1px #16428B solid;
			cursor: pointer;
			float:right;
		}
		.div_archivos_cargados{
			margin: 0 auto;
			margin-top:20px;
			margin-bottom:20px;
			text-align: center;
			width:80%;
		}
		.div_archivos_cargados div{
			border:1px solid #16428B;
			width:80%;
		}

		.tabla_archivos_cargados{
  			border-collapse:collapse;
  			border: 2px solild #ccc;
    		padding: 0; 
    		text-align: left;   	
    		width: 100%;

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
			width: 200px;
		}

		.div_historial_eliminados span{
			height: 30px;
			text-align: center;
			width: 200px;
		}
		#div_historial_eliminados{
			text-align: center;
			width: 100%;		
		}

		#div_historial_eliminados div{
			border:1px solid #16428B;
		}
	</style>

	<script type="text/javascript">
		$(document).ready(function(){
			$('.div_historial_eliminados').toggle(function(){

				var IntId 		= $('#IntId').val()				
				var criterios 	="alea="+Math.random()+"&accion_ajax=mostrar_archivos_eliminados_carga_archivos_admin&IntId="+IntId

				$('#div_historial_eliminados').load('FuncionesAjax/mostrar_archivos_eliminados_ajax.asp', criterios, function(){
					$('#img_historial_eliminados').attr('src','../Imagenes/flecha_arriba.png')
					$('.td_hover').hover(function(){
						$(this).css('background-color','#CEE3F6')
					}, function(){
						$(this).css('background-color','')
					})				
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

		function bt_eliminar(cod_cliente, nombre_archivo, pagina_origen, id_archivo)
		{
			if(confirm("¿Esta seguro que desea eliminar el archivo, posterior a esta acción no podrá recuperarlo?"))
			{
				location.href="EliminarArchivo.asp?IntId="+cod_cliente+"&VarNombreFichero="+nombre_archivo+"&pagina_origen="+pagina_origen+"&id_archivo="+id_archivo
			}
			
		}


		function bt_descargar(ruta){
		    frmSend.action = "carga_archivos_admin.asp?ruta=" + ruta + "&archivo=100";
		    frmSend.submit();
		    
		}


		function enviar(){

			var File1 = $('#File1').val()
			var IntId = $('#IntId').val()

			if(File1=="")
			{
				alert("¡Debe seleccionar archivo!")
				return
			}

			var vec = File1.split("\\");
			var cont = 0
			for(i=0;i<(vec.length);i++)
				{
					cont = cont + 1
				}

			var archivo 	=vec[cont-1]
			var extension 	=archivo.split(".");
			var contEx 		=0			

			for(i=0;i<(extension.length);i++)
			{
				contEx = contEx + 1
			}
				
			var nombre_archivo = extension[contEx-2]
			var extension_archivo = extension[contEx-1]
			//alert(archivo)

			var archivo=archivo.replace(",","");
			var archivo=archivo.replace("Á","");
			var archivo=archivo.replace("É","");
			var archivo=archivo.replace("Í","");
			var archivo=archivo.replace("Ó","");
			var archivo=archivo.replace("Ú","");
			var archivo=archivo.replace("á","");
			var archivo=archivo.replace("é","");
			var archivo=archivo.replace("í","");
			var archivo=archivo.replace("ó","");
			var archivo=archivo.replace("ú","");
			var archivo=archivo.replace("ñ","");
			var archivo=archivo.replace("Ñ","");


			var criterios ="alea="+Math.random()+"&accion_ajax=verifica_carga_archivos_admin&nombre_archivo="+archivo+"&IntId="+IntId
			$('#verifica_archivo').load('FuncionesAjax/verifica_archivo_ajax.asp', criterios, function(){
				
				var archivo_validado	=$('#archivo_validado').val()
				
				if(archivo_validado=="no_existe")
				{

					frmSend.action = "carga_archivos_admin.asp?archivo=1";
					frmSend.submit();
				
				}else{
				
					alert("El archivo que intenta subir al sistema ya existe. Si desea subirlo igualmente, elimine el archivo anterior o cambie el nombre de éste.")
					return
				}

			})

		}

	</SCRIPT>

</HEAD>


<BODY BGCOLOR='FFFFFF'>

	<input type="hidden" id="IntId" 	name="IntId" 	Value="<%=trim(IntId)%>">
	<input type="hidden" id="ruta" 		name="ruta" 	value="">
	<input type="hidden" id="archivo" 	name="archivo" 	value="">

	<div class="titulo_informe">CARGA INFORMES</div>
	<FORM name="frmSend" id="frmSend" onSubmit="return enviar(this)"  method="POST" enctype="multipart/form-data" accept-charset="utf-8" action="carga_archivos_admin.asp">

	<div class="subir_archivo">			
		<input class="input_file" name="File1" id="File1" type="file" VALUE="<%=File1%>" size="40" maxlength="40">
		<input Name="SubmitButton" Value="Cargar" class="fondo_boton_100" Type="BUTTON" onClick="enviar();">
	</div>

	<%	AbrirSCG()
		SQL_SEL ="SELECT TOP 20 id_archivo, nombre_archivo, cod_cliente, rut, convert(varchar(10), FECHA_CARGA, 103) +' '+CONVERT(VARCHAR(5),FECHA_CARGA, 108) FECHA_CARGA, "
		SQL_SEL = SQL_SEL & "ID_USUARIO_CARGA,  " 
		SQL_SEL = SQL_SEL & "isnull(nombres_usuario,'')+' '+isnull(apellido_paterno,'')+' '+isnull(apellido_materno,'') AS nombre_usuario "
		SQL_SEL = SQL_SEL & "FROM CARGA_ARCHIVOS car " 
		SQL_SEL = SQL_SEL & "INNER JOIN USUARIO usu ON usu.ID_USUARIO=car.id_usuario_carga "
		SQL_SEL = SQL_SEL & "WHERE car.activo =1 AND cod_cliente="&trim(IntId)&" AND origen = 4 "
		SQL_SEL = SQL_SEL & " ORDER BY id_archivo desc"
		set rs_sql_sel = Conn.execute(SQL_SEL)			
	%>

		<%if not rs_sql_sel.eof then%>	
			<div class="div_archivos_cargados">


			<table class="intercalado" style="width:100%;">
				<thead>
					<tr>
						<th></th>
						<th>Cont.</th>
						<th>Nombre archivo</th>
						<th>Fecha carga</th>
						<th>Usuario carga</th>
					</tr>
				</thead>
				<tbody>
				<%
				
				intCont =0
				
					do while not rs_sql_sel.eof
					
					intCont = intCont + 1
					
					   If ( i Mod 2 )= 1 Then
							bgcolor = "#F0F0F0"
					   Else
							bgcolor = "#FFFFFF"
					   End If
					   i = i + 1

				%>
						<tr class="td_hover" BGCOLOR="<%=bgcolor%>">
						
						<td>
							<a onclick="bt_eliminar('<%=trim(rs_sql_sel("cod_cliente"))%>','<%=trim(rs_sql_sel("nombre_archivo"))%>','CargaArchivos','<%=trim(rs_sql_sel("id_archivo"))%>')" href="#"><img border="0" src="../imagenes/icon_cruz_roja.jpg"></a>
						</td>

						<td align="center"><%=intCont%></td>
						
						<td>
						
							<a href="#" onclick="bt_descargar('../Archivo/CargaArchivosAdmin/<%=trim(rs_sql_sel("cod_cliente"))%>/<%=trim(rs_sql_sel("nombre_archivo"))%>')"><%=trim(rs_sql_sel("nombre_archivo"))%></a>

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

			</div>


		<%else%>
			<div class="div_archivos_cargados">
				<div>
					<br>
					<label style='font: 14px bold #000;'>Sin archivos cargados</label>
					<br>
					<br>
				</div>
			</div>
		<%end if%>

	<%
		SQL_SEL ="SELECT count(*) cantidad "
		SQL_SEL = SQL_SEL & "FROM CARGA_ARCHIVOS " 
		SQL_SEL = SQL_SEL & "WHERE activo =0 AND cod_cliente="&trim(IntId)&" AND origen = 4 "
		set rs_sql_sel = Conn.execute(SQL_SEL)
	%>
		<div class="contenedor_historial_eliminados ">
		
			<div class="div_historial_eliminados fondo_boton_130"><span >&nbsp;Historial archivos eliminados (<%=trim(rs_sql_sel("cantidad"))%>)</span> <img  id="img_historial_eliminados" height="13" width="13" src="../Imagenes/flecha_abajo.png"></div>
			
			<div id="div_historial_eliminados" >
	
			</div>

		</div>	

		<div id="verifica_archivo"></div>
	</FORM>

</BODY>
</HTML>




