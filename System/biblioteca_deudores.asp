<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

	<!--#include file="../lib/asp/comunes/general/rutinasTraeCampo.inc"-->

	<!--#include file="arch_utils.asp"-->

	<!--#include file="../lib/freeaspupload.asp" -->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link rel="stylesheet" href="../css/style_generales_sistema.css">	
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	IntId 			= session("ses_codcli")
   	strRut 			= Request("strRut")
    archivo 		= Request("archivo")
    ruta 			= Request("ruta")
    accion_archivo	= Request("accion_archivo")
	strObservaciones=Mid(Replace(request("TX_OBSERVACIONES"),";"," "),1,599)
	
	'Response.write "strObservaciones=" & strObservaciones
	
    'Response.write archivo

   AbrirSCG()

	Dim DestinationPath
	DestinationPath = Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId  & "\" & strRut

	' crear una instancia
	set Obj_FSO = createobject("scripting.filesystemobject")

	If not Obj_FSO.FolderExists(Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId) = True Then ' verifica la existencia del archivo
		Obj_FSO.CreateFolder(Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId) 

		If not Obj_FSO.FolderExists(Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId  & "\" & strRut) = True Then 
			Obj_FSO.CreateFolder(Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId  & "\" & strRut) 
		End if	
	else
		If not Obj_FSO.FolderExists(Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId  & "\" & strRut) = True Then 
			Obj_FSO.CreateFolder(Server.mapPath("../Archivo/BibliotecaDeudores") & "\" & IntId  & "\" & strRut) 
		End if	
	End if

	If Trim(accion_archivo) = "carga" Then


			'Response.write "<br>DestinationPath=" & DestinationPath

			Dim uploadsDirVar
			uploadsDirVar = DestinationPath

			function SaveFiles


				Dim Upload, fileName, fileSize, ks, i, fileKey, resumen
				Set Upload = New FreeASPUpload
				Upload.Save(uploadsDirVar)
				If Err.Number <> 0 then Exit function
				SaveFiles = ""
				ks = Upload.UploadedFiles.keys
				If (UBound(ks) <> -1) Then
					resumen = "<B>Archivos subidos:</B> "
					for each fileKey in Upload.UploadedFiles.keys

						'resumen = resumen & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "

						archivo = replace(Upload.UploadedFiles(fileKey).FileName,"á","a")
						
						'response.write replace(Upload.UploadedFiles(fileKey).FileName,"�","")

						strSql = "UPDATE DEUDOR SET FEC_SUBIDA_ULT_ARCHIVO = getdate() WHERE COD_CLIENTE = '" & session("ses_codcli") & "' AND RUT_DEUDOR = '" & strRut & "'"
						Conn.execute(strSql)

						strSql = "EXEC Proc_Audita_Archivo 1, 2, "&trim(session("session_idusuario"))&",'"&trim(strRut)&"', '"&triM(session("ses_codcli"))&"', '"&trim(archivo)&"', '',0 "
						'response.write strSql
						Conn.execute(strSql)

					next

				Else
				End if


			End function

			response.write SaveFiles()
	
		'--Inserta los registros en la Tabla Carga_Archivos_Cuota--'
		

			'----'
		
	End if


	if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(accion_archivo) = "descarga" then

		response.write DownloadFile(ruta)
		response.write ruta&""

	End if	


	AbrirSCG1()
	strSql="SELECT FORMULA_HONORARIOS, FORMULA_INTERESES, RAZON_SOCIAL, USA_SUBCLIENTE, USA_INTERESES, USA_HONORARIOS, USA_PROTESTOS, ISNULL(RETIRO_SABADO,0) AS RETIRO_SABADO, [dbo].[fun_ubicabilidad_telefono] ('" & strRut & "') as UBIC_FONO, [dbo].[fun_ubicabilidad_email] ('" & strRut & "') as UBIC_EMAIL, [dbo].[fun_ubicabilidad_direccion] ('" & strRut & "') as UBIC_DIRECCION  FROM CLIENTE WHERE COD_CLIENTE='" & IntId & "'"

	'Response.write "strSql=" & strSql
	set rsCLI=Conn1.execute(strSql)
	if not rsCLI.eof then
		strNomFormHon = ValNulo(rsCLI("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsCLI("FORMULA_INTERESES"),"C")

		strUsaSubCliente = rsCLI("USA_SUBCLIENTE")
		strUsaInteres = rsCLI("USA_INTERESES")
		strUsaHonorarios = rsCLI("USA_HONORARIOS")
		strUsaProtestos = rsCLI("USA_PROTESTOS")


		nombre_cliente=rsCLI("RAZON_SOCIAL")
		intRetiroSabado=Cint(rsCLI("RETIRO_SABADO"))
		strMsjRetiroSabado = ""
		If Trim(intRetiroSabado) = "1" Then
			strMsjRetiroSabado = "sabados,"
		End if

		strUbicFono =rsCLI("UBIC_FONO")
		strUbicEmail =rsCLI("UBIC_EMAIL")
		strUbicDireccion =rsCLI("UBIC_DIRECCION")
	end if
	rsCLI.close
	set rsCLI=nothing
	CerrarSCG1()
%>


	<TITLE>Biblioteca de Deudores</TITLE>
	<script src="../Componentes/PrettyNumber/jquery.prettynumber.js"></script>

	<script type="text/javascript">
			var format = function(cnt, cents) {
            cnt = cnt.toString().replace(/\$|\u20AC|\,/g,'');
            if (isNaN(cnt))
                return 0;    
            var sgn = (cnt == (cnt = Math.abs(cnt)));
            cnt = Math.floor(cnt * 100 + 0.5);
            cvs = cnt % 100;
            cnt = Math.floor(cnt / 100).toString();
            if (cvs < 10)
            cvs = '0' + cvs;
            for (var i = 0; i < Math.floor((cnt.length - (1 + i)) / 3); i++)
                cnt = cnt.substring(0, cnt.length - (4 * i + 3)) + ',' 
                                + cnt.substring(cnt.length - (4 * i + 3));
 
            return (((sgn) ? '' : '-') + cnt) + ( cents ?  '.' + cvs : '');
        };

		$(document).ready(function(){
			$(document).tooltip();

			$('.div_historial_eliminados').toggle(function(){

				var strRut 		= $('#strRut').val()
				var IntId 		= $('#IntId').val()				
				var criterios 	="alea="+Math.random()+"&accion_ajax=mostrar_archivos_eliminados_bibioteca_deudores&strRut="+strRut+"&IntId="+IntId

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
		
		function enviar(){

			var File1 	= $('#File1').val()
			var strRut 	= $('#strRut').val()
			var IntId 	= $('#IntId').val()
			 
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

			archivo 	=vec[cont-1]
			extension 	=archivo.split(".");
			contEx 		=0			

			for(i=0;i<(extension.length);i++)
			{
				contEx = contEx + 1
			}

				
			nombre_archivo = extension[contEx-2]
			extension_archivo = extension[contEx-1]


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
	



			var criterios ="alea="+Math.random()+"&accion_ajax=verifica_biblioteca_deudores&nombre_archivo="+encodeURIComponent(archivo)+"&strRut="+strRut+"&IntId="+IntId
			$('#verifica_archivo').load('FuncionesAjax/verifica_archivo_ajax.asp', criterios, function(){
				
				var archivo_validado	=$('#archivo_validado').val()
				
				if(archivo_validado=="no_existe")
				{
					
					var contador =0
					if($('input[type="checkbox"]:checked').size()==0)
					{
						alert("Debe seleccionar al menos un documento")
						return

					}else{

						$.each($('input[type="checkbox"]:checked'),function(){

							var TX_OBSERVACIONES	=$('#TX_OBSERVACIONES').val()
							var ID_CUOTA 			=$(this).val()
							contador = contador +1

							var criterios ="alea="+Math.random()+"&accion_ajax=CARGA_ARCHIVOS_CUOTA&strRut="+strRut+"&ID_CUOTA="+ID_CUOTA+"&TX_OBSERVACIONES="+encodeURIComponent(TX_OBSERVACIONES)
							$('#CARGA_ARCHIVOS_CUOTA').load('FuncionesAjax/detalle_gestiones_ajax.asp', criterios, function(){})

							if($('input[type="checkbox"]:checked').size()==contador)
							{
								frmSend.action = "biblioteca_deudores.asp?accion_archivo=carga&strRut="+strRut;
								frmSend.submit();
							}
						})

					}


				
				}else{
				
					alert("El archivo que intenta subir al sistema ya existe. Si desea subirlo igualmente, elimine el archivo anterior o cambie el nombre de éste.")
					return
				}

			})


			
		}

		function bt_descargar(ruta){

			frmSend.action = "biblioteca_deudores.asp?ruta="+ruta+"&accion_archivo=descarga";
			frmSend.submit();

		}

		function bt_eliminar(cod_cliente, rut, nombre_archivo, pagina_origen, id_archivo)
		{
			if(confirm("¿Esta seguro que desea eliminar el archivo, posterior a esta acción no podrá recuperarlo?"))
			{
				location.href="EliminarArchivo.asp?IntId="+cod_cliente+"&strRut="+rut+"&VarNombreFichero="+nombre_archivo+"&pagina_origen="+pagina_origen+"&id_archivo="+id_archivo
			}
			
		}

		function bt_ver_historial(ID_CUOTA)
		{

			window.open('historial_documentos_biblioteca_deudor.asp?ID_CUOTA='+ID_CUOTA,"_new","width=900, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
		}

	

	</script>
	<style type="text/css">
		form{
			text-align: center
		}
		.titulo_principal{
			background-color:#380ACD; 
			height: 20px;				
			text-align: center;
		}

		.Estilo13{
			background-color:#380ACD; 	
		}
		.subir_archivo{
			width: 90%;	
			margin: 0 auto;

		}
		.subir_archivo label{
			margin:20px;
		}
		.input_file{
			float:left;
		}	
		.boton_carga { 
			border:1px #16428B solid;
			cursor: pointer;
		} 

		.input_boton_carga{
			border:1px #16428B solid;
			cursor: pointer;
			float:right;
		}
		.div_archivos_cargados{
			margin-top:20px;
			margin-bottom:20px;
			text-align: center;
			width:100%;
		}
		.div_archivos_cargados div{
			border:1px solid #16428B;
			margin: 0 auto;
			width:90%;
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
	
		}
		.tabla_archivos_cargados td{
			text-align: left; 			
		}
		.contenedor_historial_eliminados{
			text-align: left;
			width: 90%;
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

	</style>


</HEAD>



<BODY BGCOLOR='FFFFFF'>
	<input type="hidden" name="IntId" 	id="IntId" 	Value="<%=trim(IntId)%>">
	<input type="hidden" name="strRut" 	id="strRut" Value="<%=trim(strRut)%>">
	<div class="titulo_informe">BIBLIOTECA DEUDORES</div>
	<br>
	<FORM name="frmSend" id="frmSend" onSubmit="return enviar(this)" accept-charset="utf-8" method="POST" enctype="multipart/form-data" action="biblioteca_deudores.asp">

		<div class="subir_archivo">			
			<input class="input_file" name="File1" id="File1" type="file" VALUE="<%=File1%>" onClick="cajas();" size="40" maxlength="40">
			<input Name="SubmitButton"  Value="Cargar" class="fondo_boton_100" style="float:right;" Type="BUTTON" onClick="enviar();">
		</div>
<br>
	<div name="divDocGes" id="divDocGes" style="display:none" >

		<br>
		
		<table  class="intercalado" >
			<thead>
			<tr class="">
				<td <%If Trim(strUsaSubCliente)="1" Then%> colspan="10" <%else%> colspan="8" <%end if%> align="LEFT" id="">
					<a href="#" onClick= "marcar_boxes(true);">Marcar todos</a>
					<a href="#" onClick="desmarcar_boxes(true);">Desmarcar todos</a>
				</td>
				<td align="RIGHT">
					<a href="#" onClick="cajas1();"><img border="0" alt="Eliminar archivo" src="../imagenes/NoGestionarRojo.png"></a>
				</td>						
			</tr>

			<tr class="">
				<td>&nbsp;</td>
				<%If Trim(strUsaSubCliente)="1" Then%>
					<td>RUT CLIENTE</td>
					<td>NOMBRE CLIENTE</td>
				<%End If%>				
				<td>NºDOC</td>
				<td>CUOTA</td>
				<td>FEC.VENC.</td>
				<td>TIPO DOC.</td>
				<td align="center" width="70">CAPITAL</td>
				<td align="center" width="70">ABONO</td>
				<td align="center" width="70">SALDO</td>
				<td>&nbsp;</td>
			</tr>
			</thead>
			<tbody>
			<%
			AbrirSCG()

			strSql = "SELECT dbo." & strNomFormInt & "(CUOTA.ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(CUOTA.ID_CUOTA) as HONORARIOS,"
			strSql = strSql & "	ISNULL(FACTURA_RECEPCIONADA,2) AS FACTURA_RECEPCIONADA, COD_ULT_GEST, NRO_CUOTA,"
			strSql = strSql & "	ISNULL(NOTIFICACION_RECEPCIONADA,2) AS NOTIFICACION_RECEPCIONADA,"
			strSql = strSql & "	VALOR_CUOTA, CUOTA.ID_CUOTA,RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, NRO_DOC, SALDO, NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO,"
			strSql = strSql & "	GASTOS_PROTESTOS, CUENTA, FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,"
			strSql = strSql & "	CUSTODIO, CUOTA.FECHA_ESTADO, FECHA_CREACION,"
			strSql = strSql & "	ID_ULT_GEST,ISNULL((SUBSTRING(CONVERT(VARCHAR(11),CAST(CONVERT(VARCHAR(10),CUOTA.FECHA_AGEND_ULT_GES,103) AS DATETIME),6),1,7) + '/ ' + (CASE WHEN CUOTA.HORA_AGEND_ULT_GES = '' THEN '00:00' ELSE CUOTA.HORA_AGEND_ULT_GES END)),'SIN AGEND') AS FECHA_AGEND_ULT_GES,"
			strSql = strSql & "	ISNULL(PRC.ESTADO_PRIORIZACION,1) AS ESTADO_PRIORIZACION "
			strSql = strSql & " ,( "
			strSql = strSql & " SELECT COUNT(*) "
			strSql = strSql & " FROM CARGA_ARCHIVOS_CUOTA car "
			strSql = strSql & " WHERE CAR.ID_CUOTA =CUOTA.ID_CUOTA AND car.activo=1 "
			strSql = strSql & " ) CANTIDAD_DOCUMENTOS "

			strSql = strSql & " FROM CUOTA LEFT JOIN GESTIONES_CUOTA GC ON CUOTA.ID_CUOTA = GC.ID_CUOTA AND GC.ID_GESTION = CUOTA.ID_ULT_GEST_GENERAL"
			strSql = strSql & "			   LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
			strSql = strSql & "			   LEFT JOIN GESTIONES_TIPO_GESTION ON     SUBSTRING(CUOTA.COD_ULT_GEST,1,1) = GESTIONES_TIPO_GESTION.COD_CATEGORIA"
			strSql = strSql & "													   AND SUBSTRING(CUOTA.COD_ULT_GEST,3,1) = GESTIONES_TIPO_GESTION.COD_SUB_CATEGORIA"
			strSql = strSql & "											  		   AND SUBSTRING(CUOTA.COD_ULT_GEST,5,1) = GESTIONES_TIPO_GESTION.COD_GESTION"
			strSql = strSql & "				    				  		        	AND CUOTA.COD_CLIENTE = GESTIONES_TIPO_GESTION.COD_CLIENTE"
			strSql = strSql & " 		   LEFT JOIN PRIORIZACIONES_CUOTA PRC ON CUOTA.ID_CUOTA = PRC.ID_CUOTA AND ESTADO_PRIORIZACION = '0'"

			strSql = strSql & " WHERE RUT_DEUDOR='" & strRut & "' AND CUOTA.COD_CLIENTE='" & IntId & "' AND  ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) "

			strSql = strSql & " ORDER BY RUT_SUBCLIENTE, FECHA_VENC DESC"

			'Response.write "strSql=" & strSql
			'Response.End
			set rsTemp= Conn.execute(strSql)

			intCorrelativo = 1
			strArrID_CUOTA=""
			intTotSelSaldo= 0

			strArrConcepto = ""
			strArrID_CUOTA = ""

			Do until rsTemp.eof

					intSaldo =  rsTemp("SALDO")
					intValorCuota =  rsTemp("VALOR_CUOTA")
					intAbono = intValorCuota - intSaldo
					strNroDoc = rsTemp("NRO_DOC")
					strNroCuota = rsTemp("NRO_CUOTA")
					strFechaVenc = rsTemp("FECHA_VENC")
					strTipoDoc = rsTemp("TIPO_DOCUMENTO")

					If intEstadoPrio ="0" then
					strDisabled = "disabled"
					Else
					strDisabled = ""
					End If

					'Response.write "intHonorarios=" & intHonorarios

					intTotDoc= intSaldo+intIntereses+intProtestos+intHonorarios

					intTotSelSaldo = intTotSelSaldo+intSaldo
					intTotSelAbono = intTotSelAbono+intAbono
					intTotSelValorCuota = intTotSelValorCuota+intValorCuota
					intTotSelDoc = intTotSelDoc+intTotDoc

					strArrConcepto = strArrConcepto & ";" & "CH_" & rsTemp("ID_CUOTA")
					strArrID_CUOTA = strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")

					%>
					<input name="TX_SALDO_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intTotDoc%>">
					<input name="TX_CAPITAL_<%=Replace(rsTemp("ID_CUOTA"),"-","_")%>" type="hidden" value="<%=intValorCuota%>">

					<tr class="Estilo34">					
						<TD>
							<INPUT TYPE="checkbox" ID="id_cuota" NAME="CH_<%=rsTemp("ID_CUOTA")%>" value="<%=rsTemp("ID_CUOTA")%>" <%=strDisabled%> onClick="suma_capital(this,TX_CAPITAL_<%=rsTemp("ID_CUOTA")%>.value,TX_SALDO_<%=rsTemp("ID_CUOTA")%>.value);">
						</TD>

						<%If Trim(strUsaSubCliente)="1" Then%>

							<td><%=rsTemp("RUT_SUBCLIENTE")%></td>

							<td class="Estilo4" title="<%=rsTemp("NOMBRE_SUBCLIENTE")%>">
							<%=Mid(rsTemp("NOMBRE_SUBCLIENTE"),1,30)%></td>

						<%End If%>

						<td><%=strNroDoc%></td>
						<td><%=strNroCuota%></td>
						<td><%=strFechaVenc%></td>
						<td><%=strTipoDoc%></td>
						<td ALIGN="RIGHT"><%=FN(intValorCuota,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intAbono,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intTotDoc,0)%></td>
						<td>
							<%IF trim(rsTemp("CANTIDAD_DOCUMENTOS"))>0 then%>
								<img src="../Imagenes/48px-Crystal_Clear_filesystem_folder_yellow.png" width="20" height="20" style="cursor:pointer;" alt="Historial documentos adjuntos" onclick="bt_ver_historial('<%=trim(rsTemp("ID_CUOTA"))%>')">
							<%else%>
								<img src="../Imagenes/48px-Crystal_Clear_filesystem_folder_red.png" width="20" height="20" style="cursor:pointer;" alt="Sin Historial documentos adjuntos" onclick="bt_ver_historial('<%=trim(rsTemp("ID_CUOTA"))%>')">
							<%end if%>
						</td>
					</tr>
					<%

				rsTemp.movenext
			intCorrelativo = intCorrelativo + 1
			loop

			vArrConcepto = split(strArrConcepto,";")
			vArrID_CUOTA = split(strArrID_CUOTA,";")

			intTamvConcepto = ubound(vArrConcepto)

			rsTemp.close
			set rsTemp=nothing
			CerrarSCG()

			strArrID_CUOTA = Mid(strArrID_CUOTA,2,len(strArrID_CUOTA))
	%>
		</tbody>
		<thead>
		<tr class="Estilo34">
			<td <%If Trim(strUsaSubCliente)="1" Then%> colspan = "7" <%else%> colspan = "5" <%end if%>>Totales Seleccionados:</td>
			<td ALIGN="RIGHT"><span id="span_TX_CAPITAL">0</span>
				<INPUT TYPE="hidden" ID="TX_CAPITAL" NAME="TX_CAPITAL" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)">
			</td>
			<td>&nbsp;</td>
			<td ALIGN="RIGHT"><span id="span_TX_SALDO">0</span>
				<INPUT TYPE="hidden" ID="TX_SALDO"  NAME="TX_SALDO" DISABLED style="text-align:right;width:90" size="10" size="10" onkeyup="format(this)" onchange="format(this)"></td>
					<td>&nbsp;</td>
		</tr>
				
		<INPUT TYPE="hidden" NAME="strArrID_CUOTA" VALUE="<%=strArrID_CUOTA%>">
		</thead>
		</table>

		<br>

		<table width="90%" border="0" bordercolor="#FFFFFF" align="center">
		<tr>
			<td <%If Trim(strUsaSubCliente)="1" Then%> colspan = "7" <%else%> colspan = "5" <%end if%> class="subtitulo_informe">> OBSERVACIONES (Max. 600 Caract.)</td>
		</tr>
		<tr>
		  <td <%If Trim(strUsaSubCliente)="1" Then%> colspan = "8" <%else%> colspan="4" <%end if%>  align="LEFT">
			<TEXTAREA ID="TX_OBSERVACIONES" NAME="TX_OBSERVACIONES" ROWS="4" COLS="120"></TEXTAREA>
		  </td>
		</tr>		
		</table>
		
	</div>
			
		<%
		AbrirSCG()
		
			SQL_SEL ="SELECT id_archivo, nombre_archivo, cod_cliente, rut, convert(varchar(10), FECHA_CARGA, 103) +' '+CONVERT(VARCHAR(5),FECHA_CARGA, 108) FECHA_CARGA, "
			SQL_SEL = SQL_SEL & "ID_USUARIO_CARGA,  " 
			
			SQL_SEL = SQL_SEL & "isnull(nombres_usuario,'')+' '+isnull(apellido_paterno,'')+' '+isnull(apellido_materno,'') nombre_usuario "

			SQL_SEL = SQL_SEL & "FROM CARGA_ARCHIVOS car " 
			SQL_SEL = SQL_SEL & "INNER JOIN USUARIO usu ON usu.ID_USUARIO=car.id_usuario_carga "
			SQL_SEL = SQL_SEL & "WHERE car.activo =1 AND cod_cliente="&trim(IntId)&" AND rut ='"&trim(strRut)&"' AND origen = 2 "
			SQL_SEL = SQL_SEL & " ORDER BY id_archivo desc"
			set rs_sql_sel = Conn.execute(SQL_SEL)
			'response.write SQL_SEL			
		%>
		<%if not rs_sql_sel.eof then%>	
			<div class="div_archivos_cargados">
			<div>

			<table class="intercalado" style="width:100%;">
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
							<a onclick="bt_eliminar('<%=trim(rs_sql_sel("cod_cliente"))%>','<%=trim(rs_sql_sel("rut"))%>','<%=trim(rs_sql_sel("nombre_archivo"))%>','biblioteca_deudores','<%=trim(rs_sql_sel("id_archivo"))%>')" href="#"><img border="0" alt="Eliminar archivo" src="../imagenes/icon_cruz_roja.jpg"></a>
						</td>

						<td>
						
							<a href="#" onclick="bt_descargar('../Archivo/BibliotecaDeudores/<%=trim(rs_sql_sel("cod_cliente"))%>/<%=trim(rs_sql_sel("rut"))%>/<%=trim(rs_sql_sel("nombre_archivo"))%>')"><%=trim(rs_sql_sel("nombre_archivo"))%></a>

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
		SQL_SEL = SQL_SEL & "WHERE activo =0 AND cod_cliente="&trim(IntId)&" AND rut ='"&trim(strRut)&"' " 
		SQL_SEL = SQL_SEL & "AND origen = 2 "
		set rs_sql_sel = Conn.execute(SQL_SEL)
	%>
		<div class="contenedor_historial_eliminados">
		
			<div class="div_historial_eliminados  boton_carga fondo_boton_100"><span class="">&nbsp;Historial archivos eliminados (<%=trim(rs_sql_sel("cantidad"))%>)</span> <img  id="img_historial_eliminados" height="13" width="13" src="../Imagenes/flecha_abajo.png"></div>
			
			<div id="div_historial_eliminados">
	
			</div>

		</div>	
	</FORM>



<div id="CARGA_ARCHIVOS_CUOTA"></div>
<div id="verifica_archivo"></div>



<%CerrarSCG()%>
</BODY>
</HTML>

<script type="text/javascript">
function marcar_boxes(){

	frmSend.TX_CAPITAL.value = 0;
	frmSend.TX_SALDO.value = 0;
	$('#span_TX_SALDO').text(0)
	$('#span_TX_CAPITAL').text(0)	
	desmarcar_boxes()



	<% For i=1 TO intTamvConcepto %>
			if (document.forms[0].<%=vArrConcepto(i)%>.disabled == false) {
			document.forms[0].<%=vArrConcepto(i)%>.checked=true;
			suma_capital(document.forms[0].<%=vArrConcepto(i)%>, document.forms[0].TX_CAPITAL_<%=vArrID_CUOTA(i)%>.value, document.forms[0].TX_SALDO_<%=vArrID_CUOTA(i)%>.value);
			}
	<% Next %>

}

function desmarcar_boxes(){
	frmSend.TX_CAPITAL.value = 0;
	frmSend.TX_SALDO.value = 0;

	$('#span_TX_SALDO').text(0)
	$('#span_TX_CAPITAL').text(0)	

	<% For i=1 TO intTamvConcepto %>
		document.forms[0].<%=vArrConcepto(i)%>.checked=false;
	<% Next %>
	$("#span_TX_SALDO").prettynumber();	
	$("#span_TX_CAPITAL").prettynumber();
}

function suma_capital(objeto , intValorSaldoCapital, intValorSaldo){
		//alert(objeto.checked);

		if (frmSend.TX_CAPITAL.value == '') frmSend.TX_CAPITAL.value = 0;
		if (frmSend.TX_SALDO.value == '') frmSend.TX_SALDO.value = 0;

		if (objeto.checked == true) {
			frmSend.TX_CAPITAL.value = eval(frmSend.TX_CAPITAL.value) + eval(intValorSaldoCapital);
			frmSend.TX_SALDO.value = eval(frmSend.TX_SALDO.value) + eval(intValorSaldo);

			$('#span_TX_SALDO').text($('#TX_SALDO').val())
			$('#span_TX_CAPITAL').text($('#TX_CAPITAL').val())
			
			$("#span_TX_SALDO").prettynumber();	
			$("#span_TX_CAPITAL").prettynumber();
		}
		else
		{
			frmSend.TX_CAPITAL.value = eval(frmSend.TX_CAPITAL.value) - eval(intValorSaldoCapital);
			frmSend.TX_SALDO.value = eval(frmSend.TX_SALDO.value) - eval(intValorSaldo);

			$('#span_TX_SALDO').text($('#TX_SALDO').val())
			$('#span_TX_CAPITAL').text($('#TX_CAPITAL').val())	
			$("#span_TX_SALDO").prettynumber();		
			$("#span_TX_CAPITAL").prettynumber();	
		}
	}
	
function MostrarFilas(Fila) {
var elementos = document.getElementsByName(Fila);
	for (i = 0; i< elementos.length; i++) {
		if(navigator.appName.indexOf("Microsoft") > -1){
			var visible = 'block'
		} else {
			var visible = 'table-row';
		}
		elementos[i].style.display = visible;
	}
}

function OcultarFilas(Fila) {
	var elementos = document.getElementsByName(Fila);
	for (k = 0; k< elementos.length; k++) {
		elementos[k].style.display = "none";
	}
}

function cajas()
{
	MostrarFilas('divDocGes');
}
function cajas1()
{
	OcultarFilas('divDocGes');
}
</script>
