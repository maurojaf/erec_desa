<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"  LCID = 1034%>
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/comunes/rutinas/chkFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/sondigitos.inc"-->
	<!--#include file="../lib/comunes/rutinas/formatoFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/validarFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	AbrirSCG()
	
	intTipoDiscador =  request("CMB_TIPO_DISCADOR")
	intCod_Cliente	= Replace(Replace(Replace(Request("CB_CLIENTE"), "[", ""), "]", ""), """", "")
	intEtapaCobranza=  request("CMB_ETAPA_COBRANZA")
	intNombreFoco	=  request("CB_NOMBREFOCO")
	intTipoFoco		=  request("CB_TIPOFOCO")
	dtmFechaSubFoco = REQUEST("CB_FECHA_SUB_FOCO")
	
	'response.write "dtmFechaSubFoco : " & dtmFechaSubFoco
%>

<!DOCTYPE html>
<html lang="es">
<HEAD>
   
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
   <!--#include file="sesion.asp"-->
	<%
		Id_Usuario = session("session_idusuario")
		
		resp = request("resp")
	%>



<%
	if intCod_Cliente ="" then  intCod_Cliente = session("ses_codcli")       

	intCod_Cliente		= Replace(Replace(Replace(Request("CB_CLIENTE"), "[", ""), "]", ""), """", "")
	intNombreFoco	=  request("CB_NOMBREFOCO")
	intTipoFoco		=  request("CB_TIPOFOCO")
	
%>
	<link href="../css/style_multi_select.css" rel="stylesheet"> 
    <link href="../css/normalize.css" rel="stylesheet">	
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet"> 
	<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
    <link href="../css/jquery.alerts.css" rel="stylesheet"> 
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
	<script src="../Componentes/jquery.alerts.mod.js"></script>
    
    <script src="../Componentes/jquery.multiselect.js"></script>
    

	<script language="JavaScript " type="text/JavaScript">

	    $(document).ready(function () {
		
			$(document).tooltip();
	        $.prettyLoader();
			$("#CB_CLIENTE").multiselect();
			
	    });
		
		function envia() {
		
			$.prettyLoader.show();
			
			var cliente = [];

			$('select[id="CB_CLIENTE"] option:checked').each(function () {
				cliente.push($(this).val())
			})
			
			document.datos.action = 'InformeFocos.asp?resp=si&CB_CLIENTE=' + encodeURIComponent(JSON.stringify(cliente));
			document.datos.submit();
			
		}
		
		function CargaFechas(cat)
		{
			//alert(cat);
			
			
			var comboBox = document.getElementById('CB_FECHA_SUB_FOCO');			
			comboBox.options.length = 0;
			
			var idsubfoco = cat
			
				if (cat=='1') {
					var newOption = new Option('SELECCIONE', '01/01/1900');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					strSql = "SELECT DISTINCT D.FECHA_ESTADO_FOCO"
					strSql = strSql & " FROM DEUDOR D INNER JOIN CUOTA C ON D.RUT_DEUDOR=C.RUT_DEUDOR AND D.COD_CLIENTE=C.COD_CLIENTE"
					strSql = strSql & " 		      INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"
					strSql = strSql & " WHERE ED.ACTIVO=1"
					strSql = strSql & " AND D.FECHA_ESTADO_FOCO IS NOT NULL"
					strSql = strSql & " AND D.ID_FOCO = 1"
					strSql = strSql & " AND CONVERT(VARCHAR(8),D.FECHA_ESTADO_FOCO,112)=CONVERT(VARCHAR(8),GETDATE(),112)"
					strSql = strSql & " ORDER BY D.FECHA_ESTADO_FOCO DESC"
					
					'response.write "strSql : " & strSql		
					
					set rsGestion=Conn.execute(strSql)
					If Not rsGestion.Eof Then
						Do While Not rsGestion.Eof
							%>
								var newOption = new Option('<%=rsGestion("FECHA_ESTADO_FOCO")%>', '<%=rsGestion("FECHA_ESTADO_FOCO")%>');comboBox.options[comboBox.options.length] = newOption;
							<%
							rsGestion.movenext
						Loop
					Else
					%>
						var newOption = new Option('FECHA NO DISPONIBLE', '0');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					%>
				}
				if (cat=='2') {
					var newOption = new Option('SELECCIONE', '01/01/1900');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					strSql = "SELECT DISTINCT D.FECHA_ESTADO_FOCO"
					strSql = strSql & " FROM DEUDOR D INNER JOIN CUOTA C ON D.RUT_DEUDOR=C.RUT_DEUDOR AND D.COD_CLIENTE=C.COD_CLIENTE"
					strSql = strSql & " 		      INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"
					strSql = strSql & " WHERE ED.ACTIVO=1"
					strSql = strSql & " AND D.FECHA_ESTADO_FOCO IS NOT NULL"
					strSql = strSql & " AND D.ID_FOCO = 2"
					strSql = strSql & " AND CONVERT(VARCHAR(8),D.FECHA_ESTADO_FOCO,112)=CONVERT(VARCHAR(8),GETDATE(),112)"
					strSql = strSql & " ORDER BY D.FECHA_ESTADO_FOCO DESC"
					
					'response.write "strSql : " & strSql	
					
					set rsGestion=Conn.execute(strSql)
					If Not rsGestion.Eof Then
						Do While Not rsGestion.Eof
							%>
								var newOption = new Option('<%=rsGestion("FECHA_ESTADO_FOCO")%>', '<%=rsGestion("FECHA_ESTADO_FOCO")%>');comboBox.options[comboBox.options.length] = newOption;
							<%
							rsGestion.movenext
						Loop
					Else
					%>
						var newOption = new Option('FECHA NO DISPONIBLE', '0');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					%>
				}
				if (cat=='3') {
					var newOption = new Option('SELECCIONE', '01/01/1900');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					strSql = "SELECT DISTINCT D.FECHA_ESTADO_FOCO"
					strSql = strSql & " FROM DEUDOR D INNER JOIN CUOTA C ON D.RUT_DEUDOR=C.RUT_DEUDOR AND D.COD_CLIENTE=C.COD_CLIENTE"
					strSql = strSql & " 		      INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"
					strSql = strSql & " WHERE ED.ACTIVO=1"
					strSql = strSql & " AND D.FECHA_ESTADO_FOCO IS NOT NULL"
					strSql = strSql & " AND D.ID_FOCO = 3"
					strSql = strSql & " AND CONVERT(VARCHAR(8),D.FECHA_ESTADO_FOCO,112)=CONVERT(VARCHAR(8),GETDATE(),112)"
					strSql = strSql & " ORDER BY D.FECHA_ESTADO_FOCO DESC"
					
					'response.write "strSql : " & strSql
								
					set rsGestion=Conn.execute(strSql)
					If Not rsGestion.Eof Then
						Do While Not rsGestion.Eof
							%>
								var newOption = new Option('<%=rsGestion("FECHA_ESTADO_FOCO")%>', '<%=rsGestion("FECHA_ESTADO_FOCO")%>');comboBox.options[comboBox.options.length] = newOption;
							<%
							rsGestion.movenext
						Loop
					Else
					%>
						var newOption = new Option('FECHA NO DISPONIBLE', '0');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					%>
				}
		}

    </script>

	<style type="text/css">
        .hiddencol {
		
            display:none;
			
        }
		
        .span_aviso_rojo {
		
			color:#FE2E2E;
			font-size:12px;
			width:1px;
			margin-right:2cm;
			
		}
		
		.abrir_cerrar {
		
			cursor: pointer;
			display: inline-block;
			float: right;
			position: relative;
			margin-right:2cm;
		
		}
	</style>
</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<form name="datos" method="post">
<div class="titulo_informe">INFORME FOCOS</div>
<br>
&nbsp;<table width="90%" class="estilo_columnas" align="center">
		<thead>
	      <tr height="20" >
			<td>CLIENTE</td>
			<td>TIPO DISCADOR</td>
			<td>ETAPA COBRANZA</td>
			<td>NOMBRE FOCO</td>
			<td class="style1">TIPO SUB FOCO</td>
			<td class="style1">FECHA SUB FOCO</td>
			<td class="style1">&nbsp;</td>
			<td >&nbsp;</td>
		  </tr>
		</thead>
			
		  <tr >
			
			<td>
		
                <select name="CB_CLIENTE" id="CB_CLIENTE" multiple>
				<%
				ssql="SELECT  COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
				set rsTemp= Conn.execute(ssql)
				if not rsTemp.eof then
			    
                do until rsTemp.eof
                   cadena = intCOD_CLIENTE
                   y =  rsTemp("COD_CLIENTE") 
                    strChecked = ""
                    a=Split(cadena,",")
                    for each x in a
                        IF  trim(x) = trim(y)  THEN
                            strChecked = " selected "
                        END IF 
                    next
                    
                    %>
					<option  <%=strChecked%> value="<%=rsTemp("COD_CLIENTE")%>" ><%=rsTemp("RAZON_SOCIAL")%></option>
					<%

                    
					rsTemp.movenext
					loop
				end if
				rsTemp.close
				set rsTemp=nothing
				%>
				</select>
                
                
            </td>
			
 			<td>
				<SELECT NAME="CMB_TIPO_DISCADOR" id="CMB_TIPO_DISCADOR">
					<option value="" <%If Trim(intTipoDiscador)="" Then Response.write "SELECTED"%>>SELECCIONE</option>
					<option value="1" <%If Trim(intTipoDiscador)="1" Then Response.write "SELECTED"%>>DINOMI</option>
					<option value="2" <%If Trim(intTipoDiscador)="2" Then Response.write "SELECTED"%>>PURECLOUD</option>
				</SELECT>
			</td>
			
			<td>
				<SELECT NAME="CMB_ETAPA_COBRANZA" id="CMB_ETAPA_COBRANZA">
					<option value="0" <%If Trim(intEtapaCobranza)="0" Then Response.write "SELECTED"%>>TODOS</option>
					<option value="2" <%If Trim(intEtapaCobranza)="2" Then Response.write "SELECTED"%>>PREJUDICIAL</option>
					<option value="4" <%If Trim(intEtapaCobranza)="4" Then Response.write "SELECTED"%>>CASTIGO</option>
				</SELECT>
			</td>
			<td>
				<%
					strChecked = ""
					
					if intNombreFoco = "1" then
					
						strChecked = " selected=""selected"" "
					
					end if
				%>
				<select name="CB_NOMBREFOCO" id="CB_NOMBREFOCO">
					<option value="">SELECCIONE</option>
					<option value="1" <%=strChecked %> >CALL</option>
				</select>
			</td>
			<td class="style1">
				<select name="CB_TIPOFOCO" id="CB_TIPOFOCO" onChange="CargaFechas(this.value);">
					<option value="">SELECCIONE</option>
				<%
				ssql="SELECT ID_FOCO ,TIPO_FOCO ,NOMBRE_FOCO FROM dbo.FOCOS"
				set rsTemp= Conn.execute(ssql)
				if not rsTemp.eof then
			    
                do until rsTemp.eof
                   cadena = intTipoFoco
                   y =  rsTemp("ID_FOCO") 
                    strChecked = ""
                    a=Split(cadena,",")
                    for each x in a
                        IF  trim(x) = trim(y)  THEN
                            strChecked = " selected "
                        END IF 
                    next
                    
                    %>
					<option  <%=strChecked%> value="<%=rsTemp("ID_FOCO")%>" ><%=rsTemp("NOMBRE_FOCO")%></option>
					<%

                    
					rsTemp.movenext
					loop
				end if
				rsTemp.close
				set rsTemp=nothing
				%>
				</select>
			</td>
					<td Colspan= "1">
					<SELECT NAME="CB_FECHA_SUB_FOCO" id="CB_FECHA_SUB_FOCO">
					</SELECT>
					</td>
					
			<td class="style1">&nbsp;</td>
			<td align="right">
				 <input type="Button" class="fondo_boton_100" name="Submit" value="Generar" onclick= "envia();">
				 </td>
			</tr>
    </table>
	
	<%
	
		if resp = "si" then
		
			if Trim(intTipoFoco) = "" then
				intTipoFoco = "NULL"
			end if
			
			if Trim(intNombreFoco) = "" then
				intNombreFoco = "NULL"
			end if

			strSql = "EXEC [dbo].[uspInformeFocosSelect] " &_
						"@TipoDiscador = " & intTipoDiscador & "," &_
						"@CodigoMandante = '" & intCod_Cliente & "'," &_
						"@EtapaCobranza = '" & intEtapaCobranza & "'," &_
						"@TipoFoco = " & intTipoFoco & "," &_
						"@NombreFoco = " & intNombreFoco & "," &_
						"@FechaFoco = '" & dtmFechaSubFoco & "'"
						
			'Response.Write(strSql)
			'response.end
						
			set rsTemp = Conn.execute(strSql)
			
			set rsFoco = nothing
			
			archivosGenerados = false
			
			files = ""
			
			if not rsTemp.eof then
				
				' Se llenan los parámetros de Nombre del foco, hora y fecha
				
				strSqlFoco = "SELECT NOMBRE_FOCO, CONVERT(VARCHAR(8), GETDATE(), 112) AS Fecha, SUBSTRING(REPLACE(CONVERT(VARCHAR(12), GETDATE(), 114), ':', ''), 1, 4) AS Hora  FROM dbo.FOCOS WHERE ID_FOCO = " & intTipoFoco
				
					set rsFoco = Conn.execute(strSqlFoco)
				
					if not rsFoco.eof then
					
						nombreFoco = Replace(Replace(rsFoco("NOMBRE_FOCO"), "FOCO ", ""), " ", "_")
						
					else
					
						nombreFoco = ""
					
					end if
					
					fecha = rsFoco("Fecha")
					
					hora = rsFoco("Hora")	
					
				
				' Se llenan los parámetros de Ruta del archivo			

				strSqlRutaArchivos = "SELECT Dato1, Dato2 FROM EREC4.Core.Parametro WHERE TipoParametro = 'InformeFocos' AND Codigo = '001'"
				
				set rsRutaArchivos = Conn.execute(strSqlRutaArchivos)
				
				rutaArchivos = rsRutaArchivos("Dato1")
				
				rutaArchivosUsuario = rsRutaArchivos("Dato2")
					

				' Se llenan los parámetros de Iniciales código de cliente y usuario
				
				cod_cliente = rsTemp("COD_CLIENTE")
				usuario_asig = rsTemp("USUARIO")
				etapa_cobranza = rsTemp("ETAPA_COBRANZA")

				
				'------- Se Inicia el Proceso --------
					
				'Se crean los archivos SCL y REC
				
				
				If intTipoFoco = "2" or intTipoFoco = "3" then
				
					
				strSqlParametro = "SELECT Dato3 AS NombreCliente FROM EREC4.Core.Parametro " & _
									"WHERE TipoParametro = 'Mandantes' " & _
									"AND Numerico1 = " & cod_cliente
									
				set rsParametro = Conn.execute(strSqlParametro)
				
					if not rsParametro.eof then
						
							nombreCliente = rsParametro("NombreCliente")

							If intTipoDiscador = 1 then
							
								fileName = nombreFoco & "_" & etapa_cobranza & "_" & fecha & "_" & hora & "_" & nombreCliente 

								set fso = Server.CreateObject("Scripting.FileSystemObject")
								
								files = files & fileName & ".csv" & ","
								
								set file = fso.CreateTextFile(rutaArchivos & fileName & ".csv")	
							
								file.WriteLine("TELEFONO,RUT,NOMBRE,APELLIDO,COD_CLIENTE,RUT_SUBCLIENTE,USUARIO,ETAPA_COBRANZA")
							
							ElseIf intTipoDiscador = 2 then
							
								fileName = fecha & "_" & hora & "_" & nombreFoco & "_" & etapa_cobranza & "_" & nombreCliente 												

								set fso = Server.CreateObject("Scripting.FileSystemObject")
								
								files = files & fileName & ".csv" & ","
								
								set file = fso.CreateTextFile(rutaArchivos & fileName & ".csv")	
								
								file.WriteLine("APELLIDO,TELEFONO,TELEFONO2,TELEFONO3,TELEFONO4,TELEFONO5,RUT,CODIGO_CLIENTE,SALDO_DEUDOR,PRIORIDAD,USUARIO,ETAPA_COBRANZA")
							
							End If							
						
						rsParametro.close
						
						set rsParametro = nothing
						
						do until rsTemp.eof
						
							data = ""
						
							if cod_cliente = rsTemp("COD_CLIENTE") and etapa_cobranza = rsTemp("ETAPA_COBRANZA") then
							
								For Each objField in rsTemp.Fields
								
									data = data & rsTemp(objField.Name) & ","
									
								Next
								
								data = mid(data, 1, len(data) - 1)
								
								if not file is nothing then
								
									file.WriteLine(data)
									
								end if
							
							else
							
								cod_cliente = rsTemp("COD_CLIENTE")
								usuario_asig = rsTemp("USUARIO")
								etapa_cobranza = rsTemp("ETAPA_COBRANZA")
								
								if not file is nothing then
								
									file.close
									
								end if
								
								set file = nothing
						
								strSqlParametro = "SELECT Dato3 AS NombreCliente, CONVERT(VARCHAR(8), GETDATE(), 112) AS Fecha, SUBSTRING(REPLACE(CONVERT(VARCHAR(12), GETDATE(), 114), ':', ''), 1, 4) AS Hora FROM EREC4.Core.Parametro " & _
										"WHERE TipoParametro = 'Mandantes' " & _
										"AND Numerico1 = " & cod_cliente
											
								set rsParametro = Conn.execute(strSqlParametro)
								
								if not rsParametro.eof then
								
									if not rsFoco.eof then
							
										nombreFoco = Replace(Replace(rsFoco("NOMBRE_FOCO"), "FOCO ", ""), " ", "_")
										
									else
									
										nombreFoco = ""
									
									end if
									
									If intTipoDiscador = 1 then
									
										fileName = nombreFoco & "_" & etapa_cobranza & "_" & fecha & "_" & hora & "_" & rsParametro("NombreCliente") 

										set fso = Server.CreateObject("Scripting.FileSystemObject")
										
										files = files & fileName & ".csv" & ","
										
										set file = fso.CreateTextFile(rutaArchivos & fileName & ".csv")	
									
										file.WriteLine("TELEFONO,RUT,NOMBRE,APELLIDO,COD_CLIENTE,RUT_SUBCLIENTE,USUARIO,ETAPA_COBRANZA")
									
									ElseIf intTipoDiscador = 2 then
									
										fileName = fecha & "_" & hora & "_" & nombreFoco & "_" & etapa_cobranza & "_" & rsParametro("NombreCliente") 												

										set fso = Server.CreateObject("Scripting.FileSystemObject")
										
										files = files & fileName & ".csv" & ","
										
										set file = fso.CreateTextFile(rutaArchivos & fileName & ".csv")	
										
										file.WriteLine("APELLIDO,TELEFONO,TELEFONO2,TELEFONO3,TELEFONO4,TELEFONO5,RUT,CODIGO_CLIENTE,SALDO_DEUDOR,PRIORIDAD,USUARIO,ETAPA_COBRANZA")
									
									End If	
									
									For Each objField in rsTemp.Fields
									
										data = data & rsTemp(objField.Name) & ","
										
									Next
									
									data = mid(data, 1, len(data) - 1)
									
									file.WriteLine(data)
								
								end if
								
							end if
							
							rsTemp.movenext
							
						loop
						
						set fso = nothing
						
						files = mid(files, 1, len(files) - 1)
						
						Response.Write("<br /><table width=""90%"" align=""center""><tr><td><strong>Archivos generados exitosamente en la ruta " & rutaArchivosUsuario & ".</strong></td></tr><tr><td></td></tr>")
						
						for each f in split(files, ",")
						
							Response.Write("<tr><td>" & f & "</td></tr>")
						
						next
						
						Response.Write("</table>")
						
						archivosGenerados = true
						
					end if
					
				
				elseif intTipoFoco = "1" then
				
							
					If intTipoDiscador = 1 then
					
						fileName = nombreFoco & "_" & etapa_cobranza & "_" & fecha & "_" & hora & "_" & usuario_asig

						set fso = Server.CreateObject("Scripting.FileSystemObject")
						
						files = files & fileName & ".csv" & ","
						
						set file = fso.CreateTextFile(rutaArchivos & fileName & ".csv")	
					
						file.WriteLine("TELEFONO,RUT,NOMBRE,APELLIDO,COD_CLIENTE,RUT_SUBCLIENTE,USUARIO,ETAPA_COBRANZA")
					
					ElseIf intTipoDiscador = 2 then
					
						fileName = fecha & "_" & hora & "_" & nombreFoco & "_" & etapa_cobranza & "_" & usuario_asig												

						set fso = Server.CreateObject("Scripting.FileSystemObject")
						
						files = files & fileName & ".csv" & ","
						
						set file = fso.CreateTextFile(rutaArchivos & fileName & ".csv")	
						
						file.WriteLine("APELLIDO,TELEFONO,TELEFONO2,TELEFONO3,TELEFONO4,TELEFONO5,RUT,CODIGO_CLIENTE,SALDO_DEUDOR,PRIORIDAD,USUARIO,ETAPA_COBRANZA")
					
					End If	

					'Response.Write(strSql)
					
					do until rsTemp.eof
					
						data = ""
					
						if usuario_asig = rsTemp("USUARIO") then
						
							For Each objField in rsTemp.Fields
							
								data = data & rsTemp(objField.Name) & ","
								
							Next
							
							data = mid(data, 1, len(data) - 1)
							
							if not file is nothing then
							
								file.WriteLine(data)
								
							end if
						
						else
						
							usuario_asig = rsTemp("USUARIO")
							etapa_cobranza = rsTemp("ETAPA_COBRANZA")
							
							if not file is nothing then
							
								file.close
								
							end if
							
							set file = nothing
					
						
								nombreFoco = Replace(Replace(rsFoco("NOMBRE_FOCO"), "FOCO ", ""), " ", "_")
								
								If intTipoDiscador = 1 then
								
									fileName = nombreFoco & "_" & etapa_cobranza & "_" & fecha & "_" & hora & "_" & usuario_asig

									set fso = Server.CreateObject("Scripting.FileSystemObject")
									
									files = files & fileName & ".csv" & ","
									
									set file = fso.CreateTextFile(rutaArchivos & fileName & ".csv")	
								
									file.WriteLine("TELEFONO,RUT,NOMBRE,APELLIDO,COD_CLIENTE,RUT_SUBCLIENTE,USUARIO,ETAPA_COBRANZA")
								
								ElseIf intTipoDiscador = 2 then
							
								fileName = fecha & "_" & hora & "_" & nombreFoco & "_" & etapa_cobranza & "_" & usuario_asig 												

								set fso = Server.CreateObject("Scripting.FileSystemObject")
								
								files = files & fileName & ".csv" & ","
								
								set file = fso.CreateTextFile(rutaArchivos & fileName & ".csv")	
								
								file.WriteLine("APELLIDO,TELEFONO,TELEFONO2,TELEFONO3,TELEFONO4,TELEFONO5,RUT,CODIGO_CLIENTE,SALDO_DEUDOR,PRIORIDAD,USUARIO,ETAPA_COBRANZA")
							
							End If	
								
								For Each objField in rsTemp.Fields
								
									data = data & rsTemp(objField.Name) & ","
									
								Next
								
								data = mid(data, 1, len(data) - 1)
								
								file.WriteLine(data)
							
						end if
						
						rsTemp.movenext
						
					loop
					
					set fso = nothing
					
					files = mid(files, 1, len(files) - 1)
					
					Response.Write("<br /><table width=""90%"" align=""center""><tr><td><strong>Archivos generados exitosamente en la ruta " & rutaArchivosUsuario & ".</strong></td></tr><tr><td></td></tr>")
					
					for each f in split(files, ",")
					
						Response.Write("<tr><td>" & f & "</td></tr>")
					
					next
					
					Response.Write("</table>")
					
					archivosGenerados = true			
			
				end if
				
			end if
			
			if not archivosGenerados then
			
				Response.Write("<br /><table width=""90%"" align=""center""><tr><td><strong>No se han generado archivos, ya que no hay datos para los parámetros indicados.</strong></td></tr></table>")
			
			end if
			
			if not rsFoco is nothing then
			
				rsFoco.close
				
			end if
			
			rsTemp.close
			
			set rsTemp = nothing
			
			set rsFoco = nothing
	
		end if
	
	%>
	
	<br>
	
</form>

</body>
</html>
<%cerrarscg()%>

	<script type="text/javascript">
	
		function InicializaInforme()
		{
				var comboBox = document.getElementById('CB_FECHA_SUB_FOCO');
				comboBox.options.length = 0;
				var newOption = new Option('SELECCIONE','');
				comboBox.options[comboBox.options.length] = newOption;
		}
		
		<%If dtmFechaSubFoco = "" then%>	
		InicializaInforme()	
		<%Else%>
		CargaFechas(document.datos.CB_TIPOFOCO.value)()
		datos.CB_FECHA_SUB_FOCO.value='<%=dtmFechaSubFoco%>';
		<%End If%>

    </script>