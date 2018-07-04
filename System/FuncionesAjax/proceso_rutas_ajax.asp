<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="../arch_utils.asp"-->
<!--#include file="../../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../../lib/lib.asp"-->
<tag>
<%

Response.CodePage = 65001
Response.charset="utf-8"
accion_ajax 		=request("accion_ajax")

abrirscg()

if trim(accion_ajax)="refresca_resumen" then

intCOD_CLIENTE 		=request("Cod_Cliente")
strEstadoProceso    =request("strEstadoProceso")
termino             =request("termino")
strhorario          =request("strhorario")	
strTipoGestion      =request("strTipoGestion")
Busqueda            =request("Busqueda")
str_empresa_rec     =request("str_empresa_rec")
Desde               =request("Desde")
Hasta               =request("Hasta")
Estado               =request("Estado")

if Desde = "" or Desde ="0" then Desde = 1
if Hasta = "" or Hasta ="0" then Hasta = 10



if Busqueda = 1 then 

                            cabecera = "<table width='100%' border='0'align='center'><tr><td class='abrir_cerrar'>" & _
                                           "<input class='fondo_boton_100' type='button' id='Procesar' value='Procesar'  onClick='MostrarVentana(2);'/>&nbsp;&nbsp;"  & _
                                           "<input Name='SubmitButton' class='fondo_boton_100' Value='Exportar' Type='BUTTON' onClick='exportar();'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
                                           "</td></tr><tr>" & _
                                           "<table id='tbl_Procesa' width='80%' border='0' bordercolor='#000000' class='intercalado' align='center'>" & _
                                           "<thead>" & _
                                           "<tr><td></td>" & _
                                           "<td><input type='checkbox' id='chckHead' onClick='MarcaChexbok();' /></td>" 
                                           if session("perfil_emp")  <> "Verdadero"  then
                                            cabecera =cabecera &   "<td>CLIENTE</td>" 
                                           end if 
                                           cabecera =cabecera &    "<td>EMP. REC.</td>" & _
                                           "<td>FECHA RUTA</td>" & _
                                           "<td>RUT CLIENTE</td>" & _
                                           "<td>NOMBRE CLIENTE</td>" & _
			                               "<td>RUT DEUDOR</td> " & _
                                           "<td>DEUDOR</td> " & _
                                           "<td>FACTURA</td>" & _
                                           "<td>MONTO</td>"   & _
                                           "<td>HORARIO PAGO</td>" &_
                                           "<td>AM/PM</td>" & _
                                           "<td>DIR.</td>" & _
                                           "<td>DOC</td>" & _
                                           "<td>OBS</td>" & _
                                           "<td>&nbsp;</td>" & _
                                           "<td>&nbsp;</td>" & _
                                           "<td class='hiddencol'>id_gestion</td>" & _
                                           "<td class='hiddencol'>id_cuota</td></tr>" & _
                                           "</thead> <tbody>" 
      
                                           strSql ="select  * from ( SELECT ROW_NUMBER() OVER(ORDER BY d.rut_deudor) id, CONVERT(VARCHAR(10),g.FECHA_COMPROMISO,103)AS FECHA_PAGOS ,c.RUT_SUBCLIENTE		RUTCLIENTE " 
                                            strSql = strSql  & ",c.NOMBRE_SUBCLIENTE	NOMBRECLIENTE,d.RUT_DEUDOR ,d.NOMBRE_DEUDOR DEUDOR ,c.NRO_DOC FACTURA " 
                                            strSql = strSql  & ",CAST(CAST(c.SALDO AS BIGINT) AS VARCHAR(10)) AS MONTO ,c.Id_Cuota,gc.Id_Gestion " 
                                            strSql = strSql  & ",g.HORA_DESDE + ' - ' + g.HORA_HASTA AS HORARIO_ENTREGA_PAGOS " 
                                            strSql = strSql  & ",CASE WHEN isnull(substring (g.HORA_DESDE,1,2),9) <= 13 and isnull(substring(g.HORA_HASTA,1,2),9) <= 14	THEN 'AM'	 ELSE 'PM'	END AS AM_PM" 
                                            strSql = strSql  & ",Isnull(UPPER(ISNULL(UPPER(FR.NOMBRE+' '+FR.UBICACION), upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))),'') AS DIRECCION" 
                                            strSql = strSql  & ",g.DOC_GESTION AS DOCUMENTOS_NECESARIOS" 
                                            strSql = strSql  & ",DT.TELEFONO"
                                            strSql = strSql  & ",ISNULL(TC.CONTACTO,EC.CONTACTO) AS CONTACTO "
                                            strSql = strSql & " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(G.observaciones_CAMPO ,'Ñ','N'),'Á','A'),'É','E'),'Í','I'),'Ó','O'),'Ú','U'),CHAR(10),' '),CHAR(13),' ') AS 'OBSERVACION'"
                                            strSql = strSql & "	,('COMUNA: ' + ISNULL(DD.COMUNA,'NO DEFINIDA')) AS COMUNA"
			                                strSql = strSql & "	,CALLE AS CALLE_GEO"
			                                strSql = strSql & "	,NUMERO AS NUMERO_GEO"
			                                strSql = strSql & "	,COMUNA AS COMUNA_GEO"
                                            strSql = strSql & " 	    	,CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL"
			                                strSql = strSql & " 	    	 THEN GC.NUEVA_EMPRESA_REC"
			                                strSql = strSql & " 	    	 WHEN ER.EMPRESA_REC IS NOT NULL"
			                                strSql = strSql & "	  			 THEN ER.EMPRESA_REC"
			                                strSql = strSql & " 	 		 WHEN DD.COMUNA IS NOT NULL"
			                                strSql = strSql & " 	  		 THEN UPPER(DD.COMUNA)"
			                                strSql = strSql & " 	 		 ELSE 'NO DEFINIDA'"
			                                strSql = strSql & " 		END AS EMPRESA_RECAUDADORA,"
                                            strSql = strSql & " 	   'FECHA INGRESO: '+ (CONVERT(VARCHAR(10),g.FECHA_INGRESO,103) + ' ' + CONVERT(VARCHAR(5),convert(datetime, g.HORA_INGRESO), 108)  +' '+"
			                                strSql = strSql & " 	   'USUARIO ESTADO RUTA: ' + ISNULL(u.LOGIN,'NO PROCE.') +"
			                                strSql = strSql & " 	   ' FECHA ESTADO: ' + ISNULL(CONVERT(VARCHAR(10),GC.FECHA_ESTADO_RUTA,103)+' '+CONVERT(VARCHAR(10),GC.FECHA_ESTADO_RUTA,108),'NO PROC.')) AS OBSERVACION_GENERAL_RUTA ,cli.NOMBRE_FANTASIA"
                                            strSql = strSql & " FROM DEUDOR D	inner join CLIENTE cli on cli.COD_CLIENTE = d.COD_CLIENTE " 
                                            strSql = strSql & " INNER JOIN CUOTA C ON D.RUT_DEUDOR = C.RUT_DEUDOR AND D.COD_CLIENTE in (" & intCOD_CLIENTE & ")"
				                            strSql = strSql & " 				  INNER JOIN GESTIONES G ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.FECHA_COMPROMISO IS NOT NULL"
				                            strSql = strSql & " 				  INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION AND C.ID_CUOTA = GC.ID_CUOTA" 
				                            strSql = strSql & " 				  INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
				                            strSql = strSql & " 							 AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA "
				                            strSql = strSql & " 							 AND G.COD_GESTION = GTG.COD_GESTION"
				                            strSql = strSql & " 							 AND GTG.COD_CLIENTE = D.COD_CLIENTE"
            	                            strSql = strSql & " 				  LEFT JOIN DEUDOR_DIRECCION DD ON G.ID_DIRECCION_COBRO_DEUDOR = DD.ID_DIRECCION AND C.RUT_DEUDOR = DD.RUT_DEUDOR "
				                            strSql = strSql & " 				  LEFT JOIN FORMA_RECAUDACION FR ON G.ID_FORMA_RECAUDACION = FR.ID_FORMA_RECAUDACION "
				                            strSql = strSql & " 				  LEFT JOIN EMPRESAS_RECAUDADORAS ER ON ISNULL(RTRIM(DD.COMUNA),FR.NOMBRE+' '+FR.UBICACION) = ER.NOMBRE_COMUNA AND ER.COD_CLIENTE = D.COD_CLIENTE"
				                            strSql = strSql & " 				  LEFT JOIN CAJA_FORMA_PAGO ON G.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO "
                                            strSql = strSql & " 				  LEFT JOIN DEUDOR_TELEFONO DT ON DT.ID_TELEFONO = G.ID_FONO_COBRO "
                                            strSql = strSql & " 				  LEFT JOIN TELEFONO_CONTACTO      TC  ON G.ID_CONTACTO_FONO_COBRO = TC.ID_CONTACTO"
 			                                strSql = strSql & " 				  LEFT JOIN EMAIL_CONTACTO         EC   ON G.ID_MEDIO_GESTION = EC.ID_CONTACTO"
                                            strSql = strSql & "                   LEFT JOIN USUARIO u ON              gc.USUARIO_ESTADO_RUTA = u.ID_USUARIO"
                                            strSql = strSql & " WHERE ((ISNULL(GTG.CONFIRMA_CP,0) = 1 AND ISNULL(GC.CONFIRMACION_CP,'N')='S' )"
 				                            strSql = strSql & " OR   (ISNULL(GTG.CONFIRMA_CP,0) = 0))"
 				                            strSql = strSql & " AND (G.ID_DIRECCION_COBRO_DEUDOR IS NOT NULL"
 				                            strSql = strSql & " OR ISNULL(FR.TIPO,'') = 'RETIRO' )"
					
					                            If strEstadoProceso = "0" then
					
					                            strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') >= 0 "
					                            strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') <= 30 "
					                            
                                                
					                            Else

 					                            strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 "
					
					                            End If

                                                if Trim(strEstadoProceso) <> "3" Then
						                            strSql = strSql & " 	AND ISNULL(GC.ESTADO_RUTA,0) = '" & strEstadoProceso & "'"
					                            End If
					
												if Trim(strTipoGestion) = "1" Then 'Recaudacion
												
													strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11))"
												
												ElseIf Trim(strTipoGestion) = "2" Then 'Notificacion
												
													strSql = strSql & " AND (GTG.GESTION_MODULOS = 13)"
												
												Else 
												
													strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11,13))"
													
												End If

                                                    strSql = strSql & " AND  G.ID_GESTION = C.ID_ULT_GEST_GENERAL"

					                            if Trim(strHorario) = "1" Then
						                            strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14)"
					                            End If

					                            if Trim(strHorario) = "2" Then
						                            strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14)"
					                            End If

 					                            strSql = strSql & " AND C.ESTADO_DEUDA IN (1,2,7,8) "
 					                            strSql = strSql & " AND C.COD_CLIENTE in (" & intCOD_CLIENTE & ")"

                                                if Trim(str_empresa_rec) <> "0" Then
						                            strSql = strSql & "		   AND (CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL"
						                            strSql = strSql & " 	    	 THEN GC.NUEVA_EMPRESA_REC"
						                            strSql = strSql & " 	    	 WHEN ER.EMPRESA_REC IS NOT NULL"
						                            strSql = strSql & "	  			 THEN ER.EMPRESA_REC"
						                            strSql = strSql & " 	 		 WHEN DD.COMUNA IS NOT NULL"
						                            strSql = strSql & " 	  		 THEN UPPER(DD.COMUNA)"
						                            strSql = strSql & " 	 		 ELSE 'NO DEFINIDA'"
						                            strSql = strSql & " 		END) = '" & str_empresa_rec & "'"
					                            End if
                                    strSql = strSql & " ) a   where id BETWEEN " & Desde & " AND " & Hasta & " order by rut_deudor"
                                    
                                   ' 
                                   'response.write(strSql)
                                   'response.end



                                    set rsDet=Conn.execute(strSql)

                                    dim i 
                                    i = Desde -1
                                   ' objeto =""

                                    if not rsDet.eof then
		                            do until rsDet.eof
                                        i = i + 1 

                                        If Trim(rsDet("DIRECCION")) <> "" Then
					                            srtAnexoMsg1 = rsDet("DIRECCION")
			                            Else
			                                    srtAnexoMsg1 = "SIN INFORMACI&OacuteN"
			                            End If

                                        If rsDet("DOCUMENTOS_NECESARIOS") <> "" Then
				                            srtAnexoMsg2 = rsDet("DOCUMENTOS_NECESARIOS")
			                            Else
				                            srtAnexoMsg2 = "SIN INFORMACI&OacuteN"
			                            End If

                                        If rsDet("TELEFONO") <> "" then
				                            strFonoCobro = "FONO COBRO: " + rsDet("TELEFONO")
			                            Else
				                            strFonoCobro = ""
			                            End If	
            
                                        If rsDet("CONTACTO") <> "" then
						                            strContactoCobro = " CONTACTO COBRO: " + rsDet("CONTACTO")
			                            Else
						                            strContactoCobro = ""
		                                End If	


                                        If rsDet("OBSERVACION") <> "" then
			                                    strObservacionRuta = " OBS: " + rsDet("OBSERVACION")
			                            Else
					                            strObservacionRuta = " OBS: SIN INFORMACI&OacuteN" 
			                            End If


                                        srtAnexoMsg3 = strFonoCobro + strContactoCobro + strObservacionRuta
             
			                            strComuna = rsDet("COMUNA")

                                        strDireccion2 = rsDet("CALLE_GEO") & " " & rsDet("NUMERO_GEO") & "," & rsDet("COMUNA_GEO")
			                            strDireccion2 = Trim(strDireccion2)

                                        strDireccion_geo = replace(ucase(strDireccion2),"CALLE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"CALLE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"POBLACION","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"POBLACIÓN","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PASAJE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AV.","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PJE.","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PSJE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PGE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIDA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"CAYE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"CALLLE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIDAS","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"V.","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AVDA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PASAGE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARCELA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARC.","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARCELAS","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARSELA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARS.","")		

			                            strDireccion_geo = replace(ucase(strDireccion_geo)," ","%")		 
                                        'strDireccion_geo =ucase(strDireccion_geo)
            
            
                                        objeto = objeto &  "<tr><td  title='" & rsDet("OBSERVACION_GENERAL_RUTA") & "'>"  & i  & "</td>" 
                                        objeto = objeto &  "<td><input  type='checkbox'    id='ChckRow_" & i  & "' /></td>" 
                                        if session("perfil_emp")  <> "Verdadero"  then
                                            objeto = objeto &  "<td>" & rsDet("NOMBRE_FANTASIA") &"</td>" 
                                        end if 
                                        objeto = objeto &  "<td>" & rsDet("EMPRESA_RECAUDADORA") &"</td>" 
                                        objeto = objeto &  "<td>"& rsDet("FECHA_PAGOS") &"</td>"   
                                        objeto = objeto &  "<td >"&rsDet("RUTCLIENTE")&"</td>"   
                                        objeto = objeto &  "<td title='" & rsDet("NOMBRECLIENTE") & "'>"& mid(rsDet("NOMBRECLIENTE"),1,20) &"</td>"   
                                        objeto = objeto &  "<td><a href='principal.asp?TX_RUT="& rsDet("RUT_DEUDOR")  &"'>"&rsDet("RUT_DEUDOR") &"</a></td>"   
                                        objeto = objeto &  "<td title='" & rsDet("DEUDOR") & "'>"&  mid(rsDet("DEUDOR"),1,20) &"</td>"   
                                        objeto = objeto &  "<td>"&rsDet("FACTURA")&"</td>"   
                                        objeto = objeto &  "<td>"& formatnumber(rsDet("MONTO"),0) &"</td>"   
                                        objeto = objeto &  "<td>"&rsDet("HORARIO_ENTREGA_PAGOS")&"</td>"   
                                        objeto = objeto &  "<td>"&rsDet("AM_PM")&"</td>"   
                                        objeto = objeto &  "<td title='" & srtAnexoMsg1 & "'>"&"<img src='../imagenes//mod_direccion_va.png' border='0'>" &"</td>"   
                                        objeto = objeto &  "<td title='" & srtAnexoMsg2 & "'>"&"<img src='../imagenes/icon_doc.png' border='0'>"&"</td>"   
                                        objeto = objeto &  "<td title='" & srtAnexoMsg3 & "'>"&"<img src='../imagenes/priorizar_normal.png' border='0'>"&"</td>" 
                                        objeto = objeto &  "<td  title='" & strComuna & "'>"&"<img src='../imagenes/winrar_view.png' border='0'>"&"</td>"   
                                        objeto = objeto &  "<td><img width='20'   style='cursor:pointer;' onclick=bt_geolocalizacion('"+ strDireccion_geo +"'); height='20' src='../Imagenes/map.png' title='Consulta Direcci&oacuten Mapa'/></td>"   
                                        objeto = objeto &  "<td class='hiddencol'> "& rsDet("id_gestion") &"</td>"   
                                        objeto = objeto &  "<td class='hiddencol'> "& rsDet("id_cuota") &"</td>" 
                                        objeto = objeto & "</tr>"
       	                            rsDet.movenext
		                            loop
                                   
                                   objeto =cabecera &   objeto 
                                   objeto =objeto & "</tbody></table>"
                                   objeto =objeto & "</td></tr></table>"
                                   objeto =objeto & "<TABLE class='intercalado'><TBODY><TR><TD >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & " <A HREF='#' onClick='consulta_resumen2("& Desde & "," & Hasta & ",0);'><img src='../imagenes/previous.gif' border='0'></A>"
                                   objeto =objeto & "&nbsp;Resultado Desde <b>" &  Desde & "</b> Hasta  <b>" & i & "</b>"
                                   if  int(Hasta) = int(i) then
                                   objeto =objeto & "&nbsp;<A HREF='#' onClick='consulta_resumen2("& Desde & "," & Hasta & ",1);'><img src='../imagenes/next.gif' border='0'></A></td>"
                                   end if 
                                   objeto =objeto & "</TD></TR></TBODY></TABLE>"
                                  
	                               else 
                                          objeto = "<table width='80%' border='0' bordercolor='#000000' class='intercalado' align='center'><tr><td>Busqueda Documentos Sin Resultados</td></tr></table>"
                                   end if
                                   
                                   response.write(objeto)
                                   
else 

                                           cabecera = "<table width='100%' border='0'><tr><td></td><td class='abrir_cerrar'>" & _
                                           "<input class='fondo_boton_100' type='button' id='Procesar' value='Cambiar Empresa'  onClick='MostrarVentana(1);'/>" & _
                                           "&nbsp;&nbsp;"  & _
                                           "<input Name='SubmitButton' class='fondo_boton_100' Value='Exportar' Type='BUTTON' onClick='exportar();'> " 
                                           if strTipoGestion = "2" then 
                                           cabecera =   cabecera + "&nbsp;&nbsp;<input class='fondo_boton_100' type='button' id='Procesar' value='Procesar'  onClick='MostrarVentana(2);'/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"  
                                           else 
                                           cabecera = cabecera  + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp"
                                           end if

                                           cabecera =   cabecera + "</td></tr><tr>" & _
                                           "<td></td><td><table id='tbl_Procesa' width='80%' border='0' bordercolor='#000000' class='intercalado' align='center'>" & _
                                           "<thead>" & _
                                           "<tr><td></td>" & _
                                           "<td><input type='checkbox' id='chckHead' onClick='MarcaChexbok();' /></td>" 
                                          if session("perfil_emp")  <> "Verdadero"  then
                                            cabecera = cabecera & "<td>CLIENTE</td>" 
                                           end if 
                                          cabecera = cabecera &  "<td>EMP. REC.</td>" & _
                                           "<td>FECHA RUTA</td>" & _
                                           "<td>RUT DEUDOR</td> " & _
                                           "<td>DEUDOR</td> " & _
                                           "<td>CANT DOC</td> " & _
                                           "<td>MONTO DOC</td> " & _
                                           "<td>HORARIO PAGO</td>" &_
                                           "<td>DIR.</td>" & _
                                           "<td>OBS</td>" & _
                                           "<td>&nbsp;</td>" & _
                                           "<td>&nbsp;</td>" & _
                                           "<td class='hiddencol'>id_GESTION</td>" & _
                                           "</tr></thead> <tbody>" 

                                           strSql ="select  * from " & _
                                                    "( SELECT ROW_NUMBER() OVER(ORDER BY d.rut_deudor) id, CONVERT(VARCHAR(10),g.FECHA_COMPROMISO,103)AS FECHA_PAGOS  " & _
                                                    " ,d.RUT_DEUDOR  " & _
                                                    " ,d.NOMBRE_DEUDOR DEUDOR " & _
                                                    " ,count(gc.Id_Cuota) CantidadDocumentos  "  & _
                                                    " ,SUM(C.SALDO)  MONTO"  & _
                                                    " ,CASE WHEN isnull(substring(g.HORA_DESDE,1,2),9) <= 13 and isnull(substring(g.HORA_HASTA,1,2),9) <= 14 THEN 'AM' ELSE 'PM' END AS AM_PM " & _
                                                    " ,Isnull(UPPER(ISNULL(UPPER(FR.NOMBRE+' '+FR.UBICACION),upper(DD.CALLE+' '+DD.NUMERO+' '+DD.RESTO+' '+DD.comuna))),'') AS DIRECCION " & _
                                                    " ,CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL " & _
			                                	    "	 THEN GC.NUEVA_EMPRESA_REC " & _
			                                	    "	 WHEN ER.EMPRESA_REC IS NOT NULL " & _
			                                	  	"		 THEN ER.EMPRESA_REC " & _
			                                	 	"	 WHEN DD.COMUNA IS NOT NULL " & _
			                                	  	"	 THEN UPPER(DD.COMUNA) " & _
			                                	 	"	 ELSE 'NO DEFINIDA' " & _
			                                		"    END AS EMPRESA_RECAUDADORA ,gc.Id_Gestion ,cli.NOMBRE_FANTASIA" 
                                            strSql = strSql & ",DT.TELEFONO"
                                            strSql = strSql & ",ISNULL(TC.CONTACTO,EC.CONTACTO) AS CONTACTO "
                                            strSql = strSql & " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(G.observaciones_CAMPO ,'Ñ','N'),'Á','A'),'É','E'),'Í','I'),'Ó','O'),'Ú','U'),CHAR(10),' '),CHAR(13),' ') AS 'OBSERVACION'"
                                            strSql = strSql & "	,DD.COMUNA,CALLE AS CALLE_GEO"
			                                strSql = strSql & "	,NUMERO AS NUMERO_GEO"
			                                strSql = strSql & "	,COMUNA AS COMUNA_GEO"
                                            strSql = strSql & " FROM DEUDOR D	  inner join CLIENTE cli on cli.COD_CLIENTE = d.COD_CLIENTE " 
                                            strSql = strSql & " INNER JOIN CUOTA C ON D.RUT_DEUDOR = C.RUT_DEUDOR AND D.COD_CLIENTE in (" & intCOD_CLIENTE & ")"
				                            strSql = strSql & " 				  INNER JOIN GESTIONES G ON G.RUT_DEUDOR = D.RUT_DEUDOR AND G.FECHA_COMPROMISO IS NOT NULL"
				                            strSql = strSql & " 				  INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION AND C.ID_CUOTA = GC.ID_CUOTA" 
				                            strSql = strSql & " 				  INNER JOIN GESTIONES_TIPO_GESTION GTG ON G.COD_CATEGORIA = GTG.COD_CATEGORIA "
				                            strSql = strSql & " 							 AND G.COD_SUB_CATEGORIA = GTG.COD_SUB_CATEGORIA "
				                            strSql = strSql & " 							 AND G.COD_GESTION = GTG.COD_GESTION"
				                            strSql = strSql & " 							 AND GTG.COD_CLIENTE = D.COD_CLIENTE"
            	                            strSql = strSql & " 				  LEFT JOIN DEUDOR_DIRECCION DD ON G.ID_DIRECCION_COBRO_DEUDOR = DD.ID_DIRECCION AND C.RUT_DEUDOR = DD.RUT_DEUDOR "
				                            strSql = strSql & " 				  LEFT JOIN FORMA_RECAUDACION FR ON G.ID_FORMA_RECAUDACION = FR.ID_FORMA_RECAUDACION "
				                            strSql = strSql & " 				  LEFT JOIN EMPRESAS_RECAUDADORAS ER ON ISNULL(RTRIM(DD.COMUNA),FR.NOMBRE+' '+FR.UBICACION) = ER.NOMBRE_COMUNA AND ER.COD_CLIENTE = D.COD_CLIENTE"
				                            strSql = strSql & " 				  LEFT JOIN CAJA_FORMA_PAGO ON G.FORMA_PAGO = CAJA_FORMA_PAGO.ID_FORMA_PAGO "
                                            strSql = strSql & " 				  LEFT JOIN DEUDOR_TELEFONO DT ON DT.ID_TELEFONO = G.ID_FONO_COBRO "
                                            strSql = strSql & " 				  LEFT JOIN TELEFONO_CONTACTO      TC  ON G.ID_CONTACTO_FONO_COBRO = TC.ID_CONTACTO"
 			                                strSql = strSql & " 				  LEFT JOIN EMAIL_CONTACTO         EC   ON G.ID_MEDIO_GESTION = EC.ID_CONTACTO"
                                            strSql = strSql & "                   LEFT JOIN USUARIO u ON              gc.USUARIO_ESTADO_RUTA = u.ID_USUARIO"
                                            strSql = strSql & " WHERE ((ISNULL(GTG.CONFIRMA_CP,0) = 1 AND ISNULL(GC.CONFIRMACION_CP,'N')='S' )"
 				                            strSql = strSql & " OR   (ISNULL(GTG.CONFIRMA_CP,0) = 0))"
 				                            strSql = strSql & " AND (G.ID_DIRECCION_COBRO_DEUDOR IS NOT NULL"
 				                            strSql = strSql & " OR ISNULL(FR.TIPO,'') = 'RETIRO' )"
					
					                            If strEstadoProceso = "0" then
					
                                                    strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') >= 0 "
					                                strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') <= 30 "
					              
					                            Else

 					                                strSql = strSql & " AND DATEDIFF(D,G.FECHA_COMPROMISO,'"& termino & "') = 0 "
					
					                            End If
                                                if Trim(strEstadoProceso) <> "3" Then
						                            strSql = strSql & " 	AND ISNULL(GC.ESTADO_RUTA,0) = '" & strEstadoProceso & "'"
        			                            End If
												if Trim(strTipoGestion) = "1" Then 'Recaudacion
													strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11))"
												ElseIf Trim(strTipoGestion) = "2" Then 'Notificacion
													strSql = strSql & " AND (GTG.GESTION_MODULOS = 13)"
												Else 
													strSql = strSql & " AND (GTG.GESTION_MODULOS IN (1,11,13))"
												End If
                                                    strSql = strSql & " AND  G.ID_GESTION = C.ID_ULT_GEST_GENERAL"
					                            if Trim(strHorario) = "1" Then
						                            strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) <= 14)"
					                            End If
					                            if Trim(strHorario) = "2" Then
						                            strSql = strSql & " 	AND (isnull(SUBSTRING(G.HORA_HASTA,1,2),9) > 14)"
					                            End If
 					                            strSql = strSql & " AND C.ESTADO_DEUDA IN (1,2,7,8) "
 					                            strSql = strSql & " AND C.COD_CLIENTE in (" & intCOD_CLIENTE & ")"

                                                	if Trim(str_empresa_rec) <> "0" Then
						                                strSql = strSql & "		   AND (CASE WHEN GC.NUEVA_EMPRESA_REC IS NOT NULL"
						                                strSql = strSql & " 	    	 THEN GC.NUEVA_EMPRESA_REC"
						                                strSql = strSql & " 	    	 WHEN ER.EMPRESA_REC IS NOT NULL"
						                                strSql = strSql & "	  			 THEN ER.EMPRESA_REC"
						                                strSql = strSql & " 	 		 WHEN DD.COMUNA IS NOT NULL"
						                                strSql = strSql & " 	  		 THEN UPPER(DD.COMUNA)"
						                                strSql = strSql & " 	 		 ELSE 'NO DEFINIDA'"
						                                strSql = strSql & " 		END) = '" & str_empresa_rec & "'"
					                                End if



                                        
                                                strSql = strSql & "  group  by	 d.RUT_DEUDOR ,d.NOMBRE_DEUDOR,g.HORA_DESDE,g.HORA_HASTA ,FR.NOMBRE,FR.UBICACION " 
			                                    strSql = strSql & " ,DD.CALLE,DD.NUMERO,DD.RESTO,DD.comuna,g.FECHA_COMPROMISO " 
                                                strSql = strSql & " ,GC.NUEVA_EMPRESA_REC,ER.EMPRESA_REC,gc.Id_Gestion,DT.TELEFONO,TC.CONTACTO,EC.CONTACTO,G.observaciones_CAMPO,COMUNA,NUMERO,CALLE,cli.NOMBRE_FANTASIA" 
                                                 strSql = strSql & " ) a   where id BETWEEN " & Desde & " AND " & Hasta & " order by rut_deudor"

                                   On error Resume Next 
                                        set rsDet=Conn.execute(strSql)

                                     'response.write(",Error," & strSql)
                                     'response.end

                                   If Err.number<>0 then
                                       response.write(",Error," & Err.Description )
                                       response.end
                                   end if 

                                   dim j 
                                   j = 0

                                    if not rsDet.eof then
		                            do until rsDet.eof
                                        j = j +1


                                        If rsDet("TELEFONO") <> "" then
				                            strFonoCobro = "FONO COBRO: " + rsDet("TELEFONO")
			                            Else
				                            strFonoCobro = ""
			                            End If	
            
                                        If rsDet("CONTACTO") <> "" then
						                            strContactoCobro = " CONTACTO COBRO: " + rsDet("CONTACTO")
			                            Else
						                            strContactoCobro = ""
		                                End If	


                                        If rsDet("OBSERVACION") <> "" then
			                                    strObservacionRuta = " OBS: " + rsDet("OBSERVACION")
			                            Else
					                            strObservacionRuta = " OBS: SIN INFORMACI&OacuteN" 
			                            End If


                                        srtAnexoMsg3 = strFonoCobro + strContactoCobro + strObservacionRuta

                                        strDireccion2 = rsDet("CALLE_GEO") & " " & rsDet("NUMERO_GEO") & "," & rsDet("COMUNA_GEO")
			                            strDireccion2 = Trim(strDireccion2)

                                        strDireccion_geo = replace(ucase(strDireccion2),"CALLE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"CALLE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"POBLACION","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"POBLACIÓN","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PASAJE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AV.","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PJE.","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PSJE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PGE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIDA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"CAYE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"CALLLE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIDAS","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AVENIA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"V.","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"AVDA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PASAGE","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARCELA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARC.","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARCELAS","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARSELA","")
			                            strDireccion_geo = replace(ucase(strDireccion_geo),"PARS.","")		

			                            strDireccion_geo = replace(ucase(strDireccion_geo)," ","%")	


                                        strComuna = rsDet("COMUNA")

                                     'response.write(",Error," & strComuna)
                                     'response.end

                                        objeto = objeto &  "<tr><td>"  & j  & "</td>" 
                                        objeto = objeto &  "<td><input  type='checkbox' id='ChckRow_" & j & "' /></td>" 
                                        if session("perfil_emp")  <> "Verdadero"  then
                                        objeto = objeto &  "<td>" & rsDet("NOMBRE_FANTASIA") &"</td>" 
                                        end if 
                                        objeto = objeto &  "<td>" & rsDet("EMPRESA_RECAUDADORA") &"</td>" 
                                        objeto = objeto &  "<td>"& rsDet("FECHA_PAGOS") &"</td>"   
                                        objeto = objeto &  "<td><a href='principal.asp?TX_RUT="& rsDet("RUT_DEUDOR")  &"'>"&rsDet("RUT_DEUDOR") &"</a></td>"   
                                        objeto = objeto &  "<td title='"&rsDet("DEUDOR")&"'>"& mid(rsDet("DEUDOR"),1,30)&"</td>"   
                                        objeto = objeto &  "<td >"& rsDet("CantidadDocumentos")&"</td>"   
                                        objeto = objeto &  "<td >"& formatnumber(rsDet("MONTO"),0) &"</td>"   
                                        objeto = objeto &  "<td>"&rsDet("AM_PM")&"</td>"  
                                        objeto = objeto &  "<td  title='"&rsDet("DIRECCION")&"'>" & mid(rsDet("DIRECCION"),1,30)&"</td>"  
                                        objeto = objeto &  "<td title='" & srtAnexoMsg3 & "'>"&"<img src='../imagenes/priorizar_normal.png' border='0'>"&"</td>" 
                                        objeto = objeto &  "<td  title='" & strComuna & "'>"&"<img src='../imagenes/winrar_view.png' border='0'>"&"</td>"   
                                        objeto = objeto &  "<td><img width='20'   style='cursor:pointer;' onclick=bt_geolocalizacion('"+ strDireccion_geo +"'); height='20' src='../Imagenes/map.png' title='Consulta Direcci&oacuten Mapa'/></td>"   
                                        objeto = objeto &  "<td class='hiddencol'>" & rsDet("id_gestion") & "</td>"   
                                        objeto = objeto & "</tr>"

                                rsDet.movenext

	                            loop
                                   objeto =cabecera &   objeto 
                                   objeto =objeto & "</tbody></table>"
                                   objeto =objeto & "</td></tr></table>"
                                   objeto =objeto & "<TABLE class='intercalado'><TBODY><TR><TD >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                                   objeto =objeto & " <A HREF='#' onClick='consulta_resumen2("& Desde & "," & Hasta & ",0);'><img src='../imagenes/previous.gif' border='0'></A>"
                                   objeto =objeto & "&nbsp;Resultado Desde <b>" &  Desde & "</b> Hasta  <b>" & j & "</b>"
                                   if  int(Hasta) = int(j) then
                                   objeto =objeto & "&nbsp;<A HREF='#' onClick='consulta_resumen2("& Desde & "," & Hasta & ",1);'><img src='../imagenes/next.gif' border='0'></A></td>"
                                   end if 
                                   objeto =objeto & "</TD></TR></TBODY></TABLE>"
                                  
	                               else 
                                          objeto = "<table width='80%' border='0' bordercolor='#000000' class='intercalado' align='center'><tr><td>Busqueda Documentos Sin Resultados</td></tr></table>"
                                   end if


                                   response.write(objeto)
                                   



end if 

   

end if

if trim(accion_ajax)="Procesa_Rutas" then

    Observacion         = request("Observacion")
    Empresa             = request("Empresa")
    Estado              = request("Estado")
    id_gestion          = request("id_gestion")
    id_cuota            = request("id_cuota")  
    EmpresaR            = request("EmpresaR")    
    id                  = request("id")    
    valor               = request("valor") '' si valor es 1 actualiza las empresas
    strTipoGestion      = request("strTipoGestion") '' si es 1 recaudacion/2  notificacion
    FECHA_COMPROMISO    = request("FECHA_COMPROMISO") 
    HORA_HASTA          = request("HORA_HASTA")
    HORA_DESDE          = request("HORA_DESDE")
  
   On error Resume Next 
   if valor = 2  and (strTipoGestion= 1 OR strTipoGestion= 0) then
        strSql = "UPDATE GESTIONES_CUOTA SET OBSERVACION_RUTA = '" & Observacion & "',ESTADO_RUTA = " & Estado & ",FECHA_ESTADO_RUTA = GETDATE(), USUARIO_ESTADO_RUTA = '" & session("session_idusuario") & "'  WHERE ID_CUOTA = " & id_cuota & " AND ID_GESTION = " & id_gestion & " AND ISNULL(ESTADO_RUTA,0) <> " & Estado
        set rsUpdate=Conn.execute(strSql)
   end if  
   
  if valor = 2  and strTipoGestion= 2 then
        strSql = "UPDATE GESTIONES_CUOTA SET OBSERVACION_RUTA = '" & Observacion & "',ESTADO_RUTA = " & Estado & ",FECHA_ESTADO_RUTA = GETDATE(), USUARIO_ESTADO_RUTA = '" & session("session_idusuario") & "'  WHERE ID_GESTION = " & id_gestion & " AND ISNULL(ESTADO_RUTA,0) <> " & Estado
        set rsUpdate=Conn.execute(strSql)

        strSql = "UPDATE GESTIONES SET FECHA_COMPROMISO = '" & FECHA_COMPROMISO & "',HORA_HASTA = '" & HORA_HASTA & "',HORA_DESDE = '" & HORA_DESDE & "'  WHERE ID_GESTION = " & id_gestion 
        set rsUpdate=Conn.execute(strSql)
   end if  
   
   
       
   If Err.number<>0 then
       response.write(",Error," & Err.Description )
       response.end
   end if 


   On error Resume Next 
   if valor = 1 then
        strSql = "UPDATE GESTIONES_CUOTA SET NUEVA_EMPRESA_REC = '" & Empresa & "' WHERE ID_GESTION = " & id_gestion & " AND '" & EmpresaR & "' <> '" & Empresa &"'"
        set rsUpdate2=Conn.execute(strSql)
   end if 

   If Err.number<>0 then
       response.write(",Error," & Err.Description )
       response.end
   end if 


                                 'response.write(",Error," & strSql)
                                 'response.end
    
   response.write(",OK," & "Correlativo "& id &" Modificado Correctamente <br/>")
   response.end
   


end if
cerrarscg()

%>

